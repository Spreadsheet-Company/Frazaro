# CONTEMPLATIONS — thought experiments about Frazaro and VLA

*An ongoing, append-only home for the speculative half of the project: the
questions that are not yet items, the arguments that are not yet decisions,
and the mental models that let the telos be held in one hand instead of a
cloud. Nothing here is a commitment. When a contemplation hardens into a
decision it goes to the standing-decision register; when it hardens into
work it goes to the roadmap; the contemplation stays here as the record of
how it was thought about first.*

*House rule, borrowed from every other ledger in `docs/`: a contemplation is
never edited to look smarter than it was. Corrections are appended, dated,
under the original.*

---

## Contemplation 1 — Standard or register? *(2026-08-27)*

*Prompted by `CONSULTANT.md`'s addendum A.2: is Frazaro's English a
standard (frozen, versioned, governed by a document, changed by process —
TeX's π) or a living register (grown by users, curated after the fact,
occasionally contradictory — the OED)? The owner's own read: the project
has been subconsciously synthesizing — shipping a standard prelude with
overrides that allow extension and revision.*

### The two positions, honestly priced

**Frazaro as a standard** (TeX's π: frozen core, versioned editions,
changed only by a documented process)

Pros:
- Hyrum's Law becomes tractable. If the language only changes by process,
  every dependency on an observable behaviour is a dependency on something
  you *chose*, and CO.1/SD-4 are cheap to honour.
- It stays teachable. A finite, published grammar is what LE.1's palette,
  LE.3's autocomplete, dictation, and LE.7's AI bridge all require. You
  cannot constrain a model to a mood, and you cannot autocomplete a
  register.
- It stays translatable. G-RENDER's whole premise ("write in Spanish, read
  in Japanese") works *only because* the vocabulary is finite and owned.
  Open the vocabulary and rendering degrades to "best effort."
- Auditability holds. The auditor reads a sentence and can look it up. That
  is the product's strongest positioning and it depends on there being a
  place to look.
- One curator can run it. A standard is exactly the governance shape a solo
  developer can sustain.

Cons:
- It can only say what you thought of. The corpus is 122 rules; a real
  month-end will refuse constantly, and every refusal is a support ticket
  routed to a single curator.
- It dies of rigidity. Esperanto is a standard; it is also mostly a hobby.
  Standards that cannot absorb usage get routed around — users fall back to
  recorded macros the moment the dialect says no twice.
- The curator is the throughput ceiling and the bus factor, forever. Every
  "can I say X?" is an appeal to one court.
- It gets the *taste* of one person. Business English varies by trade,
  country, and department; a single-curator standard encodes one office's
  idiolect and calls it the language.

**Frazaro as a living register** (the OED: grown by users, curated after
the fact, occasionally contradictory)

Pros:
- It grows toward its readers. Hart's memo, product form: every user is a
  potential author of the next phrasing, and the language is shaped by the
  people who actually have the procedures.
- Coverage scales with users instead of with the curator's evenings. The
  gap log stops being a queue and becomes a corpus.
- It's how natural languages actually work, which is the mission's whole
  claim. A language "people prefer to think in" is one that accepts how
  they already talk.
- Community is a moat that a compiler is not. Nobody clones a body of
  tested phrasings with adjudication history attached.

Cons:
- Contradictions are guaranteed, not occasional. Two departments *will*
  define "close out" differently, and a program carried between them
  silently changes meaning — the F.4 failure class at social scale.
- Auditability erodes. "Which rule fired?" now depends on load order,
  layer, and which phrasebooks happened to be present that morning.
- Translatability erodes. Open vocabulary means a sentence may have no
  rendering in the reader's dialect.
- The AI bridge weakens: the grammar is no longer finite, so a model's
  output can't be mechanically refused; you're back to "trust the model."
- Hyrum's Law goes vertical at every layer simultaneously, and "retire a
  spelling" becomes a negotiation with strangers.
- It needs a community, and a community is a second product built with
  different tools by a different kind of person — usually not the
  compiler's author.

### Why the subconscious synthesis is right, and where it's incomplete

The kernel-plus-override shape isn't a compromise between the two; it's
what every surviving language did. Scheme is a tiny frozen standard with an
unbounded macro register on top. Statute law is a standard; case law is a
register; the system works because a case must *cite* the statute it
interprets. Even the OED — the emblem of the register — is itself a
standard *of record*: the language lives, the dictionary freezes editions,
and a word enters an edition by citation, not by vote. So the question was
never "standard *or* register." It is **what freezes and what grows, and
what mechanism moves a thing from one to the other.**

The incomplete part: "a standard with overrides" degrades into "a register
with a standard's pretensions" if overrides are free. Overrides accumulate;
each is a small fork; after a year nobody can say what the standard *is*
without loading a particular stack of files. The synthesis needs three
things the project has only partly:

1. **A precedence rule with provenance** — it exists (GO.1: last-loaded
   wins, nothing overrides silently, `Explain` names the winner). That is
   the case-law citation rule.
2. **A promotion path** — it does not exist. How does a phrasing in an org
   phrasebook become part of the next base edition? The OED has citation
   files; Frazaro has a gap log that doesn't exist yet. Without promotion,
   the register never feeds the standard, and the standard slowly becomes
   irrelevant to actual usage while the overrides become the real language.
3. **A retirement path** — CO.1 is the mechanism, unbuilt in practice; more
   importantly, there's no *pressure* to retire. The equivalent of the dot
   count is needed: an override count, published, that people feel bad
   about in the productive way.

### The orthogonal reframing: what "English is now an extensible programming language" actually means

The sentence for the wall, because it dissolves the dichotomy rather than
balancing it:

> **Meaning is a standard. Wording is a register. Frazaro is the machine
> that keeps them separate.**

Unpacked through the layers that already exist, because each layer has a
*different natural change regime* and the confusion comes from treating
"the language" as one thing:

| Layer | What it is | Natural-language analogue | Change regime |
|---|---|---|---|
| Core forms (`set!`, `if`, `for`, `.`) | semantics | syntax — changes over centuries | **constitution**: MAJOR only, by amendment |
| Prelude (`when`, `blank?`, `with-fast-excel`) | standard verbs | function words / grammar words | **statute**: MINOR with deprecation |
| Base phrasebook (`english.vla`) | the dictionary | the dictionary — editions | **edition**: curated, versioned, promoted into |
| Org / community phrasebooks | dialect | trade jargon | **usage**: last-loaded wins, provenance visible |
| `To …:` definitions in a program | idiolect | coining a phrase by using it | **speech**: no governance, local scope |

Read down the table and notice: it's *already Lisp*. A fixed `eval`, a
small standard library, macros all the way up. Lisp survived both failure
modes — rigidity and incoherence — with precisely this shape, and Frazaro
is structurally a Lisp, so inherit Lisp's answer rather than TeX's or the
OED's. What natural languages add to Lisp's answer is the empirical fact
that **vocabulary changes daily and syntax changes over centuries**. So
"extensible" should mean, precisely: *lexically extensible, syntactically
governed.* Users add ways to *say*; they don't add things to *mean* unless
they descend a layer, and descending a layer is a different act with a
different reviewer (this is what REBUILD's `VLA_EnglishWords` /
`VLA_Sentences` split is really for — structural grammar is syntax and
lives with the standard; phrase rules are lexicon and live with the
register).

Now the keystone, the part that makes the register safe to unleash: **the
standard is the forms, and the program must be stored as forms.** Today the
sheet holds English text and F.2 only made the *pipeline* form-native. If
the program of record is the form and the English is a *rendering* into
the reader's phrasebook, then:

- Two departments' contradictory "close out" definitions don't matter: each
  program carries its own forms, and G-RENDER renders it in whoever's
  dialect is reading. The register can be as chaotic as it likes *above*
  the forms because contradictions only exist in wording, never in meaning.
- Auditability survives: the auditor reads a rendering, but signs the form
  hash.
- Translatability is total by construction, not best-effort.
- The AI bridge is preserved: the model must produce a sentence that
  *unifies with a known template into forms*, which is a mechanical check
  regardless of how many templates exist.
- Hyrum's Law is confined to the constitution layer, where MAJOR bumps and
  an amendment process absorb it.

That single move — program of record is forms, text is a view — is the
Hegelian synthesis made mechanical. It is also the last unfinished half of
F.2 (the "surrounding assembly is still text+indentation" note) plus
G-RENDER's "whole procedure" v2. Everything else on this question is
governance; that one is engineering, and it's the load-bearing piece.

### A few more handles for the monkey brain

**Phrasebooks are per-community, not per-language.** "Spanish" isn't a
phrasebook; *Madrid-accounting-Spanish* is. This reframes LX.10 /
DIALECT-REGEN: the seven demo dialects were proving the wrong thing (one
file per tongue). The real unit is one file per *speech community*, and
English gets several too. The mission sentence should probably read
"whatever language *your office* thinks in."

**The promotion mechanism is citation, and the gap log is the citation
file.** Borrow the dictionary's rule wholesale: a phrasing enters the next
base edition when it has been (a) used by N independent communities, (b)
tested, (c) non-conflicting with the edition under F.4, and (d) rendered
both ways. Publish the criteria. Then the curator stops being a court and
becomes an editor, which is a job one person *can* do — editors don't coin
words, they admit them.

**Overrides should cost something visible.** An org phrasebook that
overrides a base rule (rather than adding one) is a fork of the standard.
Count them, print them in `Explain`, and make the promotion path the way to
get the count down. Same instinct as the dot count: a number that falls.

**Contradiction is fine when it's cited.** The OED tolerates contradictory
senses by attaching quotations. The mechanism exists — provenance,
`Explain`, `override:` refusing to be silent. The register's "occasionally
contradictory" con disappears entirely if every contradiction can say where
it came from and which layer won. What's intolerable is a *silent*
contradiction, and that is exactly the F.4 class — which is why F.4 matters
more than its `~weeks` suggests: it is the thing that makes a register
auditable.

**Speech is the bottom rung, and Cut A is how registers actually grow.** In
the wild, phrases are coined by being used; only later do they get
dictionary entries. `To close out {month}: …` written in English inside a
program is exactly that act. It's the right primitive because it's the
natural one, and it's already half-built. The reason to build it before the
community exists is that it *is* how the community will exist: a user who
defined a phrase is one promotion away from being a contributor.

**Two cautionary tales.** Esperanto: a standard so complete it never
needed anyone, and so nobody needed it. Perl: a register so accommodating
that no two programs read alike, and readability — its own founding pitch —
was what it lost. Frazaro's pitch is readability by non-authors. That means
the register must be kept *above* a standard that renders, or Perl is the
destination.

### The telos, in one paragraph

Frazaro is not "English as a programming language." It is **a fixed
semantic core with a natural-language lexicon that any speech community can
extend, where programs are stored as meaning and rendered as wording.** The
core is a standard and is governed like a constitution. The base phrasebook
is an edition, curated like a dictionary and fed by citation from usage.
The org and user layers are a living register, free to grow and
contradict, tolerable because every wording carries provenance and no
wording is the program of record. The one engineering decision that makes
the whole thing coherent — and that isn't finished — is that the thing on
the sheet is a view, not the truth. Build that, and the synthesis stops
being a balancing act and becomes a design.

---

## Contemplation 2 — How to make the sheet a view, not the truth *(2026-08-27)*

*Follow-up to Contemplation 1's closing sentence. Grounded in the code as
it stood that day, not the essay.*

### What "the truth" is right now, in the code

- **The program of record is column A.** `ProgramText` (`VLA_IDE.bas`)
  joins column-A cells with CRLF; `DoCheck`, `RunProgram`,
  `InterpretProgram` all start from that string. Column B is marks. There
  is no second representation anywhere in the workbook.
- **`EnglishToVla`'s output is not an AST — it's already lowered.** In
  `instructions_golden.vla` every statement is wrapped in
  `(set! vla-step N)`, `(if (vlatraceon) …)`, `(at-line N …)`,
  `(raw "' ---- …")`; every `To …:` is a `(sub …)` with `on-error` /
  `label vla-fail` / `vla-report-error` scaffolding. That's the
  *compilation* of the program, not its *meaning*. It cannot be rendered
  back to sentences without first stripping the scaffold, and it should not
  be stored as the record because it bakes in one backend's trace
  convention.
- **G-RENDER renders phrase rules only.** `EnglishRenderForm` unifies
  against `mPatForms` — the phrase-rule templates. The structural grammar
  (`If … :`, `Count … :`, `Repeat`, `To …:`, `When …:`, `Otherwise`,
  `Done.`) is parsed by hard-coded `Case` arms in `ParseStmt` and has no
  template to unify against, so there is no way today to render an
  `(if …)` or `(sub …)` back to English. `EnglishRenderText` refuses
  anything but exactly one form.
- **Storage precedent exists.** Snapshots are very-hidden sheets
  (`TakeRunSnapshot`); the build embeds text as very-hidden sheets
  (`EmbedTextAsSheet`); DevRig uses workbook `Names`. Any of those can hold
  a forms blob.

So "the sheet is a view" is three separate pieces of work with a
dependency order.

### The build, in order

**Step 1 — Define the *surface form* layer (the thing that will be the
truth).** There are two representations today: English text and lowered
VLA. A third is needed between them: **surface forms** — one form per
sentence, un-lowered, backend-neutral, structure preserved.

`EnglishToVla` today emits `(at-line 90 (set! (cells row-number "f")
value))` inside a scaffolded `(sub stamp …)`. The surface form of that
program is:

```
(program
  (define stamp ((row-number 1) (value "ok"))
    (at-line 90 (set! (cells row-number "f") value)))
  (define-function tax ((amount))
    (at-line 187 (return (* amount 0.08))))
  …)
```

No `vla-step`, no `vlatraceon`, no `on-error`, no `raw` section markers —
those are lowerings the *backend* adds. Rules for the layer:

- Every phrase-rule sentence is exactly its template form (already true —
  that's `TryFormPath`'s output before the scaffold is wrapped).
- Every structural sentence has a canonical form with a fixed head:
  `if`/`cond` blocks, `count` loops, `repeat`, `define`,
  `when-sheet-change`, `when-click`. Some already exist as core forms; the
  ones lowered straight to `(sub on:click:N …)` need a surface head that
  remembers it *was* a click handler.
- `at-line` stays (provenance, not scaffold). `gen-row` stays.

In code: split `EnglishToVla` into `EnglishToSurfaceForms(text) →
Collection` and `LowerSurfaceForms(forms) → VLA text`, where the second is
the existing scaffold-wrapping assembly F.2's note calls "still
text+indentation, untouched on purpose." Golden-diff-empty is the
acceptance criterion: `Lower(ToSurface(text))` byte-identical to today's
`EnglishToVla(text)`. An F.2-second-half pass; factoring, not redesign.

**Step 2 — Extend rendering to structural forms.** Add `RenderStructural`
for each structural head, mirroring `ParseStmt`'s `Case` arms in reverse —
the same way `RenderExprForm` mirrors `ParseSum`/`ParseCond`.
`(if c (then …) (else …))` → `If c,` + indented body + `Otherwise` +
`Done.`; `(define name params body)` → `To name, with x of 1:` + body;
`(count v from a to b step s body)` → `Count v from a to b:`. The
operator/comparison words already exist in `ExprOpWord`.

Then `EnglishRenderProgram(forms) → String` walks the tree, renders each
node, and emits one sentence per row with block structure (blank line =
paragraph = block close, which the grid already encodes). This is G-RENDER
v2 — "whole procedure assembly" — the largest genuinely new piece. The
self-check is v1's: `ToSurface(Render(forms))` must equal `forms`,
structurally, for the whole corpus. Run it over `instructions.txt` and the
prover exists.

REBUILD's layering constraint holds: structural rendering needs the English
*words* (`If`, `Otherwise`, `Done`), which is `VLA_EnglishWords`'
territory — a Spanish structural renderer is a sibling table, not a fork.
Step 2 is a forcing function for the LX.5 split, because every structural
keyword has to be enumerated to render it.

**Step 3 — Store the forms in the workbook; make column A a projection.**
The actual inversion, small once 1 and 2 exist.

- On every successful Check/Run/Interpret: `forms = ToSurface(columnA)`;
  write `VlaWriteForm(forms)` (linted through `VlaLintFormat`) into a
  very-hidden sheet `Frazaro_Program_<tag>` (the `TakeRunSnapshot`
  precedent), plus a hash in a workbook `Name`. That's the record.
- Column A becomes **derived**: `Render(forms)` in the reader's loaded
  phrasebook. Initially rendered into the same phrasebook the text was
  parsed from, so the round trip is the identity modulo casing and noise
  words — which is exactly what has to be fixed first, and the honest
  reason to do this: the moment column A is a projection, its infidelities
  (lowercased `SalesTable`, dropped articles, `ascend` for `ascending`)
  become bugs you can see instead of accepted v1 gaps. Casing is the
  concrete one: `Fold` at tokenize time throws it away; the surface form
  needs to carry the user's spelling of names (`(name "SalesTable")` or a
  casing side-table keyed by folded symbol) so it survives.
- Editing: the user edits column A → Check → re-derive forms → overwrite
  the record. The *edit surface* stays text; the *record* is forms. The
  user never touches forms.
- The audit log (`CONSULTANT.md` §2.5) falls out for free: the record is a
  hash of forms, and a run log cites it.

**Step 4 — Now the register can be unleashed, safely.** With forms as the
record, three things impossible today become one-liners:

- **"Show me this program in Spanish"** = `Render(forms)` after loading
  `espanol.vla`. Any phrasebook, any layer, provided every head in the
  program has a template in the target phrasebook — and when one doesn't,
  that sentence renders as raw VLA (already the `:expr` fallback policy),
  which is visibly a *gap in the phrasebook*, not a broken program.
- **Org override chaos is harmless.** Two departments' contradictory
  phrasings for `(close-out month)` render differently and mean the same
  thing, because the program stored is the form.
- **Retiring a spelling (CO.1) becomes a rendering change**, not a
  migration: old programs are forms; they render in the new spelling next
  time they're opened. SD-4's "keeps its meaning" is honoured by
  construction because meaning is what's stored.

### The smallest honest first cut

Not all four. The order that proves the thesis cheapest:

1. Factor `EnglishToVla` into `ToSurface` + `Lower` with
   golden-diff-empty. (`~days`. Zero product change. Buys the layer.)
2. Structural render for exactly three heads — `define`, `if`, `count` —
   and `EnglishRenderProgram` over a tree. Prove it on `instructions.txt`'s
   own `stamp`/`tax`/`tidy-up`. (`~days` to `~week`.)
3. A single ribbon action, **"Render program"**: reads column A, writes
   forms to the hidden sheet, re-renders them into a *second* sheet beside
   the original. Don't overwrite column A yet. Diff the two by eye.
   (`~hours` once 1–2 exist.)

That third step is where the thesis is tested: if the rendered sheet reads
like something a person would have written, flip it — column A becomes the
projection and the hidden sheet becomes the record. If it doesn't, the
result is exactly which infidelities the surface-form layer has to carry
(casing, articles, optional words, alternation surface), and each is a
field on a form, not a redesign.

One warning, since it's the trap this project has walked into before: the
structural renderer will be tempting to over-build (every head, every
dialect, pretty-printing policy). Three heads, one corpus, one self-check.
The register argument only needs the *record* to be forms; it doesn't need
rendering to be beautiful yet — it needs it to round-trip.

---

## Contemplation 3 — The maximally-minimal interstitial *(2026-08-27)*

*The owner's question, turned into speculation: taken to its logical
extreme, Frazaro/VLA is a "maximally-minimal interstitial between natural
languages and programming languages" — the distant end-state where any
natural language can be deterministically translated into a
usefully-meaningful subset of any programming language. What would that
look like next to Frazaro today, and what would the current architecture
have to reconsider to extend in both directions without hard-coding any one
language in a place that would have to be duplicated for every other?*

### The shape has a name, and it has failed before

The end-state is an hourglass. N natural languages on top, M programming
languages (and hosts) on the bottom, one waist in the middle that both
sides meet at and neither side owns. Compilers call the waist an IR (LLVM,
MLIR); machine translation called it an *interlingua* and spent thirty
years failing to build one, because the interlingua tried to represent
**meaning**, and natural-language meaning is open.

VLA does not have that problem, and it is worth saying precisely why,
because it is the property that makes the extreme reachable at all:
**VLA is an interlingua of actions, not of meaning.** It does not
represent what a sentence *means*; it represents what a sentence *does* to
a finite object model. Effects are finite. "Put today into cell D1" has one
denotation because there is one cell and one clock. The "usefully-
meaningful subset" in the owner's phrasing is not a hedge; it is the whole
trick — the waist is small *because* the domain is closed, and every
language on top only has to say things that can be done, not things that
can be thought. An interlingua of doing is buildable. An interlingua of
saying is not. Keep the waist executional and the extreme stays honest.

### Today, measured against the extreme

Today is a 1 × 1 × 1: English × VBA × Excel, with the waist proven to hold
for a second backend (the interpreter — a real second implementation of the
same forms) and gestured at for other natural languages (seven demo
phrasebooks, one wired into a test, all phrase-rule-only). Most of the
distance to N × M × K is not more phrasebooks or more emitters. It is that
the waist today is not one thing — it is three things fused, and two of
them leak.

**There are three axes, not two.** The question as posed has a natural
side and a programming side. The architecture has a third that the current
code conflates with the second: the **host** — Excel's object model,
`range`/`cells`/`worksheets`, the `xl*` constants, the runtime helper zoo.
`VLA_HeadTable` is a table of *language* forms (`set!`, `if`, `for`) and
*host* forms (`range`, `cells`, `make-button`) in one list. A Python
backend targeting Excel via `openpyxl` shares the host axis with VBA and
not the language axis; a VBA backend targeting Word shares the language
axis and not the host. Until those are separate columns — better, separate
tables — every new backend re-implements the host and every new host
re-implements the backend. The waist has to split into:

- a **core** — variables, control flow, procedures, calls, comparison,
  literals, `quote`. Domain-free. Small enough to implement in an
  afternoon for any target (the walking skeleton was 372 lines for ten
  forms; that number is the right order of magnitude for the *whole* core);
- **domain profiles** — the spreadsheet profile (`range`, `cells`, `sheet`,
  `table`, `pivot`, and the slot categories `:cell`/`:range`/`:column`
  that go with them), a document profile, a mail profile. A profile is a
  declared verb set plus a declared slot-type set, the F.1 ABI generalized;
- **host bindings** — what a given backend does for a given profile's
  verbs. VBA-for-Excel is one binding; the interpreter-for-Excel is a
  second binding of the same profile; VBA-for-Word would be a binding of a
  different profile on the same backend.

Adding a natural language should touch only the top. Adding a backend
should touch only a binding column. Adding a host should add a profile and
its bindings. **The test of the waist is that no one of those three acts
requires touching the other two.** Today all three do.

### Where the hard-coding lives, by axis

An inventory, because "don't hard-code any one language" is a slogan until
the sites are named. Every entry below is a literal that belongs to one
language, one backend, or one host, sitting in a place that every other
language, backend, or host would have to duplicate.

**Natural-language axis (English, in engine code):**
- The structural grammar: `If`, `Otherwise`, `Done.`, `Repeat`, `Count …
  from … to`, `To …:`, `When …:` as `Case` arms in `ParseStmt` — the
  demo-versus-product line REBUILD Layer 3 already drew.
- The lexical machinery below the phrase rules: `IsNoiseWord`
  (`a`/`an`/`the`/`please`), `SkipArticles`, `NumberWord`, `OrdinalWord`,
  `ExprOpWord`, the Oxford-comma-required list grammar, sentence
  capitalization, `"of"` as the function-application word.
- The tokenizer's assumptions: whitespace-delimited words, a period ends a
  sentence, hyphens join identifiers, case is meaningless.
- The refusal catalogue's *text* (IDs are neutral now; the templates are
  English), and every ribbon caption, dialog, sheet name, and column header.

**Programming-language axis (VBA, in the core):**
- `SymName`'s mangling rules and `IsReservedName`'s keyword list.
- The lowering scaffold: `vla-step`, `vlatraceon`, `on-error`/`label
  vla-fail`, `Optional … Variant` defaults — the emitter's trace convention
  baked into what looks like the program.
- `raw`, which is by construction a hole through the waist to one backend.
- `Val`/`Str$` locale discipline, `CallByName` dispatch, `Application.Run`
  helper reach — all of which are the *interpreter's* backend-specific
  choices wearing the core's clothes.

**Host axis (Excel, everywhere):**
- `range`/`cells`/`rows`/`columns`/`worksheets`/`activesheet` as core
  heads in the head table and as global receivers in the interpreter.
- The `xl*`/`vb*` constant resolver.
- Slot categories `:cell`/`:range`/`:column`/`:sheet`/`:color` inside the
  *sentence matcher* — the grammar engine knows what a cell is.
- The runtime helper zoo, and the fact that the prelude's verbs
  (`with-no-alerts`, `with-fast-excel`) are Excel verbs in a file that
  presents itself as the standard library.

None of this is a mistake. Every one was the correct decision for a
1 × 1 × 1 built by one person against one host. The contemplation is only
about which of them would have to move for the count to go up on any axis,
and the answer is: the ones that sit in engine code rather than in data
owned by the axis they belong to.

### What "reconsider" would mean, per axis

**Top — the natural-language side becomes a grammar phrasebook, not a
parser.** The phrase rules are already data. The extreme needs the
*structural* grammar to be data too: a per-language table declaring the
block-opening words, the block-closing word, the comparator surfaces, the
list separator policy, the article set, the number and ordinal words, the
function-application word, the sentence terminator, and the morphology
hooks (`ascend/ing` generalized — stem/suffix is English; German needs
case endings, Latin needs declension, Finnish needs agglutination). Then
`ParseStmt`'s `Case "if"` becomes `Case Lang.BlockOpen("if")`. That is
LX.5 stated as its end-state.

But the honest speculation goes further than word tables, because word
tables assume English's *shape*. The real test of language-neutrality is
typological distance, not vocabulary: a language with SOV word order
(Japanese), with grammatical case instead of prepositions (Latin — the
demo file exists and dodges this by staying phrase-rule-only), with no
articles (Russian), with no word spaces (Chinese), with pro-drop
(Spanish, where "Pon hoy" has no subject). Each breaks a different
assumption below the word tables — tokenization, slot ordering, what
"noise word" even means. LX.10 falsified "a second phrasebook is cheap."
The typological version — one Romance, one Germanic, one SOV, one
case-marking, one unspaced — would falsify "the matcher is
language-neutral," and until it runs, the claim is the demo-shaped one
REBUILD warned about. The extreme probably requires the tokenizer and the
slot matcher themselves to be pluggable per language, with the English
ones as the reference implementation, not the engine.

**Bottom — backends become bindings, and the core gets smaller, not
bigger.** SD-5/R9 already force every backend to support or refuse every
core form. The extreme adds the converse pressure: **anything that not
every backend can support is not core.** Run that rule over the head table
today and `range`, `cells`, `make-button`, `deflambda`, `raw` all fall out
of the core into a profile or a declared backend extension. What remains is
the thing the walking skeleton implemented in an afternoon. The core should
be small enough that its complete implementation in a new target is a
weekend, and that smallness is a feature to protect, not a gap to fill —
Scheme's R7RS-small, not Common Lisp.

Programs then carry a **portability measure**: the fraction of their forms
that are core versus profile versus backend-extension. A program that is
100% core plus spreadsheet-profile runs on every binding of the
spreadsheet profile; one that uses `raw` runs on exactly one. That number
is the honest version of "usefully-meaningful subset" — not a global
claim, a per-program fact, published the way the dot count is.

**Middle — the waist has canonical names, and nobody ever sees them.** The
core forms are named in English (`set!`, `if`, `for`). That is fine — it
is how Unicode names characters and how the OED spells lemmas — *provided
no user ever reads a canonical name*. The alias table (LX.4) handles input;
rendering (Contemplation 2) handles output. The extreme is the state where
`set!` is as invisible to a Spanish user as a Unicode codepoint is to a
reader: it exists, it is the truth, and every surface is a projection of it
in the reader's language. Which is exactly why Contemplation 2's inversion
is the load-bearing engineering item for this contemplation as well: an
interlingua that anyone has to read is not an interlingua; it is a fourth
language.

### Two symmetric lints, which is what "don't hard-code" cashes out to

REBUILD already proposes one: **no English word literal below
`VLA_EnglishWords`.** The extreme needs two more, of identical shape:

- **No host literal below the domain profile.** No `range`, no `xl*`, no
  `activesheet`, no `:cell` in the core language or the sentence matcher.
- **No backend literal above the binding.** No `vla-step`, no `Optional …
  Variant`, no `raw` in anything that presents itself as a program's
  meaning.

Three greps, one runner, no Excel needed — the `tools/*.ps1` shape. A
language, a host, or a backend that has been fully separated is one whose
lint passes. Until the lints exist the separation is a belief, and until
the second natural language, the second host, and the third backend each
arrive without touching the other two axes, the belief is untested. The
architecture does not need to be built for the extreme. It needs three
seams that the extreme could be built through, and a check that says
whether they are still seams.

### What this is not

It is not a plan. It is the direction that decides, when two designs both
work today, which one to prefer: the one that puts a literal in data owned
by its axis over the one that puts it in the engine. The extreme is never
reached — the interval between a language and its speaker is dense, as the
essay this project keeps citing already proved — but every bisection
toward it is a real product: a second host is a product, a Python export
is a product, a Japanese phrasebook is a product. What the hourglass buys
is that each of those is a *sibling* of what exists rather than a *fork*
of it, which is the only version of "extensible in both directions" a
solo developer can afford.

---

## Contemplation 4 — Event Horizons *(2026-09-08)*

*The owner's question, 2026-09-08, filed the next day. Take the Excel
add-in as Frazaro's proof of concept. Then, in the owner's proposed
order: an Office Scripts runtime for future-proofing; a web interface or
fat client so translation is decoupled from any one environment; Google
Sheets via Apps Script; Apple's spreadsheet; Apache OpenOffice Calc. Then
orthogonally: English to SQL, for database administration and inside
Joomla. And the real question: what are the implications of Frazaro/VLA
not yet considered, and what does it unlock, beyond spreadsheets and
databases, when people can code deterministically — safely, reliably —
in their native tongue?*

*Sources read in full before answering: `README.md`, `SUBSTRATE.md`,
`VENTURE.md`, `VIABILITY.md`, `THREAT_MODEL.md`, `SECURITY.md`,
`ADVOCATUS.md`, `METAMETAMETALISP.md`, and Contemplations 1–3 above.
Read in part: `BETA_ROADMAP1.md` (the mission, the ordering rules, the
standing-decision register, the two-neutralities preamble, the closing
short answer, `LE.7`, `GO.3`, `PORT.1`–`PORT.3`), `BETA_ROADMAP2.md`'s
preamble, `CONSULTANT.md` §0–§2.6, `PREMORTEM.md` through D.2,
`MARKETING.md` §0–§3.2, `INTRINSICS.md`'s scope and first intrinsics,
`GRAMMAR_SINCE.md`'s rule, `LESSONS.md`'s opening, and `sententiae.txt`.
The shelf already holds the hourglass (Contemplation 3),
programs-as-forms (Contemplation 2), the AI drafting bridge (`LE.7`),
and the accidental object-capability property (`SD-15`); this
contemplation builds on those rather than restating them. Everything
below is what the shelf did not yet say. Like its three predecessors it
is speculation, not commitment: nothing here is an item until the
betting table says so.*

### The platform sequence, corrected by the architecture it stands on

**Six platforms is two engines.** The VBA emitter that exists today
already reaches LibreOffice Calc, because LibreOffice Basic runs a VBA
compatibility mode when a module opens with `Option VBASupport 1`. Apache
OpenOffice receives maintenance releases only, so LibreOffice is the
living target and OpenOffice comes along behind it. The other engine is
the JavaScript port of the translator — and eventually the interpreter —
that `SD-18` and `PORT.1`–`PORT.3` already anticipate. Office Scripts is
TypeScript, Apps Script is JavaScript, and a browser is JavaScript, so
the "web interface" step, the "Office Scripts runtime" step, and the
"Google Sheets" step are one engineering step with three bindings.
Whether each binding then interprets or emits — TypeScript for Office
Scripts, JavaScript for Apps Script — is a per-binding choice, and the
emitter route is the cheap one, as it is for the database below. Apple's
spreadsheet is Numbers, its automation surface is AppleScript and
Shortcuts, and it is the thinnest of the six. It belongs last, not
fourth.

**The fat client dissolves the ark objection.** `SUBSTRATE.md` H.3
declined a speculative JavaScript backend because it would be an ark
built from a forecast. A local-first web translator gives the same
engine a user-facing reason to exist, so the hedge and the product
coincide. That reorders the list: the JavaScript engine as the web
client comes second, and Office Scripts and Sheets fall out of it as
bindings.

**Cloud hosts invert the no-network pitch.** Apps Script runs on Google's
servers, and `VENTURE.md` §8 calls Python-in-Excel a non-starter for
exactly that reason. On cloud-native hosts the sentence changes from "no
network" to "no *additional* network — the procedure runs where the data
already lives." That is a doctrine amendment in `SD-13`'s own append-only
style, and it should be written before the port rather than discovered
by a reviewer after it.

**Capabilities are per binding, and spellings are forever.** Office
Scripts cannot touch the file system or mail on its own, so half of
`SEC.7` is solved by that host and none of it by Excel. The portability
measure from Contemplation 3 should therefore count capabilities as well
as forms. And because `SD-4` freezes every shipped spelling, an
Excel-only rule written today — anything about pivot caches or freeze
panes — is a spelling that will be carried into Sheets, where it must
refuse. Prefer intersection-profile rules now and mark Excel-only rules
as extensions at authoring time, not at porting time.

**The real product of the fat client is a file format.** Once the program
of record is forms with provenance and a hash (Contemplation 2), the
natural container is a host-neutral, language-neutral, signed procedure
document. Nothing like it exists: BPMN draws processes, RPA vendors hoard
proprietary formats, and no format lets the same procedure open in
Excel, Sheets, and Calc and render in the reader's language. An
interchange format for procedures is the standards-shaped artifact, and
the fat client is where it gets born. The same client gives exact
dry-run: because semantics are deterministic, "what would this do?" can
produce the effect list in English before anything runs, on every host.

### The database is not orthogonal; it is a homecoming

SQL was born at IBM as SEQUEL, the Structured English Query Language, in
1974. English-to-SQL is the second attempt at this idea, and the first
attempt became a priesthood. The warning that follows is specific: the
user of an English-to-SQL Frazaro is not the DBA, who already speaks the
priesthood's language, but the person who has to ask the DBA. Pitch and
phrasebook accordingly, or rebuild SQL.

What the database gives for free is everything being built by hand for
Excel:

- **Undo is a transaction.** Snapshot sheets become `BEGIN` and `ROLLBACK`.
- **Slot types are the schema.** A column's declared type is the
  `{n:number}` annotation, and a foreign key is a relationship word.
- **Refusal-before-run is a constraint.** A `CHECK` constraint is a
  `fail:` proof the DBA wrote years ago.
- **The capability model is `GRANT`.** `SEC.7`'s permission problem,
  unsolved in Excel because Excel has no permission model, maps onto
  row-level security and role grants the host already enforces.
- **There is no interpreter to write.** The query planner is the
  deterministic interpreter. Frazaro on a database is emitter-only, which
  makes it cheaper than the spreadsheet host, not more expensive.
- **Dry-run is `EXPLAIN` plus an affected-row count.** "This will delete
  rows from Customers; proceed?" is the same feature as the effect list
  above.

**Schema introspection writes the org phrasebook.** Every table and column
name is a noun the organization already chose. The concierge stream that
today needs founder hours to grow a vocabulary from an SOP can, for a
database, read `information_schema` and draft most of the phrasebook
mechanically. The DBA authored the corpus without knowing it.

**Joomla is the administration profile in miniature.** Its ACL is a
capability system, its language packs are per-community phrasebooks, its
User Actions Log is the run log `U.18` wants, and its REST API means the
JavaScript engine in the admin's browser needs no PHP port. A site admin
in Lisbon writes "Unpublish every article in category News older than a
year" in Portuguese, and the CMS profile has articles, categories, users,
and permissions as its finite object model. That profile is the first one
that is not a spreadsheet, and it will teach what the spreadsheet profile
smuggled into the core.

**The blast radius changes.** A deterministic wrong sentence in Excel
damages one workbook with an undo button behind it. The same sentence
over a production database is reliably wrong at scale, in one
transaction, with a commit. Dry-run with counts is therefore a core form
of the language, not an Excel convenience, and it should be designed
before the database emitter, not after.

### The deep horizon: what falls out when administration becomes legible

The interlingua-of-actions insight (Contemplation 3) bounds the horizon
honestly: deterministic native-language coding works wherever the domain
has a finite object model. The reframing is that most administrative
work already has one — forms, ledgers, eligibility rules, schedules,
tariffs, permits, compliance checks. That is the bureaucratic layer of
civilization, it is most white-collar labor, and it is exactly what LLM
agents are being pointed at. The counter-future is one where the rules
that run people's lives are readable, in their language, and either run
or refuse. Concretely:

1. **English stops being a prerequisite for computation.** Nearly every
   programming language keywords in English, so a bookkeeper in Lagos
   learns English before logic. Language-neutrality is not localization
   of a ribbon. It is the removal of a gate in front of most of humanity.

2. **Legible automation becomes due process.** GDPR Article 22, together
   with its "meaningful information about the logic involved" provisions,
   amounts to a right to explanation of automated decisions, and the only
   automation that satisfies it trivially is one where the explanation
   *is* the sentence. Catala already compiles French tax law from a
   controlled language, and the Bank of England piloted machine-readable
   regulation. The `DATALOG` and `PROLOG` engines are the formalism
   computational law uses. What legislative drafting lacks is the shadow
   audit that refuses overlapping rules at load and the refusals that
   teach, because today contradictory clauses are discovered in court.

3. **Policy about procedures, written in the same language.** Because
   there is no open-ended dynamic dispatch (`SD-15`, `SEC.1`), a
   program's effects are statically enumerable. "No procedure may clear
   the Ledger sheet" or "no procedure may email outside our domain" is
   itself a checked sentence that Check enforces over every other
   program. That is the compliance pack's real form: user-authored policy
   over user-authored procedures, provable before a run. Nobody can do
   that over VBA or over model-generated code.

4. **The consent layer between AI and the world.** `LE.7` makes Frazaro
   the safe target for AI-written spreadsheet automation. Generalized, a
   Frazaro profile is a human-readable tool schema: the sentence the
   model emits is the sentence the human approves, and
   one-sentence-one-meaning means the approval is of the actual action.
   The interpreter is a sandbox by accident, with no network and no
   first-class functions to smuggle a capability through. Every
   agent-actuated domain needs this layer, and none of them has it.

5. **Programming by voice, reliably.** A finite grammar is the only kind a
   speech recognizer can be constrained to, and constrained recognition
   is near-exact where open dictation is not. That opens programming to
   blind and motor-impaired users and to anyone whose hands are busy on a
   warehouse floor. It is the accessibility story the mission sentence
   already implies.

6. **A verified natural-language surface.** A core small enough to
   implement in a weekend, plus a grammar with no backtracking (`SD-16`),
   is small enough to verify formally, the way CompCert verified a C
   compiler. A machine-checked interpreter would make Frazaro the first
   natural-language-surface language with proven semantics, and that is
   the entry ticket to clinical order sets, batch recipes, and
   safety-rated procedures. Home automation is the low-stakes trainer for
   the same shape.

7. **Cross-lingual institutions.** The European Union treats every
   official-language version of a law as equally authentic and needs
   court doctrine to reconcile them when they diverge. Forms-as-record
   makes the form authentic and the languages renderings, which is the
   answer that body never had. For a multinational: Madrid writes the
   procedure, Osaka reads it, London audits it, and all three sign one
   hash. That is the end of English as the mandatory language of
   operations.

8. **Procedures that outlive platforms, people, and dialects.** VBA from
   1997 still runs; Python 2 scripts no longer run on a current
   interpreter. A procedure stored as forms under a constitution renders
   into the English of fifty years from now. The same property makes
   spreadsheet science reproducible, which the Reinhart–Rogoff error and
   the gene names Excel silently turned into dates both argue is overdue.

9. **Change control in sentences.** A semantic diff of two form trees
   renders as "this version now also clears column C." Managers and
   change boards approve diffs they can read, and the procedure's history
   is a readable ledger rather than a macro's binary blob.

10. **Standardization is the sixth destiny.** SQL, HTML, and C won when
    the specification escaped the vendor and competitors implemented it.
    `VIABILITY.md`'s five destinies stop at the Register; the end-state
    beyond it is one where Excel, Sheets, and Calc implement the core
    natively and Frazaro is the trademark on the reference
    implementation. That flips the sequence from "we port to six hosts"
    to "six hosts implement the spec."

11. **A new profession.** The automation role bifurcates into the
    procedure author, who is the operator, and the phrasebook
    lexicographer, who is a linguist-programmer nobody currently trains.
    Contemplation 1's curator-as-editor is the first job description for
    it.

### The uncomfortable implications, so the scales fall both ways

- **A shared phrasebook is a systemic risk.** One wrong base rule at
  month-end in a thousand organizations is a library vulnerability
  wearing a dictionary's clothes. `SD-4` protects spellings, not
  meanings, and `OUTPUT-EXCEPTION.md` covers generated code, not the rule
  that generated it. The Register needs a CVE process for grammar rules
  and a written liability position.
- **Meaning disputes migrate to translation.** When Madrid and Osaka
  disagree about what a form means, that is a treaty problem, not a
  compiler problem, and someone must own canonical meaning.
- **Determinism is not correctness.** A reliably wrong procedure is worse
  than an unreliably wrong one at scale. Dry-run and
  policy-over-procedures are the mitigations, and they belong in the
  core.
- **The boundary is real.** Finite object models cover most
  administration and almost none of thought. The horizon is the
  self-executing bureaucracy that citizens can read, not a general
  programming language for everything.

### One doctrinal note on the sequence itself

`SD-7`'s spirit, as `SUBSTRATE.md` H.3 applies it to backends, says no
backend without a user who needs it, and `SD-18` says a second host
follows the goldens rather than leading them. "Office Scripts for
future-proofing" is therefore the one item in the proposed sequence the
register argues against — and the fat client is what turns it into an
item the register allows.

---

## Contemplation 5 — A Theory of Categories *(2026-09-09)*

*The owner's question, 2026-09-09, filed the same day. Out of sheer
curiosity: what other orthogonal functions could Frazaro include in the
vein that produced `DATALOG`, `SQL` and `PROLOG`, with `SOLVE` scoped
behind them? A sentential API over an underlying DSL democratizes a whole
domain of programming — set theory, logic programming — by templating the
questions that can be asked of a set of data structures, and thinking in
category-theory terms (a field the owner has admired from afar) suggests
other domains could be exposed the same way. Or is that vertical thinking
where lateral is wanted? Frazaro began as English-to-VBA because
procedural macros were the obvious win, but "manager SOPs" —
sanity-checking values, querying workbook diffs, explaining* why *a number
changed by tracing formula logic — are a dollar-per-hour more important
automation the analyst never thinks about. And what about CFO and
executive SOPs: what auditing, high-level questions could an executive
ask in a Frazaro SOP that today are answered by clicking through large
workbooks?*

*Read in full before answering: `README.md`, Contemplations 1–4 above,
`BETA_ROADMAP2.md`'s QUERY AND LOGIC tranche, its standing-decision
register, THE MIDDLE LAYER, THE METAMETAMACRO LINE, THE EDITION LINE,
`U.18` and `SOP.1`–`SOP.5`; `MARKETING.md` §3.8–§3.11 and §4;
`VENTURE.md` §6; `sententiae.txt`; `PHRASEBOOK-TERMS.md`; the headers of
`VLA_Provenance.bas`, `VLA_Digest.bas` and `VLA_Identity.bas`; the
procedure lists of `VLA_Relation.bas` and `VLA_Datalog.bas`. Read by
section header only: `CUTS.md`, `SUBSTRATE.md`, the rest of `VENTURE.md`
and `MARKETING.md`. Measured, not assumed, because the reflection
candidate below depends on it: the engine can read and write ONE cell's
formula through the member form `(. obj formula)` (`Formula2` in both
backends), and `deflambda` emits one, but nothing walks a workbook's
formulas as data — a grep over `src/` finds no reference to
`Precedents`, `Dependents` or `HasFormula` at all — and `english.vla` has
no rule whose surface is an audit verb (precedent, dependent, explain,
compare, reconcile, variance, verify). Like its four predecessors this is
speculation, not commitment: nothing here is an item until the betting
table says so.*

The four engines are not four random picks, and that is why "what else"
feels both obvious and hard. They are the classic rungs of one ladder.
The genuinely orthogonal directions are not further up that ladder; they
are sideways — change what an *answer* is, or change what is being
*asked about*, and keep the evaluators already built. The owner is not
overshooting with category theory: there are exactly three places where
it earns its keep here, named below rather than waved at.

### Where the orthogonality actually is

**The vertical is mostly climbed.** Descriptive complexity classifies
query languages by expressive power. First-order logic is SQL without
recursion. Add least fixpoint and you have Datalog, which stays in
polynomial time. Add guess-and-check and you have answer set programming,
the NP rung. Add unbounded recursion over terms and you have Prolog,
which is everything computable. Frazaro has one engine per rung, with
`SOLVE` the only one unbuilt. Above Prolog is undecidability, which the
refusal doctrine cannot price. So the fifth engine should not be "a
stronger logic."

**Axis two is the codomain: what an answer is.** This is the
category-theory handle that is both real and implementable on the
relation substrate. Green, Karvounarakis and Tannen (*Provenance
Semirings*, PODS 2007) showed that relational and Datalog evaluation is
parametric in a semiring: join multiplies annotations, union adds them,
and the same evaluator yields a different kind of answer depending on
the semiring plugged in. Booleans give set semantics. Counts give bag
semantics. Sets of input-row identities give *why* a row is in the
answer. Polynomials give *how* it was derived, every path. The tropical
semiring gives shortest paths and minimum cost. `VLA_Relation` is the
evaluator; provenance is one extra column on every tuple and a few lines
in `RelJoin` and the union step. One evaluator, many products — and the
most valuable product for a manager is *why*, because that is the
question asked after every exception report. One caveat a Datalog person
will raise first: recursion needs the finite semirings. The
how-polynomial is infinite on a cycle, so that one is refused or
truncated by name, never left spinning — the step-ceiling doctrine
again.

**Axis three is the domain: what is being asked about.** Every engine
today reads a Table. But a workbook contains other relational structures
nobody has exposed as tables: which cells hold formulas, which cells
each formula reads, what names exist, which sheets are hidden, what the
last snapshot held, what a program's effect list was. Expose those as
tables and every existing engine answers questions about the workbook
itself. `DATALOG` over a precedents relation is "every cell that
ultimately depends on the tax-rate input," recursion included, at zero
engine cost. Spivak's functorial data migration (*Information and
Computation*, 2012) is the honest categorical handle here: a schema is a
category, an instance is a functor into sets, and a change of layout
between two months is a functor between schemas. That matters for
reconciliation across layout changes, below.

So there are three ways to add a question. A new logic, of which two are
worth having. A new answer type: provenance, verdicts, diffs,
attributions. A new source: the workbook about itself. The second and
third are where the manager and executive questions live, and both are
cheaper than a fifth logic. The one earned category-theory sentence: the
engines are functors, and you can change the object they are applied to
or the semiring they are valued in without touching the functor.

### The candidates

| Candidate | The formal thing | The question it answers | What it reuses | Shape |
|---|---|---|---|---|
| **Reflection** | the workbook as relations | what is hardcoded in a formula column; what depends on this input | all four engines, unchanged | ribbon action, not a worksheet function |
| **Check** | integrity constraints with witnesses | does every subtotal foot; is every invoice number unique; is this column monotone | the `SQL` and `DATALOG` evaluators; the new part is a verdict result type carrying counterexample rows | either |
| **Diff and reconcile** | key-matched symmetric difference; matching under a tolerance relation | what changed since last month; which bank lines match which ledger lines | `RelJoin`; the snapshots Undo already takes | either |
| **Why** | semiring provenance | why is this row in the answer | the relation substrate | worksheet function |
| **Attribute** | sequential attribution over a formula graph; Shapley is the exact version | why did this cell change, by input, as a bridge table | reflection, snapshots, the host's own recalculation | ribbon action |
| **Row patterns** | regular expressions over rows, Kleene algebra; SQL:2016's `MATCH_RECOGNIZE` | three consecutive declining months; a spike then a reversal | relation rows plus a small automaton | worksheet function |
| **Decide** | decision tables with a hit policy, DMN semantics | which tier, which discount, which approver | the load-time shadow audit, which is already a uniqueness hit-policy check | worksheet function |

Notes on the ones with a catch:

- **Reflection cannot be a worksheet function.** A function recalculates
  when its arguments change, and Excel does not treat a formula elsewhere
  as an argument. Making it volatile only makes the answer depend on when
  Excel last recalculated — the exact objection `PROLOG.19` is about to
  record for `assert`. So reflection engines take the Check and Interpret
  shape, or take an explicit snapshot as their argument. Worth stating
  before building. The real cost is small: `Range.Precedents` stops at
  the sheet boundary, so a tiny reference tokenizer over formula text
  does the cross-sheet half.
- **Attribute needs the host.** Recomputing with one input swapped at a
  time means copying the sheet, setting inputs, reading outputs. It is
  deterministic only when no formula on the path is volatile or
  external, and reflection can detect those and refuse by name — the
  screen `SEC.15` already wants, widened to volatility. A waterfall is
  order-dependent, so the sentence must state the order; Shapley removes
  the dependence at exponential cost, affordable for the handful of
  inputs a real bridge has.
- **Two whys share one word.** *Why is this row in the result* is
  provenance. *Why is this value in this cell* is precedents plus
  attribution. The sentence layer will receive both and should refuse
  the ambiguous one.
- **Check is on the path to `SOLVE`, not a detour.** `SOLVE.2` checks
  constraints against one derived world and reports "no answer set."
  Check is that same step with zero choices, pointed at the data,
  rendering the violating rows instead of a verdict. Build it first and
  `SOLVE.2` inherits it.
- **Decide is where the shadow audit becomes a product.** Business
  people already write decision tables in Excel. What they cannot do is
  prove a table complete and non-overlapping, and over interval
  conditions that is a sweep, not a satisfiability problem.
- **Row patterns are the purer of the two new logics, and cheap.** A
  deterministic automaton with a stated policy — leftmost,
  non-overlapping — and no backtracking. It parses rows rather than
  sentences, so it sits inside `SD-16`'s spirit.

### Manager and executive SOPs

The manager's review, as engines:

- **Plug hunting.** A hardcoded number in a formula column. Reflection
  plus one Check sentence. One of the most common findings in the
  spreadsheet-risk literature.
- **Footing and cross-footing.** Subtotals equal their rows; the summary
  page equals the detail. Check.
- **Structure drift.** Is this the same model as last month with only
  inputs changed? Diff the formulas and expect nothing; diff the values
  and expect something. Reflection plus Diff.
- **Reconciliation.** Bank to ledger, subledger to general ledger,
  intercompany. Diff with a stated tolerance and a stated first-match
  rule, so the matching is auditable rather than fuzzy.
- **The change question.** Why did gross margin move? Attribute, as a
  bridge table.
- **Sign-off with evidence.** The verdicts and the diff are the
  evidence; `U.18`'s run log is where they persist.

The executive does not open the workbook. Their SOP is a questionnaire
the workbook answers itself, and nearly every question is a manager
engine with a policy sentence on top:

- **Does every number on the summary page trace to a source sheet with
  no manual override on the path?** `DATALOG` reachability over
  precedents, joined to the plug table.
- **What are the top drivers of variance to budget?** Attribute.
- **Which assumptions is the forecast most sensitive to?** Attribute over
  hypothetical deltas, one input at a time.
- **Are we compliant with our own rules?** Contemplation 4's
  policy-over-procedures for programs, plus Check for data.
- **What are the staffing options and which is cheapest?** `SOLVE`, as
  already scoped.
- **What would it take to hit the target?** Goal seek: deterministic root
  finding, but approximate, so it would state its tolerance the way
  `PROLOG` states its step ceiling.

The pattern is that executive questions are not new engines. They are
the manager's engines under a policy sentence, rendered as one page of
verdicts. That page is the compliance pack of `VENTURE.md` §6.3, with an
artifact behind it at last.

### Refusals, and the order that pays

What should be refused by name, in `PROLOG.19`'s shape — a message
giving the reason, never implying "not yet": probabilistic logic,
because the brand is certainty; random sampling without a stated seed;
anything with a model in the loop.

Ranked by what each unlocks, the house rule: reflection first, because
it turns four engines loose on the workbook itself for almost nothing.
Check second, because it is the manager's engine and `SOLVE.2`'s base
case at once. Diff third; then Attribute as the distinctive one; then
row patterns and Decide as the two cheap new logics. None of this is an
item: `SD-7` wants a sentence that needs each, and the QUERY AND LOGIC
tranche's own contract is aspirational and appetite-boxed. What this
contemplation buys is the direction that decides, when a lateral ask
arrives, whether it is a fifth engine, a fifth answer type, or a fifth
table — and the bet is that it is almost never the first.

---

## Contemplation 6 — Notes from the Coffee Grounds *(2026-09-09)*

*The owner's question, 2026-09-09, filed the same day, in the owner's
own framing: for yet another dash of characteristic meta-self-awareness,
what questions should the owner be asking about this project that have
not been asked, as evidenced by the copious documentation? Existential,
orthogonal, observational, marketable, or otherwise useful questions
that do not quite fit the mental filters `docs/` has set up for itself.*

*The title is the owner's, and it is doing at least four jobs: the
underground it echoes is Dostoevsky's, whose narrator is the patron
saint of the question; the grounds are what a filter keeps out of the
cup, which is precisely what was asked for; they are also what a
fortune-teller reads the future in, the coroner's own trade; and they
are the ground the project stands on, which the historian already
surveyed.*

*Correction, appended 2026-09-09, same day, at the owner's prompting:
the count was low. "Notes" are also what a taster reports in the cup,
so the contemplation is an aromatic experience — and the one place the
text below smells anything is Hermans' spreadsheet smells, cited as the
manager's own checks. Five jobs, not four. The paragraph above stands
as written, per the house rule against editing oneself wiser after the
fact.*

*Filed as a contemplation rather than as a sixth execution of SD-17,
deliberately: it occupies no persona's seat, it proposes seats, and its
one concrete proposal (the collector, below) is a candidate for the
roadmap and is not minted here. Like its predecessors it is speculation
until the betting table says otherwise — with one difference the reader
should hold onto: the first section is not speculation but a
measurement, and it is the reason the rest was worth writing.*

*Read in full before answering: `PREMORTEM.md`, `ADVOCATUS.md`,
`CONTINUITY.md`, `VIABILITY.md`, `SUBSTRATE.md`, `CONSULTANT.md`
(including its addendum), `THREAT_MODEL.md`, `TESTING.md`,
`SECURITY.md`, `SUPPORT.md`, `docs/README.md`, the root `README.md`,
and Contemplations 1–5 above. Read in part: `BETA_ROADMAP2.md` (the
departments, the standing-decision register, SIGNATORY, PATIENT, SOP
IMPORT, ENVIRONMENT, THE MIDDLE LAYER through QUERY AND LOGIC, and
INTERFACE through the closing short answer); `AUDIT.md` (I.0–I.4 and
Part III); `MARKETING.md` (§0–§2, §3.8–§3.11, §4, §7–§8); `VENTURE.md`
(§2, §6, §8, §11–§13); `LESSONS.md` (the chapter headings, X and XVI);
`RELEASES.md` (the head); `ID_REGISTRY.md` (the prefix table);
`TRENCHES.md` and the `metameta/` essays by heading only. Everything
counted below was counted by grep over the working tree on 2026-09-09,
and the last section says how to count it again.*

The shelf asks harder questions of itself than most funded teams ever
do, so the gap is not rigor. It is shape. Every file on the shelf has a
slot for a decision, an item, an incident, a death, or a speculation.
Three kinds of question fit none of those slots: questions whose answer
is a fact about the world rather than a thing to build, questions only a
third party can answer, and the question of whether the questions
already asked were ever answered. The third has a measurement behind it,
so it goes first.

### The question the shelf cannot ask about itself

The five reviews of 2026-08-31 each ended with a register, and every
register said the same thing in its own voice: the fork is collections
day, and an uncollected register is death by documentation in costume.
The collection was checked on 2026-09-09, nine days and five tagged
releases later, by grepping each proposed candidate ID against both
roadmap files.

| Proposed | By | Where it stands on 2026-09-09 |
|---|---|---|
| `SIG.6`, one practicing auditor's verbatim reaction | ADVOCATUS | Not minted. Cited as if real twice in `VENTURE.md` (§6.3, §7) and once in `MARKETING.md` (§3.10). `ID_REGISTRY.md` still lists it as the SIG family's next free slot. |
| `SIG.7`, the win condition with an expiry | VIABILITY | Not minted. `VENTURE.md` was written afterward and also declines to choose. |
| `EN.9`, the substrate watch-list | PREMORTEM, hydrated by SUBSTRATE | The ID was spent on a different item — `Workbook.Path` as a cloud URL breaking `Dir$` checks. The watch-list, its contents already written, has no home. |
| `EN.10`, the factory's trust dependencies | SUBSTRATE | Not minted. |
| `CN.1`, the bus drill with a pass bar | PREMORTEM, CONTINUITY | Not minted. |
| `CN.2`, the second address | CONTINUITY | Done in substance: `git remote -v` now names an origin. No ID. |
| `CN.3`, the owner's cold-start brief | CONTINUITY | Not minted, and no file plays the role. |
| `DZ.1`, the drizzle ledger | PREMORTEM | Not minted. TERRARIUM is the nearest thing and demands a repro and a root cause — the opposite of a shrug, by its own design note. |
| `DO.7`, the front door | VIABILITY | Done in substance: the root `README.md` exists. No ID. |

Neither `RELEASES.md` nor `BETA_ROADMAP2.md` contains the words
"premortem", "drizzle", or "receivable"; the roadmap's single
"tripwire" is `PROLOG.9`'s own name for a PowerShell check, not a
re-read of the coroner's. `PREMORTEM.md`'s Ω predicted exactly this
fate for itself — filed, read once with interest, never reopened, its
tripwires unmonitored — and the prediction came true inside nine days.
It is also `AUDIT.md` I.2's unbound-symbol finding recurring one level
up: `SIG.6` is referenced in two outward-facing documents and defined
nowhere, the precise shape "first external user" had before THE PATIENT
was seated.

So the first unasked question is **who collects?** The shelf has a
minting discipline, a review cadence, and an ID registry, and no
mechanism that reads a review's register against the roadmap at the
next fork. The house-style answer is one more check in `tools/`, the
shape the eighteen existing ones already take: for every "proposed
mint — candidate X" in a review, X is minted in the roadmap, or done
under another ID with a pointer back, or declined with a written reason.
Anything else is red. It would have been red every day since the beta
shipped. Call it the collector; it is a check, not a persona, and it is
the only proposal in this contemplation.

The second is structural. **Where do open questions live?** The
register holds decisions. `LESSONS.md` holds incidents.
`CONTEMPLATIONS.md` holds speculation. The roadmap holds things to
build. `ADVOCATUS.md` invented the right shape — a claim, the evidence
that would settle it, and who must produce it — used it six times, and
never gave it a file. The owner's question of 2026-09-09 has no home on
the shelf, which is why it had to be asked in a chat window.

### Questions with no seat

**Existential.**

- **Who has the incentive to write the procedure down?** `VENTURE.md`
  §2 names the manager's want as procedures that survive people. The
  person asked to author the procedure is the person it makes
  survivable without. `ADVOCATUS.md` A.5 asks whether the clerk wants to
  program; nobody asks whether the clerk wants to encode their own moat.
  The answer changes who the tutorial is for.
- **What is a number?** `PROLOG.17` made doubles permanent for host
  parity — but `VLA_Relation.CompareValues` compares two numerics with
  VBA's exact `=`, and so does the interpreter's own equals. In VBA,
  `0.1 + 0.2 = 0.3` is False, by construction. Excel's own `=` is
  widely documented to forgive the last bit, so a cell computing a tenth
  plus two tenths equals three tenths in the formula bar and may not in
  a `SQL()` `WHERE` or an `is equal to`. No document on the shelf
  mentions floating point, epsilon, or tolerance. For a finance-first
  checked language that is the founding question, and it is one cell
  away from being measured.
- **Does automating a spreadsheet promote it into a heavier control
  regime?** Under model-risk and end-user-computing policies, a
  spreadsheet that produces reporting numbers by rule can be
  reclassified from "EUC" to "application" or "model", which brings
  validation, change control, and a named owner. The compliance pitch
  assumes automation lowers the buyer's burden. It may raise it at
  exactly the buyer it targets. "Model risk" appears nowhere on the
  shelf.
- **Which destiny?** Asked by the accountant, never minted, never
  answered. `VENTURE.md` §13 repeats the question rather than the
  answer.

**Orthogonal.**

- **The field that studied this user for thirty years is absent.**
  Nardi's *A Small Matter of Programming* argues spreadsheet users
  succeed because the formula language is task-specific and hides the
  general machinery. Blackwell's attention-investment model prices the
  analyst's decision to automate — the exact decision A.5 speculates
  about. Panko's error-rate studies are the denominator that
  `sententiae.txt` says every estimate needs. Hermans' spreadsheet
  smells are Contemplation 5's manager checks, already catalogued.
  Hermans and Panko appear once each on the shelf, as audiences to sell
  to; Nardi, Blackwell, Burnett, Ko, and the phrase "end-user
  programming" appear zero times.
- **Cucumber is the closest living relative and is not on the shelf.**
  Gherkin is a checked English whose step definitions are phrasebook
  rules by another name, with an "undefined step" refusal and nearly two
  decades of published post-mortems: step sprawl, the
  imperative-versus-declarative wars, and the non-technical authors who
  did not, in the end, write the steps. That is a premortem someone else
  already paid for. Zero mentions.
- **The landlord's own natural-language record is unread.**
  `SUBSTRATE.md`'s exhibits are all about VBA. Excel shipped
  natural-language labels in formulas in 1997 and removed them in 2007;
  Analyze Data answers English questions over a table inside Excel
  today and is Contemplation 5's executive question box with
  Microsoft's name on it; Flash Fill is programming by example. What
  Microsoft learned and abandoned about English in cells is the exhibit
  the historian did not pull.

**Observational.** The repository holds no observation of a human using
the product: no screenshot, recording, transcript, or timed session
(`MARKETING.md` §1 already notes the absence of any image). Every
document reasons from other documents and from code. Five instruments
need no user:

- **The reader test.** Five strangers, one program, ten minutes each,
  asked to say what it does. The claim that a reviewer can read what ran
  has never met a reviewer; `ADVOCATUS.md` A.2 asked only for a
  stranger's *writing* hand, and reading is cheaper and is the pitch.
- **The acceptance benchmark, asked once and never run.**
  `CONSULTANT.md` §2.13 proposed measuring the refusal rate against
  external text. Forum question titles are a free corpus of what people
  want to say to Excel, in their own words. A day of work, and the first
  coverage number that is not a prediction.
- **Engine capacity.** `PROLOG_MAX_STEPS` is a total budget of 120,
  charged per candidate, and `PROLOG.9`'s own entry records the usable
  range of a single `between` goal as 119. No document says what size of
  real table each engine answers before refusing. A rota with fifty
  shifts and twenty staff may already be over the line.
- **The number experiment.** One cell holding `=0.1+0.2`, one `SQL()`
  comparing it to `0.3`, one formula doing the same. One minute. It
  settles the existential item above.
- **The judgment step as evidence.** The grammar can ask (`Ask … and
  put answer into …`), branch, and `Stop.`, so a "check with Priya"
  step can be composed today. Nothing records that Priya said yes: the
  run log (`U.18`) is unbuilt, and an auditor's question about a manual
  step is who approved it and when, not whether the program paused.

Two more observations are about the shelf itself, and both are counts:

| Measured over `docs/*.md`, 2026-09-09 | Count |
|---|---|
| Occurrences of "the owner" | 237 |
| Occurrences of "the user" | 117 |
| Occurrences of "the pilot" | 43 |
| Lines in `docs/*.md` | 27,948 |
| Lines in `src/*.bas` and `src/*.cls` | 66,510 |
| Lines in `english.vla` plus `prelude.vla` | 3,654 |

The owner appears twice as often as the user, and every review is in one
voice, written by the dyad under review — which `ADVOCATUS.md`'s Ω
confessed and `VENTURE.md` §11 presents as diligence. **Has any human
other than the owner read the shelf, and how long did it take them?**
`CONTINUITY.md` assumed the shelf explains itself. That is a forecast
with no reader behind it, and the docs' own specimen count is therefore
the same number as the product's.

**Marketable.**

- **Positioning by exclusion.** `MARKETING.md` serves fifteen
  audiences, `VENTURE.md` three constituencies, `VIABILITY.md` five
  destinies. No document names an audience the project refuses to serve
  this year. A beta that says yes to everyone has not said who it is
  for.
- **Two products in one box.** The English SOP tool and the query and
  logic engines share a download and do not share a buyer: the
  accountant did not ask for Prolog and the Prolog person did not ask
  for month-end. Whether the box is the product or the confusion is
  unasked.
- **Whimsy and the signatory.** The pirate and alien phrasebooks
  recruit engineers and ship in the same file as the SOX control.
  Whether the edition a CFO installs should carry them is a question
  about the buyer, not about taste.
- **Where the moat is thickest.** Against Copilot the edge is largest
  where English is weakest, because Copilot emits English-keyword code
  and Frazaro renders in the reader's language. The beta is
  English-first — aiming, on this reading, at the market where the
  differentiator is thinnest.

### Seats SD-17 has not filled

The rule demands a persona nobody has used, and `SUBSTRATE.md`'s Ω
already said the cheap seats are taken. The questions above imply these,
each with the finding it would own:

- **The labor economist**, for the incentive inversion and the
  attention-investment arithmetic.
- **The end-user-programming researcher**, to read Nardi, Blackwell,
  Panko, and Hermans against the shelf and say which premise they
  already falsified.
- **The reader** — a stranger handed one program and a stopwatch. Not a
  user. The pitch is about readers.
- **The model-risk officer**, for the reclassification paradox and for
  what "a control" means to the people who define the word.
- **The competitor's product manager**, who knows why natural-language
  labels were removed and what Analyze Data taught Redmond.

And one that is not a persona but a check: the collector. If only one
thing in this contemplation becomes work, it should be that check,
because it is what keeps every other review — this one included — from
becoming a monument.

### The numbers, and how to re-run them

All counts above come from the working tree on 2026-09-09; no Excel is
needed, and every line below is a plain shell command:

```
# Which proposed mints exist as roadmap items (bold ID at bullet head)?
for id in SIG.6 SIG.7 EN.9 EN.10 CN.1 CN.2 CN.3 DZ.1 DO.7; do
  printf '%-6s %s %s\n' "$id" \
    "$(grep -c "\*\*$id\b" docs/BETA_ROADMAP2.md)" \
    "$(grep -c "\*\*$id\b" docs/BETA_ROADMAP1.md)"; done
# Where are they referenced at all?
grep -n "SIG\.6\|SIG\.7\|CN\.[123]\b\|DZ\.1\|EN\.10" docs/*.md README.md
# Did any fork re-read the coroner?
grep -nic "premortem\|drizzle\|receivable" docs/RELEASES.md docs/BETA_ROADMAP2.md
# The shelf's vocabulary
for t in "the owner" "the user" "the pilot"; do grep -roi "$t" docs/*.md | wc -l; done
# The literature, and the landlord's own record
for t in Nardi Blackwell Burnett "end-user programming" Panko Hermans \
         Gherkin Cucumber "Analyze Data" "natural language label" \
         "Flash Fill" epsilon floating "model risk"; do
  printf '%-24s %s\n' "$t" "$(grep -rli -- "$t" docs/ README.md | wc -l)"; done
# Any observation of a human?
find . -iname '*.png' -o -iname '*.gif' -o -iname '*.mp4' -o -iname '*transcript*'
```

---

## Contemplation 7 — Bergson's Burger *(2026-09-10)*

*The owner's question, 2026-09-10, after three rounds of `VOX_POPS.md`:
why were the vox pops funny, and how is that humor replicated in future
contexts — a specific combination of constraints, or a happy accident?
The answer given in conversation, re-summarized here at the owner's
request. The title is the owner's. Bergson's essay on laughter holds
that we laugh at the mechanical encrusted upon the living; a burger is
the layered thing this contemplation ends with, a recipe stacked in a
fixed order. Both readings are used below.*

*Source: `VOX_POPS.md` itself — seventy-two invented quotations in three
rounds, written 2026-09-09 and 2026-09-10 — and the session that
produced them. Like every contemplation this is a mental model, not a
commitment; unlike most, it has its specimen already on the shelf.*

### Not an accident: four constraints

**The question was a machine.** "If you could program spreadsheets in
English, what would you do?" is a fixed frame with one free slot, the
persona. Everyone answers from their own job, so every joke has the same
shape — the gap between what the tool promises and what that job
actually wants — and a new persona yields a new joke for free. The frame
does not deplete, which is why round three was as strong as round one.

**The product is a straight man.** Frazaro's defining behaviour is that
it says no, and a tool that refuses is a deadpan partner. Most of the
quotes are built on it: the persona asks for the thing they really want,
which is impossible or dishonest — "Make Q3 look better," "Allocate
overhead fairly," "Approve everyone who looks fine" — and the refusal
reveals the person. Nobody invented the punchline mechanism. The refusal
doctrine supplied it, and a refusal that teaches is already a comic
form.

**The material was pre-loaded with true tensions.** The jokes were
written after a day inside the shelf, and the funniest lines are
compressed findings: the warehouse supervisor is "determinism is not
correctness," the senior accountant is Contemplation 6's incentive
inversion, Cabo is the bus factor, the helpdesk technician is the install
cliff. Comedy is the shortest route to an uncomfortable true sentence,
and `sententiae.txt` is already joke-shaped — "Near money, say less" is
a punchline with a moral. The jokes were the dividend on the reading;
without it they would have been jokes about spreadsheets in general,
which exist in bulk and are worse.

**The form mirrored the product.** A Frazaro program is a list of
sentences run in order, and the best quotes are exactly that: "Undo
whatever Kevin did. Find out what Kevin did." "Balance the drawer. Find
the eleven cents. Stop looking for the eleven cents." The last command
reveals the person. The product's core metaphor handed the jokes their
native rhythm.

### Bergson, applied

Bergson's thesis is that laughter answers *the mechanical encrusted upon
the living*: a person behaving like a mechanism, rigid where life should
be supple. The vox pops run the thesis in both directions at once. The
tool is the mechanism and the persona is the living thing that wants —
that is the straight-man structure above. But the sharpest lines are the
ones where the *person* has become the procedure: the branch manager who
cannot stop looking for eleven cents, the owner of The Master
Spreadsheet whose whole system is two tabs named DO NOT DELETE, the
payroll clerk standing very still. The joke is a human who already runs
in sentences, meeting a tool that finally does too.

Two more of Bergson's clauses explain two of the rules below. Laughter
is a *social gesture*, a corrective: the cynic who concedes the pitch
while mocking it — "wrong the same way every time; we call that a
process" — is the audience correcting the marketing and agreeing with it
in the same breath, which is the only form in which a product's argument
is safe to post. And laughter needs *a momentary anesthesia of the
heart*: momentary, not permanent — which is why the jokes may punch at
situations and never at competence. The actuary's spreadsheet is right.
The adjuster's own claim is lying. The person is never the fool; the
doctrine already decided who is.

### The recipe, stacked in order

- **Write the tag as a scene, in five words.** "Bank branch manager,
  after close." The tag carries the context so the quote can start
  mid-thought and end on the punch.
- **Replace every abstraction with the one noun the persona would say.**
  Not "the person who knows the spreadsheet" but "She's in Cabo." Doreen,
  4B, Gary, the red truck, DO NOT DELETE 2. Specificity is where
  "written" turns into "overheard."
- **Get the in-group tell right.** The SRE is four words long. The
  consultant speaks in slide numbers and a fee. The IB analyst says
  "Buddy." If the quote could be said by anyone, the tag is doing
  nothing.
- **Let one in six reject the premise.** "Program it to do what? Like,
  emotionally?" Without those the list is a brochure with jokes in it.
- **Concede the pitch while mocking it.** The move that makes the humor
  postable.
- **Never lie about the product inside a joke.** The compliance
  officer's run log does not exist yet, and the field notes say so. Jokes
  may exaggerate people; they may not exaggerate the tool, or the brand's
  one asset goes with them.
- **Punch at situations, not competence.** Bergson's anesthesia,
  momentary.
- **Ration the callbacks.** Kevin once per round. A running gag rewards
  reading the whole list; a mascot ends it.
- **Cut to the shortest version and end on the concrete word.** "Ours is
  called Thursday." "'Assume a spreadsheet.'"
- **Make a fifth of them secretly findings, and write the notes.** The
  field-notes section is what makes the jokes feel earned, and it is how
  a reader learns the doctrine without being taught it.

### To get it again

Three inputs and a register. The fixed question; a list of jobs with a
native grievance; and the domain's true tensions, either as a document
to read or as one sentence each. The register set in the prompt matters
as much: "guerrilla journalism" and "man on the street" granted
permission for irreverence, and the pirate enthusiast would not have
appeared under "write marketing copy." The one part honestly owed to
luck is that the writer had been inside the shelf's voice for hours, and
that voice is understatement. Understatement is most of comedy, and the
shelf has been rehearsing it for a year without meaning to.

---
