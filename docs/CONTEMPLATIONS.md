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
