Attribute VB_Name = "VLA_SentenceEngine"
Option Explicit
Public Const VLA_SENTENCEENGINE_VERSION As String = "PPROF.0"

' SEC.2: EnglishLoadVocabulary's raw-consent gate result codes. Plain
' Long + Const, not an Enum - this codebase's own 27 modules use Enum
' nowhere else. Declared HERE, in the module's own top-of-file
' Declarations block, not near where they're used (EnglishLoadVocabulary,
' far below) - three live compiles (owner-caught) demonstrated in
' sequence that VBA requires every bare module-level Const/Dim/Enum/Type
' declaration to live together in this one block, before the first
' Sub/Function/Property in the module, full stop - not merely "declared
' above its own first use," which is what the second attempt (still
' misplaced, just relocated to a different mid-file spot) got wrong.
Private Const RawConsentDeclined As Long = 0
Private Const RawConsentWorkbook As Long = 1
Private Const RawConsentDevice As Long = 2
' SEC.11, 0.5.3: where a remembered raw-consent grant is filed - the
' registry section under "Frazaro" for device scope, and the prefix of
' the CustomDocumentProperty name for workbook scope.
'
' The "V2" is the whole point, and it is a deliberate owner decision
' rather than a naming accident. Every grant written before 0.5.3 is
' keyed by the OLD 32-bit polynomial value; after SEC.11 no such key can
' ever match again, so every one of them is dead weight. Giving the new
' generation its own section means old and new keys CANNOT collide even
' in principle, the re-prompt is guaranteed rather than merely likely,
' and the stale entries sit inert under a name that says what they are.
'
' Nothing migrates and nothing is deleted, both on purpose. Re-prompting
' is the fail-safe direction - the worst case is that a person is asked
' once more about a phrasebook they already trust, which is exactly the
' question SEC.2 exists to ask. And SEC.10 is about to move workbook-
' scope consent device-side anyway, so any migration machinery written
' here would be thrown away before it ever paid for itself.
Private Const RAW_CONSENT_SECTION As String = "SEC2RawConsentV2"
' PPROF.0: EnglishToVla gains the translate-side half of P-PROF's
' per-phase timing (VLA.bas's own PPROF.0 note has the full mechanism
' and gating reasoning) - trans-tokenize/trans-build, behind the same
' VLA.mProfileOn switch VlaTranspile's compile-side phases check, added
' specifically because P-TOK's own named hotspot (TokAt, this file -
' positional Collection indexing, confirmed still present) lives here,
' not in VlaTranspile, and the archived compile-side phase list never
' covered it. Zero cost when the switch is off: Timer is never read.
' EDITIONMANIFEST.1: EnglishAuditPhrasebook now also accepts a single
' pre-built Variant array (VLA_Build.bas's own per-edition phrasebook
' chains, assembled at runtime) alongside its existing ParamArray of
' literal paths - VBA has no array-unpacking syntax, so a caller
' holding an array couldn't spread it across positional arguments the
' way every existing caller already does. See the function's own
' header, just above its body, for the detection rule.
' LX5.2: split out of VLA_English.bas (BETA_ROADMAP2.md's LX.5, phase 2)
' - this module is the language-neutral sentence machinery: grammar
' registration, the DCG phrase-matching engine, G-RENDER, statement/
' condition/expression parsing, vocabulary-file loading/diff/profiling/
' audit tooling, sub-assembly bookkeeping, diagnostics, and the Public
' front door (EnglishToVla and friends - kept their exact names; VBA
' resolves an unqualified Public call project-wide regardless of which
' module hosts it, so this move needed no caller-side change anywhere
' in the project). VLA_English.bas keeps only this language's own
' vocabulary tables (NumberWord/OrdinalWord/ExprOpWord/IsNoiseWord/
' SkipArticles/IsColorWord/SlotDesc/StrayCharHint) and LX5.1's
' keyword-alias seam - see its own header for the full reasoning and
' the variable-usage audit that decided this exact boundary. A
' phrasebook's (keyword-alias "si" "if") directive still resolves
' through VLA_English.RegisterKeywordAlias; CanonicalizeStructuralWords
' still runs (from VLA_English.bas) right after every EnTokenize call
' below that feeds ParseStmt.
' Below this point is VLA_English.bas's own pre-split header,
' unedited: the DCG design spec and the full development changelog
' this module's code carries forward from where it was written.

' PROLOG.1: UnifyForm (G-RENDER's own one-way unifier, the FormSubstitute
' inverse) is now a thin pass-through to VLA_Unify.UnifyOneWay
' (VLA_Unify.bas, new module) rather than its own implementation -
' BETA_ROADMAP2.md's own PROLOG entry, "the unifier is substrate, not
' engine-private." Read in full before touching it: every one of
' UnifyForm's four cases (bare {name} slot, glued slot make-{d}/xl{d},
' literal atom, list recursion) turned out to be completely DSL-
' ignorant on inspection - the glued-slot branch is pure prefix/suffix
' string arithmetic over its own two arguments, nothing English-
' specific about it - so there was no atom-level hook left to keep
' here; UnifyForm is a one-line call now, not a partial reimplementation.
' One real behavior change rides along, found by reading UnifyForm and
' its own RenderBoundLookup together rather than assumed: the original
' never checked that a slot NAME repeated within one template bound the
' same value both times (RenderBoundLookup silently returned whichever
' occurrence it scanned first); VLA_Unify.UnifyOneWay now enforces that,
' real unification's own rule. See VLA_Unify.bas's own header for the
' real corpus case this protects (english_expanded.vla's own
' "paste-values", two templates - one with distinct {a}/{b}, one
' reusing {r} twice for the paste-in-place shorthand) and
' VLA_Tests_Query.bas's own TestGRenderUnify (via TestDSLs) for the pin.
' LX2.0: this module's 96 of 97 raw Err.Raise refusal sites now route
' through VLA_Messages.RaiseMsg with a stable id - SD-2/LX.2. The 97th
' (EnglishAuditText's cleanup, ~line 7606) is a bare re-raise of an
' already-caught error, not an origination of new English text, and
' stays raw by design, same category as VLA_IDE.bas/VLA_Interpreter.bas's
' own exceptions. Several ids share text with a sibling call site
' elsewhere in this file (the "extra words after the first statement"
' and "Program file not found" refusals each fire from 2-3 places).
' Several templates needed literal curly braces moved out of the stored
' template text and into a passed value (e.g. "example", "{name}") since
' this file's own domain is slot syntax, which uses {name}/{name:category}
' literally in several messages - VLA_Messages.bas's own header explains
' why that is safe (substitution is a single, non-recursive scan).
' Rendered text and Err.Number/Err.Source are unchanged (verified by
' manual trace at migration time for every site).
' RULECOVERAGE.0: BETA_ROADMAP.md's AS.1, VBA-native this time - see the
' owner's own "why isn't this shipped with Frazaro" challenge to the
' PowerShell-only version. EnglishRuleCoverageReport (new, public) does
' a fresh EnglishResetGrammar+EnglishLoadVocabulary, the same contract
' EnglishExpandedVocabularyText already uses, then walks mPatItems in
' registration order bucketing each rule by mRuleTestCounts (new state,
' populated in RunVocabTest via BumpUsage's own remove-and-readd
' counting idiom, keyed by mLastRuleIdx - the REAL rule TryPhrase
' matched, not a text-position guess). Strictly stronger than tools/
' check_rule_coverage.ps1's own positional heuristic: a shadowed rule
' whose "own" test-success actually fires an earlier, more general rule
' (F.4's own found instance) shows a real 0 here, where the PowerShell
' scanner - reading only where the text SITS - would still show 1.
' "Phrasebook Test Coverage" ribbon button (VLA_Build.bas) and
' EnglishIdeRuleCoverageReport (VLA_IDE.bas, Save-As export, same shape
' as Export Expanded Phrasebook) wire it to the UI. Not yet
' owner-verified.
' G6.0: list-valued slots - {name:cat-list} (cat one of text/range/
' cell/column/sheet/color) parses one or more comma-separated
' MatchRefToken items (the extracted, shared body of every singular
' typed slot's own token-consumption logic - TryPhrase used to inline
' this once per category; now every category and every "-list" sibling
' call the same function), lowering to a new runtime primitive,
' `(array item1 item2 ...)` (VLA.bas's EmitExpr, VLA_Interpreter.bas's
' EvalExpr) - deliberately NOT a LISTOPS primitive, since a pivot's
' field names are a real value the emitted/interpreted PROGRAM uses,
' never expand-time compiler data (LISTOPS-STDLIB's own "different
' family from the runtime accessors" line, the mirror image of it).
' Oxford comma REQUIRED, by construction rather than a rejection rule:
' "and" is only ever consumed as a silent no-op immediately after an
' already-consumed comma, never as a standalone separator - so "rows of
' A, B and columns of C" (pareto.txt section 10's own motivating
' sentence) stops the rows list cleanly at "B" with zero lookahead
' against the surrounding pattern, while a genuinely dropped Oxford
' comma on a real 3-item list produces the ordinary "I understood... -
' then expected X but found Y" refusal every other slot failure already
' gives, not a silent misparse. Unblocks G-PIVOT's own still-open
' pivot-rows/-columns/-filters and L-PIVOT-HELPERS (both named G6 as
' their dependency) and dropdown-list; also lets a future multi-column
' sort generalize past sort-two's own fixed two-slot shape. No caller
' wired into the real corpus yet - shipped as infrastructure, matching
' how LISTOPS-STDLIB's own functions shipped ahead of any real caller.
' Two item categories pinned (TestG6, VLA_Tests_Grammar.bas), not one,
' to prove the mechanism is generic: text-shaped and column-shaped
' (RefShapeOk's own IsColLetters, reused unchanged).
' GEXPANDERLINT.0: BETA_ROADMAP.md's AS.1 scoping pass - "Export Expanded
' Vocabulary" now returns house-style, pretty-printed text (VLA_Lint.
' VlaLintFormat over GEXPANDER.0's own flat blob) instead of one line per
' directive forever. Owner's own standing preference: every .vla this
' project writes is linted by default, not just the two hand-edited
' corpus files VlaLintCheck already gates. See EnglishExpandedVocabularyText's
' own comment for the reuse argument (SplitSegments already accepts
' GEXPANDER.0's flat text as valid input, so this is a re-flow of forms
' already built, not a second formatter). TestGExpander's three text-shape
' assertions (splice-flattening, row-tag placement) updated to match the
' new multi-line rendering, hand-derived against VLA_Lint.bas's own PpForm
' (english-vla/test-success are IsDirectiveShaped, so both ALWAYS
' verticalize as "(head arg1" / "  arg2)", regardless of width) - not yet
' live-verified, needs an owner run.
' ANTONYMSWEEP.0: BETA_ROADMAP.md's ANTONYM-SWEEP - the five hand-
' copied verb-antonym macro pairs (hide|unhide column, wrap|unwrap
' text, merge|unmerge range, protect|unprotect sheet, hide|unhide row)
' refactored through two real generators, both pure vocabulary
' additions (scripts/english.vla) - no VBA engine change at all, same
' as COLOR-FAMILY. `bool-antonym-family` (target-and-property shape,
' a boolean flip: hide-column, wrap-text, hide-row - 3 of 5) and
' `method-antonym-family` (target-and-method shape, the antonym is two
' method names: merge-range - 1 of 5), both deriving the negative
' macro's own name via QUASIQUOTE's `(symbol "un" name)` fusion, never
' needing LISTOPS at all (this is a generator CALLED ONCE PER PAIR,
' TABLE-FAMILY's own shape, not a table-walking recursive one -
' ANTONYM-SWEEP's roadmap entry names both QUASIQUOTE and LISTOPS as
' its gate, but only the former turned out load-bearing). Every
' generated macro is byte-for-byte what the hand-written pair used to
' emit; the english-vla rule and test-success proofs for each pair are
' UNTOUCHED - G11r's own {d:hide|unhide}-style glued-identifier
' dispatch already reaches whichever macro name results, hand-written
' or generated, so only the macro PAIR itself needed replacing.
' protect-sheet/unprotect-sheet (1 of 5) stays hand-written on purpose,
' table-to-range's own precedent - its keyword-argument-plus-extra-
' parameter shape genuinely does not fit either family. `TestAntonymSweep`
' (VLA_Tests_Grammar.bas) wired into VlaSelfTest right after
' TestListopsDepthSafety.
' ATROW.0: BETA_ROADMAP.md's LISTOPS-PROVENANCE - a rule/test/macro a
' generator produces from row N of a table must be traceable to row N,
' not just to the generator's own call site (today every spliced form
' inherits the call site's line only - fine for a five-line antonym
' pair, not fine for fifty rules from one call). Needed real design,
' not just a wider error string, so a new (at-row label form) directive
' joins "begin" as a second transparent wrapper DispatchVocabForm
' recognizes structurally: label rides through untouched (rowTag, a new
' parameter threaded everywhere startLine already was - begin passes it
' to every sibling unchanged, at-row is the only place it ever changes,
' ExpandVocabMacroCall/RegisterVocabMacro/RunVocabTest/RunVocabFailTest
' all take it too now), and ProvLoc appends "(label)" to every failure
' inside that form - test-success/test-fail proof failures, a macro's
' own arity mismatch, its own reserved-name refusal, an unrecognized
' directive - all of it, not just one message string. rowTag="" (every
' existing rule, untouched) reproduces the OLD wording byte-for-byte -
' ProvLoc's own guard is `If Len(rowTag) > 0`. test-success/test-fail
' are QUEUED at dispatch time and RUN later (EnglishLoadVocabularyText's
' own second pass) - rowTag is captured into a new parallel collection
' (testRowTags) the same moment testLines.Add startLine already is, not
' read live from a global at failure time, for the identical reason
' startLine itself was never a global. Not LISTOPS-only: usable today,
' by hand, in any large hand-written block - TestAtRow's own last pin
' proves the actual future shape (a generator's own expansion emitting
' several at-row-wrapped children, nested begin inside each) works
' today, without LISTOPS existing yet, because the mechanism only cares
' about splicing and dispatch, never about how the forms got generated.
' The load-time half above was the first cut; the owner's own follow-up
' asked for the emitted-CODE half too - a row's own provenance visible
' in the COMPILED VBA's trailing comment, not just in an error message
' at load time. RegisterVocabMacro now rewraps a row-tagged macro's own
' TEMPLATE half (signature/docstring untouched) in (gen-row "label"
' ...), a new form VLA.bas's EmitStmt recognizes structurally alongside
' at-line - so EVERY future call to that macro, in ANY program, emits
' with a trailing " vla-row:label" (separate from, never combined into,
' the existing " vla:N" line-tag). Deliberately carries NO ObjPtr/
' TagLine involvement at all: label is ordinary DATA inside the
' template, copied by Substitute like any other literal, so gen-row can
' never suffer the S3.1 heap-address staleness bug the line-tag system
' already had to fix once - see VLA.bas's own mGenRow declaration and
' MapTag's widened gate for the full reasoning. mGenRow joins mAtLine
' in VlaFrame's snapshot (twenty fields now, was nineteen) for the same
' reentrancy discipline every other per-compile module variable already
' gets. TestGenRow (VLA_Tests.bas, 2 assertions) pins the emitter
' mechanism directly; TestAtRow's own two new pins (now 9 total) prove
' the full round trip - an at-row-tagged macro's real compiled output
' carries the tag on every call, and an ordinary macro's does not.
' TABLEFAMILY.0: BETA_ROADMAP.md's TABLE-FAMILY, the first real generator
' macro METAVOCAB's own mechanism carries - not a test of the mechanism
' (TestMetaVocab already did that with synthetic macros), the mechanism
' put to real use in scripts/english.vla. table-property-family
' (english.vla) generates table-style/table-totals-on/table-totals-off's
' own (defmacro (name n ...) doc (set! (. (activesheet.listobjects n)
' property) value)) / (english-vla ...) / (test-success ...) triple.
' `name` and `property`, bare symbols, substitute as a WHOLE value into
' the nested defmacro's own head position and its set!'s member-access
' position respectively - exactly the "bare symbol passed as a whole
' argument into a nested (defmacro (nameparam n) ...)" TABLE-FAMILY's
' own roadmap entry named as already working. The three macros do NOT
' share one arity (table-style's VALUE is a second user-supplied slot;
' table-totals-on/-off's is a fixed true/false with no second slot at
' all), so the generator threads it rather than assuming a uniform
' shape: a trailing `& extras` rest param collects zero or more bare
' param names from the generator's own call and splices them, in
' order, after `n` in the generated macro's signature - zero items
' vanishes cleanly (no dangling "&"), never surfacing in any generated
' macro's own signature since none of these three instances actually
' need a rest param there. `value` is a separate, ordinary argument,
' bound to whatever the generator's call passes - the bare symbol that
' names the just-spliced extra param, or a fixed literal, either way
' whole-value substitution, never gluing. Every generated macro, call
' template, and test-success form is byte-for-byte what the hand-
' written macros used to emit - scripts/instructions_golden.vla's own
' table-style/table-totals-on/-off lines are untouched by this pass.
' METAVOCAB.0: vocabulary that writes vocabulary. The loader's dispatch
' (formerly one flat If/ElseIf inside EnglishLoadVocabularyText's own
' loop) is now DispatchVocabForm, a Sub that calls itself: its final
' Else used to raise "unrecognized top-level directive" immediately;
' now it first tries ExpandVocabMacroCall, which expands the form as a
' call to a macro this file has carried so far (mVocabMacros) plus the
' prelude, via VLA.VlaExpandText's existing toFixpoint expansion - no
' new macro engine, the one VLA.bas already runs for defmacro/deflambda
' everywhere else. If that expansion changes nothing, the ORIGINAL
' message fires unchanged (a genuine typo looks exactly as it always
' did). If it changes something and the result is (begin dir1 dir2
' ...), each child is spliced back through DispatchVocabForm itself,
' recursively, so a generator macro's output is indistinguishable from
' rules a human typed by hand - including running through the SAME
' test-success/test-fail proof machinery. Nested (begin ...) - one
' generator calling another - is capped at 20 levels, a splice-depth
' guard only; VlaExpandText's own fixpoint expansion (VLA.bas's own
' 200-deep macro-recursion guard) does the actual macro-expansion
' work, once, per top-level form. Two real limits, stated plainly, not
' hidden: a generator can only build the TEMPLATE half of a rule from
' whole pre-written forms passed as arguments - it cannot glue a bound
' word into part of a new identifier the way make-{d} does, because
' defmacro's own Substitute has no such operation (only the English
' matcher's FormSubstitute does, and only for {slot} bindings); and
' generated rules carry no test-success proofs of their own unless the
' generator's own body emits them alongside its rules, same as any
' hand-written rule would need.
' APROPOSPLUS.0: apropos gains the English-sentence tier. AproposRulesBlob
' builds one "pattern -> template" line per loaded rule from mPatTexts/
' mPatForms (the same pair EnglishListRefSlots already reads), and both
' load-time call sites of VlaAproposCarry now pass it as a second
' argument (VLA.bas's own new optional parameter) alongside the macro
' text already carried - so a search for a macro name now also
' surfaces the English sentence that reaches it, and a search for a
' word that lives only in a sentence's own wording now finds it too,
' where before only macro names and docstrings were searchable. See
' VLA.bas's own APROPOSPLUS.0 note for the carry-mechanism half of
' this (VlaFrame's new field, threaded through VlaPushContext/
' VlaPopContext the same way AproposCarry already was).
' VOCABDIFF.0: structural vocabulary diff. EnglishDiffVocabularyText/
' EnglishDiffVocabulary compare two vocabulary sources by walking both
' through VLA.VlaReadFormsWithLines (the same reader every other vocab
' tool uses) and keying each top-level directive by its own stable
' identity - a rule's/override's pattern text, a test's sentence, a
' macro's name, a function word's phrase - rather than by line
' position, so a reflow or a reordering that touches no rule's content
' reports as no change (VLA.VlaWriteForm decides CHANGED vs UNCHANGED
' for a matched key). No new equality walker: the reader and the
' writer are the only machinery, both already load-bearing elsewhere.
' Deliberately does not attempt cross-language rule matching (a
' dialect file diffed against its own base reports as a full add/
' remove, correctly - recognizing "the same rule, different words"
' would need matching template forms up to slot-renaming, out of
' scope here); the intended use is two revisions of one vocabulary,
' or a staged candidate against a committed one. Read-only and
' state-independent - no EnglishResetGrammar, no rule registration,
' safe to run without touching whatever grammar is currently loaded.
' GROWLOOP.0: the row-loop family. Checked against the live grammar
' before writing anything: "Count <name> [down] from <a> to <b>:"
' already covered ascending and the safe bottom-up descending shape -
' a first draft invented a parallel "For each row from ... to ...:"
' surface before that check caught the duplication, reverted whole
' (see VLA_Tests_Grammar.bas's GROWLOOP.0 note for the fuller account).
' The one shape actually missing was a custom step for skip-N row
' loops ("every other row"), so Case "count" gained an optional "step
' <s>" clause: ascending steps by {s} as given; "down" negates it in
' the emitted form, (- 0 {s}), rather than asking the user to type a
' negative number - this grammar has no unary minus in words. Both
' EmitFor (VLA.bas) and ExecFor (VLA_Interpreter.bas) already evaluate
' the step slot as an arbitrary expression, not just a literal, so
' this needed no emitter or interpreter change at all - F.1's grammar/
' emitter ABI paying for itself exactly as designed.
' G8.0: number words and ordinals. NumberWord (the tokenizer's existing
' zero..twelve table) gains thirteen..twenty - safe unconditionally,
' checked against the whole vocab first, none of those thirteen words
' claimed anywhere else. Ordinals (first..twentieth) are deliberately
' NOT at the tokenizer: "first" is already reserved by mFnOf's "first
' of X" list-accessor idiom (first of found-items -> vlafirst) and by a
' shipped pattern-literal rule ("first {n} letters of"); a blanket
' rewrite the way cardinals get one would silently break both every
' time "first" appeared anywhere. New OrdinalWord is instead checked
' only inside ParsePrimCore's expr grammar, as a new ElseIf positioned
' right after the existing "<word> of <value>" check - reached only
' once that check has already failed, so "first of X" keeps first
' claim whenever "of" actually follows, and "first {n} letters of"
' (a pattern literal, never routed through this expr grammar at all)
' is unaffected regardless. Shipped example: a second delete-row rule,
' `delete {n:expr} row`, alongside the existing `delete row {n:expr}` -
' genuinely different word orders (quantity-after-noun vs. ordinal-
' before-noun), so two rules sharing one macro, not an alternation on
' one. "the" is a dropped noise word (IsDroppedWord) either way, so it
' has no place in either pattern.
' G7.0: slot defaults - {name:category=text}. IsSlotTok splits an
' optional '=text' off the category (last '=' wins) and hands it back
' through two new Optional ByRef params, hasDefault/defaultVal, that
' existing callers who don't ask for them never see - cat comes back
' stripped either way, so ValidateRuleItems/LintRule/EnglishListRefSlots
' etc. are unchanged. TryPhrase's slot Select Case now sets a matchOk
' flag instead of NoteFail-and-Exit-Function on every failure branch;
' after the Select Case, a default softens matchOk=False into val =
' defaultVal (pure substitution, never re-parsed) with p left wherever
' the failed parse left it - true for every parser here (ParseExpr/
' ParseCond/RefShapeOk all leave pos untouched on their own failure
' path, checked one call down to ParsePrimCore). ExpandedSignatures
' grows a "" (absent) branch for a defaulted slot alongside its normal
' piece, same present/absent duality IsOptTok already gets - so a
' defaulted slot audits like an optional, not like a second rule.
' BuildDispatchIndex needed no change: any typed slot, defaulted or
' not, already marks a rule universal (any token can start it), which
' is the same conservative bucket a default would otherwise earn.
' LX3.0: every identity-deciding LCase$ site now folds through
' VLA_Identity.Fold (invariant ASCII, never locale-dependent) - R6/SD-8.
' GEXPANDER.0: BETA_ROADMAP.md's G-EXPANDER. Every EnglishLoadVocabulary
' file load now writes a fully macro-expanded sibling next to the
' source (english_expanded.vla for english.vla) - every real directive
' DispatchVocabForm actually dispatches (english-vla/override/test-
' success/test-fail/defmacro/english-function/keyword-alias [LX5.1]),
' in the exact order it reaches them, rendered back to literal text via
' VLA.VlaWriteForm -
' AproposRulesBlob's own precedent (walk a loaded-forms collection,
' write each, join), applied to the whole directive stream instead of
' just phrase-rule templates. DispatchVocabForm already flattens both
' structural wrappers before a leaf is ever recorded: a generator's
' (begin ...) splice and an (at-row label ...) wrap are never
' themselves directives, so RecordExpandedForm only fires from the
' branches that ARE - meaning the artifact shows the POST-expansion
' grammar, never a generator's own one-line call. rowTag (the at-row
' label, threaded the same way ATROW.0 already threads it everywhere
' else) prints as a "; row: <label>" comment directly above the form it
' labels when non-empty, same LISTOPS-PROVENANCE half ATROW.0 shipped
' for load-time failures, now visible in the artifact itself. expTexts/
' expTags are new parallel Collections threaded through
' DispatchVocabForm/ExpandVocabMacroCall exactly the way testSents/
' testExps/.../testRowTags already are - no new module state, so nested
' or repeated loads can never see another load's leftovers. The write
' itself (WriteExpandedSibling) happens in EnglishLoadVocabulary alone,
' strictly after EnglishLoadVocabularyText returns - never reachable
' from inside ExpandMacros/VlaExpandText, which stay pure by LISTOPS-
' PURITY's own wall - and is best-effort: a write failure (read-only
' install dir, a locked file) never fails the vocabulary load itself,
' VlaRefreshBetaCopy's own "never fails the build" precedent for an
' auxiliary side artifact. The sibling's own header stamps the source
' file's mtime+size at read time, so a future static consumer (AS.1)
' can compare that stamp against the CURRENT source file on disk and
' refuse loudly on a mismatch instead of silently reading a stale
' expansion - the same fail-loud-over-silent-misreport instinct AS.6/
' F.12 already hold elsewhere in this file.
' LX5.1: BETA_ROADMAP2.md's LX.5, phase 1 - the structural-keyword
' canonicalization seam. ParseStmt/ParseCond/ParseExpr/ParsePrim/the
' Oxford-comma list logic (887, 967, 4306, 4510-4540, 4916-4949) compare
' tokens against dozens of literal English words (if/repeat/until/times/
' while/for/each/in/stop/loop/increase/decrease/add/get/give/try/that/
' fails/when/clicked/create/define/done/otherwise/and/or/is/equals/
' greater/less/at/least/most/does/not/plus/minus/divided/joined/with/
' percent/of), not just the four headline ones LX.1 named - none of it
' goes through VLA_HeadTable's alias map, which only ever resolves
' EMITTED s-expression heads inside VLA.bas, never English sentence
' tokens. Rather than touch every comparison site, one new
' CanonicalizeStructuralWords pass runs right after every EnTokenize
' call that feeds ParseStmt (EnglishToVla's own translate path, and the
' five internal EnTokenize+ParseStmt pairs test-success/test-fail
' proofs, G-RENDER's self-check, and EnglishExplain each already ran):
' it rewrites any bare-word token a phrasebook has aliased (via a new
' `(keyword-alias "si" "if")` directive,
' DispatchVocabForm) to its canonical English spelling, so every
' downstream comparison keeps comparing against English unchanged -
' language-neutral by construction, not by rewrite. Same alias-map
' shape VlaHeadTableAliasMap already proved, staged one step earlier;
' same IsWordTok-only protection SkipArticles/IsNoiseWord already rely
' on, so quoted text and numbers are never touched. English's own load
' path registers zero aliases, so this is a no-op for the shipped
' product until a language file asks for it - verified by an unchanged
' VlaSelfTests()/VlaGoldens() diff. The Oxford-comma list grammar's own
' POSITIONAL consumption rule stays universal, fixed core syntax, per
' LX.1's own "keep, structural" ruling - only its trigger words go
' through the same table as everything else.

' =====================================================================
'  VLA_English - a DCG-style controlled-English front-end for VLA.
'
'  Pipeline:   English sentences -> VLA forms -> VBA
'              (every layer is plain text you can inspect)
'
'  The grammar is a set of phrase rules, each a DCG production:
'
'      sentence -> "set" var "to" expr        =>  (set! {v} {e})
'
'  written here as a pattern string with typed slots:
'
'      EnglishAddPhrase "set {v:var} to {e:expr}", "(set! {v} {e})"
'
'  Slot categories (the grammar's nonterminals):
'      {x:name}   one word, used as an identifier
'      {x:var}    like name, but marks the word as an assigned
'                variable (feeds auto-declaration)
'      {x:expr}   an English expression: numbers, "strings", names,
'                plus/minus/times/divided by/joined with
'      {x:cond}   a comparison: is / equals / is greater than /
'                is at least / is divisible by / ... with and, or
'
'  Sentences end with "." â€” one sentence, one statement. Control flow
'  is structural (built in), ending blocks with "Done.":
'      To <name>: ... Done.              define an action (a Sub)
'      To <name> of <param>: ... Give back <value>.   (a Function -
'          the name then composes everywhere: "tax of subtotal")
'      To [get] <name> using <p> of <default> and <q>: ...   (named
'          parameters; call: "tax using income of 5 and rate of 2",
'          omitted parameters take their defaults)
'      If <cond>: ... Otherwise, if <cond>: ... Otherwise: ... Done.
'      If <cond>, <sentence>.            one-line if
'      When <value> is <case>: ... When it is <case>: ... Otherwise: ...
'      Try: ... If that fails: ...       (recovery paragraph optional)
'      Repeat <n> times: ... Done.       (loop counter: "counter")
'      Repeat until <cond>: ... Done.
'      While <cond>: ... Done.
'      For each <x> in <coll>: ... Done.
'      Create a number/text/value/list/lookup called <x>.
'
'  Friendliness rules: "the/a/an/please" are noise words; number words
'  zero..twelve work; assigned variables are auto-declared; loose
'  sentences outside any "To ..." become (sub main () ...).
'
'  Public API:
'      EnglishToVla(text)                 -> VLA source
'      EnglishToVba(text)                 -> VBA source
'      EnglishCompileToModule(text, name) -> inject into a module
'      EnglishAddPhrase(pattern, template)-> extend the grammar
' =====================================================================

Private mPatItems As Collection    ' each item: Collection of pattern tokens
Private mPatForms As Collection    ' F.2: parallel - template pre-parsed
                                    ' into forms (Collection of top-level
                                    ' forms). AddPhraseRule refuses to
                                    ' register a rule whose template
                                    ' doesn't parse as at least one form -
                                    ' forms are the only path a rule's
                                    ' template ever takes (the dual-path
                                    ' era, and its Replace()-text fallback
                                    ' mPatTmpls, retired once F.2's own
                                    ' closure bar was met - see
                                    ' BETA_ROADMAP.md's F.2 entry).
Private mLastRenderUsedFallback As Boolean  ' G-RENDER: set by
                                    ' RenderExprForm whenever an
                                    ' :expr/:cond binding fell outside
                                    ' the known operator table and was
                                    ' spliced in as raw VLA text instead
                                    ' - EnglishRenderSelfCheck reads this
                                    ' right after each render call to
                                    ' know which results are round-trip
                                    ' testable and which aren't.
Private mInited As Boolean

Private mPreludeCount As Long      ' rules 1..mPreludeCount are built-in
Private mPatTexts As Collection    ' original pattern strings (for messages)
Private mPatSigs As Collection     ' rule shape signatures (for lint checks)
Private mSigOwner As Collection    ' G3 perf: one exact signature string -> the
                                    ' earliest rule index that claims it (the
                                    ' same "first match always wins" semantics
                                    ' real dispatch enforces) - turns the
                                    ' override-target lookup and the duplicate-
                                    ' shape check from O(corpus size) per
                                    ' registration into O(this rule's own
                                    ' signature count). See AddSigOwner/
                                    ' RebuildSigOwner for how it stays correct
                                    ' across overrides and EnglishTryRule's
                                    ' snapshot/restore.
Private mLintWarnings As Collection ' accumulated lint warnings

' Step mapping: each translated statement gets a number and its
' sentence text, so runtime errors can point at the user's sentence.
Private mStepCount As Long
Private mStepTexts As Collection
Private mStepTracking As Boolean   ' set True in EnsureInit (default on)

' Action signatures and recorded calls, for translation-time call
' checking (converts would-be VBA compile errors into clean messages).
Private mActNames As Collection    ' lcase action names
Private mActParams As Collection   ' parallel: Collection of lcase param names
Private mActReq As Collection      ' parallel: Collection of "1"(required)/"0"
Private mCallNames As Collection   ' lcase action name per recorded call
Private mCallArgs As Collection    ' parallel: Collection of lcase arg names
Private mCallTexts As Collection   ' parallel: the call sentence, for messages
Private mCallLines As Collection   ' parallel: the call sentence's line

' Aliases: "Define hot-pink as "#FF69B4"." at top level becomes a
' module-level Const, visible in every sub and unassignable.
Private mAliasNames As Collection  ' keyed lcase alias names
Private mAliasDefs As String       ' the (const ...) lines to prepend

' Worked examples: the first test sentence that matches each rule,
' keyed "r<index>". Populated as vocabulary tests run, shown beside
' patterns in EnglishListPhrases.
Private mRuleExamples As Collection
' AS.1: the REAL per-rule test-success count, keyed "r<index>" like
' mRuleExamples - but a genuine count, not just one worked example.
' Populated in RunVocabTest from mLastRuleIdx, the same real matcher
' TryPhrase already runs - this is a firing count, not a text-position
' heuristic (that's tools/check_rule_coverage.ps1's own, weaker, no-
' live-host substitute). Never suspended by mUsageSuspended - unlike
' BumpUsage/mUsageCounts (real-PROGRAM usage, S2's own reason those ARE
' suspended during test loading), a proof firing IS exactly what this
' counts.
Private mRuleTestCounts As Collection
Private mLastRuleIdx As Long       ' rule matched by the last TryPhrase
Private mPatSources As Collection  ' G3: per-rule provenance, lockstep with mPatTexts
Private mLoadSource As String      ' G3: the file being loaded ("" = API/direct)
Private mAuditMode As Boolean      ' G3: audits report what a load would refuse
Private mVocabMacros As String     ' L4: defmacro text carried by phrasebooks, load order
Private mIncludeLines As String    ' G12: top-matter (include ...) lines
Private mIncludeSeen As Collection ' G12: dedup keys, case-insensitive
Private mVocabMacroNames As Collection ' L4: keyed macro name -> carrying file
Private mTestsRun As Long          ' U.10: pass-proofs run since reset
Private mFailsRun As Long          ' U.10: fail-proofs run since reset
Private mLastLoadSource As String  ' U.10: the last load's source name
Private mLastExpandedBlob As String ' GEXPANDER.0: the last load's own expanded-form text
Private mUsageCounts As Collection ' S2: firing counts, keyed by display
                                   ' key ("rule: <pattern>" for phrase
                                   ' rules, "form: <construct>" for the
                                   ' structural form that claimed a
                                   ' top-level sentence). Keyed by TEXT,
                                   ' so counts survive grammar resets
                                   ' and vocabulary reloads.
Private mUsageKeys As Collection   ' first-seen order (a Collection
                                   ' hides its keys from enumeration)
Private mUsageSuspended As Boolean ' S2: True while vocab test: proofs
                                   ' run - a proof firing is evidence
                                   ' of loading, not of usage
Private mClaim As String           ' A4: the construct that claimed the
                                   ' current top-level sentence, for the
                                   ' Explain trace. First claim wins within
                                   ' a sentence (the OUTERMOST form is the
                                   ' headline); reset per top-level sentence
                                   ' by the translation loop and by Explain.
Private mLoopStack As Collection   ' enclosing loop kinds ("for"/"do"),
                                   ' so "Stop the loop." emits the right Exit

' B4: value-returning actions. mInFuncDef is true while a
' "To <name> of <param>:" body parses (gates "Give back" and flips
' Stop./failure exits to Exit Function). mUserFnWords remembers the
' function words a translation registered, so the NEXT translation
' can unregister them first - a program-defined word must not leak
' into programs that never defined it. mFnActNames feeds the
' duplicate-definition guard.
Private mInFuncDef As Boolean
Private mInRecovery As Boolean     ' B7: gates 'the problem' to Try recovery
Private mSawSheetChangeEvent As Boolean   ' IN.7: one 'When the sheet
                                    ' changes:' per program - see its
                                    ' dispatch site for why a second
                                    ' one is refused rather than
                                    ' silently overwriting the first
' IN.7 (button-click half): mClickCaptions guards one handler per
' button NAME (unlike sheet-change's single program-wide flag, a
' program may declare several - "one name, one definition" per
' caption, CheckDupAction's own discipline). The internal proc name
' each 'When "<caption>" is clicked:' compiles to is CAPTION-derived
' ("on:click:" & ClickHandlerSlug(caption), not a plain counter as
' first shipped) - compiled parity's own EmitStmt "make-button" case
' (VLA.bas) needs to compute the SAME target name independently, from
' nothing but the caption sitting in its own make-button call, and a
' counter has no way to be recomputed from outside the parse that
' handed it out. Two captions whose slug collides (punctuation aside,
' e.g. "Go!" and "Go?") would otherwise produce the same raw proc
' name and ride into VBA's own "duplicate declaration" compile error
' the moment Compile actually ran - mClickSlugs (below) catches that
' HERE instead, with words and a line number, the moment the second
' colliding caption is parsed, rather than leaving it to the
' pre-existing "VBA's own compile step is the loud judge" gap
' RecordSubDoc's own doc comment accepts for a LITERAL duplicate name
' (hand-authored VLA repeating a sub name outright is still that
' gap - this only closes the NEW way English itself can introduce one).
' mClickNames/mClickProcs are the parallel pair VLA_IDE.bas's
' InterpretProgram reads right after translating, to register each
' declared handler with
' VLA_Events - VLA_Events.bas's own mWb/mSrc idiom, at the English
' layer instead of the registry's.
Private mClickCaptions As Collection
Private mClickSlugs As Collection     ' Fold(slug) -> the FIRST caption that claimed it
Private mClickNames As Collection
Private mClickProcs As Collection
Private mUserFnWords As Collection
Private mFnActNames As Collection

' B6: using-style value actions. Keyed by lowercase name: the ordered
' parameter-name list and matching "1"/"0" required flags. Purely
' per-translation state (reset at each EnglishToVla), so program-
' defined names cannot leak between programs.
Private mFnUsingNames As Collection
Private mFnUsingReq As Collection

' V3: lookup names declared by "Create a lookup called ..." in the
' CURRENT translation. Purely per-translation state, reset at each
' EnglishToVla like the action registries, so a program's lookups
' never leak into programs that never declared them. This registry
' is what SCOPES the keyed-read syntax: "<name> for <key>" is a
' value only when <name> is a declared lookup - everywhere else,
' "for" stays inert in expressions and no existing rule's greediness
' contract changes.
Private mDictNames As Collection

' B3: each Try block numbers its generated labels (vla-tryf-N /
' vla-tryr-N / vla-tryd-N) so several Trys - nested or sequential -
' can share a procedure. Reset per translation.
Private mTryCount As Long

' Function words: expression-level vocabulary. Unary words read as
' "<word> of <value>" (length of code); nullary words are values
' themselves (today). Extensible from vocab files via function: lines.
Private mFnOf As Collection        ' unary: word -> VBA function name
Private mFnNullary As Collection   ' nullary: word -> VBA function name
Private mFnDisplay As Collection   ' display strings for the phrase list

' A1: tokens carry their source line. mTokLines parallels the token
' Collection of the most recent EnTokenize; mErrLine is the line of
' the most recent translation error (0 = unknown), exposed via
' EnglishLastErrorLine for exact IDE row attribution.
Private mTokLines() As Long        ' P-TOK: was a Collection - see
                                   ' VLA.bas's own mTokLines for the
                                   ' identical materialize-once shape
Private mErrLine As Long

' Paragraph marker emitted by the tokenizer for blank lines: closes
' all open blocks, means nothing at top level. "|" can never be
' produced any other way (it is not a word character).
Private Const PARA_TOK As String = "|"

' Best partial match across rule attempts for the current sentence,
' used to build helpful "I understood X, expected Y" errors.
Private mBestProgress As Long      ' furthest token position any rule reached
Private mBestExpect As String      ' human description of what it wanted there
Private mBestTieIdx As Collection  ' LE.2: every rule idx tied at mBestProgress (leader included), up to 3

' Per-sub bookkeeping for auto-declaration.
Private mDeclared As Collection    ' names declared via Create ...
Private mAssigned As Collection    ' names bound by :var slots / loops

' V1: the first line each name was assigned on, keyed lcase name,
' translation-global (first assignment wins, which AddKeyed-style
' duplicate-key silence gives for free). Exists so the ONE Check
' refusal raised AFTER parsing - BuildSub's assigned-a-Defined-alias
' - can still name its sentence's line now that the IDE attributes
' from a single translation instead of re-translating prefixes.
' mCurLine is the source line of the statement currently being
' parsed, stamped by ParseTracked.
Private mAssignedLines As Collection
Private mCurLine As Long
Private mSrcLines As Variant       ' S3.2: the translation's original
                                   ' source, split per line - the step
                                   ' table quotes the USER's sentence,
                                   ' not the parser's rendition

' V7: first-token rule dispatch. Rules are bucketed by every token
' that can BEGIN a match - a first literal; every surface of every
' branch of a first bare or braced alternation (expanded by the same
' BareSurfaces/SurfaceForms every signature uses - never a parallel
' expansion); the surfaces of any leading optionals PLUS whatever
' follows them (an optional is free when absent, so the next item
' also begins the rule). A rule whose first non-optional item is a
' value slot (:var, :expr, :cell ...) can begin at ANY token, so it
' lives in the universal list and is always tried. Dispatch merges
' the sentence's bucket with the universal list in ascending rule
' index, so first-match order is preserved by construction: every
' rule NOT in the merge provably cannot match (its entire first-token
' set excludes the sentence's key), and its only observable side
' effect - the near-miss NoteFail at the article-skipped position,
' with its first item's description - is simulated exactly, in rule
' order, so error wording cannot drift (the G4 pins stay honest).
' The index is invalidated by every rule-store mutation (add,
' override, reset, G5 restore) and rebuilt lazily at the next
' dispatch - which during a vocabulary load is the next test: proof,
' i.e. rebuilt at load, as the roadmap specifies.
Private mDspValid As Boolean       ' False whenever the rule store mutates
Private mDspBuckets As Collection  ' keyed "k:"&token -> Collection of rule
                                   ' indices (ascending, built 1..N)
Private mDspUniversal As Collection ' always-try rule indices, ascending
Private mDspFailDesc() As String   ' per bucketed rule: its first item's
                                   ' NoteFail description ("" = universal)
Private mDspN As Long

' V8: the helper manifest, cached per session (the checklist item).
' The manifest derives from VLA_Runtime's source, which cannot change
' under a running session - a VBE edit wipes module state (clearing
' this cache with it) and a rebuilt add-in is a new session - so one
' successful read serves every resolve check. Cached only when
' non-empty: a failed read (no VBE trust, missing sheet) retries per
' call, exactly the old behavior for the soft-failure mode.
Private mManifestCache As String

' V8: the ADODB stream, created once and reused (the checklist
' item) - CreateObject is the cost, Open/Close per read is not.
' Any failure drops the cache so a wedged stream cannot persist.
Private mVocabStream As Object              ' rule count the index was built over

' PPROF.0: the translate-side half of P-PROF's per-phase timing,
' checking the SAME switch VlaTranspile's own compile-side phases do
' (VLA.mProfileOn, VLA.bas) rather than a second one - the roadmap's own
' "behind one switch" - so both files light up together. Two buckets,
' not four: EnglishToVla has no clean expand/emit seam of its own (it
' builds VLA source text while it parses, fused, not staged), so
' everything past tokenize is one "trans-build" bucket. Added because
' P-TOK's own named hotspot (TokAt's positional Collection indexing,
' this file) lives on THIS side, not inside VlaTranspile, and the
' compile-side phase list (ALPHA6_ROADMAP's own tokenize/parse/expand/
' emit) never covered it.
Public mProfMsTransTokenize As Double
Public mProfMsTransBuild As Double
Public mProfTransCalls As Long

' PPROF.0: Rule-12 duplicate of VLA.bas's own Private ProfMs (itself a
' duplicate of VLA_DevRig.bas's DevMs) - same midnight-wrap handling,
' not visible across modules.
Private Function ProfMs(ByVal t0 As Double) As Double
    Dim d As Double
    d = Timer - t0
    If d < 0 Then d = d + 86400#
    ProfMs = d * 1000#
End Function

' =====================================================================
'  Public API
' =====================================================================

Public Function EnglishToVla(ByVal text As String) As String
    EnsureInit
    mSrcLines = Split(Replace(text, vbCrLf, vbLf), vbLf)   ' S3.2
    ' PPROF.0: trans-tokenize, then trans-build (everything below, to
    ' this function's own return line) - see the header comment above
    ' mProfMsTransTokenize for why this stays a two-way split. Honest
    ' limit, not fixed: unlike VlaTranspile's own emitfail (VLA.bas),
    ' this function has no shared error handler to credit a mid-loop
    ' raise, so a call that fails partway through the Do While loop
    ' below loses its own trans-build time rather than crediting a
    ' partial one - acceptable for a corpus expected to translate
    ' cleanly (the self-test suite's real vocabulary/instructions.txt),
    ' not the deliberately-failing shape TestListopsBudget's 5500-deep
    ' case exercises on the compile side.
    Dim ptProf As Double
    If VLA.mProfileOn Then ptProf = Timer
    Dim toks() As String
    toks = CanonicalizeStructuralWords(EnTokenize(text))   ' LX5.1
    If VLA.mProfileOn Then
        mProfMsTransTokenize = mProfMsTransTokenize + ProfMs(ptProf)
        mProfTransCalls = mProfTransCalls + 1
        ptProf = Timer
    End If

    mStepCount = 0
    mErrLine = 0
    Set mStepTexts = New Collection
    Set mActNames = New Collection
    Set mActParams = New Collection
    Set mActReq = New Collection
    Set mCallNames = New Collection
    Set mCallArgs = New Collection
    Set mCallTexts = New Collection
    Set mCallLines = New Collection
    Set mAliasNames = New Collection
    mAliasDefs = ""
    Set mLoopStack = New Collection
    mTryCount = 0
    mInFuncDef = False
    mInRecovery = False
    mSawSheetChangeEvent = False
    Set mClickCaptions = New Collection
    Set mClickSlugs = New Collection
    Set mClickNames = New Collection
    Set mClickProcs = New Collection
    ' Unregister the previous translation's user-defined function
    ' words: they belong to that program, not to the grammar.
    If Not mUserFnWords Is Nothing Then
        Dim ufw As Variant
        For Each ufw In mUserFnWords
            On Error Resume Next
            mFnOf.Remove CStr(ufw)
            mFnNullary.Remove CStr(ufw)
            On Error GoTo 0
        Next
    End If
    Set mUserFnWords = New Collection
    Set mFnActNames = New Collection
    Set mFnUsingNames = New Collection
    Set mFnUsingReq = New Collection
    Set mAssignedLines = New Collection
    mCurLine = 0
    Set mDictNames = New Collection
    mIncludeLines = ""               ' G12
    Set mIncludeSeen = New Collection

    Dim pos As Long
    pos = 1
    Dim outParts As String
    Dim mainStmts As New Collection
    Dim mainDecl As New Collection
    Dim mainAsgn As New Collection
    Set mDeclared = mainDecl
    Set mAssigned = mainAsgn

    Dim subName As String
    Dim body As Collection
    Dim paramSpec As String
    Dim pOpt As Boolean
    Dim pName As String
    Dim pDef As String
    Dim sigParams As Collection
    Dim sigReq As Collection

    Do While pos <= UBound(toks)
        mClaim = ""                      ' claim tracking is per top-level sentence
        If TokAt(toks, pos) = PARA_TOK Then
            pos = pos + 1                ' blank lines at top level are decoration
        ElseIf PeekWord(toks, pos) = "define" Then
            ParseDefine toks, pos
        ElseIf (PeekWord(toks, pos) = "use" Or PeekWord(toks, pos) = "import") And _
               (WordAt(toks, pos + 1) = "library" Or WordAt(toks, pos + 1) = "code") Then
            ' G12: gated on the second word, so sentences merely
            ' starting with use/import keep falling through to the
            ' machinery that owns them (the To-get gate's pattern).
            ParseUseLibrary toks, pos
        ElseIf AtSheetChangeEvent(toks, pos) Then
            ' IN.7: "When the sheet changes:" - a top-level-only
            ' declaration, registered like a "To ...:" action rather
            ' than folded into mainStmts (a handler is not a step in
            ' the main script; it runs later, on a real Excel event,
            ' maybe never during this run at all). "the" is dropped by
            ' the tokenizer (IsDroppedWord), so this is the literal
            ' 3-word shape when/sheet/changes plus the block colon -
            ' AtSheetChangeEvent gates on that exact 4-token lookahead
            ' so it can never shadow "When <value> is <case>:", the
            ' choices form ParseStmt's own Case "when" already owns.
            pos = pos + 4                    ' when sheet changes :
            If mSawSheetChangeEvent Then
                VLA_Messages.RaiseMsg "english-sheet-change-dup", "loc", LineTag(pos - 1)
            End If
            mSawSheetChangeEvent = True
            Set mDeclared = New Collection
            Set mAssigned = New Collection
            Set body = ParseBlock(toks, pos)
            outParts = outParts & BuildSub("on:sheet-change", body) & vbCrLf
            Set mDeclared = mainDecl
            Set mAssigned = mainAsgn
        ElseIf AtButtonClickEvent(toks, pos) Then
            ' IN.7 (button-click half): 'When "<caption>" is clicked:'
            ' - a top-level-only declaration, the same shape as
            ' AtSheetChangeEvent's own 'When the sheet changes:' just
            ' above (a handler, not a step in the main script; it
            ' runs later, on a real click, maybe never during this
            ' run). Multiple named buttons per program means this
            ' gate cannot be a fixed "one program, one handler" flag
            ' like mSawSheetChangeEvent - each caption gets its own
            ' registration, refused only on an exact repeat (CheckName/
            ' CheckDupAction's own "one name, one definition" reflex).
            ' AtButtonClickEvent's own exact 5-token lookahead keeps
            ' this from ever shadowing ParseStmt's 'When <value> is
            ' <case>:' choices form - see its comment for the one
            ' contrived case (a case literally named "clicked") that
            ' shape sharing still allows, accepted on the same terms
            ' IN.7's sheet-change gate already accepted for "sheet"/
            ' "changes".
            Dim btnCaption As String
            btnCaption = Mid$(TokAt(toks, pos + 1), 2)   ' strip the token's leading quote mark
            pos = pos + 5                     ' when "caption" is clicked :
            If CollHasKey(mClickCaptions, btnCaption) Then
                VLA_Messages.RaiseMsg "english-click-handler-dup", "caption", btnCaption, "loc", LineTag(pos - 1)
            End If
            AddKeyed mClickCaptions, btnCaption
            Dim clickSlug As String
            clickSlug = ClickHandlerSlug(btnCaption)
            Dim slugFound As Boolean, priorCaption As Variant
            priorCaption = CollGet(mClickSlugs, clickSlug, slugFound)
            If slugFound Then
                VLA_Messages.RaiseMsg "english-click-handler-slug-collision", "a", btnCaption, "b", priorCaption, "loc", LineTag(pos - 1)
            End If
            On Error Resume Next
            mClickSlugs.Add btnCaption, VLA_Identity.Fold(clickSlug)
            On Error GoTo 0
            Dim clickProcName As String
            clickProcName = "on:click:" & clickSlug
            Set mDeclared = New Collection
            Set mAssigned = New Collection
            Set body = ParseBlock(toks, pos)
            outParts = outParts & BuildSub(clickProcName, body) & vbCrLf
            Set mDeclared = mainDecl
            Set mAssigned = mainAsgn
            mClickNames.Add btnCaption
            mClickProcs.Add clickProcName
        ElseIf PeekWord(toks, pos) = "to" Then
            ' "To <name>[, with [optional] p [of default] [and ...]]: ..."
            ' or, B4: "To <name> of <param>: ... Give back <value>."
            ' or, B6: "To [get] <name> using <p> [of <default>] ...: ..."
            pos = pos + 1
            ' "get" is header sugar marking a value-returning action.
            ' Recognized only when a name follows and then "using",
            ' "of", or ":". (An action named get was never possible -
            ' Get is VBA-reserved and CheckName refuses it, as a
            ' self-test pin discovered when it asserted otherwise;
            ' the gate's real job is letting sentences that merely
            ' START with get - vocabulary rules, the standalone-Get
            ' refusal - fall through to the machinery that owns
            ' them.) "To get answer:" is B7.2's NULLARY value action:
            ' no parameters, the name becomes a value word usable
            ' anywhere ("If answer is 42, ...").
            Dim getNullary As Boolean
            getNullary = False
            If PeekWord(toks, pos) = "get" And Len(WordAt(toks, pos + 1)) > 0 Then
                If TokAt(toks, pos + 2) = "using" Or TokAt(toks, pos + 2) = "of" Then
                    pos = pos + 1
                ElseIf TokAt(toks, pos + 2) = ":" Then
                    pos = pos + 1
                    getNullary = True
                End If
            End If
            subName = ExpectWord(toks, pos, "an action name after 'To'")
            CheckName subName
            CheckDupAction subName, pos - 1
            If getNullary Then
                ExpectTok toks, pos, ":", "':' after 'To get " & subName & "'"
                Set mDeclared = New Collection
                Set mAssigned = New Collection
                AddFnEntry mFnNullary, subName, subName
                mUserFnWords.Add VLA_Identity.Fold(subName)
                mFnActNames.Add VLA_Identity.Fold(subName)
                mInFuncDef = True
                Set body = ParseBlock(toks, pos)
                mInFuncDef = False
                outParts = outParts & BuildSub(subName, body, "", True) & vbCrLf
                Set mDeclared = mainDecl
                Set mAssigned = mainAsgn
            ElseIf TokAt(toks, pos) = "using" Then
                ' B6: named parameters, "of" giving a default exactly
                ' as in sub-action "with" clauses. Registered before
                ' the body parses (recursion composes). VBA requires
                ' optional parameters last, so a required parameter
                ' after a defaulted one is refused here with words
                ' instead of surfacing as a compile error.
                pos = pos + 1
                Set mDeclared = New Collection
                Set mAssigned = New Collection
                paramSpec = ""
                Set sigParams = New Collection
                Set sigReq = New Collection
                Dim sawDefault As Boolean
                sawDefault = False
                Do
                    pName = ExpectWord(toks, pos, "a parameter name after 'using'")
                    If IsFnWord(mFnOf, VLA_Identity.Fold(pName)) Or IsFnWord(mFnNullary, VLA_Identity.Fold(pName)) Then
                        VLA_Messages.RaiseMsg "english-param-name-taken", "name", pName, "loc", LineTag(pos - 1)
                    End If
                    pDef = ""
                    If TokAt(toks, pos) = "of" Then
                        pos = pos + 1
                        pDef = ParseExprReq(toks, pos)
                    End If
                    If Len(pDef) > 0 Then
                        sawDefault = True
                    ElseIf sawDefault Then
                        VLA_Messages.RaiseMsg "english-param-missing-default", "name", pName, "loc", LineTag(pos - 1)
                    End If
                    If Len(paramSpec) > 0 Then paramSpec = paramSpec & " "
                    If Len(pDef) > 0 Then
                        paramSpec = paramSpec & "(optional " & pName & " Variant " & pDef & ")"
                    Else
                        paramSpec = paramSpec & "(" & pName & " Variant)"
                    End If
                    MarkDeclared pName
                    sigParams.Add VLA_Identity.Fold(pName)
                    sigReq.Add IIf(Len(pDef) > 0, "0", "1")
                    If TokAt(toks, pos) = "and" Then
                        pos = pos + 1
                    Else
                        Exit Do
                    End If
                Loop
                ExpectTok toks, pos, ":", "':' after the 'using' parameters"
                mFnUsingNames.Add sigParams, VLA_Identity.Fold(subName)
                mFnUsingReq.Add sigReq, VLA_Identity.Fold(subName)
                mFnActNames.Add VLA_Identity.Fold(subName)
                mInFuncDef = True
                Set body = ParseBlock(toks, pos)
                mInFuncDef = False
                outParts = outParts & BuildSub(subName, body, paramSpec, True) & vbCrLf
                Set mDeclared = mainDecl
                Set mAssigned = mainAsgn
            ElseIf TokAt(toks, pos) = "of" Then
                ' Value-returning action: one parameter, named in the
                ' header the same way it will be used at call sites
                ' ("tax of subtotal") - the definition rhymes with the
                ' call. Registered BEFORE the body parses, so the
                ' action can use itself (recursion composes too).
                pos = pos + 1
                pName = ExpectWord(toks, pos, "a parameter name after 'of'")
                ExpectTok toks, pos, ":", "':' after 'To " & subName & " of " & pName & "'"
                Set mDeclared = New Collection
                Set mAssigned = New Collection
                MarkDeclared pName          ' the parameter is never auto-dimmed
                AddFnEntry mFnOf, subName, subName
                mUserFnWords.Add VLA_Identity.Fold(subName)
                mFnActNames.Add VLA_Identity.Fold(subName)
                mInFuncDef = True
                Set body = ParseBlock(toks, pos)
                mInFuncDef = False
                outParts = outParts & BuildSub(subName, body, "(" & pName & " Variant)", True) & vbCrLf
                Set mDeclared = mainDecl
                Set mAssigned = mainAsgn
            Else
            Set mDeclared = New Collection
            Set mAssigned = New Collection
            paramSpec = ""
            Set sigParams = New Collection
            Set sigReq = New Collection
            sawDefault = False
            If TokAt(toks, pos) = "," Then
                pos = pos + 1
                ExpectWordIs toks, pos, "with"
                Do
                    pOpt = False
                    If TokAt(toks, pos) = "optional" Then
                        pOpt = True
                        pos = pos + 1
                    End If
                    pName = ExpectWord(toks, pos, "a parameter name after 'with'")
                    pDef = ""
                    If TokAt(toks, pos) = "of" Then
                        pos = pos + 1
                        pDef = ParseExprReq(toks, pos)
                    End If
                    If Len(paramSpec) > 0 Then paramSpec = paramSpec & " "
                    ' A default makes the parameter optional even
                    ' without the keyword (VBA requires it anyway).
                    ' B6 retrofit: VBA requires optional parameters
                    ' last - refuse the ordering here with words
                    ' instead of a compile crash (the hole existed
                    ' since with-clauses shipped).
                    If pOpt Or Len(pDef) > 0 Then
                        sawDefault = True
                    ElseIf sawDefault Then
                        VLA_Messages.RaiseMsg "english-param-required-after-optional", "name", pName, "loc", LineTag(pos - 1)
                    End If
                    If pOpt Or Len(pDef) > 0 Then
                        paramSpec = paramSpec & "(optional " & pName & " Variant" & _
                                    IIf(Len(pDef) > 0, " " & pDef, "") & ")"
                    Else
                        paramSpec = paramSpec & "(" & pName & " Variant)"
                    End If
                    MarkDeclared pName          ' params are never auto-dimmed
                    sigParams.Add VLA_Identity.Fold(pName)
                    sigReq.Add IIf(pOpt Or Len(pDef) > 0, "0", "1")
                    If TokAt(toks, pos) = "and" Then
                        pos = pos + 1
                    Else
                        Exit Do
                    End If
                Loop
            End If
            ExpectTok toks, pos, ":", "':' after the action name (or its 'with' clause)"
            Set body = ParseBlock(toks, pos)
            outParts = outParts & BuildSub(subName, body, paramSpec) & vbCrLf
            mActNames.Add VLA_Identity.Fold(subName)
            mActParams.Add sigParams
            mActReq.Add sigReq
            Set mDeclared = mainDecl
            Set mAssigned = mainAsgn
            End If
        Else
            mainStmts.Add ParseTracked(toks, pos, 1)
        End If
    Loop

    ' Every definition is now known: check the recorded calls.
    ValidateActionCalls

    If mainStmts.Count > 0 Then
        outParts = outParts & BuildSub("main", mainStmts) & vbCrLf
    End If

    ' Step infrastructure: the step variable, the error reporter, and
    ' the step-number -> sentence lookup, all generated as VLA.
    If mStepTracking And mStepCount > 0 Then
        outParts = "(dim vla-step Long)" & vbCrLf & vbCrLf & outParts & BuildStepInfra()
    End If
    If mTryCount > 0 Then
        ' B7: the problem - what went wrong, captured before the
        ' Resume hop clears Err, readable in recovery paragraphs.
        outParts = "(dim vla-problem String)" & vbCrLf & vbCrLf & outParts
    End If
    If Len(mAliasDefs) > 0 Then
        outParts = mAliasDefs & vbCrLf & outParts
    End If

    If Len(mIncludeLines) > 0 Then
        ' G12.1 (owner-incident correction): library imports ride at
        ' the BOTTOM. The G12.0 top placement put spliced library
        ' PROCEDURES above the program's module-level declarations,
        ' and VBA ignores declarations after the first procedure -
        ' Option Explicit's "Variable not defined" modal at Run, the
        ' brief's own hazard list in generated-module form. Bottom
        ' placement is topology-safe (procedures after procedures)
        ' and absolutely source-map neutral (nothing above it moves,
        ' even in importing programs - stronger than the L4.2 claim
        ' the top placement borrowed). The core's new TopoGuard
        ' refuses the broken shape in words at Check regardless.
        outParts = outParts & vbCrLf & "; ---- libraries (G12) ----" & vbCrLf & mIncludeLines
    End If
    If Len(mVocabMacros) > 0 Then
        ' L4.2 (owner-incident revision): phrasebook macros ride at
        ' the BOTTOM of every translation. Pass 1 collects top-level
        ' defmacros order-independently, so bottom is semantically
        ' identical to top - and it is SOURCE-MAP NEUTRAL: prepending
        ' shifted every ' vla:N tag in the emitted VBA by the block's
        ' height (the owner's golden diff, +4 on every line), so any
        ' future macro change would have re-shifted the whole .vba
        ' golden - noise exactly where the instrument shows signal.
        ' Appended, macro changes touch only the goldens' tail.
        outParts = outParts & vbCrLf & "; ---- carried by phrasebooks (L4) ----" & vbCrLf & _
                   mVocabMacros & vbCrLf
    End If
    EnglishToVla = outParts
    If VLA.mProfileOn Then mProfMsTransBuild = mProfMsTransBuild + ProfMs(ptProf)
End Function

' Parse one statement, assigning it a step number and recording its
' sentence text; the emitted statement is preceded by a step marker.
' Tests and comma-bodies call ParseStmt directly and stay unmarked.
' S3.2: the original text of source line ln, trimmed; falls back to
' the parsed rendition when the line isn't recoverable (defensive -
' tracked steps always carry a line, and direct ParseStmt callers
' never reach the step table).
Private Function SrcLineText(ByVal ln As Long, toks() As String, ByVal pos As Long) As String
    If IsArray(mSrcLines) Then
        If ln - 1 >= LBound(mSrcLines) And ln - 1 <= UBound(mSrcLines) Then
            Dim s As String
            s = Trim$(CStr(mSrcLines(ln - 1)))
            If Len(s) > 0 Then
                SrcLineText = s
                Exit Function
            End If
        End If
    End If
    SrcLineText = RenderSentenceAt(toks, pos)
End Function

Private Function ParseTracked(toks() As String, ByRef pos As Long, ByVal ind As Long) As String
    mCurLine = TokLine(pos)            ' V1: assignment-line bookkeeping
    If Not mStepTracking Then
        ParseTracked = ParseStmt(toks, pos, ind)
        Exit Function
    End If
    mStepCount = mStepCount + 1
    Dim myN As Long
    myN = mStepCount
    Dim stpLn As Long
    stpLn = TokLine(pos)
    ' S3.2 (owner catch, first standalone dialog): the step table
    ' quotes the user's ORIGINAL sentence - their case, their
    ' spacing - not the parser's lowercased rendition, which read
    ' as a stranger's paraphrase of what they wrote.
    If stpLn > 0 Then
        mStepTexts.Add SrcLineText(stpLn, toks, pos) & " [line " & stpLn & "]"
    Else
        mStepTexts.Add RenderSentenceAt(toks, pos)
    End If
    Dim stmt As String
    stmt = ParseStmt(toks, pos, ind)
    ' V4: wrap the sentence's form(s) in (at-line N ...) so every VBA
    ' statement they become carries ' vla:X src:N - the three-layer
    ' round trip, VBA statement -> vla line -> sentence line, readable
    ' in the generated module itself. One wrapper covers everything
    ' the sentence produced (Try translates to a train of sibling
    ' forms; the emitter's at-line loops over all of them), and block
    ' bodies re-wrap per sentence, overriding within their extent.
    ' The step MARKER stays outside the wrapper: it is machinery, and
    ' the map marks the user's statements - the same rationale that
    ' keeps macro-template lines unmapped. Textual wrapping is safe
    ' because s-expressions are their own terminators.
    Dim pad As String
    pad = String$(ind * 2, " ")
    If stpLn > 0 Then
        ' S5: the guarded trace call rides between the step marker and
        ' the sentence - one Boolean check per step when the trace is
        ' off, the step's number and ORIGINAL text logged when on.
        ParseTracked = pad & "(set! vla-step " & myN & ")" & vbCrLf & _
                       pad & "(if (vlatraceon) (then (vlatracestep " & myN & " (vla-step-text " & myN & "))))" & vbCrLf & _
                       pad & "(at-line " & stpLn & vbCrLf & stmt & ")"
    Else
        ParseTracked = pad & "(set! vla-step " & myN & ")" & vbCrLf & _
                       pad & "(if (vlatraceon) (then (vlatracestep " & myN & " (vla-step-text " & myN & "))))" & vbCrLf & stmt
    End If
End Function

' Render the sentence (or block header) starting at a token position,
' for step lookup: stops after the first "." or ":".
Private Function RenderSentenceAt(toks() As String, ByVal start As Long) As String
    Dim r As String, i As Long, t As String
    ' L0: a raw VLA token IS the sentence - render it alone (trimmed
    ' for the step table) instead of walking into the next sentence.
    t = TokAt(toks, start)
    If Left$(t, 1) = "(" Then
        If Len(t) > 60 Then t = Left$(t, 57) & "..."
        RenderSentenceAt = t
        Exit Function
    End If
    For i = start To start + 29
        t = TokAt(toks, i)
        If Len(t) = 0 Then Exit For
        If t = PARA_TOK Then Exit For
        If IsStrTok(t) Then t = """" & Mid$(t, 2) & """"
        If Len(r) > 0 And t <> "." And t <> "," And t <> ":" Then r = r & " "
        r = r & t
        If t = "." Or t = ":" Then Exit For
    Next
    RenderSentenceAt = r
End Function

' The generated error reporter and step lookup, emitted as VLA so the
' whole program remains inspectable at every layer.
Private Function BuildStepInfra() As String
    ' V8: builder - the per-step case loop below grows with the
    ' program (one case per sentence), the pass's textbook quadratic.
    Dim sb As String, sbU As Long
    ' S3: the step-failure dialog goes through the message seam - a
    ' runtime helper resolvable in BOTH topologies (Frazaro_EN_Runtime beside
    ' the generated module in a user workbook; VLA_Runtime in the
    ' add-in and the dev workbook) - so the dialog's CONTENT is
    ' readable by the test harness. Styling (vbexclamation, title)
    ' now lives in the helper: one place, not every generated program.
    SbAdd sb, sbU, "(sub vla-report-error ()" & vbCrLf
    SbAdd sb, sbU, "  (vlashowerror (& ""Something went wrong at step "" vla-step "":"" vbcrlf vbcrlf" & vbCrLf
    SbAdd sb, sbU, "                   (vla-step-text vla-step) vbcrlf vbcrlf" & vbCrLf
    SbAdd sb, sbU, "                   ""Excel says: "" err.description)))" & vbCrLf & vbCrLf
    SbAdd sb, sbU, "(function vla-step-text ((byval n Long)) String" & vbCrLf
    SbAdd sb, sbU, "  (select n" & vbCrLf
    Dim i As Long
    For i = 1 To mStepTexts.Count
        SbAdd sb, sbU, "    (case (" & i & ") (return " & VlaStringLit(CStr(mStepTexts.Item(i))) & "))" & vbCrLf
    Next
    SbAdd sb, sbU, "    (case-else (return ""an unknown step""))))" & vbCrLf
    ' (S4.1: a "sacrificial compile probe" no-op briefly lived here -
    ' the theory being that entering one procedure compiles the whole
    ' module inside the IDE's error handling. The owner's deliberate-
    ' defect test disproved BOTH assumptions on a real machine: VBA's
    ' default Compile On Demand compiles per PROCEDURE, so the no-op
    ' verified nothing beyond itself - and when a defective procedure
    ' did compile under Application.Run, the error crossed as VBA's
    ' untrappable MODAL, not as a trappable Err. No in-VBA probe
    ' survives that second fact; S4 is re-scoped to a static resolve
    ' check on the generated text, per the Alpha 3 roadmap.)
    BuildStepInfra = SbText(sb, sbU)
End Function

' Enable/disable step mapping for subsequent translations (default on).
Public Sub EnglishStepTracking(ByVal enabled As Boolean)
    EnsureInit
    mStepTracking = enabled
End Sub

Public Function EnglishToVba(ByVal text As String) As String
    EnglishToVba = VlaTranspile(EnglishToVla(text))
End Function

Public Sub EnglishCompileToModule(ByVal text As String, ByVal moduleName As String, _
                                  Optional ByVal targetWb As Workbook)
    ' D1: the IDE passes its captured host so the injection cannot be
    ' misdirected by a focus change; omitted, ActiveWorkbook applies.
    VlaCompileToModule EnglishToVla(text), moduleName, targetWb
End Sub

' Extend the grammar with one production. First matching rule wins;
' rules are tried in registration order after the built-in prelude.
Public Sub EnglishAddPhrase(ByVal pattern As String, ByVal template As String)
    EnsureInit
    AddPhraseRule pattern, template
End Sub

' =====================================================================
'  Grammar registration
' =====================================================================

Private Sub EnsureInit()
    If mInited Then Exit Sub
    Set mPatItems = New Collection
    Set mPatSources = New Collection
    mVocabMacros = ""
    Set mVocabMacroNames = New Collection
    Set VLA_English.mKeywordAlias = New Collection
    Set mPatForms = New Collection
    Set mPatTexts = New Collection
    Set mPatSigs = New Collection
    Set mSigOwner = New Collection
    Set mLintWarnings = New Collection
    Set mRuleExamples = New Collection
    Set mRuleTestCounts = New Collection
    RegisterBuiltinFuncWords
    mStepTracking = True
    ' The prelude vocabulary. Each line is one DCG production.
    mLoadSource = "(built-in)"
    AddPhraseRule "set {v:var} to {e:expr}", "(set! {v} {e})"
    AddPhraseRule "add {e:expr} to {v:var}", "(add! {v} {e})"
    AddPhraseRule "increase {v:var} by {e:expr}", "(add! {v} {e})"
    AddPhraseRule "decrease {v:var} by {e:expr}", "(set! {v} (- {v} {e}))"
    ' B7: the multiplicative verbs. "Grow total by 10%." means x1.10 -
    ' the reading every typist intends - while Increase/Decrease/Add
    ' stay purely additive and REFUSE percent-tailed amounts loudly
    ' (see the sniff cases in ParseStmt). One verb, one mechanism.
    AddPhraseRule "grow {v:var} by {e:expr}", "(set! {v} (* {v} (+ 1 {e})))"
    AddPhraseRule "shrink {v:var} by {e:expr}", "(set! {v} (* {v} (- 1 {e})))"
    ' Lists (B1). "Append" claims a fresh verb so arithmetic "add"
    ' never means two mechanisms. The list slot is :var deliberately,
    ' not :name: appending to a name the program never Created then
    ' auto-declares it, so the mistake surfaces as a step-numbered
    ' runtime message ("Object variable not set" at YOUR sentence)
    ' instead of a VBA compile error - the failure class no runtime
    ' handler can catch (the renamed-parameter incident).
    AddPhraseRule "append {e:expr} to {v:var}", "(. {v} add {e})"
    AddPhraseRule "show {e:expr}", "(msgbox {e})"
    AddPhraseRule "say {e:expr}", "(msgbox {e})"
    AddPhraseRule "log {e:expr}", "(debug-print {e})"
    AddPhraseRule "stop", "(exit-sub)"
    mLoadSource = ""
    mPreludeCount = mPatItems.Count
    mInited = True
End Sub

' G0: slot spelling is {name} / {name:category} - braces both sides
' (pattern and template), owner revision. The closing brace is load-
' bearing, not cosmetic: it delimits the slot completely, so template
' substitution of {t} can never touch {total} - the prefix-collision
' class the old ?-sigil carried a longest-first workaround for.
' Returns True and fills name/cat when t is a well-formed slot;
' raises with directions on a malformed one; False for a literal.
Private Function IsSlotTok(ByVal t As String, ByRef slotName As String, ByRef cat As String, _
                            Optional ByRef hasDefault As Boolean, Optional ByRef defaultVal As String) As Boolean
    hasDefault = False
    defaultVal = ""
    If Left$(t, 1) <> "{" Then
        If Left$(t, 1) = "?" Then
            VLA_Messages.RaiseMsg "english-old-slot-spelling", "t", t, "suggested", "{" & Mid$(t, 2) & "}", "example", "{name} (no category)"
        End If
        Exit Function
    End If
    If Right$(t, 1) <> "}" Or Len(t) < 3 Then
        VLA_Messages.RaiseMsg "english-slot-not-closed", "t", t, "ex1", "{name}", "ex2", "{name:category}"
    End If
    Dim inner As String
    inner = Mid$(t, 2, Len(t) - 2)
    Dim cp As Long
    cp = InStr(inner, ":")
    If cp = 0 Then
        slotName = inner
        cat = "name"
    Else
        slotName = Left$(inner, cp - 1)
        cat = Mid$(inner, cp + 1)
    End If
    If Len(slotName) = 0 Or (cp > 0 And Len(cat) = 0) Then
        VLA_Messages.RaiseMsg "english-malformed-slot", "t", t, "ex1", "{name}", "ex2", "{name:category}"
    End If
    ' G7: an optional default - {name:category=text} - splits off after
    ' the category, last '=' wins. Pure substitution: whatever follows
    ' '=' is the exact VLA text used when the slot goes unmatched, never
    ' re-parsed or re-typed. Every caller that only wants slotName/cat
    ' (lint, validation, ref-slot listing) sees the same stripped cat
    ' as before - defaults are invisible unless a caller asks for them.
    Dim eq As Long
    eq = InStrRev(cat, "=")
    If eq > 0 Then
        hasDefault = True
        defaultVal = Mid$(cat, eq + 1)
        cat = Left$(cat, eq - 1)
        If Len(cat) = 0 Or Len(defaultVal) = 0 Then
            VLA_Messages.RaiseMsg "english-malformed-default", "t", t, "ex", "{name:category=text}"
        End If
    End If
    IsSlotTok = True
End Function

Private Sub AddPhraseRule(ByVal pattern As String, ByVal template As String, _
                          Optional ByVal isOverride As Boolean = False)
    Dim items As New Collection
    Dim parts() As String
    Dim i As Long, w As String
    parts = Split(pattern, " ")
    For i = LBound(parts) To UBound(parts)
        w = VLA_Identity.Fold(Trim$(parts(i)))
        If Len(w) > 0 Then
            If Not IsNoiseWord(w) Then items.Add w
        End If
    Next
    If items.Count = 0 Then VLA_Messages.RaiseMsg "english-empty-pattern"

    ValidateRuleItems pattern, items

    Dim src As String
    src = mLoadSource
    If Len(src) = 0 Then src = "(added directly)"

    If isOverride Then
        ' G3: an override REPLACES exactly one earlier same-shape
        ' rule, in place - position (and thus match order relative to
        ' every other rule) is preserved, and the slot's provenance
        ' records both files. The marker must be earned: matching
        ' nothing is an error (a stale override after the base rule
        ' was reworded should scream, not silently append), matching
        ' several is an error (one override, one rule), and the
        ' built-in core is not overridable (an overridden prelude
        ' slot would survive EnglishResetGrammar as a stale ghost -
        ' write a differently-worded rule instead).
        ' G3 perf: mSigOwner turns "which already-registered rule(s)
        ' does this override's shape match" from an O(corpus size) scan
        ' into O(this override's own signature count) - see AddSigOwner/
        ' RebuildSigOwner's own header for the full reasoning. Dedupe by
        ' OWNER INDEX, not by signature: an override with several
        ' branches routinely matches the SAME earlier rule through
        ' several of its own signatures, and that must count as one
        ' hit, not several - "ambiguous" means genuinely different
        ' target rules, never the same one found twice.
        Dim newSigs As Collection
        Set newSigs = ExpandedSignatures(items)
        Dim hits As New Collection
        Dim k As Long
        Dim e As Variant
        For Each e In newSigs
            If CollHasKey(mSigOwner, CStr(e)) Then
                k = CLng(mSigOwner.Item(CStr(e)))
                If Not CollHasKey(hits, "i" & k) Then hits.Add k, "i" & k
            End If
        Next
        If hits.Count = 0 Then
            VLA_Messages.RaiseMsg "english-override-no-match", "pattern", pattern
        End If
        If hits.Count > 1 Then
            Dim lst As String
            For k = 1 To hits.Count
                If Len(lst) > 0 Then lst = lst & "; "
                lst = lst & "'" & mPatTexts.Item(CLng(hits.Item(k))) & "'"
            Next
            VLA_Messages.RaiseMsg "english-override-ambiguous", "pattern", pattern, "count", hits.Count, "list", lst
        End If
        Dim tgt As Long
        tgt = CLng(hits.Item(1))
        If tgt <= mPreludeCount Then
            VLA_Messages.RaiseMsg "english-override-matches-builtin", "pattern", pattern, "builtin", mPatTexts.Item(tgt)
        End If
        LintRule pattern, items, tgt

        Dim oldSrc As String
        oldSrc = mPatSources.Item(tgt)
        ' G3 perf: capture tgt's own OLD signature lines before ReplaceAt
        ' overwrites mPatSigs.Item(tgt) - mSigOwner needs to stop
        ' crediting tgt for a shape it no longer has, not just start
        ' crediting it for its new one.
        Dim oldSigLines() As String
        oldSigLines = Split(mPatSigs.Item(tgt), vbLf)
        ReplaceAt mPatItems, tgt, items
        ReplaceAt mPatForms, tgt, TemplateForms(template)
        ReplaceAt mPatTexts, tgt, pattern
        ReplaceAt mPatSigs, tgt, JoinSigs(newSigs)
        ReplaceAt mPatSources, tgt, src & " (overrides " & oldSrc & ")"
        Dim osl As Variant
        For Each osl In oldSigLines
            ' only clear an entry tgt itself still owns - a signature
            ' tgt merely SHARED with something else (mAuditMode's own
            ' tolerated-duplicate case) is not tgt's to remove.
            If CollHasKey(mSigOwner, CStr(osl)) Then
                If CLng(mSigOwner.Item(CStr(osl))) = tgt Then mSigOwner.Remove CStr(osl)
            End If
        Next
        Dim nsv As Variant
        For Each nsv In newSigs
            AddSigOwner CStr(nsv), tgt
        Next
        Dim exf As Boolean
        exf = False
        CollGet mRuleExamples, "r" & tgt, exf
        If exf Then mRuleExamples.Remove "r" & tgt   ' the old rule's worked example is stale
        mDspValid = False   ' V7: an override can reword the pattern's
                            ' head in place - count unchanged, so the
                            ' flag is the only honest invalidation
        Exit Sub
    End If

    Dim dupMsg As String
    dupMsg = LintRule(pattern, items)
    If Len(dupMsg) > 0 And Not mAuditMode Then
        ' G3: a silent same-shape shadow was the linter's oldest
        ' complaint recast as a mechanism; now it refuses the load
        ' unless the replacement is earned with the marker. In audit
        ' mode the warning is recorded and the walk continues - the
        ' audit's contract is to report everything a load would refuse.
        ' F.13: the suggested fix used to be a literal, copy-pastable
        ' "override: <pattern>" line - the old DSL's own single
        ' spelling. The new grammar has no single spelling to suggest
        ' (english-vla-override, deutsche-vla-override, ... - whichever
        ' <lingua>-vla this rule's own directive already used), so the
        ' message names the shape instead of a snippet.
        VLA_Messages.RaiseMsg "english-rule-shadow", "dupMsg", dupMsg
    End If

    Dim appendSigs As Collection
    Set appendSigs = ExpandedSignatures(items)
    mPatItems.Add items
    mPatForms.Add TemplateForms(template)
    mPatTexts.Add pattern
    mPatSigs.Add JoinSigs(appendSigs)
    mPatSources.Add src
    mDspValid = False   ' V7: rebuilt lazily at the next dispatch
    Dim asv As Variant
    For Each asv In appendSigs
        AddSigOwner CStr(asv), mPatItems.Count   ' G3 perf: this rule's own index, just appended
    Next
End Sub

' G3: in-place replacement for the parallel rule collections.
Public Sub ReplaceAt(coll As Collection, ByVal idx As Long, ByVal v As Variant)
    coll.Remove idx
    If idx > coll.Count Then
        If IsObject(v) Then coll.Add v Else coll.Add v
    Else
        If IsObject(v) Then coll.Add v, , idx Else coll.Add v, , idx
    End If
End Sub

' F.2: pre-parse a template into forms, once, at registration time -
' the reader half of the double round-trip (English substitutes text,
' VlaTranspile re-tokenizes it) never has to run twice per sentence.
' Forms are the ONLY path a rule's template takes now (the dual-path
' era's Replace()-text fallback is retired - see BETA_ROADMAP.md's F.2
' entry), so a template that doesn't parse as at least one well-formed
' VLA form is a load-time bug, not a degradation to fall back from -
' raises here, with the template text in the message, at the exact
' same load-time gate every other malformed-rule refusal already goes
' through (AddPhraseRule's own ValidateRuleItems/LintRule checks,
' caught the same way by EnglishTryRule's own audition tool). A
' template that DOES parse may still yield more than one top-level
' form (a multi-statement template, several sibling forms) -
' FormSubstitute and TryFormPath below splice every one of them, in
' order.
Private Function TemplateForms(ByVal template As String) As Collection
    On Error GoTo fail
    Set TemplateForms = VLA.VlaReadForms(template)
    If TemplateForms Is Nothing Then GoTo fail
    If TemplateForms.Count = 0 Then GoTo fail
    Exit Function
fail:
    Dim d As String
    d = Err.Description
    VLA_Messages.RaiseMsg "english-template-not-well-formed", "template", template, "detail", IIf(Len(d) > 0, " (" & d & ")", "")
End Function

' F.2: deep-copy tmplForm, substituting bound values in as FORMS, never
' text - so a bound value can never collide with a slot delimiter or
' corrupt a sibling slot's braces the way the old Replace()-based text
' path could (retired; see BETA_ROADMAP.md's F.2 entry). Two shapes,
' both real in the shipped phrasebook: an atom that reads EXACTLY
' "{slotname}" splices the
' bound value's FULL form in structurally (a LIST, when the slot is an
' :expr/:cond - "total plus 1" binds a (+ total 1) list, not text); a
' slot GLUED onto a larger atom - "make-{d}", "xl{d}", BETA_ROADMAP.md's
' own documented idiom for building a keyword from an alternation's
' matched word - text-splices just that bound word into the identifier,
' the one place a text-level operation is still correct because the
' result must stay ONE atom, not become a list.
'
' G0 hazard, the reason BoundLookup below is a Sub with a ByRef out
' parameter and not a Variant-returning Function: a VBA Collection's
' hidden default member is Item, which takes an index. "target =
' expr" where expr evaluates to an object with no zero-argument
' default member raises 450 ("wrong number of arguments") - exactly
' what the form path's first live run surfaced, on the FIRST :expr
' slot whose bound value was a list rather than a bare atom. Set is
' the only safe way to move an object reference; BoundLookup Sets
' directly into the caller's variable, and the caller below branches
' on IsObject(that variable) - never on the raw call expression -
' before choosing Set or Let for its OWN assignment.
Private Function FormSubstitute(tmplForm As Variant, bn As Collection, bv As Collection) As Variant
    If Not IsObject(tmplForm) Then
        Dim s As String
        s = CStr(tmplForm)
        If Left$(s, 1) = "{" And Right$(s, 1) = "}" And Len(s) > 2 And InStr(2, s, "{") = 0 Then
            Dim bf As Variant
            BoundLookup Mid$(s, 2, Len(s) - 2), bn, bv, bf
            If IsObject(bf) Then
                Set FormSubstitute = bf
            Else
                FormSubstitute = bf
            End If
        ElseIf InStr(s, "{") > 0 Then
            FormSubstitute = SpliceEmbeddedSlots(s, bn, bv)
        Else
            FormSubstitute = tmplForm
        End If
        Exit Function
    End If
    Dim src As Collection, dst As New Collection
    Set src = tmplForm
    Dim e As Variant
    For Each e In src
        dst.Add FormSubstitute(e, bn, bv)
    Next
    Set FormSubstitute = dst
End Function

' F.2: bv's own form for slotName, whole - a list when the slot is
' :expr/:cond, an atom otherwise - Set into outVal (see FormSubstitute's
' own comment for why a Function return would be unsafe here). Every bv
' entry is already independently-valid single-form VLA text today
' (VlaStringLit's quoted literals, ParseExpr's/ParseCond's built forms,
' an alternation's matched surface word) - the same guarantee the text
' path has always silently depended on - so VLA.VlaReadForms(val) below
' always yields exactly one form. Falls back to the literal "{slotname}"
' text, unresolved, when bn has no entry by that name - the same no-op
' the old Replace() loop had for a template brace naming no declared
' pattern slot.
Private Sub BoundLookup(ByVal slotName As String, bn As Collection, bv As Collection, ByRef outVal As Variant)
    Dim k As Long
    For k = 1 To bn.Count
        If CStr(bn.Item(k)) = slotName Then
            Dim bound As Collection
            Set bound = VLA.VlaReadForms(CStr(bv.Item(k)))
            If bound.Count <> 1 Then
                VLA_Messages.RaiseMsg "english-slot-value-not-one-form", "slot", slotName, "count", bound.Count, "value", CStr(bv.Item(k))
            End If
            If IsObject(bound.Item(1)) Then
                Set outVal = bound.Item(1)
            Else
                outVal = bound.Item(1)
            End If
            Exit Sub
        End If
    Next
    outVal = "{" & slotName & "}"
End Sub

' F.2: scan atomText for every "{slotname}" run and replace it with
' EmbeddedSlotText's contribution - the general form of the "make-{d}"
' idiom (a slot anywhere inside a larger atom, not just as a suffix).
Private Function SpliceEmbeddedSlots(ByVal atomText As String, bn As Collection, bv As Collection) As String
    Dim r As String, i As Long, openPos As Long, closePos As Long
    i = 1
    Do While i <= Len(atomText)
        openPos = InStr(i, atomText, "{")
        If openPos = 0 Then
            r = r & Mid$(atomText, i)
            Exit Do
        End If
        closePos = InStr(openPos, atomText, "}")
        If closePos = 0 Then
            r = r & Mid$(atomText, i)
            Exit Do
        End If
        r = r & Mid$(atomText, i, openPos - i)
        r = r & EmbeddedSlotText(Mid$(atomText, openPos + 1, closePos - openPos - 1), bn, bv)
        i = closePos + 1
    Loop
    SpliceEmbeddedSlots = r
End Function

' F.2: the text one embedded slot contributes when it is glued onto a
' larger atom. Only a bare symbol can be glued into an identifier this
' way - a list or a quoted string spliced there would not be valid VLA
' under the OLD Replace() path either (raw text glued into the middle
' of a bareword, failing one stage later at VlaTranspile's own
' tokenizer); this raises the same failure earlier, with a clearer
' message, instead of reproducing the old path's garbage.
Private Function EmbeddedSlotText(ByVal slotName As String, bn As Collection, bv As Collection) As String
    Dim k As Long
    For k = 1 To bn.Count
        If CStr(bn.Item(k)) = slotName Then
            Dim bound As Collection
            Set bound = VLA.VlaReadForms(CStr(bv.Item(k)))
            If bound.Count <> 1 Then
                VLA_Messages.RaiseMsg "english-slot-value-not-one-form", "slot", slotName, "count", bound.Count, "value", CStr(bv.Item(k))
            End If
            ' VBA's Or does not short-circuit - CStr() on an object
            ' argument raises its own confusing error, so the object
            ' check must be a separate branch, never the left side of
            ' an Or evaluated alongside a CStr() on the same value.
            Dim badGlue As Boolean
            If IsObject(bound.Item(1)) Then
                badGlue = True
            Else
                badGlue = (Left$(CStr(bound.Item(1)), 1) = Chr$(34))
            End If
            If badGlue Then
                VLA_Messages.RaiseMsg "english-slot-glued-to-identifier", "quoted", "{" & slotName & "}", "value", CStr(bv.Item(k))
            End If
            EmbeddedSlotText = CStr(bound.Item(1))
            Exit Function
        End If
    Next
    EmbeddedSlotText = "{" & slotName & "}"
End Function

' F.2: the form path for one rule match. AddPhraseRule refuses to
' register a rule whose template doesn't parse as at least one form
' (TemplateForms raises there), so every loaded rule's forms collection
' is always non-empty by the time a match reaches here. Multiple
' top-level template forms join on their own lines, matching how
' EnglishToVla joins multi-statement text.
Private Function TryFormPath(ByVal idx As Long, bn As Collection, bv As Collection) As String
    Dim forms As Collection
    Set forms = mPatForms.Item(idx)
    Dim r As String
    Dim f As Variant
    For Each f In forms
        If Len(r) > 0 Then r = r & vbCrLf
        r = r & VLA.VlaWriteForm(FormSubstitute(f, bn, bv))
    Next
    TryFormPath = r
End Function

' F.2's dual-path era ended here. EnglishSetFormPath/EnglishFormPathEnabled
' (the dev/test on-off toggle) and EnglishFormPathSelfCheck (the A/B
' proof that comparing the two paths agreed) are retired along with
' mPatTmpls and TryPhrase's own Replace()-text fallback - there is only
' one path now, so there is nothing left to toggle or compare against.
' See BETA_ROADMAP.md's F.2 entry for the full retirement story;
' docs/TESTING.md's old Pass 2 (which called EnglishFormPathSelfCheck)
' is retired for the same reason - ordinary vocabulary loading already
' re-verifies every test:/fail: proof through the one remaining path.

' =====================================================================
'  G-RENDER: forms back to sentences - the reverse of this file's own
'  FormSubstitute. Given one concrete top-level VLA form (already past
'  its own {slot} placeholders - a real macro call, not a template),
'  find the loaded rule whose cached template form it structurally
'  matches and render that rule's English pattern with the concrete
'  bindings spliced back in.
'
'  v1 scope (BETA_ROADMAP.md's G-RENDER entry has the full design):
'  a rule is a render candidate only when its template compiles to
'  EXACTLY ONE top-level form (mPatForms(idx).Count = 1) - a
'  multi-statement template is not renderable yet. Ambiguity is real
'  and already present in the shipped corpus (delete-row's two
'  identical-shape rules; put-today/put-expr's literal-vs-wildcard
'  overlap), so a match is scored by specificity - how much of the
'  concrete form a template pins down with literal text rather than a
'  wildcard slot - and the most specific match wins; a true tie keeps
'  whichever rule was found first (registration order), the same
'  precedent G3's own duplicate-shape audit already uses. Original
'  word casing can never be recovered (VLA_Identity.Fold lowercases
'  every bareword at English-tokenize time, before any binding ever
'  exists) - a :range/:cell/:column value is canonicalized to
'  uppercase on render as a readability choice, not a restoration;
'  every other category prints its stored (lowercase) text as-is.
' =====================================================================

' The FormSubstitute inverse - now a thin pass-through to
' VLA_Unify.UnifyOneWay (PROLOG.1; see this file's own header note).
' bn/bv keep their G-RENDER-only convention unchanged: bn holds slot
' names, bv holds the RAW bound form (an atom or a Collection, never
' VLA text) - distinct from FormSubstitute's own bn/bv, which hold
' text; the two are never passed to each other's functions. Callers
' must still pass a FRESH bn/bv pair per attempt (UnifyOneWay leaves
' them however far the walk got the moment anything fails to unify) -
' FindRenderRule's own comment already names why ("As New" inside a
' loop only auto-instantiates once).
Private Function UnifyForm(tmplForm As Variant, concreteForm As Variant, _
                            bn As Collection, bv As Collection) As Boolean
    UnifyForm = VLA_Unify.UnifyOneWay(tmplForm, concreteForm, bn, bv)
End Function

' G0 hazard (same one BoundLookup's own comment names): a Collection's
' hidden default member takes an index, so "target = bv.Item(k)" for an
' object-valued item raises 450. A ByRef-out Sub, branching on
' IsObject(bv.Item(k)) before choosing Set or Let, is the only safe shape.
Private Sub RenderBoundLookup(ByVal slotName As String, bn As Collection, bv As Collection, ByRef outVal As Variant)
    Dim k As Long
    For k = 1 To bn.Count
        If CStr(bn.Item(k)) = slotName Then
            If IsObject(bv.Item(k)) Then
                Set outVal = bv.Item(k)
            Else
                outVal = bv.Item(k)
            End If
            Exit Sub
        End If
    Next
End Sub

' Match quality: the count of tmplForm's literal leaf atoms (no braces
' at all) - a template that pins down more of the concrete form with
' its OWN vocabulary, rather than a wildcard slot, is the more specific
' match. An :expr/:cond slot's bare-brace atom unifies with ANY
' subform, so a general rule (put {e:expr} into cell {r:cell}) always
' also matches whatever a specific one does (put today into cell
' {r:cell}) - this score is what lets the specific one win instead.
Private Function TemplateSpecificity(ByVal tmplForm As Variant) As Long
    If Not IsObject(tmplForm) Then
        If InStr(CStr(tmplForm), "{") = 0 Then TemplateSpecificity = 1
        Exit Function
    End If
    Dim lst As Collection
    Set lst = tmplForm
    Dim e As Variant
    For Each e In lst
        TemplateSpecificity = TemplateSpecificity + TemplateSpecificity(e)
    Next
End Function

' Structural unification alone is blind to a slot's DECLARED category -
' two rules with the identical template shape but different typed slots
' (put into cell {r:cell} vs. put into range {r:range}; ...column
' {c:column}... vs. ...column number {c:expr}...) tie on both form and
' specificity score, since a brace is a brace regardless of what
' category follows the colon in the PATTERN text (invisible to the
' template FORM entirely). This gate re-runs each candidate's OWN typed
' slots through the same RefShapeOk check the forward grammar itself
' enforces, against the bound text UnifyForm actually extracted, and
' disqualifies a candidate outright when a bound value doesn't fit the
' shape its own rule declares - e.g. a multi-cell "a1:b2" bound to a
' :cell slot, or a bare number "1" bound to a :column slot. A slot bound
' to a LIST (not a single atom) can never satisfy any typed-reference
' category, so it disqualifies immediately without calling RefShapeOk.
Private Function CandidateShapeOk(ByVal idx As Long, bn As Collection, bv As Collection) As Boolean
    Dim items As Collection
    Set items = mPatItems.Item(idx)
    Dim it As Variant, t As String, slotName As String, cat As String
    Dim hasDefault As Boolean, defaultVal As String
    For Each it In items
        t = CStr(it)
        If IsSlotTok(t, slotName, cat, hasDefault, defaultVal) Then
            Select Case cat
                Case "cell", "range", "column", "sheet", "color"
                    Dim bound As Variant
                    RenderBoundLookup slotName, bn, bv, bound
                    If IsObject(bound) Then Exit Function
                    If Not RefShapeOk(cat, StripQuoteSigil(CStr(bound))) Then Exit Function
            End Select
        End If
    Next
    CandidateShapeOk = True
End Function

' Scans every single-top-level-form rule for one whose template
' unifies against concreteForm AND whose own typed slots actually fit
' the bound text's shape (CandidateShapeOk), keeping the most specific
' match found (a strictly higher score replaces the current best; a
' tie keeps whichever rule was found FIRST, i.e. registration order -
' the same first-loaded-wins precedent G3's duplicate-shape audit
' already uses).
Private Function FindRenderRule(ByVal concreteForm As Variant, ByRef bestIdx As Long, _
                                 ByRef bestBn As Collection, ByRef bestBv As Collection) As Boolean
    Dim idx As Long, bestScore As Long
    bestIdx = 0
    bestScore = -1
    Dim bn As Collection, bv As Collection
    For idx = 1 To mPatItems.Count
        Dim forms As Collection
        Set forms = mPatForms.Item(idx)
        If forms.Count = 1 Then
            ' G0: a FRESH pair every attempt - "Dim ... As New" inside
            ' this loop would auto-instantiate lazily and then reuse the
            ' SAME Collection for every rule (As New only allocates once
            ' per procedure, on first Nothing access), silently leaking
            ' one rule's bindings into the next rule's unify attempt.
            ' An explicit Set after a plain Dim is the only shape that
            ' actually allocates a new object each iteration.
            Set bn = New Collection
            Set bv = New Collection
            If UnifyForm(forms.Item(1), concreteForm, bn, bv) Then
                If CandidateShapeOk(idx, bn, bv) Then
                    Dim score As Long
                    score = TemplateSpecificity(forms.Item(1))
                    If score > bestScore Then
                        bestScore = score
                        bestIdx = idx
                        Set bestBn = bn
                        Set bestBv = bv
                    End If
                End If
            End If
        End If
    Next
    FindRenderRule = (bestIdx > 0)
End Function

' The first branch of a bare/braced alternation spec, stem only - the
' deterministic choice render makes wherever the form carries no trace
' of which branch/surface was actually typed (a bare "into|in" binds
' nothing at all; a bound alternation slot's value IS its stem, already
' unambiguous - see RenderSlotValue).
Private Function FirstAltSurface(ByVal spec As String) As String
    Dim alts() As String
    alts = Split(spec, "|")
    Dim first As String
    first = alts(LBound(alts))
    Dim sp As Long
    sp = InStr(first, "/")
    If sp > 0 Then
        FirstAltSurface = Left$(first, sp - 1)
    Else
        FirstAltSurface = first
    End If
End Function

Private Function StripQuoteSigil(ByVal s As String) As String
    If Left$(s, 1) = Chr$(34) Then
        StripQuoteSigil = Mid$(s, 2)
    Else
        StripQuoteSigil = s
    End If
End Function

' Formats one resolved slot binding for splicing into English text, per
' its pattern-declared category (read from mPatItems, never from
' mPatForms - the template form carries no category info at all, only
' the pattern text does).
Private Function RenderSlotValue(ByVal cat As String, ByVal bound As Variant) As String
    If IsAltCat(cat) Then
        ' bound is the matched STEM (TryPhrase's own Case Else binds
        ' the stem, never the full surface) - print it as-is. Which of
        ' the branch's one or two surfaces (stem alone, or stem+suffix)
        ' produced it is not recoverable from the form, so the stem is
        ' what prints: always legal VLA, occasionally the less fluent
        ' of two English spellings ("ascend" for a stem/suffix branch
        ' like ascend/ing, never "ascending").
        RenderSlotValue = CStr(bound)
        Exit Function
    End If
    Select Case cat
        Case "range", "cell", "column"
            RenderSlotValue = UCase$(StripQuoteSigil(CStr(bound)))
        Case "sheet", "color", "text"
            RenderSlotValue = StripQuoteSigil(CStr(bound))
        Case "var", "name"
            RenderSlotValue = CStr(bound)
        Case "expr", "cond"
            RenderSlotValue = RenderExprForm(bound)
        Case Else
            VLA_Messages.RaiseMsg "english-render-no-category-renderer", "cat", cat
    End Select
End Function

' The one place a slot's bound value can be an arbitrary sub-form
' rather than a single atom - its own small, separate grammar, mirroring
' ParseSum/ParseProd/ParseCond's operator table in reverse. Two English
' surfaces can parse to the same operator symbol ("times" and
' "multiplied by" both read as *) so render commits to exactly one
' canonical word per symbol. Anything outside this known table (the
' blank?/contains/divisible/mod-based idioms ParseCondSimple's own
' special forms expand into) falls back to raw VLA text spliced into
' the sentence - always legal-looking, never a rendering failure, just
' occasionally inelegant - and sets mLastRenderUsedFallback so a caller
' auditing round-trip fidelity (EnglishRenderSelfCheck) knows this
' particular render was never round-trip tested.
Private Function RenderExprForm(ByVal v As Variant) As String
    If Not IsObject(v) Then
        RenderExprForm = StripQuoteSigil(CStr(v))
        Exit Function
    End If
    Dim lst As Collection
    Set lst = v
    If lst.Count = 3 Then
        Dim head As String
        head = CStr(lst.Item(1))
        Dim opWord As String
        opWord = ExprOpWord(head)
        If Len(opWord) > 0 Then
            RenderExprForm = RenderExprForm(lst.Item(2)) & " " & opWord & " " & RenderExprForm(lst.Item(3))
            Exit Function
        End If
        If head = "and" Or head = "or" Then
            RenderExprForm = RenderExprForm(lst.Item(2)) & " " & head & " " & RenderExprForm(lst.Item(3))
            Exit Function
        End If
    End If
    mLastRenderUsedFallback = True
    RenderExprForm = VLA.VlaWriteForm(v)
End Function

' The public entry point: render one concrete top-level form as an
' English sentence. Raises when no loaded rule's cached template form
' unifies against it - a caller wanting a softer "can this be rendered"
' check should trap this rather than pre-testing, since FindRenderRule
' is the only thing that actually knows.
Public Function EnglishRenderForm(ByVal concreteForm As Variant) As String
    EnsureInit
    mLastRenderUsedFallback = False
    Dim idx As Long, bn As Collection, bv As Collection
    If Not FindRenderRule(concreteForm, idx, bn, bv) Then
        VLA_Messages.RaiseMsg "english-render-no-matching-rule", "form", VLA.VlaWriteForm(concreteForm)
    End If
    Dim items As Collection
    Set items = mPatItems.Item(idx)
    Dim r As String
    Dim it As Variant, t As String, slotName As String, cat As String
    Dim hasDefault As Boolean, defaultVal As String, ow As String
    Dim word As String
    For Each it In items
        t = CStr(it)
        word = ""
        If IsSlotTok(t, slotName, cat, hasDefault, defaultVal) Then
            Dim bound As Variant
            RenderBoundLookup slotName, bn, bv, bound
            word = RenderSlotValue(cat, bound)
        ElseIf IsOptTok(t, ow) Then
            ' v1 policy: an optional literal never prints - the form
            ' carries no trace of whether one was present in the
            ' original sentence, so render always takes the shortest
            ' legal reading.
            word = ""
        ElseIf InStr(t, "|") > 0 Or InStr(t, "/") > 0 Then
            ' A bare alternation ("into|in") binds nothing - same
            ' reasoning, print the first branch deterministically.
            word = FirstAltSurface(t)
        Else
            word = t
        End If
        If Len(word) > 0 Then
            If Len(r) > 0 Then r = r & " "
            r = r & word
        End If
    Next
    If Len(r) > 0 Then r = UCase$(Left$(r, 1)) & Mid$(r, 2)
    EnglishRenderForm = r & "."
End Function

' Convenience wrapper over EnglishRenderForm for a caller holding VLA
' source text rather than an already-parsed form.
Public Function EnglishRenderText(ByVal vlaText As String) As String
    Dim forms As Collection
    Set forms = VLA.VlaReadForms(vlaText)
    If forms.Count <> 1 Then
        VLA_Messages.RaiseMsg "english-render-not-single-form", "count", forms.Count
    End If
    EnglishRenderText = EnglishRenderForm(forms.Item(1))
End Function

' Builds a synthetic-but-legal binding for every slot in rule idx's own
' pattern, one representative value per category - not real user data,
' just enough to instantiate the rule's template into a concrete form
' via FormSubstitute (this is FormSubstitute's OWN bn/bv convention,
' text values, NOT UnifyForm's raw-form convention - the two never mix).
' Returns False (bn/bv left partial) the moment a slot's category isn't
' one of the ones this function knows how to synthesize - today that
' never happens (the list matches ValidateRuleItems's own known-category
' set exactly), but a future category addition should fail this loudly
' rather than synthesize nonsense.
Private Function BuildSyntheticBindings(ByVal idx As Long, bn As Collection, bv As Collection) As Boolean
    Dim items As Collection
    Set items = mPatItems.Item(idx)
    Dim it As Variant, t As String, slotName As String, cat As String
    Dim hasDefault As Boolean, defaultVal As String
    For Each it In items
        t = CStr(it)
        If IsSlotTok(t, slotName, cat, hasDefault, defaultVal) Then
            Dim val As String
            If IsAltCat(cat) Then
                val = FirstAltSurface(cat)         ' a bare stem atom
            Else
                Select Case cat
                    Case "range":            val = """a1:b2"""
                    Case "cell":             val = """a1"""
                    Case "column":           val = """a"""
                    Case "sheet", "color", "text": val = """stub"""
                    Case "var", "name":      val = "stubvar"
                    Case "expr":             val = "1"
                    Case "cond":             val = "(> 1 0)"
                    Case Else
                        BuildSyntheticBindings = False
                        Exit Function
                End Select
            End If
            bn.Add slotName
            bv.Add val
        End If
    Next
    BuildSyntheticBindings = True
End Function

' Re-parses a single already-rendered sentence through the SAME
' machinery RunVocabTest uses for every test:/fail: line (state
' save/restore included - this exact minimal save/restore is already
' exercised at full corpus scale on every vocabulary load, not a novel
' pattern) and reads the resulting VLA text back into a form. Returns
' False on any failure to translate or to read back as exactly one
' form - never raises.
Private Function ReparseSentenceToForm(ByVal sentence As String, ByRef outForm As Variant) As Boolean
    Dim savedD As Collection, savedA As Collection
    Set savedD = mDeclared
    Set savedA = mAssigned
    Set mDeclared = New Collection
    Set mAssigned = New Collection
    Dim savedSusp As Boolean
    savedSusp = mUsageSuspended
    mUsageSuspended = True

    Dim toks() As String, pos As Long, got As String
    On Error GoTo fail
    toks = CanonicalizeStructuralWords(EnTokenize(sentence))   ' LX5.1
    pos = 1
    got = ParseStmt(toks, pos, 0)
    On Error GoTo 0

    Set mDeclared = savedD
    Set mAssigned = savedA
    mUsageSuspended = savedSusp

    Dim forms As Collection
    On Error Resume Next
    Set forms = VLA.VlaReadForms(got)
    On Error GoTo 0
    If forms Is Nothing Then Exit Function
    If forms.Count <> 1 Then Exit Function
    If IsObject(forms.Item(1)) Then
        Set outForm = forms.Item(1)
    Else
        outForm = forms.Item(1)
    End If
    ReparseSentenceToForm = True
    Exit Function

fail:
    Set mDeclared = savedD
    Set mAssigned = savedA
    mUsageSuspended = savedSusp
End Function

' G-RENDER's own corpus-wide prover, in EnglishFormPathSelfCheck's own
' spirit: for every single-top-level-form rule, synthesize a legal
' instance of its template, render it, and (when the render didn't fall
' back to raw VLA text for an out-of-scope :expr/:cond shape) re-parse
' the rendered sentence and confirm it reproduces the IDENTICAL form -
' never a text-identity check, since original word casing is never
' recoverable by construction (see this section's own opening comment).
' Deliberately does not require the WINNING rule to be idx itself:
' under real ambiguity (delete-row's pair, put-today/put-expr's overlap)
' a different, equally-or-more-specific rule may legitimately render the
' same synthetic instance - the invariant this checks is that SOME
' render-then-reparse cycle preserves the form, not that idx's own
' English wins the vote, which is the correct, general thing to assert
' about an ambiguity-tolerant renderer. Run after a normal vocabulary
' load; operates on whatever grammar is currently loaded.
Public Function EnglishRenderSelfCheck() As String
    EnsureInit
    Dim idx As Long
    Dim total As Long, rendered As Long, roundTripped As Long, usedFallback As Long, failedCount As Long
    Dim report As String
    Dim bn As Collection, bv As Collection
    For idx = 1 To mPatItems.Count
        Dim forms As Collection
        Set forms = mPatForms.Item(idx)
        If forms.Count = 1 Then
            total = total + 1
            ' G0: same "As New in a loop" hazard FindRenderRule's own
            ' comment names - a fresh pair every rule, explicitly.
            Set bn = New Collection
            Set bv = New Collection
            If BuildSyntheticBindings(idx, bn, bv) Then
                mLastRenderUsedFallback = False
                Dim sentence As String
                On Error Resume Next
                Err.Clear
                sentence = EnglishRenderForm(FormSubstitute(forms.Item(1), bn, bv))
                Dim errDesc As String
                errDesc = Err.Description
                Dim hadErr As Boolean
                hadErr = (Err.Number <> 0)
                On Error GoTo 0
                If hadErr Then
                    failedCount = failedCount + 1
                    report = report & "  [" & idx & "] '" & mPatTexts.Item(idx) & "' - render FAILED: " & errDesc & vbCrLf
                Else
                    rendered = rendered + 1
                    If mLastRenderUsedFallback Then
                        usedFallback = usedFallback + 1
                    Else
                        Dim reForm As Variant
                        If Not ReparseSentenceToForm(sentence, reForm) Then
                            failedCount = failedCount + 1
                            report = report & "  [" & idx & "] '" & mPatTexts.Item(idx) & "' -> """ & sentence & _
                                """ did not re-parse at all" & vbCrLf
                        ElseIf VLA.VlaWriteForm(reForm) <> VLA.VlaWriteForm(FormSubstitute(forms.Item(1), bn, bv)) Then
                            failedCount = failedCount + 1
                            report = report & "  [" & idx & "] '" & mPatTexts.Item(idx) & "' -> """ & sentence & _
                                """ re-parsed to a DIFFERENT form (got " & VLA.VlaWriteForm(reForm) & ")" & vbCrLf
                        Else
                            roundTripped = roundTripped + 1
                        End If
                    End If
                End If
            End If
        End If
    Next
    EnglishRenderSelfCheck = "G-RENDER self-check: " & total & " single-form rules, " & _
        rendered & " rendered, " & roundTripped & " round-tripped to an identical form, " & _
        usedFallback & " used the raw expr/cond fallback (not round-trip tested), " & _
        failedCount & " failed." & vbCrLf & report
End Function

' G1: registration-time validation. Before this, an unknown slot
' category raised at MATCH time - during the first Check of any
' sentence - which is the wrong moment and the wrong audience. A
' vocabulary's shape errors now refuse at load, where the test:
' gates already live. Checks: a non-alternation category must be a
' known one; every alternation branch must be a word (empty branches
' come from typos like {d:left|}); optional literals must be closed
' and non-empty (IsOptTok raises its own directions).
Private Sub ValidateRuleItems(ByVal pattern As String, items As Collection)
    Dim it As Variant
    Dim t As String, sn As String, ct As String, ow As String
    Dim alts() As String
    Dim ai As Long
    Dim altv As Variant
    For Each it In items
        t = it
        If IsSlotTok(t, sn, ct) Then
            If IsAltCat(ct) Then
                alts = Split(ct, "|")
                For ai = LBound(alts) To UBound(alts)
                    If Len(Trim$(alts(ai))) = 0 Then
                        VLA_Messages.RaiseMsg "english-alternation-empty-branch", "pattern", pattern, "token", t, "example", "{name:word|word}"
                    End If
                    ValidateSurfaceSpec pattern, "branch", VLA_Identity.Fold(alts(ai))
                Next
            Else
                Select Case ct
                    Case "name", "var", "text", "expr", "cond", _
                         "range", "cell", "column", "sheet", "color", "path"
                        ' known category (the five before "path" are
                        ' G2's typed reference slots - :text plus a
                        ' shape check; "path" is G-PATH's own, quoted-
                        ' or-variable, see MatchPathToken's header note)
                    Case "text-list", "range-list", "cell-list", _
                         "column-list", "sheet-list", "color-list"
                        ' G6: list-valued slots - same six shapes, one
                        ' or more comma-separated items instead of one.
                        ' This whitelist is a SEPARATE gate from
                        ' TryPhrase's own Select Case cat dispatch
                        ' (below in this file) - found live, the hard
                        ' way, the first time TestG6 actually ran: this
                        ' registration-time check (runs before a
                        ' sentence is ever matched) has to know about a
                        ' new category too, or a real rule using one
                        ' never gets past load, regardless of whether
                        ' TryPhrase itself is correct.
                    Case Else
                        VLA_Messages.RaiseMsg "english-unknown-slot-category", "pattern", pattern, "cat", ct, "example", "{name:word|word}"
                End Select
            End If
        Else
            If IsOptTok(t, ow) Then
                For Each altv In Split(ow, "|")
                    If Len(Trim$(CStr(altv))) = 0 Then
                        VLA_Messages.RaiseMsg "english-optional-empty-branch", "pattern", pattern, "token", t
                    End If
                    ValidateSurfaceSpec pattern, "optional literal", VLA_Identity.Fold(CStr(altv))
                Next
            ElseIf InStr(t, "|") > 0 Or InStr(t, "/") > 0 Then
                ' G10: bare surface tokens validate branch by branch,
                ' the same rules as braced alternations.
                For Each altv In Split(t, "|")
                    If Len(Trim$(CStr(altv))) = 0 Then
                        VLA_Messages.RaiseMsg "english-bare-alternation-empty-branch", "pattern", pattern, "token", t
                    End If
                    ValidateSurfaceSpec pattern, "branch", VLA_Identity.Fold(CStr(altv))
                Next
            End If
        End If
    Next
End Sub

' G1: optional literal spelling is [word] - one word, matched when
' present, free when absent. Returns True and fills word for a
' well-formed optional; raises with directions on a malformed one;
' False for anything not starting with "[".
Private Function IsOptTok(ByVal t As String, ByRef word As String) As Boolean
    If Left$(t, 1) <> "[" Then Exit Function
    If Right$(t, 1) <> "]" Or Len(t) < 3 Then
        VLA_Messages.RaiseMsg "english-optional-not-closed", "token", t
    End If
    word = VLA_Identity.Fold(Mid$(t, 2, Len(t) - 2))
    IsOptTok = True
End Function

' G1.1 (owner design): a branch or optional literal may carry ONE
' slash - stem/suffix - matching the stem surface OR the stem+suffix
' surface, and the STEM is what binds: {d:left|right|center/ed}
' accepts "center" and "centered" and binds "center" either way, so
' a template's xl{d} stays xlcenter whichever surface the writer
' chose. A slash can never collide with sentence text: "/" is not a
' word character in the tokenizer, so no sentence token contains one.
' G1.2 (maiden-run fix): the alternation classifier. "Contains a
' pipe" was too narrow - a SINGLE-branch surface form {d:center/ed}
' (one branch, two surfaces) has no pipe, yet is the documented
' idiom for a bare word with two spellings. A category is a surface
' slot when it carries either marker; real categories (name, var,
' text, expr, cond) contain neither, so nothing legitimate moves.
Private Function IsAltCat(ByVal cat As String) As Boolean
    IsAltCat = (InStr(cat, "|") > 0 Or InStr(cat, "/") > 0)
End Function

Private Function SurfaceMatch(ByVal spec As String, ByVal tok As String, ByRef stem As String) As Boolean
    Dim sp As Long
    sp = InStr(spec, "/")
    If sp = 0 Then
        stem = spec
        SurfaceMatch = (tok = spec)
    Else
        stem = Left$(spec, sp - 1)
        SurfaceMatch = (tok = stem) Or (tok = stem & Mid$(spec, sp + 1))
    End If
End Function

' G10: every surface of every branch of a bare spec - left|center/ed
' yields left, center, centered. Serves lint and signatures.
Private Function BareSurfaces(ByVal spec As String) As Collection
    Dim r As New Collection
    Dim alts() As String
    Dim ai As Long
    Dim f As Variant
    alts = Split(spec, "|")
    For ai = LBound(alts) To UBound(alts)
        For Each f In SurfaceForms(VLA_Identity.Fold(alts(ai)))
            r.Add CStr(f)
        Next
    Next
    Set BareSurfaces = r
End Function

' G10: match a token against a bare surface spec - branches split on
' |, each branch a word or stem/suffix - binding nothing. Serves the
' bare-alternation pattern tokens and alternation-carrying optionals.
Private Function BareAltMatch(ByVal spec As String, ByVal tok As String) As Boolean
    Dim alts() As String
    Dim ai As Long
    Dim dummy As String
    alts = Split(spec, "|")
    For ai = LBound(alts) To UBound(alts)
        If SurfaceMatch(VLA_Identity.Fold(alts(ai)), tok, dummy) Then
            BareAltMatch = True
            Exit Function
        End If
    Next
End Function

' G1.1: the surface forms a branch/optional spec can present - one
' for a plain word, two for stem/suffix. Serves the near-miss
' description, the operator-word lint, and signature expansion.
Private Function SurfaceForms(ByVal spec As String) As Collection
    Dim r As New Collection
    Dim sp As Long
    sp = InStr(spec, "/")
    If sp = 0 Then
        r.Add spec
    Else
        r.Add Left$(spec, sp - 1)
        r.Add Left$(spec, sp - 1) & Mid$(spec, sp + 1)
    End If
    Set SurfaceForms = r
End Function

' G1.1: validation shared by branches and optionals - at most one
' slash, both sides non-empty. Raises with directions.
Private Sub ValidateSurfaceSpec(ByVal pattern As String, ByVal where As String, ByVal spec As String)
    Dim first As Long
    first = InStr(spec, "/")
    If first = 0 Then Exit Sub
    If InStr(first + 1, spec, "/") > 0 Then
        VLA_Messages.RaiseMsg "english-surface-spec-multi-slash", "pattern", pattern, "where", where, "spec", spec
    End If
    If first = 1 Or first = Len(spec) Then
        VLA_Messages.RaiseMsg "english-surface-spec-empty-side", "pattern", pattern, "where", where, "spec", spec
    End If
End Sub

' G3 perf: record ONE rule as the owner of ONE exact signature string -
' idempotent (never overwrites an existing owner), which is what makes
' this safe under mAuditMode's own tolerance for an unmarked duplicate
' actually registering anyway: the FIRST rule to claim a shape keeps
' owning it in mSigOwner, exactly matching what real dispatch would
' actually do (first match wins, in registration order) - a later
' rule sharing the same shape is real, load-bearing information (it's
' the whole reason the duplicate check exists), not a bug to silently
' overwrite away.
' ---------------------------------------------------------------------
'  Registration-time linting. Warnings accumulate (they never block
'  registration); read them with EnglishLintReport.
' ---------------------------------------------------------------------
Private Sub AddSigOwner(ByVal sig As String, ByVal idx As Long)
    If Not CollHasKey(mSigOwner, sig) Then mSigOwner.Add idx, sig
End Sub

' G3 perf: rebuild mSigOwner from scratch against whatever mPatItems/
' mPatSigs currently hold. Only two callers, both cold paths where an
' O(corpus size) rebuild is the right trade for correctness over
' incremental bookkeeping: EnsureInit's own fresh-grammar case (empty,
' trivially correct either way) and RestoreRules (EnglishTryRule's
' snapshot/discard flow) - a candidate rule's AddPhraseRule call
' already added its OWN temporary entries into mSigOwner via the
' incremental path below, and RestoreRules's whole job is discarding
' that candidate, so mSigOwner has to be re-derived from the RESTORED
' table, not patched - the same reasoning RestoreRules's own
' `mDspValid = False` line already states for the dispatch index.
Private Sub RebuildSigOwner()
    Set mSigOwner = New Collection
    Dim i As Long, sigv As Variant
    For i = 1 To mPatItems.Count
        For Each sigv In ExpandedSignatures(mPatItems.Item(i))
            AddSigOwner CStr(sigv), i
        Next
    Next
End Sub

' F.4 (shape 2): a noise word ("a"/"an"/"the"/"please") is stripped from
' every PATTERN at registration (IsNoiseWord, called from AddPhraseRule)
' before this function or the matcher ever sees it. When one of those
' words sits immediately before a {...} slot in the raw pattern, the
' slot silently swallows whatever token the input leaves in the noise
' word's place instead of its intended value - confirmed live in
' espanol.vla's "envia un correo a {who:expr} ..." (the French "a"/
' Spanish "a" preposition collides with the English article list this
' engine shares across every <lingua>-vla dialect). Purely lexical: a
' scan of the raw pattern string, before noise-word stripping, needs
' no grammar reasoning.
Private Sub LintNoiseWordBeforeSlot(ByVal pattern As String)
    Dim parts() As String
    Dim i As Long
    Dim w As String, nxt As String
    parts = Split(pattern, " ")
    For i = LBound(parts) To UBound(parts) - 1
        w = VLA_Identity.Fold(Trim$(parts(i)))
        If IsNoiseWord(w) Then
            nxt = Trim$(parts(i + 1))
            If Left$(nxt, 1) = "{" Then
                mLintWarnings.Add "pattern '" & pattern & "': noise word '" & w & _
                    "' immediately precedes slot " & nxt & " - '" & w & _
                    "' is stripped from the pattern at registration, so the slot will silently swallow whatever token the input leaves in its place"
            End If
        End If
    Next
End Sub

Private Function LintRule(ByVal pattern As String, items As Collection, _
                          Optional ByVal skipIdx As Long = 0) As String
    Dim it As Variant
    Dim t As String, prevCat As String
    Dim sn As String, ct As String, ow As String
    Dim alts() As String
    Dim ai As Long
    Dim sfv As Variant
    LintNoiseWordBeforeSlot pattern
    prevCat = ""
    For Each it In items
        t = it
        If IsSlotTok(t, sn, ct) Then
            If IsAltCat(ct) Then
                ' G1: an alternation consumes a literal token, so a
                ' branch that IS an operator word after an expression
                ' slot is unreachable - same class as the literal
                ' warning below, per branch.
                If prevCat = "expr" Or prevCat = "cond" Then
                    alts = Split(ct, "|")
                    For ai = LBound(alts) To UBound(alts)
                        ' G1.1: check every surface a branch presents.
                        For Each sfv In SurfaceForms(VLA_Identity.Fold(alts(ai)))
                            If IsExprOpWord(CStr(sfv)) Or _
                               (prevCat = "cond" And (CStr(sfv) = "and" Or CStr(sfv) = "or")) Then
                                mLintWarnings.Add "pattern '" & pattern & "': alternative '" & CStr(sfv) & _
                                    "' directly follows a {:" & prevCat & "} slot - the expression will consume it and that branch can never match"
                            End If
                        Next
                    Next
                End If
                prevCat = ""
            Else
                prevCat = ct
            End If
        ElseIf InStr(t, "|") > 0 Or InStr(t, "/") > 0 Then
            ' G10: a bare surface token consumes a literal - the same
            ' per-branch, per-surface operator check as a braced
            ' alternation, and prevCat resets for the same reason.
            If prevCat = "expr" Or prevCat = "cond" Then
                For Each sfv In BareSurfaces(t)
                    If IsExprOpWord(CStr(sfv)) Or _
                       (prevCat = "cond" And (CStr(sfv) = "and" Or CStr(sfv) = "or")) Then
                        mLintWarnings.Add "pattern '" & pattern & "': alternative '" & CStr(sfv) & _
                            "' directly follows a {:" & prevCat & "} slot - the expression will consume it and that branch can never match"
                    End If
                Next
            End If
            prevCat = ""
        ElseIf IsOptTok(t, ow) Then
            ' G1: an optional operator word after an expression slot
            ' is consumed whenever present; and because it may be
            ' ABSENT, the expression stays adjacent to whatever
            ' follows - prevCat deliberately persists.
            If (prevCat = "expr" Or prevCat = "cond") And IsExprOpWord(ow) Then
                mLintWarnings.Add "pattern '" & pattern & "': optional literal '" & ow & _
                    "' directly follows a {:" & prevCat & "} slot - the expression will consume it whenever it is present"
            ElseIf prevCat = "cond" And (ow = "and" Or ow = "or") Then
                mLintWarnings.Add "pattern '" & pattern & "': optional literal '" & ow & _
                    "' directly follows a {:cond} slot - the condition will consume it whenever it is present"
            End If
        Else
            If (prevCat = "expr" Or prevCat = "cond") And IsExprOpWord(t) Then
                mLintWarnings.Add "pattern '" & pattern & "': literal '" & t & _
                    "' directly follows a {:" & prevCat & "} slot - the expression will consume it and the rule can never match"
            ElseIf prevCat = "cond" And (t = "and" Or t = "or") Then
                mLintWarnings.Add "pattern '" & pattern & "': literal '" & t & _
                    "' directly follows a {:cond} slot - the condition will consume it and the rule can never match"
            End If
            prevCat = ""
        End If
    Next

    ' Duplicate check against already-registered rules, over EXPANDED
    ' signatures (G1): every alternation branch and every optional
    ' include/omit is its own shape, so two rules that can claim the
    ' same sentence are flagged even when their spellings differ -
    ' "fix {w:cell|range} {r:text}" collides with "fix cell {r:text}"
    ' on the cell branch, and "warm cell {r:text} [up]" collides with
    ' "warm cell {r:text}" on the omitted branch. (An earlier
    ' prefix-extension warning was removed: under whole-sentence
    ' matching, a general rule registered before a longer one that
    ' extends it simply fails at its period and falls through - it
    ' cannot shadow.)
    ' G3 perf: mSigOwner (one exact signature -> its earliest owner)
    ' answers "does anything already own this shape" in O(1)-ish
    ' Collection-key time, replacing an O(already-registered rules)
    ' InStr scan per signature - was O(corpus size) per registration,
    ' O(corpus size squared) per full load, unmeasured as a problem
    ' today at ~114 rules but exactly the cost class the 10-20-more-
    ' languages plan would eventually make felt (shape 1's own
    ' AuditCrossRuleShadow regression is the concrete precedent).
    ' Semantically identical to the scan it replaces - same exact-
    ' signature-string equality, same skipIdx exclusion for an
    ' override checking against its own not-yet-replaced target - with
    ' one cosmetic difference documented where it matters less: a rule
    ' that collides with MULTIPLE different earlier rules across
    ' different branches may report them in signature order here
    ' rather than earlier-rule-index order: the same complete set of
    ' warnings reaches mLintWarnings either way, just not necessarily
    ' in the same sequence, and only the single first message actually
    ' surfaces in the raised error text.
    Dim newSigs As Collection
    Set newSigs = ExpandedSignatures(items)
    Dim e As Variant
    Dim w As String
    Dim ownerIdx As Long
    For Each e In newSigs
        If CollHasKey(mSigOwner, CStr(e)) Then
            ownerIdx = CLng(mSigOwner.Item(CStr(e)))
            If ownerIdx <> skipIdx Then
                w = "pattern '" & pattern & "' duplicates '" & mPatTexts.Item(ownerIdx) & _
                    "' (same shape: " & CStr(e) & ") - the earlier rule always wins"
                mLintWarnings.Add w
                If Len(LintRule) = 0 Then LintRule = w
            End If
        End If
    Next
End Function

' G1: all signatures a rule can present, one per alternation branch
' and optional include/omit combination - literals as themselves,
' category slots as {category}. The cartesian growth is tiny in
' practice (branches <= 4, optionals <= 2) and only runs at load.
Private Function ExpandedSignatures(items As Collection) As Collection
    Dim sigs As New Collection
    sigs.Add ""
    Dim it As Variant
    Dim t As String, sn As String, ct As String, ow As String
    Dim pieces As Collection
    Dim grown As Collection
    Dim s As Variant, pc As Variant
    Dim alts() As String
    Dim ai As Long
    Dim hasDefault As Boolean, defaultVal As String
    For Each it In items
        t = it
        Set pieces = New Collection
        If IsSlotTok(t, sn, ct, hasDefault, defaultVal) Then
            If IsAltCat(ct) Then
                alts = Split(ct, "|")
                For ai = LBound(alts) To UBound(alts)
                    ' G1.1: a stem/suffix branch is TWO shapes.
                    For Each pc In SurfaceForms(VLA_Identity.Fold(alts(ai)))
                        pieces.Add CStr(pc)
                    Next
                Next
            Else
                pieces.Add "{" & ct & "}"
            End If
            ' G7: a default makes the slot omissible - the same
            ' present/absent duality IsOptTok gets below, generalized
            ' from literals to slots.
            If hasDefault Then pieces.Add ""
        ElseIf IsOptTok(t, ow) Then
            pieces.Add ""
            For Each pc In BareSurfaces(ow)
                pieces.Add CStr(pc)
            Next
        ElseIf InStr(t, "|") > 0 Or InStr(t, "/") > 0 Then
            ' G10: shape-identical to the braced spelling, so the
            ' audit sees bare and braced twins as the duplicates
            ' they are.
            For Each pc In BareSurfaces(t)
                pieces.Add CStr(pc)
            Next
        Else
            pieces.Add t
        End If
        Set grown = New Collection
        For Each s In sigs
            For Each pc In pieces
                If Len(CStr(pc)) = 0 Then
                    grown.Add CStr(s)
                ElseIf Len(CStr(s)) = 0 Then
                    grown.Add CStr(pc)
                Else
                    grown.Add CStr(s) & "|" & CStr(pc)
                End If
            Next
        Next
        Set sigs = grown
    Next
    Set ExpandedSignatures = sigs
End Function

Private Function JoinSigs(sigs As Collection) As String
    Dim e As Variant
    Dim sb As String, sbU As Long
    For Each e In sigs
        If sbU > 0 Then SbAdd sb, sbU, vbLf
        SbAdd sb, sbU, CStr(e)
    Next
    JoinSigs = SbText(sb, sbU)
End Function

' F.4 (shape 1): one representative concrete token per slot category,
' used to turn an ExpandedSignatures shape back into a real sentence
' the actual matcher can be probed with. Chosen against the real
' validators, not guessed: a quoted string satisfies MatchRefToken
' unconditionally for every reference category (IsStrTok short-
' circuits RefShapeOk's own per-category shape check), so ONE
' placeholder covers text/range/cell/column/sheet/color and their
' -list forms; "5" bottoms out ParseExpr's recursive descent; "5 is
' empty" is ParseCondSimple's one nullary comparator, the cheapest
' valid :cond; name/var just need a colon-free word.
Private Function CategoryPlaceholderValue(ByVal cat As String) As String
    Select Case cat
        Case "name", "var"
            CategoryPlaceholderValue = "x"
        Case "expr"
            CategoryPlaceholderValue = "5"
        Case "cond"
            CategoryPlaceholderValue = "5 is empty"
        Case Else
            CategoryPlaceholderValue = """x"""
    End Select
End Function

' F.4 (shape 1): rebuild one concrete sentence from one
' ExpandedSignatures shape - literal pieces travel through as-is (an
' alternation/optional branch is already a real surface word, never a
' placeholder), a bare "{category}" piece becomes that category's
' representative token, and an omitted optional (the empty piece)
' contributes nothing. The trailing period is TryPhrase's own hard
' requirement (it will not set mLastRuleIdx without one - confirmed by
' reading the real matcher, not assumed).
Private Function SignatureToSentence(ByVal sig As String) As String
    Dim pieces() As String
    pieces = Split(sig, "|")
    Dim i As Long
    Dim piece As String
    Dim sb As String, sbU As Long
    For i = LBound(pieces) To UBound(pieces)
        piece = pieces(i)
        If Len(piece) = 0 Then
            ' an omitted optional - nothing to add
        ElseIf Left$(piece, 1) = "{" And Right$(piece, 1) = "}" Then
            If sbU > 0 Then SbAdd sb, sbU, " "
            SbAdd sb, sbU, CategoryPlaceholderValue(Mid$(piece, 2, Len(piece) - 2))
        Else
            If sbU > 0 Then SbAdd sb, sbU, " "
            SbAdd sb, sbU, piece
        End If
    Next
    SignatureToSentence = SbText(sb, sbU) & "."
End Function

' F.4 (shape 1) performance: which rules could dispatch even TRY
' before this one, for this exact shape - the real basis for whether a
' probe can possibly find a shadow, not a heuristic. Mirrors
' BuildDispatchIndex's own walk (a leading optional contributes its
' own surface as a key without stopping - "the NEXT item can also
' begin the rule" - and a bare, non-alternation slot leading the
' pattern is "universal", always a candidate) but reads it off the
' already-expanded SIGNATURE string instead of raw pattern items:
' ExpandedSignatures already turned "included"/"omitted" into separate
' concrete strings, so an empty leading piece IS the omitted case, and
' skipping forward to the next piece reproduces the real dispatcher
' exactly. Literal pieces already arrive Fold-ed (AddPhraseRule folds
' every pattern word at registration; SurfaceForms folds every
' alternation branch) - no re-folding needed here.
'
' A universal (bare-slot-leading) signature still gets a real key,
' not an empty one: the SYNTHESIZED sentence for that shape has a
' concrete first word too (the category's own CategoryPlaceholderValue
' - "5" for :expr, "x" for :name/:var), and real dispatch for THAT
' SPECIFIC sentence would also try whatever bucket that word maps to,
' not only the universal list. Skipping this would be a genuine (if
' narrow) false-negative risk: an earlier LITERAL rule whose own first
' word happened to equal a placeholder like "x" would be a real
' candidate the caller's anyUniversalSeen check alone could not see.
' Harmless for the six quoted-string categories (text/range/cell/
' column/sheet/color, +"-list"): a quoted VALUE token can never equal
' a bucket key, which is always drawn from bare pattern literals - so
' this key simply never collides there, exactly as it should.
Private Function SignatureLeadKey(ByVal sig As String, ByRef isUniversal As Boolean) As String
    Dim pieces() As String
    pieces = Split(sig, "|")
    Dim i As Long
    isUniversal = True
    For i = LBound(pieces) To UBound(pieces)
        If Len(pieces(i)) > 0 Then
            If Left$(pieces(i), 1) = "{" And Right$(pieces(i), 1) = "}" Then
                isUniversal = True
                SignatureLeadKey = CategoryPlaceholderValue(Mid$(pieces(i), 2, Len(pieces(i)) - 2))
            Else
                SignatureLeadKey = pieces(i)
                isUniversal = False
            End If
            Exit Function
        End If
    Next
    ' every piece omitted (a pattern of nothing but optionals) - the
    ' only honest answer is "could match anything," same as universal.
End Function

' F.4 (shape 1) performance: fold one rule's own signatures into the
' running "what could an EARLIER rule look like" state AuditCrossRuleShadow
' walks in ascending index order - once a key (or universal) is seen,
' every LATER rule sharing it must be probed for real; nothing yet
' seen means no earlier rule could have claimed the shape.
Private Sub RecordDispatchKeys(sigs As Collection, seenKeys As Collection, ByRef anyUniversalSeen As Boolean)
    Dim s As Variant
    Dim k As String, u As Boolean
    For Each s In sigs
        k = SignatureLeadKey(CStr(s), u)
        If u Then
            anyUniversalSeen = True
        ElseIf Not CollHasKey(seenKeys, k) Then
            seenKeys.Add True, k
        End If
    Next
End Sub

' F.4 (shape 1): does an EARLIER-registered rule's slot swallow a
' LATER rule's own literal continuation, for every shape (branch/
' optional combination) it can present? Automates EnglishTryRule's
' (G5) own one-at-a-time verdict ("an EARLIER rule matched - ... - the
' candidate never fired") for the whole corpus: synthesize a real
' sentence per signature, dispatch it through the actual matcher, and
' read mLastRuleIdx - the same field AS.1's firing counts and G5's
' verdict already trust as the real answer, not a re-implementation of
' matching semantics that could drift from it.
'
' AUDIT-TIME ONLY, deliberately not hooked into every registration:
' an early version hooked this into AddPhraseRule itself, so every
' rule paid for a dispatch probe proportional to however much of the
' corpus was already loaded - O(N) per rule, O(N^2) per full corpus
' load, paid again on every test function's own EnglishAddPhrase/
' EnglishLoadVocabulary call across the whole suite (most of which
' never audit anything). Measured live: VlaSelfTest went from ~9s to
' ~37s. F.4 was always framed as something "a human can still
' audition by hand," not a load-time gate (unlike shape 2, cheap
' enough to run everywhere) - so this now runs exactly once per
' EnglishAuditText/EnglishAuditPhrasebook call, over the FINISHED
' corpus, not once per rule as it registers.
Private Sub AuditCrossRuleShadow()
    Dim savedD As Collection, savedA As Collection
    Set savedD = mDeclared
    Set savedA = mAssigned
    Set mDeclared = New Collection
    Set mAssigned = New Collection
    Dim savedSusp As Boolean
    savedSusp = mUsageSuspended
    mUsageSuspended = True
    ' A synthesized sentence is not a real translation, so
    ' mCallNames/mCallArgs/mCallTexts/mCallLines are set aside (set to
    ' Nothing, which RecordCall's own guard - "vocab tests outside a
    ' translation" - already treats as a safe no-op) for the duration:
    ' an arbitrary synthesized sentence is far more likely than
    ' authored test prose to fall through to the bare-call-to-an-
    ' undefined-action path, and letting that pollute the real
    ' corpus's call-vs-definition check would be a new bug in service
    ' of catching an old one.
    Dim savedCN As Collection, savedCA As Collection, savedCT As Collection, savedCL As Collection
    Set savedCN = mCallNames
    Set savedCA = mCallArgs
    Set savedCT = mCallTexts
    Set savedCL = mCallLines
    Set mCallNames = Nothing
    Set mCallArgs = Nothing
    Set mCallTexts = Nothing
    Set mCallLines = Nothing

    ' F.4 (shape 1) performance: measured live at ~25-30s for a single
    ' full audit of the real ~102-rule corpus (owner's own timed
    ' VlaSelfTest runs isolated it to this one call, after the O(N^2)
    ' registration-time hook was already removed) - too slow to be
    ' worth pinning into the routine self-test even once (see
    ' TestF4RealCorpusShadow, VLA_Tests_Grammar.bas, for where that
    ' real-corpus confirmation now lives instead). seenKeys/
    ' anyUniversalSeen (RecordDispatchKeys, SignatureLeadKey) skip the
    ' expensive tokenize-and-dispatch probe whenever NO earlier rule
    ' could possibly claim this exact shape - provably correct, not a
    ' sampling shortcut, since it uses the identical ascending-index/
    ' first-key logic the real dispatcher's own BuildDispatchIndex
    ' relies on. Most of the ~102 rules have a first verb no other
    ' rule shares, so most probes are skipped outright; only the
    ' genuinely-collidable ones (like "show"/"show", "set"/"set" -
    ' the four real findings this pass turned up) pay the real cost.
    Dim idx As Long
    Dim sigv As Variant
    Dim sentence As String
    Dim toks() As String
    Dim pos As Long
    Dim seen As Collection
    Dim ruleSigs As Collection
    Dim seenKeys As New Collection
    Dim anyUniversalSeen As Boolean
    Dim leadKey As String, isUniv As Boolean
    For idx = 1 To mPatItems.Count
        Set ruleSigs = ExpandedSignatures(mPatItems.Item(idx))
        If idx > mPreludeCount Then
            Set seen = New Collection   ' dedupe identical synthesized sentences across one rule's own branches
            For Each sigv In ruleSigs
                leadKey = SignatureLeadKey(CStr(sigv), isUniv)
                If anyUniversalSeen Or CollHasKey(seenKeys, leadKey) Then
                    sentence = SignatureToSentence(CStr(sigv))
                    If Not CollHasKey(seen, sentence) Then
                        seen.Add True, sentence
                        mLastRuleIdx = 0
                        On Error Resume Next
                        toks = CanonicalizeStructuralWords(EnTokenize(sentence))   ' LX5.1
                        pos = 1
                        ParseStmt toks, pos, 0
                        On Error GoTo 0
                        If mLastRuleIdx > 0 And mLastRuleIdx <> idx Then
                            mLintWarnings.Add "pattern '" & mPatTexts.Item(idx) & "': its own shape '" & CStr(sigv) & _
                                "' (synthesized sentence: """ & sentence & """) is claimed by an earlier rule, '" & _
                                mPatTexts.Item(mLastRuleIdx) & "' [from " & RuleSourceOf(mLastRuleIdx) & _
                                "] - this rule can never fire for that shape"
                        End If
                    End If
                End If
            Next
        End If
        RecordDispatchKeys ruleSigs, seenKeys, anyUniversalSeen
    Next

    Set mDeclared = savedD
    Set mAssigned = savedA
    mUsageSuspended = savedSusp
    Set mCallNames = savedCN
    Set mCallArgs = savedCA
    Set mCallTexts = savedCT
    Set mCallLines = savedCL
End Sub

' =====================================================================
'  English tokenizer
'  Tokens: lowercased words (incl. numbers), "." "," ":", and string
'  literals tagged with a leading quote char (same trick as VLA.bas).
'  Noise words are dropped; ! and ? count as full stops; number words
'  zero..twelve become digits.
' =====================================================================

Private Function EnTokenize(ByVal text As String) As String()
    ' B5: curly double quotes (Word's autocorrect) are silently
    ' accepted as straight quotes - defense in depth below the
    ' importer's own cleanup, so pasted text behaves too.
    text = Replace(text, ChrW$(8220), """")
    text = Replace(text, ChrW$(8221), """")
    Dim outc As New Collection
    Dim outLn As New Collection
    Dim lineNo As Long
    lineNo = 1
    Dim i As Long, n As Long
    Dim c As String, w As String
    Dim nlRun As Long                ' consecutive line breaks seen;
                                     ' starts at 1 so char 1 counts as
                                     ' a line start (L0 uses this)
    Dim qEnd As Long, refEnd As Long, rc As String   ' 'Sheet Name'! scan
    nlRun = 1
    n = Len(text)
    i = 1
    Do While i <= n
        c = Mid$(text, i, 1)
        If c = vbLf Then
            nlRun = nlRun + 1
            lineNo = lineNo + 1
            ' A blank line (two breaks with only whitespace between)
            ' becomes one paragraph marker: "close any open blocks".
            If nlRun = 2 And outc.Count > 0 Then
                If outc.Item(outc.Count) <> PARA_TOK Then
                    outc.Add PARA_TOK
                    outLn.Add lineNo
                End If
            End If
            i = i + 1
        ElseIf c = vbCr Or c = " " Or c = vbTab Then
            i = i + 1
        ElseIf c = "#" Then          ' comment to end of line
            Do While i <= n
                c = Mid$(text, i, 1)
                If c = vbCr Or c = vbLf Then Exit Do
                i = i + 1
            Loop
            nlRun = 0                ' a comment line counts as content
        ElseIf c = """" Then
            nlRun = 0
            i = i + 1
            w = ""
            Do While i <= n
                c = Mid$(text, i, 1)
                If c = """" Then
                    ' B7: VBA's own doubling escape - "" inside a
                    ' literal reads as one quote, so SOP text can
                    ' quote terms: "He said ""hi""".
                    If Mid$(text, i + 1, 1) = """" Then
                        w = w & """"
                        i = i + 2
                    Else
                        i = i + 1
                        Exit Do
                    End If
                Else
                    w = w & c
                    i = i + 1
                End If
            Loop
            outc.Add Chr$(34) & w
            outLn.Add lineNo
        ElseIf c = "." Or c = "!" Or c = "?" Then
            nlRun = 0
            outc.Add "."
            outLn.Add lineNo
            i = i + 1
        ElseIf c = "," Then
            nlRun = 0
            outc.Add ","
            outLn.Add lineNo
            i = i + 1
        ElseIf c = "%" Then
            ' B6.3, owner decision: % is a real token - a postfix
            ' operator handled at the primary level (15% reads as
            ' 15 divided by 100, exactly as Excel reads it).
            nlRun = 0
            outc.Add "%"
            outLn.Add lineNo
            i = i + 1
        ElseIf c = ":" Then
            nlRun = 0
            outc.Add ":"
            outLn.Add lineNo
            i = i + 1
        ElseIf c = "(" And nlRun >= 1 Then
            ' L0.2 (this pass): a row that BEGINS with an opening
            ' parenthesis is a raw VLA form - the middle layer as a
            ' first-class citizen of the sheet. Captured as ONE token
            ' (recognizable ever after by its "(" first character),
            ' balanced across as many rows as it takes - L0's own
            ' comment (preserved in spirit below) already named this as
            ' unblocked: the one-row rule existed only because the OLD
            ' progressive-prefix Check could show false-red interior
            ' rows of a still-incomplete form, and V1 replaced that with
            ' a translate-once Check, which has no such interior state
            ' to be wrong about - a row is marked only after the WHOLE
            ' program has already translated. Balancing honors VLA's
            ' own strings (backslash escapes) and its ; comments, so a
            ' ')' inside either never miscounts. Parentheses are the
            ' terminator - no period wanted, though a trailing period or
            ' ; comment on the row is politely consumed. A '(' anywhere
            ' ELSE stays a stray character (see the hint below).
            '
            ' Narrow, deliberate exception: a double-quoted STRING that
            ' itself breaks across a row boundary still ends the row (the
            ' inner string-scanning loop below is unchanged) - spanning
            ' rows mid-string is a much rarer, odder thing to author than
            ' spanning rows between forms, and out of scope for this
            ' pass; a split string still surfaces as a normal unclosed-
            ' form/stray-token refusal rather than silently misreading.
            Dim vDepth As Long, vStart As Long, vForm As String
            vStart = lineNo
            vDepth = 0
            vForm = ""
            Do While i <= n
                c = Mid$(text, i, 1)
                If c = vbLf Then
                    ' Still open (vDepth is >0 by the time any newline
                    ' is reached - the branch's very first character
                    ' always bumps it from 0 to 1 before this check is
                    ' ever revisited). Swallow the break and keep
                    ' reading; lineNo must still advance so every token
                    ' after this one keeps reporting its true row.
                    vForm = vForm & vbLf
                    lineNo = lineNo + 1
                    i = i + 1
                ElseIf c = vbCr Then
                    i = i + 1   ' matches the outer loop's own bare-vbCr handling
                ElseIf c = """" Then
                    vForm = vForm & c
                    i = i + 1
                    Do While i <= n
                        c = Mid$(text, i, 1)
                        If c = vbCr Or c = vbLf Then Exit Do
                        If c = "\" Then
                            vForm = vForm & c & Mid$(text, i + 1, 1)
                            i = i + 2
                        Else
                            vForm = vForm & c
                            i = i + 1
                            If c = """" Then Exit Do
                        End If
                    Loop
                ElseIf c = ";" Then
                    Do While i <= n
                        c = Mid$(text, i, 1)
                        If c = vbCr Or c = vbLf Then Exit Do
                        vForm = vForm & c
                        i = i + 1
                    Loop
                ElseIf c = "(" Then
                    vDepth = vDepth + 1
                    vForm = vForm & c
                    i = i + 1
                ElseIf c = ")" Then
                    vDepth = vDepth - 1
                    vForm = vForm & c
                    i = i + 1
                    If vDepth = 0 Then Exit Do
                Else
                    vForm = vForm & c
                    i = i + 1
                End If
            Loop
            If vDepth <> 0 Then
                ' L0.2: reachable now only by running out of PROGRAM
                ' text entirely with a form still open (the row-bounded
                ' version of this message is retired - a form may
                ' legitimately span rows now), so this names the real
                ' remaining cause plainly instead of the old, now-wrong
                ' "close it on this row" advice.
                mErrLine = vStart
                VLA_Messages.RaiseMsg "english-form-never-closes", "depth", vDepth, "loc", LineSuf(vStart)
            End If
            outc.Add vForm
            outLn.Add vStart
            nlRun = 0
            ' politely consume a trailing period and/or ; comment
            Do While i <= n And (Mid$(text, i, 1) = " " Or Mid$(text, i, 1) = vbTab)
                i = i + 1
            Loop
            If Mid$(text, i, 1) = "." Then i = i + 1
            Do While i <= n And (Mid$(text, i, 1) = " " Or Mid$(text, i, 1) = vbTab)
                i = i + 1
            Loop
            If Mid$(text, i, 1) = ";" Then
                Do While i <= n
                    c = Mid$(text, i, 1)
                    If c = vbCr Or c = vbLf Then Exit Do
                    i = i + 1
                Loop
            End If
        ElseIf IsWordChar(c) Then
            nlRun = 0
            w = ""
            Do While i <= n
                c = Mid$(text, i, 1)
                If IsWordChar(c) Then
                    w = w & c
                    i = i + 1
                ElseIf c = ":" And IsWordChar(Mid$(text, i + 1, 1)) Then
                    ' A colon glued between word characters is a range
                    ' colon (A1:B10), not a block colon (which is
                    ' always followed by whitespace or end of line).
                    w = w & ":"
                    i = i + 1
                ElseIf c = "!" And IsWordChar(Mid$(text, i + 1, 1)) Then
                    ' B5.2, owner decision: a bang glued between word
                    ' characters is Excel's sheet qualifier (Data!B2),
                    ' not a full stop - a full stop is never glued to
                    ' the next word. Free-standing "!" still ends a
                    ' sentence, so "Show total!" keeps working; only
                    ' the formula-habit reference binds. Chains with
                    ' the range colon: Data!A1:B10 is one token.
                    w = w & "!"
                    i = i + 1
                ElseIf c = "." And Len(w) > 0 Then
                    ' A period glued between digits is a decimal point
                    ' (3.14159), not a full stop (which is never
                    ' digit-flanked on both sides).
                    If Right$(w, 1) Like "[0-9]" And Mid$(text, i + 1, 1) Like "[0-9]" Then
                        w = w & "."
                        i = i + 1
                    Else
                        Exit Do
                    End If
                ElseIf c = "," And Len(w) > 0 Then
                    ' A comma is a thousands separator only in the
                    ' strict shape digit,DDD(non-digit) - then it is
                    ' swallowed, so 1,000,000 reads as 1000000.
                    ' Anything looser (a clause comma, or a typo like
                    ' 1,00) stays a comma token and fails loudly
                    ' rather than silently changing a magnitude.
                    If Right$(w, 1) Like "[0-9]" _
                       And Mid$(text, i + 1, 1) Like "[0-9]" _
                       And Mid$(text, i + 2, 1) Like "[0-9]" _
                       And Mid$(text, i + 3, 1) Like "[0-9]" _
                       And Not Mid$(text, i + 4, 1) Like "[0-9]" Then
                        i = i + 1              ' swallow the separator
                    Else
                        Exit Do
                    End If
                Else
                    Exit Do
                End If
            Loop
            w = VLA_Identity.Fold(w)
            If Not IsDroppedWord(w) Then
                outc.Add NumberWord(w)
                outLn.Add lineNo
            End If
        ElseIf c = "'" Then
            ' B5.3, owner decision: Excel's own spaced-sheet syntax
            ' works bare - 'Q1 Data'!A1 - the single quotes already
            ' delimit the name, so double-quoting it again was belt
            ' over suspenders. Complete-or-refuse: the quote must
            ' close on this line AND be followed by !reference, or
            ' the apostrophe is refused exactly as before (so
            ' contractions still fail loudly, with the same hint).
            qEnd = i + 1
            Do While qEnd <= n
                rc = Mid$(text, qEnd, 1)
                If rc = "'" Then Exit Do
                If rc = vbCr Or rc = vbLf Then
                    qEnd = 0
                    Exit Do
                End If
                qEnd = qEnd + 1
            Loop
            If qEnd > n Then qEnd = 0
            If qEnd > 0 Then
                If Mid$(text, qEnd + 1, 1) = "!" And IsWordChar(Mid$(text, qEnd + 2, 1)) Then
                    refEnd = qEnd + 2
                    Do While refEnd <= n
                        rc = Mid$(text, refEnd, 1)
                        If IsWordChar(rc) Then
                            refEnd = refEnd + 1
                        ElseIf rc = ":" And IsWordChar(Mid$(text, refEnd + 1, 1)) Then
                            refEnd = refEnd + 1   ' range colon chains
                        Else
                            Exit Do
                        End If
                    Loop
                    ' String-tagged, case preserved: flows through
                    ' every :text slot and ReadTextRef unchanged,
                    ' and emits as Range("'Q1 Data'!A1") verbatim.
                    outc.Add Chr$(34) & Mid$(text, i, refEnd - i)
                    outLn.Add lineNo
                    nlRun = 0
                    i = refEnd
                Else
                    qEnd = 0
                End If
            End If
            If qEnd = 0 Then
                mErrLine = lineNo
                VLA_Messages.RaiseMsg "english-unknown-character", "char", c, "hint", StrayCharHint(c), "loc", LineSuf(lineNo)
            End If
        Else
            ' B5: an unrecognized character is refused with words,
            ' never skipped. Silent skipping ran "Set price to $5."
            ' as 5 and turned "5 + 3" into the baffling "5 3" - the
            ' dangerous branch. Formula habits get the word to use
            ' instead; the message names the character and its line.
            mErrLine = lineNo
            VLA_Messages.RaiseMsg "english-unknown-character", "char", c, "hint", StrayCharHint(c), "loc", LineSuf(lineNo)
        End If
    Loop
    ' P-TOK: materialize once into arrays for O(1) positional access -
    ' the scan above is untouched, still building via cheap sequential
    ' Collection.Add calls; only the destination changes. ReDim to
    ' size 0 (not left undimmed) even for an empty sentence, so
    ' TokAt/TokLine's own UBound checks never hit an unallocated-array
    ' error on a zero-token input.
    ' MUST be For Each, not "For k = 1 To .Count: x = .Item(k)" - a
    ' Collection's own Item(k) is itself an O(k) positional walk, so an
    ' indexed copy loop is O(n^2), the exact cost this item exists to
    ' remove (caught live on the compile side: it relocated parse's own
    ' slowdown into Tokenize instead of eliminating it). For Each uses
    ' the Collection's real enumerator - true O(n) for the whole copy.
    Dim toksArr() As String
    ReDim toksArr(1 To outc.Count)
    ReDim mTokLines(1 To outLn.Count)
    Dim k As Long, v As Variant
    k = 0
    For Each v In outc
        k = k + 1
        toksArr(k) = v
    Next v
    k = 0
    For Each v In outLn
        k = k + 1
        mTokLines(k) = v
    Next v
    EnTokenize = toksArr
End Function

Private Function IsWordChar(ByVal c As String) As Boolean
    IsWordChar = (c Like "[A-Za-z0-9_-]")
End Function

' ---------------------------------------------------------------------
'  Token access helpers (all total: safe past end of stream)
' ---------------------------------------------------------------------

' G12: "Use library "helpers.vla"." - the English entry point for
' the reader's (include ...). Surfaces: use|import library|code
' [at|from] "name". The wording adjudication (owner deferred):
' "code" joins "library" so the sentence says it grabs CODE, not
' books; "VLA" was dropped as jargon - "Import code from ..." says
' the same thing in words a non-programmer owns. The name MUST be
' quoted, and the refusal teaches why: file names contain periods,
' and a period ends a sentence - the same rule as spaced sheet
' names. Sentences collect into BOTTOM-matter wherever they appear
' (an import applies to the whole program); repeats of one file
' splice once (case-insensitive, first-mention order), because a
' double splice would define every library procedure twice and die
' at VBA's duplicate-definition compile step.
Private Sub ParseUseLibrary(toks() As String, ByRef pos As Long)
    pos = pos + 2                          ' use|import library|code
    If WordAt(toks, pos) = "at" Or WordAt(toks, pos) = "from" Then pos = pos + 1
    Dim ftok As String
    ftok = TokAt(toks, pos)
    If Not IsStrTok(ftok) Then
        VLA_Messages.RaiseMsg "english-library-name-not-quoted", "context", SentenceContext(toks, pos), "loc", LineTag(pos)
    End If
    pos = pos + 1
    Dim fname As String
    fname = Mid$(ftok, 2)
    ExpectTok toks, pos, ".", "'.' after the library name"
    Dim k As String
    k = VLA_Identity.Fold(fname)
    On Error Resume Next
    mIncludeSeen.Add k, k
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Sub                           ' seen before - splice once
    End If
    On Error GoTo 0
    mIncludeLines = mIncludeLines & "(include """ & fname & """)" & vbCrLf
End Sub

Public Function TokAt(toks() As String, ByVal p As Long) As String
    If p < 1 Or p > UBound(toks) Then Exit Function
    TokAt = toks(p)
End Function

Private Function IsStrTok(ByVal t As String) As Boolean
    IsStrTok = (Left$(t, 1) = Chr$(34))
End Function

Private Function IsNumTok(ByVal t As String) As Boolean
    If Len(t) = 0 Then Exit Function
    If IsStrTok(t) Then Exit Function
    IsNumTok = IsNumeric(t)
End Function

Public Function IsWordTok(ByVal t As String) As Boolean
    If Len(t) = 0 Then Exit Function
    IsWordTok = (Left$(t, 1) Like "[a-z]")
End Function

' Returns the word at p, or "" if the token there isn't a usable word.
Private Function WordAt(toks() As String, ByVal p As Long) As String
    Dim t As String
    t = TokAt(toks, p)
    If Not IsWordTok(t) Then Exit Function
    If t = "done" Or t = "otherwise" Then Exit Function
    WordAt = t
End Function

Private Function PeekWord(toks() As String, ByVal p As Long) As String
    Dim t As String
    t = TokAt(toks, p)
    If IsWordTok(t) Then PeekWord = t
End Function

' Match a space-separated word sequence; advances pos only on success.
Private Function MatchWords(toks() As String, ByRef pos As Long, ByVal phrase As String) As Boolean
    Dim parts() As String
    Dim i As Long
    parts = Split(phrase, " ")
    For i = LBound(parts) To UBound(parts)
        If TokAt(toks, pos + i) <> parts(i) Then Exit Function
    Next
    pos = pos + (UBound(parts) - LBound(parts) + 1)
    MatchWords = True
End Function

Private Function SentenceContext(toks() As String, ByVal pos As Long) As String
    Dim r As String, i As Long, t As String
    For i = pos To pos + 11
        t = TokAt(toks, i)
        If Len(t) = 0 Then Exit For
        If t = PARA_TOK Then Exit For
        If IsStrTok(t) Then t = """" & Mid$(t, 2) & """"
        ' Punctuation attaches to the preceding word, as written.
        If Len(r) > 0 And t <> "." And t <> "," And t <> ":" Then r = r & " "
        r = r & t
        If t = "." Then Exit For
    Next
    SentenceContext = r
End Function

Private Function ExpectWord(toks() As String, ByRef pos As Long, ByVal what As String) As String
    Dim w As String
    w = WordAt(toks, pos)
    If Len(w) = 0 Then VLA_Messages.RaiseMsg "english-expected-near", "what", what, "context", SentenceContext(toks, pos), "loc", LineTag(pos)
    If InStr(w, ":") > 0 Then VLA_Messages.RaiseMsg "english-name-has-colon", "word", w, "context", SentenceContext(toks, pos), "loc", LineTag(pos)
    pos = pos + 1
    ExpectWord = w
End Function

' The source line of the token at pos (EOF errors report the last
' line; 0 = no token info).
Private Function TokLine(ByVal pos As Long) As Long
    On Error GoTo notReady   ' mTokLines never allocated - no
                             ' EnTokenize has run yet (mirrors the old
                             ' "Is Nothing" guard for a Collection)
    Dim hi As Long
    hi = UBound(mTokLines)
    If pos >= 1 And pos <= hi Then
        TokLine = mTokLines(pos)
    ElseIf pos > hi And hi > 0 Then
        TokLine = mTokLines(hi)
    End If
    Exit Function
notReady:
End Function

' Formats " (line N)" for error messages, recording N as the last
' error line. Only ever evaluated on failure paths.
Private Function LineTag(ByVal pos As Long) As String
    mErrLine = TokLine(pos)
    LineTag = LineSuf(mErrLine)
End Function

Private Function LineSuf(ByVal n As Long) As String
    If n > 0 Then LineSuf = " (line " & n & ")"
End Function

' The line of the most recent translation error (0 = unknown).
Public Function EnglishLastErrorLine() As Long
    EnglishLastErrorLine = mErrLine
End Function

' IN.7 (button-click half): the parallel pair naming every
' 'When "<caption>" is clicked:' handler the LAST EnglishToVla call
' declared - VLA_IDE.bas's InterpretProgram reads these right after
' translating (the same moment it already checks
' VLA_Interpreter.VlaHasProc("on:sheet-change")) to register each one
' with VLA_Events, caption bound to its internal proc name.
Public Function EnglishClickHandlerNames() As Collection
    Set EnglishClickHandlerNames = mClickNames
End Function

Public Function EnglishClickHandlerProcs() As Collection
    Set EnglishClickHandlerProcs = mClickProcs
End Function

' A4: the construct that claimed the last top-level sentence parsed -
' a structural case, a value/action call, or "" when a phrase rule
' (or nothing) handled it. Read by EnglishExplain and the self-test.
Public Function EnglishLastClaim() As String
    EnglishLastClaim = mClaim
End Function

' First claim wins within a sentence: the outermost form that commits
' to a sentence owns the headline, and inner forms (a one-line If's
' body, a branch block's sentences) cannot overwrite it. Set only at
' COMMIT points - never in a case that can still fall through to the
' phrase rules, or the fall-through would carry a false claim.
Private Sub Claim(ByVal what As String)
    If Len(mClaim) = 0 Then
        mClaim = what
        BumpUsage "form: " & what   ' S2: one per top-level sentence
                                    ' (mClaim resets per sentence, and
                                    ' first claim wins, so nested
                                    ' constructs ride their sentence)
    End If
End Sub

Private Sub ExpectWordIs(toks() As String, ByRef pos As Long, ByVal word As String)
    If TokAt(toks, pos) <> word Then VLA_Messages.RaiseMsg "english-expected-word-near", "word", word, "context", SentenceContext(toks, pos), "loc", LineTag(pos)
    pos = pos + 1
End Sub

Private Sub ExpectTok(toks() As String, ByRef pos As Long, ByVal t As String, ByVal what As String)
    If TokAt(toks, pos) <> t Then VLA_Messages.RaiseMsg "english-expected-near", "what", what, "context", SentenceContext(toks, pos), "loc", LineTag(pos)
    pos = pos + 1
End Sub

' =====================================================================
'  Statement and block parsing
' =====================================================================

' Statements until the block closes: "Done." (consumed - closes just
' this block), a paragraph marker (NOT consumed - unwinds every open
' block up the stack; the top level discards it), or end of input.
Private Function ParseBlock(toks() As String, ByRef pos As Long) As Collection
    Dim outc As New Collection
    Do
        If pos > UBound(toks) Then Exit Do                 ' EOF closes blocks
        If TokAt(toks, pos) = PARA_TOK Then Exit Do      ' leave marker for outer blocks
        If TokAt(toks, pos) = "done" Then
            pos = pos + 1
            ExpectTok toks, pos, ".", "'.' after 'Done'"
            Exit Do
        End If
        outc.Add ParseTracked(toks, pos, 2)
    Loop
    Set ParseBlock = outc
End Function

' Branch body of a chained block (If/Otherwise or When/When it is):
' stops at a continuation word from contWords - "otherwise"/"when" -
' whether it appears mid-paragraph or after blank lines (the natural
' prose layout), leaving pos ON that word and reporting it in
' contWord. Also stops at "Done." (consumed, contWord = ""), at a
' paragraph marker with no continuation beyond it (marker left for
' outer blocks), or at EOF. contWords is space-delimited with spaces
' at both ends, e.g. " when otherwise ".
Private Function ParseBranchBlock(toks() As String, ByRef pos As Long, _
                                  ByVal contWords As String, ByRef contWord As String) As Collection
    Dim outc As New Collection
    contWord = ""
    Do
        If pos > UBound(toks) Then Exit Do                 ' EOF closes blocks
        If TokAt(toks, pos) = PARA_TOK Then
            ' Look past paragraph markers: a continuation word there
            ' means this break separates branches rather than ending
            ' the statement; absorb it and hand over.
            contWord = ContinuationAhead(toks, pos, contWords)
            Exit Do
        End If
        If InStr(contWords, " " & TokAt(toks, pos) & " ") > 0 Then
            contWord = TokAt(toks, pos)
            Exit Do
        End If
        If TokAt(toks, pos) = "done" Then
            pos = pos + 1
            ExpectTok toks, pos, ".", "'.' after 'Done'"
            Exit Do
        End If
        outc.Add ParseTracked(toks, pos, 2)
    Loop
    Set ParseBranchBlock = outc
End Function

' After a branch: does a continuation word follow, possibly across
' blank lines? If so, advance pos to it and return it; otherwise pos
' is untouched (the paragraph marker stays for outer blocks).
Private Function ContinuationAhead(toks() As String, ByRef pos As Long, _
                                   ByVal contWords As String) As String
    Dim j As Long
    j = pos
    Do While TokAt(toks, j) = PARA_TOK
        j = j + 1
    Loop
    Dim t As String
    t = TokAt(toks, j)
    If Len(t) > 0 And InStr(contWords, " " & t & " ") > 0 Then
        pos = j
        ContinuationAhead = t
    End If
End Function

' B3: body of a Try block. Like ParseBranchBlock, but the continuation
' is the three-word phrase "If that fails" (followed by ':') rather
' than a single word - checked at statement boundaries only, so a
' genuine If sentence inside the body is untouched. Consumes the
' phrase and its colon itself; hasRecovery reports whether one exists
' (a Try without recovery means: on failure, just skip to the end).
Private Function ParseTryBody(toks() As String, ByRef pos As Long, ByRef hasRecovery As Boolean) As Collection
    Dim outc As New Collection
    Dim j As Long
    hasRecovery = False
    Do
        If pos > UBound(toks) Then Exit Do                 ' EOF closes blocks
        If TokAt(toks, pos) = PARA_TOK Then
            j = pos
            Do While TokAt(toks, j) = PARA_TOK
                j = j + 1
            Loop
            If AtFailIntro(toks, j) Then
                pos = j                                  ' blank line separates body and recovery
            Else
                Exit Do                                  ' leave marker for outer blocks
            End If
        End If
        If AtFailIntro(toks, pos) Then
            pos = pos + 3                                ' "if" "that" "fails"
            ExpectTok toks, pos, ":", "':' after 'If that fails'"
            hasRecovery = True
            Exit Do
        End If
        If TokAt(toks, pos) = "done" Then
            pos = pos + 1
            ExpectTok toks, pos, ".", "'.' after 'Done'"
            Exit Do
        End If
        outc.Add ParseTracked(toks, pos, 2)
    Loop
    Set ParseTryBody = outc
End Function

Private Function AtFailIntro(toks() As String, ByVal p As Long) As Boolean
    AtFailIntro = (TokAt(toks, p) = "if" And TokAt(toks, p + 1) = "that" _
                   And TokAt(toks, p + 2) = "fails")
End Function

' IN.7: read-only 4-token lookahead for "When the sheet changes:" -
' "the" never reaches the token stream (IsDroppedWord), so the literal
' shape is when/sheet/changes/":" . Deliberately exact (not a prefix
' match) so no other "when ..." sentence - least of all ParseStmt's
' own "When <value> is <case>:" choices form - is ever mistaken for it.
Private Function AtSheetChangeEvent(toks() As String, ByVal p As Long) As Boolean
    AtSheetChangeEvent = (TokAt(toks, p) = "when" And TokAt(toks, p + 1) = "sheet" _
                   And TokAt(toks, p + 2) = "changes" And TokAt(toks, p + 3) = ":")
End Function

' IN.7 (button-click half): read-only 5-token lookahead for
' 'When "<caption>" is clicked:' - deliberately exact (not a prefix
' match), AtSheetChangeEvent's own reasoning: a quoted string in the
' second slot is what tells this apart from ParseStmt's 'When <value>
' is <case>:' choices form, which also reads "when"/word/"is"/.../
' ":" - the two shapes only collide if a program names a choices-form
' case literally "clicked", a contrived corner this project already
' accepts for "sheet"/"changes" just above.
Private Function AtButtonClickEvent(toks() As String, ByVal p As Long) As Boolean
    AtButtonClickEvent = (TokAt(toks, p) = "when" And IsStrTok(TokAt(toks, p + 1)) _
                   And TokAt(toks, p + 2) = "is" And TokAt(toks, p + 3) = "clicked" _
                   And TokAt(toks, p + 4) = ":")
End Function

' The VLA that re-arms the surrounding error story after a Try: the
' per-sub step handler when step tracking is on, VBA's default when
' it is off (a developer running untracked wants raw errors back).
Private Function RestoreHandlerVla() As String
    If mStepTracking Then
        RestoreHandlerVla = "(on-error goto vla-fail)"
    Else
        RestoreHandlerVla = "(on-error goto 0)"
    End If
End Function

' One or more case values joined by "or": "North" or "South" or edge.
' Full expressions are allowed (quoted text is typical); "or" cannot
' be consumed by ParseExprReq (it is a condition joiner, not an
' expression operator), so it reliably separates the values.
Private Function ParseCaseValues(toks() As String, ByRef pos As Long) As String
    Dim r As String
    r = "(" & ParseExprReq(toks, pos)
    Do While TokAt(toks, pos) = "or"
        pos = pos + 1
        r = r & " " & ParseExprReq(toks, pos)
    Loop
    ParseCaseValues = r & ")"
End Function

' L0: Check-time validation of a raw VLA row. Two gates: top-level
' definition heads are refused with directions (a row is a statement
' inside the program; a (sub ...) would emit as a garbage call), and
' the form is probe-transpiled inside a scratch Sub so the reader,
' macro expander, and emitter all get their say - their message
' arrives line-attributed to the row.
Private Sub ValidateRawVla(ByVal formText As String, ByVal pos As Long)
    Dim j As Long, hd As String, ch As String
    j = 2
    Do While j <= Len(formText)
        ch = Mid$(formText, j, 1)
        If ch = " " Or ch = vbTab Or ch = "(" Or ch = ")" Then Exit Do
        hd = hd & ch
        j = j + 1
    Loop
    Select Case VLA_Identity.Fold(hd)
        Case "sub", "function", "defmacro", "type", "enum", "public", "private"
            VLA_Messages.RaiseMsg "english-top-level-definition-in-row", "head", hd, "loc", LineTag(pos)
    End Select
    ' L12: the balance count in words BEFORE the probe - an
    ' unbalanced form reaching here gets the friendly arithmetic
    ' instead of the parser's structural message. (The row CAPTURE
    ' already refuses unbalanced rows with its own count - this
    ' covers every other door into raw validation.)
    Dim balHint As String
    balHint = VlaBalanceHint(formText)
    If Len(balHint) > 0 Then
        VLA_Messages.RaiseMsg "english-form-balance-hint", "hint", balHint, "loc", LineTag(pos)
    End If
    On Error Resume Next
    Dim junk As String
    junk = VlaTranspile("(sub vla-check-probe () " & formText & ")")
    Dim en As Long, d As String
    en = Err.Number
    d = Err.Description
    On Error GoTo 0
    If en <> 0 Then
        VLA_Messages.RaiseMsg "english-form-doesnt-transpile", "detail", d, "loc", LineTag(pos)
    End If
End Sub

Private Function ParseStmt(toks() As String, ByRef pos As Long, ByVal ind As Long) As String
    Dim pad As String
    pad = String$(ind * 2, " ")
    If TokAt(toks, pos) = PARA_TOK Then
        VLA_Messages.RaiseMsg "english-expected-sentence-blank-line"
    End If

    ' L0: a token beginning with "(" is a raw VLA form the tokenizer
    ' captured whole. It is a statement wherever a sentence can stand
    ' (top level, block bodies, action bodies); Check validates it by
    ' probe-transpiling, so a bad form goes red in the sheet instead
    ' of erupting at Run - first-class means check-before-run parity.
    ' Step tracking wraps it like any sentence, so a runtime failure
    ' still says which row.
    If Left$(TokAt(toks, pos), 1) = "(" Then
        Claim "a raw VLA form"
        Dim rawF As String
        rawF = TokAt(toks, pos)
        ValidateRawVla rawF, pos
        pos = pos + 1
        ParseStmt = pad & rawF
        Exit Function
    End If

    Dim w As String
    w = PeekWord(toks, pos)

    Dim c As String, e As String, v As String, kind As String
    Dim inner As String
    Dim thenC As Collection, elseC As Collection, body As Collection
    Dim contW As String
    Dim tn As Long
    Dim hasRec As Boolean
    Dim r As String
    Dim i As Long
    Dim txt As String

    Select Case w
        Case "if"
            If AtFailIntro(toks, pos) Then
                Claim "the Try recovery intro ('If that fails:')"
                VLA_Messages.RaiseMsg "english-if-fails-misplaced", "context", SentenceContext(toks, pos), "loc", LineTag(pos)
            End If
            Claim "the If form"
            pos = pos + 1
            c = ParseCondReq(toks, pos)
            If TokAt(toks, pos) = "," Then
                ' one-line if: "If <cond>, <sentence>."
                pos = pos + 1
                inner = ParseStmt(toks, pos, 0)
                ParseStmt = pad & "(if " & c & " (then " & inner & "))"
                Exit Function
            End If
            ExpectTok toks, pos, ":", "',' or ':' after the If condition"
            Set thenC = ParseBranchBlock(toks, pos, " otherwise ", contW)
            r = pad & "(if " & c & vbCrLf & pad & "  (then" & vbCrLf & JoinStmts(thenC) & ")"
            ' B2: "Otherwise, if <cond>:" chains (VLA elseif); a final
            ' plain "Otherwise:" closes the chain. Each branch may end
            ' at a blank line - the next Otherwise continues the same
            ' If, exactly as the two-branch form always worked.
            Do While contW = "otherwise"
                pos = pos + 1                            ' the word itself
                If TokAt(toks, pos) = "," Then pos = pos + 1
                If TokAt(toks, pos) = "if" Then
                    pos = pos + 1
                    c = ParseCondReq(toks, pos)
                    If TokAt(toks, pos) = "," Then
                        ' one-line branch: "Otherwise, if <cond>, <sentence>."
                        pos = pos + 1
                        inner = ParseTracked(toks, pos, 2)
                        contW = ContinuationAhead(toks, pos, " otherwise ")
                    Else
                        ExpectTok toks, pos, ":", "',' or ':' after the Otherwise-if condition"
                        Set thenC = ParseBranchBlock(toks, pos, " otherwise ", contW)
                        inner = JoinStmts(thenC)
                    End If
                    r = r & vbCrLf & pad & "  (elseif " & c & vbCrLf & inner & ")"
                Else
                    ExpectTok toks, pos, ":", "':' (or ', if <condition>:') after 'Otherwise'"
                    Set elseC = ParseBlock(toks, pos)
                    r = r & vbCrLf & pad & "  (else" & vbCrLf & JoinStmts(elseC) & ")"
                    contW = ""
                End If
            Loop
            ParseStmt = r & ")"
            Exit Function

        Case "repeat"
            Claim "the Repeat loop"
            pos = pos + 1
            If TokAt(toks, pos) = "until" Then
                ' B2: "Repeat until <cond>:" - keep going until it
                ' holds, tested before each pass (While with the sense
                ' flipped). No implicit counter: nothing is counted.
                ' Feeds the loop stack as "do", so "Stop the loop."
                ' emits Exit Do.
                pos = pos + 1
                c = ParseCondReq(toks, pos)
                If TokAt(toks, pos) = "," Then
                    pos = pos + 1
                    PushLoop "do"
                    inner = ParseStmt(toks, pos, 0)
                    PopLoop
                    ParseStmt = pad & "(do-until " & c & " " & inner & ")"
                    Exit Function
                End If
                ExpectTok toks, pos, ":", "',' or ':' after 'Repeat until ...'"
                PushLoop "do"
                Set body = ParseBlock(toks, pos)
                PopLoop
                ParseStmt = pad & "(do-until " & c & vbCrLf & JoinStmts(body) & ")"
                Exit Function
            End If
            e = ParseExprReq(toks, pos)
            ExpectWordIs toks, pos, "times"
            MarkAssigned "counter"
            If TokAt(toks, pos) = "," Then
                ' one-line: "Repeat <n> times, <sentence>."
                pos = pos + 1
                PushLoop "for"
                inner = ParseStmt(toks, pos, 0)
                PopLoop
                ParseStmt = pad & "(dotimes counter " & e & " " & inner & ")"
                Exit Function
            End If
            ExpectTok toks, pos, ":", "',' or ':' after 'Repeat ... times'"
            PushLoop "for"
            Set body = ParseBlock(toks, pos)
            PopLoop
            ParseStmt = pad & "(dotimes counter " & e & vbCrLf & JoinStmts(body) & ")"
            Exit Function

        Case "count"
            ' "Count <name> [down] from <a> to <b> [step <s>][:|,]" - the
            ' loop for when the variable matters: named in the sentence,
            ' any start, any direction. (Repeat keeps the implicit
            ' "counter", which always runs 1..N.)
            ' G-ROWLOOP: "step <s>" (the missing shape for skip-N row
            ' loops - "every other row," bottom-up in twos) rides this
            ' form rather than a new one - "Count row from 2 to
            ' f-last:" already covers ascending and "Count row down
            ' from f-last to 2:" already covers the safe bottom-up
            ' shape (nothing new needed there, checked directly against
            ' this Case before adding anything). A "down" step is
            ' negated in the emitted form, (- 0 s), rather than asking
            ' the user to type a negative literal - this grammar has no
            ' unary minus in words.
            Claim "the Count loop"
            pos = pos + 1
            v = ExpectWord(toks, pos, "a name after 'Count'")
            MarkAssigned v
            Dim cntDown As Boolean
            If TokAt(toks, pos) = "down" Then
                cntDown = True
                pos = pos + 1
            End If
            ExpectWordIs toks, pos, "from"
            e = ParseExprReq(toks, pos)
            ExpectWordIs toks, pos, "to"
            c = ParseExprReq(toks, pos)
            kind = ""
            If TokAt(toks, pos) = "step" Then
                pos = pos + 1
                Dim cntStep As String
                cntStep = ParseExprReq(toks, pos)
                kind = IIf(cntDown, " (- 0 " & cntStep & ")", " " & cntStep)
            ElseIf cntDown Then
                kind = " -1"
            End If
            If TokAt(toks, pos) = "," Then
                pos = pos + 1
                PushLoop "for"
                inner = ParseStmt(toks, pos, 0)
                PopLoop
                ParseStmt = pad & "(for (" & v & " " & e & " " & c & kind & ") " & inner & ")"
                Exit Function
            End If
            ExpectTok toks, pos, ":", "',' or ':' after 'Count ... from ... to ...'"
            PushLoop "for"
            Set body = ParseBlock(toks, pos)
            PopLoop
            ParseStmt = pad & "(for (" & v & " " & e & " " & c & kind & ")" & vbCrLf & JoinStmts(body) & ")"
            Exit Function

        Case "stop"
            If TokAt(toks, pos + 1) = "loop" Then
                Claim "the Stop form"
                If Len(CurrentLoop()) = 0 Then
                    VLA_Messages.RaiseMsg "english-stop-loop-outside-loop", "context", SentenceContext(toks, pos)
                End If
                pos = pos + 2
                ExpectTok toks, pos, ".", "'.' after 'Stop the loop'"
                ParseStmt = pad & IIf(CurrentLoop() = "do", "(exit-do)", "(exit-for)")
                Exit Function
            ElseIf TokAt(toks, pos + 1) = "." Then
                ' B4: inside a value-returning action this must be
                ' Exit Function - Exit Sub there is a compile error.
                Claim "the Stop form"
                pos = pos + 2
                ParseStmt = pad & IIf(mInFuncDef, "(exit-function)", "(exit-sub)")
                Exit Function
            End If
            ' anything else starting with "stop" tries the phrase rules

        Case "while"
            Claim "the While loop"
            pos = pos + 1
            c = ParseCondReq(toks, pos)
            If TokAt(toks, pos) = "," Then
                ' one-line: "While <cond>, <sentence>."
                pos = pos + 1
                PushLoop "do"
                inner = ParseStmt(toks, pos, 0)
                PopLoop
                ParseStmt = pad & "(while " & c & " " & inner & ")"
                Exit Function
            End If
            ExpectTok toks, pos, ":", "',' or ':' after the While condition"
            PushLoop "do"
            Set body = ParseBlock(toks, pos)
            PopLoop
            ParseStmt = pad & "(while " & c & vbCrLf & JoinStmts(body) & ")"
            Exit Function

        Case "for"
            Claim "the For each loop"
            pos = pos + 1
            ExpectWordIs toks, pos, "each"
            v = ExpectWord(toks, pos, "a name after 'For each'")
            MarkAssigned v
            ExpectWordIs toks, pos, "in"
            ' V3.1 (owner revision of V3): "For each pair in
            ' prices:" walks the lookup's ENTRIES - each loop value
            ' is a (key, value) pair, read with "key of pair" /
            ' "value of pair" - the baseline dictionary iteration of
            ' every mainstream language, granted on the owner's ask.
            ' V3 REFUSED the direct walk because Excel's two possible
            ' representations disagree about what it yields (a
            ' Dictionary enumerates keys, a Collection values); the
            ' revision DISSOLVES that ambiguity rather than guarding
            ' it - the English layer never emits a raw walk over the
            ' lookup object, it emits VlaDictPairs, which answers
            ' identically under both representations (insertion
            ' order; values as of the walk's start). "keys of"
            ' remains for key-only walks.
            If IsDictName(TokAt(toks, pos)) And _
               (TokAt(toks, pos + 1) = "," Or TokAt(toks, pos + 1) = ":") Then
                e = "(vladictpairs " & TokAt(toks, pos) & ")"
                pos = pos + 1
            Else
                e = ParseExprReq(toks, pos)
            End If
            If TokAt(toks, pos) = "," Then
                ' one-line: "For each <x> in <coll>, <sentence>."
                pos = pos + 1
                PushLoop "for"
                inner = ParseStmt(toks, pos, 0)
                PopLoop
                ParseStmt = pad & "(for-each (" & v & " " & e & ") " & inner & ")"
                Exit Function
            End If
            ExpectTok toks, pos, ":", "',' or ':' after 'For each ... in ...'"
            PushLoop "for"
            Set body = ParseBlock(toks, pos)
            PopLoop
            ParseStmt = pad & "(for-each (" & v & " " & e & ")" & vbCrLf & JoinStmts(body) & ")"
            Exit Function

        Case "increase", "decrease", "add"
            ' B7.4 (owner revision of B7): "Increase total by 10%."
            ' is the idiomatic spelling of growth, so a SIMPLE share -
            ' one number or name wearing % or the word percent - now
            ' MEANS Grow/Shrink, emitting their exact template. The
            ' original hazard (a % amount silently composing as 0.1
            ' into the additive rules) is guarded HARDER, not softer:
            ' any other amount carrying % or percent - arithmetic
            ' around the share, like "5 plus 5%" or "10% times 2" -
            ' refuses with directions. (B7 refused the first shape
            ' but let the second fall through to a silent additive
            ' read; B7.4 closes that.) Amounts without % fall through
            ' to the phrase rules untouched, so "Add sheet called
            ' ..." and every vocabulary rule keep working. The sniff
            ' is a pure peek: pos moves only when a share commits.
            Dim shN As String, shTgt As String, shOp As String
            Dim shScan As Long, shStop As String, shHit As Boolean
            If w = "add" Then
                ' Add <N><%|percent> to <name>.
                shN = TokAt(toks, pos + 1)
                If (TokAt(toks, pos + 2) = "%" Or TokAt(toks, pos + 2) = "percent") _
                   And TokAt(toks, pos + 3) = "to" _
                   And Len(WordAt(toks, pos + 4)) > 0 _
                   And TokAt(toks, pos + 5) = "." _
                   And (IsNumTok(shN) Or Len(WordAt(toks, pos + 1)) > 0) Then
                    shTgt = WordAt(toks, pos + 4)
                    Claim "the percent form of Increase/Decrease/Add"
                    MarkAssigned shTgt
                    pos = pos + 6
                    ParseStmt = pad & "(set! " & shTgt & " (* " & shTgt & " (+ 1 (/ " & shN & " 100))))"
                    Exit Function
                End If
                shStop = "to"
            Else
                ' Increase/Decrease <name> by <N><%|percent>.
                shTgt = WordAt(toks, pos + 1)
                shN = TokAt(toks, pos + 3)
                If Len(shTgt) > 0 And TokAt(toks, pos + 2) = "by" _
                   And (TokAt(toks, pos + 4) = "%" Or TokAt(toks, pos + 4) = "percent") _
                   And TokAt(toks, pos + 5) = "." _
                   And (IsNumTok(shN) Or Len(WordAt(toks, pos + 3)) > 0) Then
                    shOp = IIf(w = "increase", "+", "-")
                    Claim "the percent form of Increase/Decrease/Add"
                    MarkAssigned shTgt
                    pos = pos + 6
                    ParseStmt = pad & "(set! " & shTgt & " (* " & shTgt & " (" & shOp & " 1 (/ " & shN & " 100))))"
                    Exit Function
                End If
                shStop = "."
            End If
            ' Not the simple share: if % or percent appears anywhere
            ' in the AMOUNT - after the verb up to "to" for Add,
            ' after "by" up to the period for the others - refuse
            ' with directions rather than let either reading win
            ' silently. Only the amount region is scanned, so a
            ' variable merely NAMED percent stays usable as a target
            ' ("Increase percent by 5." / "Add 5 to percent.").
            shHit = False
            If w = "add" Then
                shScan = pos + 1
            ElseIf TokAt(toks, pos + 2) = "by" Then
                shScan = pos + 3
            Else
                shScan = pos + 1
            End If
            Do While shScan <= UBound(toks)
                If TokAt(toks, shScan) = "." Or TokAt(toks, shScan) = shStop _
                   Or TokAt(toks, shScan) = PARA_TOK Then Exit Do
                If TokAt(toks, shScan) = "%" Or TokAt(toks, shScan) = "percent" Then
                    shHit = True
                    Exit Do
                End If
                shScan = shScan + 1
            Loop
            If shHit Then
                Claim "the percent guard on Increase/Decrease/Add"
                VLA_Messages.RaiseMsg "english-percent-mixed-amount", "sentence", RenderSentenceAt(toks, pos), "suggestion", IIf(w = "add", "'Add step to <name>.'", "'" & StrConv(w, vbProperCase) & " <name> by step.'"), "loc", LineTag(pos)
            End If
            ' no % in the amount: hand the sentence to the rules

        Case "get"
        Case "get"
            ' B7: a bare using-call is a value, not an instruction.
            If IsUsingFn(TokAt(toks, pos + 1)) And TokAt(toks, pos + 2) = "using" Then
                Claim "the standalone-Get guard"
                VLA_Messages.RaiseMsg "english-standalone-get", "fn", TokAt(toks, pos + 1), "loc", LineTag(pos)
            End If
            ' otherwise fall through: variables and rules named get

        Case "give"
            ' B4: "Give back <value>." - the result of a value-
            ' returning action. Meaningless anywhere else, so refuse
            ' it there with directions rather than generating VBA
            ' that cannot compile.
            Claim "the Give back form"
            pos = pos + 1
            ExpectWordIs toks, pos, "back"
            If Not mInFuncDef Then
                VLA_Messages.RaiseMsg "english-give-back-outside-action", "context", SentenceContext(toks, pos - 2), "loc", LineTag(pos - 2)
            End If
            e = ParseExprReq(toks, pos)
            ExpectTok toks, pos, ".", "'.' at the end of the sentence"
            ParseStmt = pad & "(return " & e & ")"
            Exit Function

        Case "try"
            ' B3: error tolerance. The body runs under a local handler;
            ' the first failing sentence jumps - silently, that is the
            ' point - to the recovery paragraph ("If that fails:"),
            ' or straight past the block when no recovery exists.
            '
            ' The step-tracking interaction, designed deliberately:
            '   - step numbers keep advancing inside the body, so a
            '     failure AFTER the Try still reports its exact sentence;
            '   - a failure IN the body is not reported (Try's purpose);
            '   - the recovery paragraph runs under the RESTORED step
            '     handler, so a failure there is reported normally;
            '   - both exits re-arm the step handler ((on-error goto 0)
            '     when tracking is off, so untracked runs stay raw).
            ' The (resume ...) hop matters: VBA cannot trap an error
            ' raised while a handler is active, so recovery must not
            ' run inside the handler - Resume exits that state (and
            ' clears Err) before any recovery sentence executes.
            Claim "the Try block"
            pos = pos + 1
            ExpectTok toks, pos, ":", "':' after 'Try'"
            mTryCount = mTryCount + 1
            tn = mTryCount
            Set body = ParseTryBody(toks, pos, hasRec)
            If hasRec Then
                mInRecovery = True
                Set elseC = ParseBlock(toks, pos)
                mInRecovery = False
            Else
                Set elseC = New Collection
            End If
            r = pad & "(on-error goto vla-tryf-" & tn & ")" & vbCrLf
            If body.Count > 0 Then r = r & JoinStmts(body) & vbCrLf
            r = r & pad & "(goto vla-tryd-" & tn & ")" & vbCrLf
            r = r & pad & "(label vla-tryf-" & tn & ")" & vbCrLf
            r = r & pad & "(set! vla-problem err.description)" & vbCrLf
            r = r & pad & "(resume vla-tryr-" & tn & ")" & vbCrLf
            r = r & pad & "(label vla-tryr-" & tn & ")" & vbCrLf
            r = r & pad & RestoreHandlerVla() & vbCrLf
            If elseC.Count > 0 Then r = r & JoinStmts(elseC) & vbCrLf
            r = r & pad & "(label vla-tryd-" & tn & ")" & vbCrLf
            r = r & pad & RestoreHandlerVla()
            ParseStmt = r
            Exit Function

        Case "when"
            ' B2: choices. "When <value> is <case>:" opens the block
            ' and names the value being examined; following paragraphs
            ' continue it with "When it is <case>:" (several values
            ' join with "or") and an optional "Otherwise:" catches
            ' everything else. English Select Case.
            Claim "the When choices form"
            pos = pos + 1
            If TokAt(toks, pos) = "it" Then
                VLA_Messages.RaiseMsg "english-when-it-is-first", "context", SentenceContext(toks, pos - 1), "loc", LineTag(pos - 1)
            End If
            ' IN.7 (button-click half), owner-found live gap: reaching
            ' HERE with a bare (unquoted) name followed by "is clicked:"
            ' means AtButtonClickEvent's own 5-token lookahead (top of
            ' ParseDo) already tried the click-handler shape and missed
            ' ONLY on the missing quotes - IsStrTok is the one thing
            ' that lookahead checks that this generic path does not.
            ' Almost certainly a forgotten pair of quotes, not a real
            ' choices-form case that happens to be spelled "clicked" -
            ' AtButtonClickEvent's own comment already named this
            ' "contrived corner" as accepted, but never taught a way
            ' out of it: left alone, ParseExprReq below treats the bare
            ' name as a variable reference, emitting a real
            ' "Select Case <name>" VBA statement against a variable
            ' that was never declared - an untrappable "Variable not
            ' defined" VBA compile error the first time anyone actually
            ' clicked Compile (S4's own territory), not a Frazaro-level
            ' refusal at all. Caught here instead, with words, before
            ' ParseExprReq ever turns the name into an expression.
            If Not IsStrTok(TokAt(toks, pos)) And TokAt(toks, pos + 1) = "is" And _
               TokAt(toks, pos + 2) = "clicked" And TokAt(toks, pos + 3) = ":" Then
                VLA_Messages.RaiseMsg "english-click-handler-needs-quotes", "tok", TokAt(toks, pos), "loc", LineTag(pos)
            End If
            ' Sibling gap, same owner-found live shape, the other half:
            ' the name QUOTED correctly but "clicked" itself misspelled
            ' ("clocked", "cliked", ...). AtButtonClickEvent's own exact
            ' 4-token shape (quoted name/"is"/"clicked"/":") misses on
            ' word 3 only, same as the guard above misses on word 1 -
            ' and a quoted STRING LITERAL as a choices-form scrutinee is
            ' never sensible on its own (the case never varies), so this
            ' exact shape is always the click-handler idiom, never a
            ' real 'When <value> is <case>:' - left alone, ParseCaseValues
            ' below treats the misspelled word as a bare, undeclared
            ' variable reference, the same untrappable VBA "Variable not
            ' defined" compile error the guard above already heads off.
            If IsStrTok(TokAt(toks, pos)) And TokAt(toks, pos + 1) = "is" And _
               IsWordTok(TokAt(toks, pos + 2)) And TokAt(toks, pos + 2) <> "clicked" And _
               TokAt(toks, pos + 3) = ":" Then
                VLA_Messages.RaiseMsg "english-clicked-misspelled", "tok", RenderTok(TokAt(toks, pos)), "misspelled", TokAt(toks, pos + 2), "loc", LineTag(pos)
            End If
            e = ParseExprReq(toks, pos)
            ExpectWordIs toks, pos, "is"
            c = ParseCaseValues(toks, pos)
            ExpectTok toks, pos, ":", "':' after 'When ... is ...'"
            Set body = ParseBranchBlock(toks, pos, " when otherwise ", contW)
            r = pad & "(select " & e & vbCrLf & pad & "  (case " & c & vbCrLf & JoinStmts(body) & ")"
            Do While Len(contW) > 0
                pos = pos + 1                            ' "when" or "otherwise"
                If contW = "when" Then
                    ExpectWordIs toks, pos, "it"
                    ExpectWordIs toks, pos, "is"
                    c = ParseCaseValues(toks, pos)
                    ExpectTok toks, pos, ":", "':' after 'When it is ...'"
                    Set body = ParseBranchBlock(toks, pos, " when otherwise ", contW)
                    r = r & vbCrLf & pad & "  (case " & c & vbCrLf & JoinStmts(body) & ")"
                Else
                    ExpectTok toks, pos, ":", "':' after 'Otherwise'"
                    Set body = ParseBlock(toks, pos)
                    r = r & vbCrLf & pad & "  (case-else" & vbCrLf & JoinStmts(body) & ")"
                    contW = ""
                End If
            Loop
            ParseStmt = r & ")"
            Exit Function

        Case "create"
            Claim "the Create declaration"
            pos = pos + 1
            SkipArticles toks, pos
            kind = ExpectWord(toks, pos, "'number', 'text', 'value', 'list', or 'lookup' after 'Create'")
            ' IN.7 (button-click half), owner-found live gap: 'Create a
            ' button called "..." at cell ...:' is a natural sentence
            ' to reach for (it matches every other 'Create a <kind>
            ' called <name>.' sentence's own shape), but a button is
            ' never one of this handler's five real kinds, and its own
            ' shape diverges immediately after 'called' anyway (a
            ' QUOTED caption, then 'at cell <cell>', never a bare name
            ' plus a period) - the generic Case Else below would only
            ' ever be reached AFTER those two mismatches already threw
            ' a confusing 'expected a name'/'expected .' error, so the
            ' redirect belongs HERE, before either one fires, not down
            ' there. The working sentence is 'Make a button ...'
            ' (english.vla's own phrasebook rule) - a genuinely
            ' different grammar shape, not a synonym this keyword-
            ' dispatched handler could ever fall through to (phrase
            ' rules are only tried when the sentence's first word
            ' matches none of this Select Case's own labels, and
            ' "create" always does).
            If kind = "button" Then
                VLA_Messages.RaiseMsg "english-button-not-created-with-create", "loc", LineTag(pos - 1)
            End If
            ' G-PIVOT: the same collision as "button" above, hit live
            ' the same way - pareto.txt's own surface text for
            ' pivot-create literally starts with "Create a pivot table
            ' from ... called ...", which reads as this handler's own
            ' five-kind shape right up until "table" where "called"
            ' would be. Redirected here, before that mismatch fires,
            ' to english.vla's actual phrase rule: "Make a pivot table
            ' from <range> at <cell> called <name>."
            If kind = "pivot" Then
                VLA_Messages.RaiseMsg "english-pivot-not-created-with-create", "loc", LineTag(pos - 1)
            End If
            ExpectWordIs toks, pos, "called"
            v = ExpectWord(toks, pos, "a name after 'called'")
            ExpectTok toks, pos, ".", "'.' at the end of the sentence"
            MarkDeclared v
            Select Case kind
                Case "number": ParseStmt = pad & "(dim " & v & " Double)"
                Case "text": ParseStmt = pad & "(dim " & v & " String)"
                Case "value": ParseStmt = pad & "(dim " & v & ")"
                Case "list"
                    ' A list is a VBA Collection, declared and born in
                    ' one sentence - Set-vs-Let (obj-set!) resolved
                    ' here, once, so users never meet it.
                    ParseStmt = pad & "(begin (dim " & v & " Collection) (obj-set! " & v & " (new Collection)))"
                Case "lookup"
                    ' V3: keyed memory - store values at keys and
                    ' read them back by key, the "prices by part
                    ' number" shape (the store verb is V3.2's owner
                    ' revision: Remember binds names, Store fills
                    ' containers - one verb, one mechanism). Declared Object because the
                    ' representation is the VlaDict helper family's
                    ' business (Scripting.Dictionary where it exists,
                    ' a pair-Collection where it doesn't); Set-vs-Let
                    ' and that split are both resolved once, in the
                    ' helpers. Registering the name here is what
                    ' turns on the "<name> for <key>" read syntax.
                    If mDictNames Is Nothing Then Set mDictNames = New Collection
                    AddKeyed mDictNames, v
                    ParseStmt = pad & "(begin (dim " & v & " Object) (obj-set! " & v & " (vladictnew)))"
                Case Else
                    VLA_Messages.RaiseMsg "english-create-unknown-kind", "kind", kind
            End Select
            Exit Function

        Case "to"
            Claim "the To definition"
            VLA_Messages.RaiseMsg "english-to-not-top-level"

        Case "define"
            Claim "the Define declaration"
            VLA_Messages.RaiseMsg "english-define-not-top-level"

        Case "done", "otherwise"
            Claim "the block structure (Done/Otherwise placement)"
            VLA_Messages.RaiseMsg "english-unexpected-token-block", "tok", w
    End Select

    ' Phrase rules (DCG productions), in order; first match wins.
    ' Track the furthest any rule gets so failures can explain themselves.
    ' V7: instead of trying every rule, jump to the bucket keyed by the
    ' sentence's first non-article token, merged with the universal
    ' (slot-first) list in ascending rule index - first-match order
    ' preserved within and across buckets. A rule outside the merge
    ' would have failed at its first item, at exactly the article-
    ' skipped position, with its recorded description; within one gap
    ' all such failures share that position, and NoteFail keeps only
    ' the first at any new furthest position, so ONE simulated call
    ' per gap - with the gap's FIRST rule's description - reproduces
    ' the old loop's near-miss record exactly.
    mBestProgress = -1
    mBestExpect = ""
    Set mBestTieIdx = Nothing
    If (Not mDspValid) Or mDspN <> mPatItems.Count Then BuildDispatchIndex
    Dim dspStar As Long
    dspStar = pos
    SkipArticles toks, dspStar
    Dim dspB As Collection
    Set dspB = DspBucketFor("k:" & VLA_Identity.Fold(TokAt(toks, dspStar)))
    Dim dspBi As Long, dspUi As Long, dspLast As Long
    Dim dspNb As Long, dspNu As Long, dspNext As Long, dspSim As Long
    dspBi = 1
    dspUi = 1
    dspLast = 0
    Do
        dspNb = mDspN + 1
        If Not dspB Is Nothing Then
            If dspBi <= dspB.Count Then dspNb = dspB.Item(dspBi)
        End If
        dspNu = mDspN + 1
        If dspUi <= mDspUniversal.Count Then dspNu = mDspUniversal.Item(dspUi)
        dspNext = dspNb
        If dspNu < dspNext Then dspNext = dspNu
        ' Simulate the skipped rules between the last candidate and
        ' this one (or the tail, when no candidate remains).
        dspSim = dspNext - 1
        If dspSim > mDspN Then dspSim = mDspN
        If dspLast + 1 <= dspSim Then
            If Len(mDspFailDesc(dspLast + 1)) > 0 Then NoteFail dspStar, mDspFailDesc(dspLast + 1), dspLast + 1
        End If
        If dspNext > mDspN Then Exit Do
        If TryPhrase(dspNext, toks, pos, txt) Then
            ParseStmt = pad & txt
            Exit Function
        End If
        dspLast = dspNext
        If dspNext = dspNb Then dspBi = dspBi + 1 Else dspUi = dspUi + 1
    Loop

    ' Fallback: a name is a call to a defined action - bare ("Greet.")
    ' or with named arguments ("Stamp with row of 2 and value of "x".").
    ' Every call is recorded and checked against the definitions once
    ' the whole file has parsed.
    If Len(w) > 0 Then
        If TokAt(toks, pos + 1) = "." Then
            Claim "a call to the action '" & w & "'"
            RecordCall w, New Collection, RenderSentenceAt(toks, pos), TokLine(pos)
            pos = pos + 2
            ParseStmt = pad & "(" & w & ")"
            Exit Function
        ElseIf TokAt(toks, pos + 1) = "with" Then
            Claim "a call to the action '" & w & "' with arguments"
            Dim callArgs As New Collection
            Dim callTxt As String
            Dim callLine As Long
            callTxt = RenderSentenceAt(toks, pos)
            callLine = TokLine(pos)
            pos = pos + 2
            inner = ""
            Do
                v = ExpectWord(toks, pos, "a parameter name after 'with'")
                ExpectWordIs toks, pos, "of"
                e = ParseExprReq(toks, pos)
                inner = inner & " :" & v & " " & e
                callArgs.Add VLA_Identity.Fold(v)
                If TokAt(toks, pos) = "and" Then
                    pos = pos + 1
                Else
                    Exit Do
                End If
            Loop
            ExpectTok toks, pos, ".", "'.' at the end of the sentence"
            RecordCall w, callArgs, callTxt, callLine
            ParseStmt = pad & "(" & w & inner & ")"
            Exit Function
        End If
    End If

    VLA_Messages.RaiseMsg "english-parse-error", "msg", BuildParseError(toks, pos)
End Function

' V8: the string builder - rule-12 twin of the pair in VLA (see the
' full rationale there): pre-sized buffer, doubling growth, Mid$
' assignment; r = r & piece is O(n^2) in the result, this is linear.
Private Sub SbAdd(ByRef buf As String, ByRef used As Long, ByVal s As String)
    Dim n As Long
    n = Len(s)
    If n = 0 Then Exit Sub
    If used + n > Len(buf) Then
        Dim cap As Long
        cap = Len(buf)
        If cap < 64 Then cap = 64
        Do While used + n > cap
            cap = cap * 2
        Loop
        buf = buf & Space$(cap - Len(buf))
    End If
    Mid$(buf, used + 1, n) = s
    used = used + n
End Sub

Private Function SbText(ByRef buf As String, ByVal used As Long) As String
    SbText = Left$(buf, used)
End Function

Private Function JoinStmts(stmts As Collection) As String
    ' V8: builder - main's body joins here; the old loop copied the
    ' whole join once per statement.
    Dim sb As String, sbU As Long
    Dim e As Variant
    For Each e In stmts
        If sbU > 0 Then SbAdd sb, sbU, vbCrLf
        SbAdd sb, sbU, CStr(e)
    Next
    JoinStmts = SbText(sb, sbU)
End Function

' =====================================================================
'  Phrase matching (the DCG engine proper)
' =====================================================================

' G6: one reference token - quoted string, bare word, or number,
' shape-checked against cat unless cat is "text" (unconstrained). This
' is TryPhrase's own former inline "text"/"range"/"cell"/"column"/
' "sheet"/"color" case body, extracted so a list-valued slot
' ({name:cat-list}) can call it once per item in a loop instead of
' duplicating the token-shape logic. Same contract every other slot
' matcher in this file already has: on failure, p is left untouched.
Private Function MatchRefToken(ByVal cat As String, toks() As String, ByRef p As Long, ByRef val As String) As Boolean
    Dim t As String
    t = TokAt(toks, p)
    If IsStrTok(t) Then
        val = VlaStringLit(Mid$(t, 2))
        p = p + 1
        MatchRefToken = True
    ElseIf Len(WordAt(toks, p)) > 0 Or IsNumTok(t) Then
        If cat <> "text" Then
            If Not RefShapeOk(cat, t) Then Exit Function
        End If
        val = VlaStringLit(t)
        p = p + 1
        MatchRefToken = True
    End If
End Function

' G-PATH (pareto.txt SS15's own preamble: "needs a new {p:path} slot
' type (quoted, period-safe) - the G12 rule already solved exactly
' this problem for library names; reuse that decision"). Reuses G12's
' POLICY (a path must be quoted, and the refusal teaches why), not its
' MECHANISM: ParseUseLibrary is a bespoke top-level statement parser,
' fine for one special form, but SS15 has ~13 ordinary phrase-rule
' verbs that need this, so it rides the same G2 typed-ref-slot
' machinery every other reference category already uses instead of 13
' hand-written parsers.
'
' A quoted token always wins outright (same as every other category) -
' its content is opaque string data by then, so a period or backslash
' inside the quotes is no more dangerous than inside "abc" (the
' shipped password rule already proves this works). An unquoted bare
' word is accepted too, as a bound VARIABLE NAME - the six already-
' shipped file rules ({p:expr}) already lean on "Open workbook
' report-path." for a computed/reused path, and a strict quote-only
' slot would silently regress that capability for any new rule that
' adopted it.
'
' The loud refusal: EnTokenize itself (this file) already makes a
' genuine unquoted path impossible to parse correctly before this
' function ever runs - a period is NEVER glued onto a preceding
' non-digit word (a full stop is never letter-flanked), so
' "file.xlsx" unquoted is always three tokens ("file" "." "xlsx"), and
' a colon is only glued into a word when BOTH neighbors are word
' characters (IsWordChar has no "\"), so a drive letter's own colon in
' "C:\Reports\..." always arrives as its own free-standing ":" token
' immediately after the bareword. That is a reliable, low-false-
' positive signal - an ordinary sentence essentially never has a bare
' word immediately followed by a lone ":" token here - so it's caught
' with a specific, teaching refusal instead of falling through to
' whatever confusing generic mismatch the fragments would otherwise
' produce a few tokens later. Rarer unquoted shapes (a relative path,
' a UNC share) aren't specifically detected - they still fail, just
' through the ordinary near-miss reporting every other slot category
' already falls back on, not silently.
Private Function MatchPathToken(toks() As String, ByRef p As Long, ByRef val As String) As Boolean
    Dim t As String
    t = TokAt(toks, p)
    If IsStrTok(t) Then
        val = VlaStringLit(Mid$(t, 2))
        p = p + 1
        MatchPathToken = True
        Exit Function
    End If
    If Len(WordAt(toks, p)) = 0 Then Exit Function
    If TokAt(toks, p + 1) = ":" Then
        VLA_Messages.RaiseMsg "english-path-not-quoted", "context", SentenceContext(toks, p), "loc", LineTag(p)
    End If
    ' Unquoted: a VARIABLE REFERENCE (an identifier, unquoted in the
    ' emitted form), matching {:expr}'s existing "Open workbook
    ' report-path." ergonomics - deliberately NOT VlaStringLit(t) the
    ' way MatchRefToken's cell/range/sheet categories wrap a bare
    ' token, since there the bare token itself IS the reference's
    ' string-valued address/name; here it names a variable holding one.
    val = t
    p = p + 1
    MatchPathToken = True
End Function

Private Function TryPhrase(ByVal idx As Long, toks() As String, ByRef pos As Long, ByRef outText As String) As Boolean
    Dim items As Collection
    Set items = mPatItems.Item(idx)
    Dim p As Long
    p = pos
    Dim bn As New Collection      ' binding names
    Dim bv As New Collection      ' binding values (VLA text)
    Dim varNames As New Collection
    Dim it As Variant
    Dim t As String, slotName As String, cat As String
    Dim cp As Long
    Dim val As String
    Dim ok As Boolean
    Dim alts() As String
    Dim ai As Long
    Dim hit As Boolean
    Dim ow As String
    Dim osm As String
    Dim hasDefault As Boolean, defaultVal As String
    Dim matchOk As Boolean

    For Each it In items
        t = it
        If IsSlotTok(t, slotName, cat, hasDefault, defaultVal) Then
            matchOk = True
            Select Case cat
                Case "name", "var"
                    val = WordAt(toks, p)
                    If Len(val) = 0 Or InStr(val, ":") > 0 Then
                        matchOk = False
                    Else
                        p = p + 1
                        If cat = "var" Then varNames.Add val
                    End If
                Case "text", "range", "cell", "column", "sheet", "color"
                    ' A reference: quoted string, bare word, or number -
                    ' always becomes a VBA string literal. G2: the typed
                    ' categories ride the same binding, then shape-check
                    ' the BARE token - a quoted reference asserts (named
                    ' ranges, spaced sheets: the author's explicit form
                    ' is never second-guessed), and a shape failure
                    ' falls through (to a default when the slot has one,
                    ' G7; otherwise to a later rule), the teaching text
                    ' riding the I-understood frame when nothing does.
                    matchOk = MatchRefToken(cat, toks, p, val)
                Case "path"
                    ' G-PATH: quoted literal (any content - periods,
                    ' backslashes, colons are just string content once
                    ' inside quotes) OR a bound variable name (matches
                    ' {:expr}'s existing ergonomics on the six
                    ' already-shipped file rules, so a computed/reused
                    ' path stays possible - a strict quote-only slot
                    ' would silently regress that). MatchPathToken's own
                    ' header note has the full reasoning for the loud
                    ' refusal on the unquoted case.
                    matchOk = MatchPathToken(toks, p, val)
                Case "text-list", "range-list", "cell-list", "column-list", "sheet-list", "color-list"
                    ' G6: list-valued slots. One or more MatchRefToken
                    ' items (shape-checked against the category before
                    ' the "-list" suffix - {c:column-list} reuses the
                    ' exact IsColLetters check {c:column} already has),
                    ' comma-separated. Oxford comma REQUIRED, by
                    ' construction rather than by a special rejection
                    ' rule: "and" is only ever consumed as a silent
                    ' no-op immediately after an already-consumed comma
                    ' (SkipArticles's own "a"/"an" skip, same shape) -
                    ' never as a standalone separator. This is what
                    ' makes the design unambiguous with NO lookahead
                    ' against the surrounding pattern's own next
                    ' literal: "rows of A, B and columns of C" stops
                    ' the list cleanly after "B" (no comma before
                    ' "and"), while a dropped Oxford comma ("A, B and
                    ' C" meant as three items) simply produces the
                    ' ordinary "I understood '...B' - then expected X
                    ' but found 'and'" refusal every other slot failure
                    ' already gives - the missing-comma teaching moment
                    ' falls out of existing machinery, not new code.
                    ' Lowers to `(array item1 item2 ...)` - G6's own new
                    ' runtime primitive (EmitExpr/VLA.bas, EvalExpr/
                    ' VLA_Interpreter.bas), deliberately NOT a LISTOPS
                    ' primitive: this is a real value the emitted/
                    ' interpreted PROGRAM uses (a pivot's field names),
                    ' never expand-time quote-literal compiler data -
                    ' LISTOPS-STDLIB's own "different family from the
                    ' runtime accessors" line, the other direction.
                    Dim itemCat As String
                    itemCat = Left$(cat, Len(cat) - Len("-list"))
                    matchOk = MatchRefToken(itemCat, toks, p, val)
                    If matchOk Then
                        val = "(array " & val
                        Dim listTryP As Long, listItemVal As String
                        Do While TokAt(toks, p) = ","
                            listTryP = p + 1
                            If TokAt(toks, listTryP) = "and" Then listTryP = listTryP + 1
                            If Not MatchRefToken(itemCat, toks, listTryP, listItemVal) Then Exit Do
                            val = val & " " & listItemVal
                            p = listTryP
                        Loop
                        val = val & ")"
                    End If
                Case "expr"
                    val = ParseExpr(toks, p, ok)
                    If Not ok Then matchOk = False
                Case "cond"
                    val = ParseCond(toks, p, ok)
                    If Not ok Then matchOk = False
                Case Else
                    If Not IsAltCat(cat) Then
                        VLA_Messages.RaiseMsg "english-unknown-slot-category-runtime", "cat", cat
                    End If
                    ' G1: alternation - the sentence token must be one
                    ' of the branch literals, and the MATCHED literal
                    ' binds into the template ({d:left|right} with a
                    ' template writing xl{d} becomes xlleft/xlright).
                    ' Branches try in written order; articles skip
                    ' first, exactly as before any literal.
                    SkipArticles toks, p
                    alts = Split(cat, "|")
                    hit = False
                    For ai = LBound(alts) To UBound(alts)
                        ' G1.1: a branch may be stem/suffix - either
                        ' surface matches, the STEM binds.
                        If SurfaceMatch(VLA_Identity.Fold(alts(ai)), TokAt(toks, p), val) Then
                            p = p + 1
                            hit = True
                            Exit For
                        End If
                    Next
                    If Not hit Then matchOk = False
            End Select
            If Not matchOk Then
                ' G7: a default makes the slot's absence a successful
                ' match, not a failure - the pure-substitution text
                ' binds and no token is consumed (every parser above
                ' leaves p untouched on its own failure path).
                If hasDefault Then
                    val = defaultVal
                ElseIf IsAltCat(cat) Then
                    NoteFail p, AltDesc(cat), idx
                    Exit Function
                Else
                    NoteFail p, SlotDesc(cat), idx
                    Exit Function
                End If
            End If
            bn.Add slotName
            bv.Add val
        ElseIf IsOptTok(t, ow) Then
            ' G1: optional literal - consumed when present, free when
            ' absent, and it can never fail the match. G1.1: the word
            ' may be stem/suffix. G10: it may also be a bare
            ' alternation - [in|into] - any surface of any branch.
            SkipArticles toks, p
            If BareAltMatch(ow, TokAt(toks, p)) Then p = p + 1
        ElseIf InStr(t, "|") > 0 Or InStr(t, "/") > 0 Then
            ' G10 (owner design): a BARE surface token - in|into, or
            ' sort/ed - matches one branch surface and binds NOTHING.
            ' Braces now mean exactly one thing: the matched word
            ' travels into the template; bare means just match one
            ' of these. A pipe or slash can never appear in a
            ' sentence token (not word characters), so the notation
            ' costs the sentence side nothing, the G1.1 proof reused.
            SkipArticles toks, p
            If Not BareAltMatch(t, TokAt(toks, p)) Then
                NoteFail p, AltDesc(t), idx
                Exit Function
            End If
            p = p + 1
        Else
            SkipArticles toks, p
            If TokAt(toks, p) <> t Then
                NoteFail p, "'" & t & "'", idx
                Exit Function
            End If
            p = p + 1
        End If
    Next

    SkipArticles toks, p
    If TokAt(toks, p) <> "." Then
        NoteFail p, "'.' to end the sentence", idx
        Exit Function
    End If
    p = p + 1

    ' Success: substitute bindings into the template via the cached
    ' form - the only path now (F.2's dual-path era is retired; see
    ' BETA_ROADMAP.md's F.2 entry).
    For Each it In varNames
        MarkAssigned CStr(it)
    Next
    pos = p
    outText = TryFormPath(idx, bn, bv)
    mLastRuleIdx = idx
    BumpUsage "rule: " & mPatTexts.Item(idx)   ' S2: every firing counts
    TryPhrase = True
End Function

' V7: build the first-token dispatch index over the current rule
' store. Per rule, walk the pattern items exactly as TryPhrase would
' at position zero: leading optionals contribute their surfaces AND
' fall through (free when absent); the first non-optional item
' decides - a literal contributes itself, a bare or braced
' alternation contributes every surface of every branch (via
' BareSurfaces, the same expansion ExpandedSignatures rides - never
' a parallel one), and a value slot makes the rule universal (any
' token can begin it, so it is always tried for real). A pattern of
' nothing but optionals is classified universal too - conservative,
' and exact, because universal rules run through TryPhrase itself.
' For each bucketed rule the description its first item would
' NoteFail with is precomputed here, once, so dispatch can simulate
' a skipped rule's near-miss without matching it.
Private Sub BuildDispatchIndex()
    Set mDspBuckets = New Collection
    Set mDspUniversal = New Collection
    mDspN = mPatItems.Count
    If mDspN > 0 Then
        ReDim mDspFailDesc(1 To mDspN)
    End If
    Dim i As Long
    Dim items As Collection
    Dim it As Variant, f As Variant, kv As Variant
    Dim t As String, sn As String, ct As String, ow As String
    Dim keys As Collection
    Dim b As Collection
    Dim universal As Boolean, done As Boolean
    Dim desc As String
    For i = 1 To mDspN
        Set items = mPatItems.Item(i)
        Set keys = New Collection
        universal = False
        done = False
        desc = ""
        For Each it In items
            t = it
            If IsSlotTok(t, sn, ct) Then
                If IsAltCat(ct) Then
                    For Each f In BareSurfaces(ct)
                        DspAddKey keys, CStr(f)
                    Next
                    desc = AltDesc(ct)
                Else
                    universal = True
                End If
                done = True
            ElseIf IsOptTok(t, ow) Then
                ' Free when absent - keep walking: the NEXT item can
                ' also begin the rule.
                For Each f In BareSurfaces(ow)
                    DspAddKey keys, CStr(f)
                Next
            ElseIf InStr(t, "|") > 0 Or InStr(t, "/") > 0 Then
                For Each f In BareSurfaces(t)
                    DspAddKey keys, CStr(f)
                Next
                desc = AltDesc(t)
                done = True
            Else
                DspAddKey keys, t
                desc = "'" & t & "'"
                done = True
            End If
            If done Then Exit For
        Next
        If Not done Then universal = True
        If universal Then
            mDspUniversal.Add i
            mDspFailDesc(i) = ""
        Else
            mDspFailDesc(i) = desc
            For Each kv In keys
                Set b = DspBucketFor("k:" & CStr(kv))
                If b Is Nothing Then
                    Set b = New Collection
                    mDspBuckets.Add b, "k:" & CStr(kv)
                End If
                b.Add i
            Next
        End If
    Next
    mDspValid = True
End Sub

' V7: collect a distinct key - a keyed self-add is the duplicate
' guard (branches like left|left contribute one bucket entry, not
' two, so a bucket list can never carry the same index twice).
Private Sub DspAddKey(keys As Collection, ByVal k As String)
    On Error Resume Next
    keys.Add k, "k:" & k
    On Error GoTo 0
End Sub

' V7: keyed bucket lookup - Nothing when the token heads no rule.
Private Function DspBucketFor(ByVal key As String) As Collection
    On Error Resume Next
    Set DspBucketFor = mDspBuckets.Item(key)
    On Error GoTo 0
End Function

' =====================================================================
'  English expressions -> VLA expression text
'  Precedence (loosest to tightest):
'    joined with / followed by   ->  &
'    plus / minus                ->  + -
'    times / divided by / multiplied by  ->  * /
'    primary: number | "string" | name
'  Each operator level backtracks if the right operand fails, so words
'  like "times" in "Repeat 10 times:" are left for the sentence.
' =====================================================================

Private Function ParseExpr(toks() As String, ByRef pos As Long, ByRef ok As Boolean) As String
    Dim a As String, b As String
    Dim ok2 As Boolean
    Dim sp As Long
    a = ParseSum(toks, pos, ok)
    If Not ok Then Exit Function
    Do
        sp = pos
        If MatchWords(toks, pos, "joined with") Then
            ' matched
        ElseIf MatchWords(toks, pos, "followed by") Then
            ' matched
        Else
            Exit Do
        End If
        b = ParseSum(toks, pos, ok2)
        If Not ok2 Then
            pos = sp
            Exit Do
        End If
        a = "(& " & a & " " & b & ")"
    Loop
    ParseExpr = a
    ok = True
End Function

Private Function ParseSum(toks() As String, ByRef pos As Long, ByRef ok As Boolean) As String
    Dim a As String, b As String
    Dim ok2 As Boolean
    Dim sp As Long
    Dim w As String, op As String
    a = ParseProd(toks, pos, ok)
    If Not ok Then Exit Function
    Do
        w = TokAt(toks, pos)
        If w = "plus" Then
            op = "+"
        ElseIf w = "minus" Then
            op = "-"
        Else
            Exit Do
        End If
        sp = pos
        pos = pos + 1
        b = ParseProd(toks, pos, ok2)
        If Not ok2 Then
            pos = sp
            Exit Do
        End If
        a = "(" & op & " " & a & " " & b & ")"
    Loop
    ParseSum = a
    ok = True
End Function

Private Function ParseProd(toks() As String, ByRef pos As Long, ByRef ok As Boolean) As String
    Dim a As String, b As String
    Dim ok2 As Boolean
    Dim sp As Long
    Dim op As String
    a = ParsePrim(toks, pos, ok)
    If Not ok Then Exit Function
    Do
        sp = pos
        If TokAt(toks, pos) = "times" Then
            op = "*"
            pos = pos + 1
        ElseIf MatchWords(toks, pos, "divided by") Then
            op = "/"
        ElseIf MatchWords(toks, pos, "multiplied by") Then
            op = "*"
        Else
            Exit Do
        End If
        b = ParsePrim(toks, pos, ok2)
        If Not ok2 Then
            pos = sp
            Exit Do
        End If
        a = "(" & op & " " & a & " " & b & ")"
    Loop
    ParseProd = a
    ok = True
End Function

' B6.3: the percent postfix. Any primary followed by % is that value
' divided by 100 - Excel's own reading (15% is 0.15; A1% works in
' formulas too). Handled here, at the primary level, so it binds
' tighter than times/divided (2 times 50% is 1) and composes on
' names and numbers alike; each extra % divides again, as in Excel.
Private Function ParsePrim(toks() As String, ByRef pos As Long, ByRef ok As Boolean) As String
    ParsePrim = ParsePrimCore(toks, pos, ok)
    If Not ok Then Exit Function
    Do While TokAt(toks, pos) = "%" Or TokAt(toks, pos) = "percent"
        ' B7.2: the word "percent" is a synonym for the % postfix,
        ' so dictated sentences work: "10 percent" is 10%.
        pos = pos + 1
        ParsePrim = "(/ " & ParsePrim & " 100)"
    Loop
End Function

Private Function ParsePrimCore(toks() As String, ByRef pos As Long, ByRef ok As Boolean) As String
    Dim t As String
    Dim sp As Long
    Dim cref As String
    Dim okc As Boolean
    Dim n2 As String
    Dim ok2 As Boolean
    t = TokAt(toks, pos)
    If IsNumTok(t) Then
        pos = pos + 1
        ParsePrimCore = t
        ok = True
    ElseIf IsStrTok(t) Then
        pos = pos + 1
        ParsePrimCore = VlaStringLit(Mid$(t, 2))
        ok = True
    ElseIf IsFnWord(mFnNullary, t) Then
        pos = pos + 1
        ParsePrimCore = "(" & FnTarget(mFnNullary, t) & ")"
        ok = True
    ElseIf IsFnWord(mFnOf, t) And TokAt(toks, pos + 1) = "of" Then
        ' "<word> of <value>": binds tightly and composes -
        ' "length of cell B2", "month of today", "day of now".
        sp = pos
        pos = pos + 2
        n2 = ParsePrim(toks, pos, ok2)
        If ok2 Then
            ParsePrimCore = "(" & FnTarget(mFnOf, t) & " " & n2 & ")"
            ok = True
        Else
            pos = sp + 1             ' no operand: the word is a name
            ParsePrimCore = t
            ok = True
        End If
    ElseIf Len(OrdinalWord(t)) > 0 Then
        ' G8: ordinal words, checked only here in the expr grammar,
        ' never at the tokenizer - and only reached once the "<word> of
        ' <value>" check above has already failed. "first" is claimed
        ' by mFnOf's list-accessor idiom (first of found-items ->
        ' vlafirst) ONLY when followed by "of"; that check runs first
        ' and wins whenever "of" is actually there, so "first of X"
        ' never reaches this branch. A shipped pattern-literal rule
        ' ("first {n} letters of") is unaffected regardless - pattern
        ' literals are matched by TryPhrase directly and never pass
        ' through this expr grammar at all. A blanket tokenizer rewrite
        ' the way cardinals get one would have broken both (every
        ' "first" everywhere, unconditionally); scoping the lookup to
        ' this one call site is what makes "Delete the third row." (and
        ' any other {n:expr}-slotted rule) get transparent ordinal
        ' support without touching either.
        pos = pos + 1
        ParsePrimCore = OrdinalWord(t)
        ok = True
    ElseIf (IsUsingFn(t) And TokAt(toks, pos + 1) = "using") _
        Or (t = "get" And IsUsingFn(TokAt(toks, pos + 1)) And TokAt(toks, pos + 2) = "using") Then
        ' B6: a using-style value action call - named arguments, in
        ' any order, omitted ones taking their defaults via VBA named
        ' arguments. "get" is optional sugar. The recorded and-
        ' collision resolves by PARAMETER LOOKAHEAD: after "and", a
        ' known parameter name followed by "of" continues the
        ' argument list; anything else hands the "and" back to the
        ' condition. Unknown and missing-required parameters are
        ' refused here, at Check, with the parameter list named -
        ' before VBA could ever object obscurely.
        If t = "get" Then
            pos = pos + 1
            t = TokAt(toks, pos)
        End If
        Dim fnPos As Long
        fnPos = pos
        Claim "a value call: '" & t & " using ...'"
        pos = pos + 2                        ' the name and "using"
        Dim pNames As Collection, pReq As Collection
        Dim fOk As Boolean
        Set pNames = CollGet(mFnUsingNames, VLA_Identity.Fold(t), fOk)
        Set pReq = CollGet(mFnUsingReq, VLA_Identity.Fold(t), fOk)
        Dim kwArgs As String
        Dim seen As New Collection
        Dim aName As String
        Do
            aName = ExpectWord(toks, pos, "a parameter name after 'using'")
            If Not InParamList(pNames, aName) Then
                VLA_Messages.RaiseMsg "english-action-unknown-param", "action", t, "param", aName, "list", JoinParamNames(pNames), "loc", LineTag(pos - 1)
            End If
            If CollHasKey(seen, VLA_Identity.Fold(aName)) Then
                VLA_Messages.RaiseMsg "english-param-given-twice", "param", aName, "action", t, "loc", LineTag(pos - 1)
            End If
            seen.Add "1", VLA_Identity.Fold(aName)
            ExpectWordIs toks, pos, "of"
            kwArgs = kwArgs & " :" & aName & " " & ParseExprReq(toks, pos)
            ' Parameter lookahead: does this "and" continue the call?
            If TokAt(toks, pos) = "and" _
               And InParamList(pNames, TokAt(toks, pos + 1)) _
               And TokAt(toks, pos + 2) = "of" Then
                pos = pos + 1
            Else
                Exit Do
            End If
        Loop
        ' Every required parameter must have been supplied.
        Dim ri As Long
        For ri = 1 To pNames.Count
            If pReq.Item(ri) = "1" And Not CollHasKey(seen, CStr(pNames.Item(ri))) Then
                VLA_Messages.RaiseMsg "english-call-missing-required-param", "action", t, "param", pNames.Item(ri), "list", JoinParamNames(pNames), "loc", LineTag(fnPos)
            End If
        Next
        ParsePrimCore = "(" & t & kwArgs & ")"
        ok = True
        Exit Function

    ElseIf t = "problem" And mInRecovery Then
        ' B7: inside an "If that fails:" paragraph, "the problem" is
        ' what went wrong ("the" is a noise word, so the token is just
        ' "problem"). Elsewhere, problem stays an ordinary name.
        pos = pos + 1
        ParsePrimCore = "vla-problem"
        ok = True
        Exit Function
    ElseIf t = "item" Then
        ' B7: "item <n> of <list>" fetches an element (lists and
        ' ranges alike, via VlaItem). Falls back to the plain name
        ' when the shape does not follow, so variables named item
        ' keep working.
        sp = pos
        pos = pos + 1
        n2 = ParseExpr(toks, pos, ok2)
        If ok2 And TokAt(toks, pos) = "of" Then
            pos = pos + 1
            cref = ParsePrim(toks, pos, okc)
            If okc Then
                ParsePrimCore = "(vlaitem " & cref & " " & n2 & ")"
                ok = True
                Exit Function
            End If
        End If
        pos = sp + 1
        ParsePrimCore = t
        ok = True
        Exit Function
    ElseIf t = "cell" Or (t = "value" And TokAt(toks, pos + 1) = "in") Then
        ' Cell references are values: "cell B2" and
        ' "cell in column C row k" work anywhere an expression does.
        ' B5.2: "value in column ..." / "value in cell B2" are the
        ' honest synonyms - 5 is the VALUE in the cell, not the cell -
        ' so conditions read as written. Cell = the place (Put/Clear
        ' keep it); value = what is in it. The word "value" only takes
        ' this reading when "in" follows, so variables named value are
        ' untouched.
        sp = pos
        pos = pos + 1
        If TokAt(toks, pos) = "in" Then
            pos = pos + 1
            If t = "value" And TokAt(toks, pos) = "cell" Then
                pos = pos + 1
                cref = ReadTextRef(toks, pos, okc)
                If okc Then
                    ParsePrimCore = "(range " & cref & ")"
                    ok = True
                    Exit Function
                End If
            ElseIf TokAt(toks, pos) = "column" Then
                pos = pos + 1
                If TokAt(toks, pos) = "number" Then
                    ' B5: "cell in column number <expr> row <expr>" -
                    ' the column chosen by the program, as everywhere
                    ' else "column number" already worked in sentence
                    ' forms; this branch makes it a VALUE too, so it
                    ' composes in conditions and arithmetic.
                    pos = pos + 1
                    cref = ParseExpr(toks, pos, okc)
                Else
                    cref = ReadTextRef(toks, pos, okc)
                End If
                If okc Then
                    If TokAt(toks, pos) = "row" Then
                        pos = pos + 1
                        n2 = ParseExpr(toks, pos, ok2)
                        If ok2 Then
                            ParsePrimCore = "(cells " & n2 & " " & cref & ")"
                            ok = True
                            Exit Function
                        End If
                    End If
                End If
            End If
        ElseIf t = "cell" Then
            cref = ReadTextRef(toks, pos, okc)
            If okc Then
                ParsePrimCore = "(range " & cref & ")"
                ok = True
                Exit Function
            End If
        End If
        pos = sp + 1                 ' no reference followed: the word is a name
        ParsePrimCore = t
        ok = True
    ElseIf Len(WordAt(toks, pos)) > 0 Then
        If InStr(t, ":") > 0 Then
            ok = False               ' bare ranges belong in reference slots
        ElseIf IsDictName(t) And TokAt(toks, pos + 1) = "for" Then
            ' V3: the keyed read - "<lookup> for <key>" is a value
            ' anywhere an expression goes, conditions included:
            ' "Set p to prices for part-code.", "If prices for code
            ' is greater than 5, ...". Scoped HARD to names declared
            ' by "Create a lookup called ..." in this program, so
            ' "for" stays inert in every other expression and no
            ' vocabulary rule's greediness changes. The key binds
            ' tightly (a primary), like "length of" - compound keys
            ' get computed into a name first - so keyed reads chain
            ' in arithmetic: "prices for a plus prices for b" is a
            ' sum of two reads. A miss is a loud step error naming
            ' the key (VlaDictGet), the form-79 lookup precedent.
            sp = pos
            pos = pos + 2
            n2 = ParsePrim(toks, pos, ok2)
            If ok2 Then
                ParsePrimCore = "(vladictget " & t & " " & n2 & ")"
                ok = True
            Else
                pos = sp + 1         ' no key followed: the word is a name
                ParsePrimCore = t
                ok = True
            End If
        Else
            pos = pos + 1
            ParsePrimCore = t
            ok = True
        End If
    Else
        ok = False
    End If
End Function

' V3: is this word a lookup declared by the current translation?
' Total on any state (Nothing-safe) - direct ParseStmt callers
' (vocab tests, Explain) run without EnglishToVla's init.
Private Function IsDictName(ByVal w As String) As Boolean
    If mDictNames Is Nothing Then Exit Function
    IsDictName = CollHasKey(mDictNames, w)
End Function

' One reference token (quoted, bare word, or number) as a VBA string
' literal - the :text slot's acceptance rule, reusable in expressions.
Private Function ReadTextRef(toks() As String, ByRef pos As Long, ByRef ok As Boolean) As String
    Dim t As String
    t = TokAt(toks, pos)
    If IsStrTok(t) Then
        ReadTextRef = VlaStringLit(Mid$(t, 2))
        pos = pos + 1
        ok = True
    ElseIf IsNumTok(t) Or IsWordTok(t) Then
        ReadTextRef = VlaStringLit(t)
        pos = pos + 1
        ok = True
    Else
        ok = False
    End If
End Function

Private Sub PushLoop(ByVal kind As String)
    If mLoopStack Is Nothing Then Set mLoopStack = New Collection
    mLoopStack.Add kind
End Sub

Private Sub PopLoop()
    If mLoopStack Is Nothing Then Exit Sub
    If mLoopStack.Count > 0 Then mLoopStack.Remove mLoopStack.Count
End Sub

Private Function CurrentLoop() As String
    If mLoopStack Is Nothing Then Exit Function
    If mLoopStack.Count = 0 Then Exit Function
    CurrentLoop = mLoopStack.Item(mLoopStack.Count)
End Function

Private Function ParseExprReq(toks() As String, ByRef pos As Long) As String
    Dim ok As Boolean
    ParseExprReq = ParseExpr(toks, pos, ok)
    If Not ok Then VLA_Messages.RaiseMsg "english-expected-value", "context", SentenceContext(toks, pos)
End Function

' =====================================================================
'  English conditions -> VLA condition text
' =====================================================================

Private Function ParseCond(toks() As String, ByRef pos As Long, ByRef ok As Boolean) As String
    Dim c As String, c2 As String
    Dim ok2 As Boolean
    Dim sp As Long
    Dim w As String
    c = ParseCondSimple(toks, pos, ok)
    If Not ok Then Exit Function
    Do
        w = TokAt(toks, pos)
        If w <> "and" And w <> "or" Then Exit Do
        sp = pos
        pos = pos + 1
        c2 = ParseCondSimple(toks, pos, ok2)
        If Not ok2 Then
            pos = sp
            Exit Do
        End If
        c = "(" & w & " " & c & " " & c2 & ")"
    Loop
    ParseCond = c
    ok = True
End Function

Private Function ParseCondSimple(toks() As String, ByRef pos As Long, ByRef ok As Boolean) As String
    Dim l As String, r As String
    Dim op As String
    Dim divisible As Boolean
    Dim kind As String
    l = ParseExpr(toks, pos, ok)
    If Not ok Then Exit Function
    ok = False
    divisible = False
    kind = ""

    ' DF1 (dogfood pass): the five idiomatic conditions read through
    ' the P-layer's predicates instead of hand-assembled comparisons -
    ' the generated VLA now says (blank? x), (zero? (mod x n)),
    ' (positive? (instr ...)) where it used to spell the plumbing.
    ' Expansion makes divisible/contains/does-not-contain emit
    ' byte-identical VBA; is-empty/is-not-empty emit the equivalent
    ' len-trim test (blank?'s own spelling), the one deliberate VBA
    ' change of the pass - predicted in the goldens, proven by the
    ' vocabulary's rewritten-in-lockstep test: lines.
    ' Nullary comparators first: no right-hand value follows.
    If MatchWords(toks, pos, "is not empty") Then
        ParseCondSimple = "(not (blank? " & l & "))"
        ok = True
        Exit Function
    End If
    If MatchWords(toks, pos, "is empty") Then
        ParseCondSimple = "(blank? " & l & ")"
        ok = True
        Exit Function
    End If

    ' Longest comparator phrases first.
    If MatchWords(toks, pos, "is greater than or equal to") Then
        op = ">="
    ElseIf MatchWords(toks, pos, "is less than or equal to") Then
        op = "<="
    ElseIf MatchWords(toks, pos, "is at least") Then
        op = ">="
    ElseIf MatchWords(toks, pos, "is at most") Then
        op = "<="
    ElseIf MatchWords(toks, pos, "is greater than") Then
        op = ">"
    ElseIf MatchWords(toks, pos, "is more than") Then
        op = ">"
    ElseIf MatchWords(toks, pos, "is less than") Then
        op = "<"
    ElseIf MatchWords(toks, pos, "does not contain") Then
        kind = "ncontains"
    ElseIf MatchWords(toks, pos, "contains") Then
        kind = "contains"
    ElseIf MatchWords(toks, pos, "starts with") Then
        kind = "starts"
    ElseIf MatchWords(toks, pos, "ends with") Then
        kind = "ends"
    ElseIf MatchWords(toks, pos, "is divisible by") Then
        divisible = True
    ElseIf MatchWords(toks, pos, "is not equal to") Then
        op = "<>"
    ElseIf MatchWords(toks, pos, "does not equal") Then
        op = "<>"
    ElseIf MatchWords(toks, pos, "is not") Then
        op = "<>"
    ElseIf MatchWords(toks, pos, "is equal to") Then
        op = "="
    ElseIf MatchWords(toks, pos, "equals") Then
        op = "="
    ElseIf MatchWords(toks, pos, "is") Then
        op = "="
    Else
        Exit Function
    End If

    r = ParseExpr(toks, pos, ok)
    If Not ok Then Exit Function
    If divisible Then
        ParseCondSimple = "(zero? (mod " & l & " " & r & "))"
    ElseIf kind = "contains" Then
        ParseCondSimple = "(positive? (instr 1 " & l & " " & r & " vbtextcompare))"
    ElseIf kind = "ncontains" Then
        ParseCondSimple = "(zero? (instr 1 " & l & " " & r & " vbtextcompare))"
    ElseIf kind = "starts" Then
        ParseCondSimple = "(= (instr 1 " & l & " " & r & " vbtextcompare) 1)"
    ElseIf kind = "ends" Then
        ParseCondSimple = "(= (lcase (right " & l & " (len " & r & "))) (lcase " & r & "))"
    Else
        ParseCondSimple = "(" & op & " " & l & " " & r & ")"
    End If
    ok = True
End Function

Private Function ParseCondReq(toks() As String, ByRef pos As Long) As String
    Dim ok As Boolean
    ParseCondReq = ParseCond(toks, pos, ok)
    If Not ok Then VLA_Messages.RaiseMsg "english-expected-condition", "context", SentenceContext(toks, pos)
End Function

' =====================================================================
'  Sub assembly and bookkeeping
' =====================================================================

Private Function BuildSub(ByVal name As String, stmts As Collection, _
                          Optional ByVal paramSpec As String = "", _
                          Optional ByVal asFunc As Boolean = False) As String
    ' V8: builder - the per-statement loop below grows with the
    ' procedure (main IS the program for a typical script).
    Dim sb As String, sbU As Long
    Dim e As Variant
    Dim instrument As Boolean
    instrument = mStepTracking And stmts.Count > 0
    SbAdd sb, sbU, "(" & IIf(asFunc, "function", "sub") & " " & name & " (" & paramSpec & ")" & vbCrLf
    ' Auto-declare anything assigned but never Created. A Create or a
    ' parameter may shadow an alias (legal VBA); assigning a bare
    ' alias would be a compile error, so refuse it here with words.
    For Each e In mAssigned
        If CollHasKey(mDeclared, CStr(e)) Then
            ' declared locally (Create / parameter) - nothing to do
        ElseIf CollHasKey(mAliasNames, CStr(e)) Then
            ' V1: name the assigning sentence's line. This refusal is
            ' raised during sub assembly, AFTER parsing - LineTag's
            ' token position is long gone - so the line was recorded
            ' at MarkAssigned time instead. Without it, the IDE's
            ' translate-once Check (which retired the prefix loop)
            ' could only attribute this error to the last row.
            Dim aFound As Boolean
            Dim aLine As Variant
            aLine = CollGet(mAssignedLines, CStr(e), aFound)
            If aFound Then mErrLine = CLng(aLine)
            VLA_Messages.RaiseMsg "english-define-value-immutable", "name", e, "loc", IIf(aFound, LineSuf(CLng(aLine)), "")
        Else
            SbAdd sb, sbU, "  (dim " & e & ")" & vbCrLf
        End If
    Next
    If instrument Then SbAdd sb, sbU, "  (on-error goto vla-fail)" & vbCrLf
    For Each e In stmts
        SbAdd sb, sbU, CStr(e) & vbCrLf
    Next
    If instrument Then
        SbAdd sb, sbU, "  " & IIf(asFunc, "(exit-function)", "(exit-sub)") & vbCrLf
        SbAdd sb, sbU, "  (label vla-fail)" & vbCrLf
        SbAdd sb, sbU, "  (vla-report-error)" & vbCrLf
    End If
    ' S1.2 (owner aesthetic revision): the procedure closer stacks on
    ' the last body line, traditional Lisp style - own-line closers
    ' read as Java, the goldens' first audit finding.
    BuildSub = StackCloser(SbText(sb, sbU), ")") & vbCrLf
End Function

' S1.2: append a closer to the LAST line of already-built text
' (strip trailing line breaks, then append). Safe by construction:
' emitted VLA carries no ';' comments, and captured raw rows cannot
' retain one (interior comments force the doesn't-close refusal;
' trailing ones are never captured) - so a stacked closer can never
' be swallowed by a comment.
Private Function StackCloser(ByVal built As String, ByVal closer As String) As String
    Do While Len(built) > 0
        Dim ch As String
        ch = Right$(built, 1)
        If ch <> vbCr And ch <> vbLf Then Exit Do
        built = Left$(built, Len(built) - 1)
    Loop
    StackCloser = built & closer
End Function

Private Sub MarkAssigned(ByVal name As String)
    CheckName name
    AddKeyed mAssigned, name
    ' V1: remember WHERE (first assignment wins - the duplicate key
    ' is silently ignored, which is exactly the semantics wanted).
    ' Nothing-guarded: direct ParseStmt callers (vocab tests, comma
    ' bodies mid-test) can arrive before any EnglishToVla ran.
    If Not mAssignedLines Is Nothing Then
        If mCurLine > 0 Then
            On Error Resume Next
            mAssignedLines.Add CStr(mCurLine), VLA_Identity.Fold(name)
            On Error GoTo 0
        End If
    End If
End Sub

Private Sub MarkDeclared(ByVal name As String)
    CheckName name
    AddKeyed mDeclared, name
End Sub

' VBA reserves these words at the language level: a variable, action,
' or alias with one of these names would be a compile error - the
' class of failure no runtime handler can catch. Refuse in English
' instead. Hyphenated names always dodge the problem (seek-row is
' fine: it becomes seek_row).
Private Function IsReservedName(ByVal n As String) As Boolean
    Dim list As String
    list = " and as boolean byref byval byte call case close const currency date declare dim do double each else elseif empty end enum eqv erase error event exit false for friend function get goto if imp implements in input integer is let like lock long loop lset me mod new next not nothing null object on open option optional or paramarray preserve print private property public put raiseevent redim rem resume rset seek select set single static stop string sub then time to true type typeof unlock until variant wend while with withevents write xor "
    IsReservedName = (InStr(list, " " & VLA_Identity.Fold(n) & " ") > 0)
End Function

Private Sub CheckName(ByVal n As String)
    If IsReservedName(n) Then
        VLA_Messages.RaiseMsg "english-reserved-word-name", "name", n
    End If
    If Not mFnNullary Is Nothing Then
        Dim f As Boolean, v As Variant
        v = CollGet(mFnNullary, n, f)
        If f Then
            VLA_Messages.RaiseMsg "english-value-word-name", "name", n
        End If
    End If
End Sub

' B4: one name, one definition. A second 'To' with a used name - sub
' or value-returning - would become two VBA procedures and crash the
' user into a compile error no handler can catch; a name that is
' already a function word ('length', 'count', a vocabulary word, or
' one this program defined) would silently shadow it. Refuse both
' with words, before anything is generated.
Private Sub CheckDupAction(ByVal name As String, ByVal namePos As Long)
    Dim n As String
    n = VLA_Identity.Fold(name)
    Dim e As Variant
    If Not mActNames Is Nothing Then
        For Each e In mActNames
            If CStr(e) = n Then VLA_Messages.RaiseMsg "english-action-name-taken", "name", name, "loc", LineTag(namePos)
        Next
    End If
    If Not mFnActNames Is Nothing Then
        For Each e In mFnActNames
            If CStr(e) = n Then VLA_Messages.RaiseMsg "english-action-name-taken", "name", name, "loc", LineTag(namePos)
        Next
    End If
    If IsFnWord(mFnOf, n) Or IsFnWord(mFnNullary, n) Then
        VLA_Messages.RaiseMsg "english-action-name-means-something", "name", name, "loc", LineTag(namePos)
    End If
    If IsUsingFn(n) Then
        VLA_Messages.RaiseMsg "english-action-name-taken", "name", name, "loc", LineTag(namePos)
    End If
End Sub

' Is this word a using-style value action defined by this program?
' Total on any state (Nothing-safe) - direct ParseStmt callers (tests,
' comma bodies) run without EnglishToVla's init.
Private Function IsUsingFn(ByVal word As String) As Boolean
    ' CollHasKey, not CollGet: the stored value is a Collection, and
    ' Let-copying an object Variant invokes a default property that
    ' Collections do not have (error 438). Set-vs-Let, absorbed once
    ' more - existence is all this asks.
    If mFnUsingNames Is Nothing Then Exit Function
    IsUsingFn = CollHasKey(mFnUsingNames, VLA_Identity.Fold(word))
End Function

Private Function InParamList(pNames As Collection, ByVal word As String) As Boolean
    If pNames Is Nothing Then Exit Function
    Dim e As Variant
    For Each e In pNames
        If CStr(e) = VLA_Identity.Fold(word) Then
            InParamList = True
            Exit Function
        End If
    Next
End Function

Private Function JoinParamNames(pNames As Collection) As String
    Dim e As Variant
    For Each e In pNames
        If Len(JoinParamNames) > 0 Then JoinParamNames = JoinParamNames & ", "
        JoinParamNames = JoinParamNames & CStr(e)
    Next
End Function

Private Sub AddKeyed(col As Collection, ByVal s As String)
    On Error Resume Next
    col.Add s, VLA_Identity.Fold(s)
    On Error GoTo 0
End Sub

' IN.7 (button-click half): turns a free-text button caption into the
' suffix of its internal handler proc name ("on:click:" & this). A
' caption is free text - spaces, punctuation, anything a real Excel
' button caption allows - but the VLA reader's own token boundaries
' (space/tab/CR/LF, "(", ")", ";", '"') cannot survive inside a
' single symbol token, and VBA's identifier rules are stricter still.
' Folding everything outside A-Z/a-z/0-9/_ to '_' satisfies both at
' once - deliberately NOT VLA.bas's own SymName (which preserves '.'
' for dotted member access; a literal '.' in a caption would silently
' become member-access syntax instead of a name character, exactly
' the untrappable-modal risk this whole feature exists to avoid).
' Twin of the identical helper in VLA.bas's EmitStmt "make-button"
' case, which must recompute this SAME slug from nothing but the
' caption in its own call, independently of this parse (CollGet's own
' neighboring comment: "each module stays self-contained").
Private Function ClickHandlerSlug(ByVal caption As String) As String
    Dim i As Long, n As Long, c As String
    Dim buf As String
    n = Len(caption)
    For i = 1 To n
        c = Mid$(caption, i, 1)
        If (c >= "a" And c <= "z") Or (c >= "A" And c <= "Z") Or (c >= "0" And c <= "9") Or c = "_" Then
            buf = buf & c
        Else
            buf = buf & "_"
        End If
    Next i
    ClickHandlerSlug = buf
End Function

' Keyed Collection lookup that reports absence instead of raising.
' (A twin of the helper in VLA.bas, which is Private there and thus
' invisible to this module - each module stays self-contained.)
Private Function CollGet(col As Collection, ByVal key As String, ByRef found As Boolean) As Variant
    On Error GoTo missing
    If IsObject(col.Item(VLA_Identity.Fold(key))) Then
        Set CollGet = col.Item(VLA_Identity.Fold(key))
    Else
        CollGet = col.Item(VLA_Identity.Fold(key))
    End If
    found = True
    Exit Function
missing:
    found = False
End Function

' G0's own hazard, a third module in: "target = expr" where expr is a
' Variant that happens to hold an object (CollGet's own return, when
' the stored item is a Collection) raises 450 ("wrong number of
' arguments") - VBA tries the object's hidden zero-argument default
' member (a Collection's is Item, which takes one) rather than just
' copying the reference. AssignVar branches on IsObject and picks Set
' or Let itself so no caller has to. (A twin of the helper in VLA.bas
' and VLA_Interpreter.bas, each Private and thus invisible here - see
' CollGet's own neighboring comment: "each module stays self-
' contained.")
Private Sub AssignVar(ByRef target As Variant, ByVal v As Variant)
    If IsObject(v) Then
        Set target = v
    Else
        target = v
    End If
End Sub

Private Function CollHasKey(col As Collection, ByVal key As String) As Boolean
    ' Object-safe existence probe. The old Let-copy (v = col.Item(k))
    ' raised 438 on OBJECT items - inside the error trap - so stored
    ' Collections reported as missing: B6.1 shipped with using-call
    ' parameter lists invisible to their own lookup, and the five
    ' using-form pins caught it on their maiden run. IsObject takes
    ' its argument as a Variant, copying an object REFERENCE without
    ' the default-property coercion, so existence is proven for
    ' object and plain items alike. (Set-vs-Let's third appearance
    ' in this registry family; see also CollGet and IsUsingFn.)
    On Error GoTo missing
    Dim isObj As Boolean
    isObj = IsObject(col.Item(VLA_Identity.Fold(key)))
    CollHasKey = True
    Exit Function
missing:
End Function

' Escape a raw string into a VLA string literal (\" and \\ escapes).
Private Function VlaStringLit(ByVal content As String) As String
    Dim c As String
    c = Replace(content, "\", "\\")
    c = Replace(c, """", "\""")
    VlaStringLit = """" & c & """"
End Function

' =====================================================================
'  Vocabulary files: load many phrase rules from a standalone file.
'
'  File format - every directive is a parenthesized form, and ";"
'  starts a comment. (Corrected while building F.10: this header
'  still described the "pattern =>" line format that F.13 RETIRED -
'  see EnglishLoadVocabularyText's own note, "No custom line/paren
'  scanning left here at all - forms are the only path." Nothing
'  reads the old shape any more, so the example below is the real
'  one.)
'      ; a comment
'      (english-vla
'          "put {e:expr} into cell {r:cell}"
'          (set! (range {r}) {e}))
'  A rule's own source language rides in the head (english-vla,
'  espanol-vla, ...); english-function's target is a bare atom,
'  because it names a function rather than English prose. Rules
'  register in file order, after the prelude and anything loaded
'  earlier. A phrasebook may declare its own preconditions with
'  (requires-version "X") - see this module's F.10 section.
'
'  Typical setup for a reload button:
'      EnglishResetGrammar
'      EnglishLoadVocabulary ThisWorkbook.Path & "\vocab\office.vocab"
'      EnglishLoadVocabulary ThisWorkbook.Path & "\vocab\finance.vocab"
' =====================================================================

' Load rules from a file. Returns the number of rules added.
Public Function EnglishLoadVocabulary(ByVal filePath As String) As Long
    EnsureInit
    If Len(Dir$(filePath)) = 0 Then
        VLA_Messages.RaiseMsg "english-vocab-file-not-found", "path", filePath
    End If
    Dim text As String
    text = VocabReadFile(filePath)

    ' F.10: both requires- gates run BEFORE SEC.2's raw prompt, so a
    ' phrasebook this build cannot run is refused outright rather than
    ' being refused only after the user has been asked to approve raw
    ' VBA access for it. The pure half runs first (version-first
    ' ordering holds across both sites, not just within one), then the
    ' consent-shaped half. EnglishLoadVocabularyText re-runs the pure
    ' half below - it is idempotent, and that is where the host-free
    ' translate path gets its own gate.
    VocabRequiresCheckPure text, filePath
    VocabRequiresCheckCapability text, filePath

    ' SEC.2: gate BEFORE a single rule from this file registers - see
    ' the raw-consent block above EnglishLoadVocabularyText for the
    ' full design and why this is the right chokepoint, not that
    ' shared, documented-host-free function.
    If VocabTextHasRawForm(text) Then
        Dim contentHash As String
        contentHash = EnglishSourceHash(filePath)

        ' D1-shaped capture (VLA_IDE.bas's own CaptureHost): ActiveWorkbook,
        ' never ThisWorkbook - inside a built .xlam, ThisWorkbook is the
        ' add-in itself ("add-ins never appear [as ActiveWorkbook], which
        ' is why [CaptureHost] was never ThisWorkbook" - that function's
        ' own comment). Using ThisWorkbook here would silently turn
        ' "approve for this workbook" into "approve for every workbook
        ' this add-in ever serves," defeating the whole point of the
        ' narrower option. Captured once, before the prompt below, not
        ' re-read after - a live dialog round-trip is exactly the kind of
        ' pause a focus change could land in.
        Dim hostWb As Workbook
        Set hostWb = ActiveWorkbook

        If Not VlaRawConsentGranted(contentHash, hostWb) Then
            Dim choice As Long
            choice = VlaAskRawConsent(filePath)
            Select Case choice
                Case RawConsentWorkbook
                    VlaRawConsentRecordWorkbook contentHash, hostWb
                Case RawConsentDevice
                    VlaRawConsentRecordDevice contentHash
                Case Else
                    VLA_Messages.RaiseMsg "english-vocab-raw-consent-declined", "source", filePath
            End Select
        End If
    End If

    EnglishLoadVocabulary = EnglishLoadVocabularyText(text, filePath)
End Function

' GEXPANDER.1: owner reversal of GEXPANDER.0's own original design -
' "every phrasebook load writes a fully macro-expanded sibling
' automatically" was adjudicated overkill (99.99% of loads never want
' one) and is GONE from EnglishLoadVocabulary above. The expanded text
' itself is still exactly what it always was - not a separate artifact
' from the grammar, a pure, deterministic function of the forms
' EnglishLoadVocabularyText just finished registering (mLastExpandedBlob,
' built fresh by the EnglishLoadVocabularyText call below, never carried
' over from an earlier load) - only WHEN it gets produced changed: now
' a deliberate, occasional, human-triggered export (VLA_IDE.bas's
' EnglishIdeExportExpandedVocabulary, a Save-As dialog on every click,
' the same category as the file-to-file Translate to VLA/VBA pipeline
' below), never an automatic side effect a self-test or VlaWriteGoldens
' run would trip over. Always
' reloads fresh (EnglishResetGrammar + EnglishLoadVocabulary) rather
' than trusting whatever happened to be loaded last, so the exported
' text is guaranteed to reflect vocabPath's own CURRENT on-disk content
' at export time, never a stale in-session snapshot.
Public Function EnglishExpandedVocabularyText(ByVal vocabPath As String) As String
    EnglishResetGrammar
    EnglishLoadVocabulary vocabPath

    ' Owner correction (kept from GEXPANDER.0): the header names the
    ' source by its bare filename (Dir$ strips the directory), never a
    ' machine-specific absolute path, and carries source-size but not
    ' mtime - see ExpandedSiblingPath's own neighbor note below and
    ' BETA_ROADMAP.md's G-EXPANDER entry for the full reasoning (both
    ' still apply unchanged; only the automatic-write half was reversed).
    ' 0.5.1 (owner call, 2026-09-07): the staleness stamp is a hash over the
    ' source's NON-WHITESPACE bytes, not its byte size. The byte size read a
    ' Lint VLA re-indent, a CRLF/LF checkout difference, or any other
    ' whitespace-only touch as "stale", and Beta's history shows what that
    ' cost: at least four commits to english_expanded.vla changed nothing
    ' but the stamp line (one titled "Re-stamp ... source-size"). Whitespace
    ' between tokens never changes what the phrasebook means, so it must
    ' not change whether the export of it counts as current. Computed from
    ' raw file bytes on both sides (here and tools/check_rule_coverage.ps1's
    ' Get-SourceHash), so no text-encoding step can make the two disagree.
    ' 0.5.3 (SEC.11): the same non-whitespace byte stream, but SHA-256 over
    ' it instead of a 32-bit polynomial, stamped with an explicit "sha256:"
    ' prefix so old and new stamps can never be confused for one another.
    ' check_rule_coverage.ps1 still READS the old 8-hex stamp (and the
    ' older byte-size one before it) - an artifact exported before this
    ' change keeps working and is upgraded by its next re-export, exactly
    ' how the byte-size stamp was retired in 0.5.1.
    Dim stamp As String
    stamp = "; GENERATED by EnglishExpandedVocabularyText from " & Dir$(vocabPath) & _
        " - do not hand-edit, re-export via Export Expanded Phrasebook if the source changes." & vbCrLf & _
        "; source-hash: " & EnglishSourceHash(vocabPath) & " - hash of the CURRENT source file's non-whitespace bytes (re-indenting or re-saving the source never changes it); a mismatch means this export is stale for it; re-export before trusting it." & vbCrLf & vbCrLf

    ' GEXPANDERLINT.0: the owner's own standing preference - every .vla
    ' this project writes is linted by default, the expanded artifact
    ' included, not just the two hand-edited corpus files VlaLintCheck
    ' already gates. mLastExpandedBlob (GEXPANDER.0's own flat, one-
    ' line-per-directive text, "; row:" tags included) is already valid
    ' house-style INPUT - VLA_Lint.SplitSegments reads a comment line or
    ' a balanced-paren form per line either way - so this is a straight
    ' reuse, not a second formatter: VLA_Lint.VlaLintFormat re-flows the
    ' same forms this pass already built, never re-derives them.
    EnglishExpandedVocabularyText = stamp & VLA_Lint.VlaLintFormat(mLastExpandedBlob)
End Function

' The identity of a phrasebook's CONTENT, used for two different jobs:
' the staleness stamp EnglishExpandedVocabularyText writes (and
' tools/check_rule_coverage.ps1 verifies), and - the load-bearing one -
' the key SEC.2's raw-consent records are filed under.
'
' 0.5.3 (SEC.11): this is now SHA-256, emitted as
' "sha256:<64 uppercase hex> over N non-whitespace bytes". Until 0.5.3 it
' was h = (h * 31 + b) mod 2^32, a Java-string-hash-class polynomial
' rendered as eight hex digits. That was fine for the staleness job and
' completely inadequate for the consent job: second preimages against a
' 31-polynomial are constructible BY HAND (flip two bytes of a ; comment
' so the running value lands back where it was), so a grant a person gave
' to a phrasebook they had actually read transferred, silently, to a
' crafted phrasebook they had never seen. See VLA_Digest.bas for the
' digest itself, and for why it is 150 lines of VBA arithmetic rather
' than the CreateObject("System.Security.Cryptography.SHA256Managed")
' SEC.11 was minted proposing - that class does not activate on the
' owner's own machine, or on a default Windows 11.
'
' WHAT IS HASHED, decided deliberately rather than inherited. Only bytes
' that are NOT tab, LF, CR or space take part - exactly the filter the
' polynomial used. Keeping it is a real decision with a real reason:
' .gitattributes checks .bas and .vla out as CRLF while the repository
' stores LF, so a raw whole-file digest would give ONE phrasebook two
' different identities on two machines - re-prompting for consent that
' was already granted, and marking every exported artifact stale, purely
' because of how a file was checked out. It costs nothing
' cryptographically: an attacker searching for a second preimage was
' always free to vary the non-whitespace bytes, and that is the entire
' space SHA-256 is hard over. Whitespace between tokens has never changed
' what a phrasebook MEANS, so it must not change what it IS.
'
' N (the count of participating bytes) is kept as a cheap human-readable
' second opinion, not as security - it was never load-bearing and is not
' now.
'
' Whole-file BYTES, never decoded text, on both sides: a UTF-8 BOM or a
' non-ASCII character hashes identically here and in PowerShell because
' neither side decodes anything.
'
' THE TWIN. Get-SourceHash in tools/check_rule_coverage.ps1 must produce
' the same digest for the same file. That agreement is no longer held
' together by this comment alone - tools/check_hash_twin.ps1 pins both
' sides to the same written-down vectors, and fails if either drifts.
Public Function EnglishSourceHash(ByVal filePath As String) As String
    Dim f As Integer
    f = FreeFile
    Open filePath For Binary Access Read As #f
    Dim n As Long
    n = LOF(f)
    Dim buf() As Byte
    If n > 0 Then
        ReDim buf(0 To n - 1)
        Get #f, , buf
    End If
    Close #f
    Dim counted As Long, hx As String
    hx = VLA_Digest.VlaSha256HexSkippingWhitespace(buf, n, counted)
    EnglishSourceHash = "sha256:" & hx & " over " & counted & " non-whitespace bytes"
End Function

' AS.1: "Phrasebook Test Coverage" - which rules have no test-success;
' which have only one. The VBA-native answer, not tools/check_rule_
' coverage.ps1's own weaker substitute: that script attributes a test to
' a rule by TEXT POSITION in an exported artifact (a heuristic, stated as
' such in its own header) - this function counts real FIRINGS, straight
' from mRuleTestCounts, which RunVocabTest already populates from
' mLastRuleIdx (TryPhrase's own real match, not a text-adjacency guess).
' Same fresh-load contract as EnglishExpandedVocabularyText - no
' staleness is possible here at all, because nothing is cached; every
' call re-derives the answer from vocabPath's own current content.
Public Function EnglishRuleCoverageReport(ByVal vocabPath As String) As String
    EnglishResetGrammar
    EnglishLoadVocabulary vocabPath

    Dim r As String
    r = "=== PHRASEBOOK TEST COVERAGE: " & Dir$(vocabPath) & " ===" & vbCrLf
    Dim ruleCount As Long
    ruleCount = mPatItems.Count - mPreludeCount   ' EnglishVocabStats's own convention
    r = r & "Rules: " & ruleCount & vbCrLf & vbCrLf

    ' firedTotal walks EVERY loaded rule, prelude included - any real
    ' TryPhrase match (an mRuleTestCounts entry) means the test fired an
    ' actual phrase rule, core built-in or not; only the true remainder
    ' (mTestsRun - firedTotal) is genuinely structural (if/repeat/create
    ' a number/..., forms TryPhrase never sees at all). The zero/one/2+
    ' BUCKETING below, by contrast, only walks THIS PHRASEBOOK'S OWN
    ' rules (index > mPreludeCount), the same convention
    ' EnglishVocabStats already uses for "loaded: N rules" - a core
    ' rule's own coverage isn't this report's job; english.vla didn't
    ' write it.
    Dim firedTotal As Long
    Dim i As Long
    For i = 1 To mPatItems.Count
        Dim found0 As Boolean
        Dim cur0 As Variant
        cur0 = CollGet(mRuleTestCounts, "r" & i, found0)
        If found0 Then firedTotal = firedTotal + CLng(cur0)
    Next

    Dim zero As String, one As String
    Dim zeroCount As Long, oneCount As Long, twoPlusCount As Long
    For i = mPreludeCount + 1 To mPatItems.Count
        Dim found As Boolean
        Dim cur As Variant
        cur = CollGet(mRuleTestCounts, "r" & i, found)
        Dim n As Long
        n = 0
        If found Then n = CLng(cur)
        Dim line As String
        line = "  " & mPatTexts.Item(i) & " [" & RuleSourceOf(i) & "]" & vbCrLf
        Select Case n
            Case 0
                zero = zero & line
                zeroCount = zeroCount + 1
            Case 1
                one = one & line
                oneCount = oneCount + 1
            Case Else
                twoPlusCount = twoPlusCount + 1
        End Select
    Next

    r = r & "--- Zero tests, " & zeroCount & " rule(s) ---" & vbCrLf & zero & vbCrLf
    r = r & "--- Exactly one test, " & oneCount & " rule(s) ---" & vbCrLf & one & vbCrLf

    r = r & "test-fail proofs (file-level, never attributed to a rule): " & mFailsRun & vbCrLf
    r = r & "test-success proofs for a structural form, not a phrase rule: " & _
            (mTestsRun - firedTotal) & vbCrLf & vbCrLf

    r = r & "=== SUMMARY: " & twoPlusCount & "/" & ruleCount & " rules have 2+ tests; " & _
            oneCount & " have exactly one; " & zeroCount & " have zero ==="
    EnglishRuleCoverageReport = r
End Function

' "english.vla" -> "english_expanded.vla"; a source with no extension
' gets a bare "_expanded" suffix rather than a bogus trailing dot.
' Public (GEXPANDER.1): now also the SUGGESTED filename for the Export
' Expanded Vocabulary save dialog (VLA_IDE.bas), a genuine cross-module
' caller, not just this module's own former auto-write path.
Public Function ExpandedSiblingPath(ByVal filePath As String) As String
    Dim dotPos As Long, slashPos As Long
    dotPos = InStrRev(filePath, ".")
    slashPos = InStrRev(filePath, "\")
    If dotPos > slashPos Then
        ExpandedSiblingPath = Left$(filePath, dotPos - 1) & "_expanded" & Mid$(filePath, dotPos)
    Else
        ExpandedSiblingPath = filePath & "_expanded"
    End If
End Function

' AS.1: the SUGGESTED filename for the Phrasebook Test Coverage save
' dialog (VLA_IDE.bas) - same sibling-naming shape as ExpandedSiblingPath
' above, but a fixed .txt extension: a coverage report is prose, never
' valid VLA source, so it should never masquerade as a re-importable
' .vla file the way an _expanded sibling legitimately can.
Public Function CoverageReportPath(ByVal filePath As String) As String
    Dim dotPos As Long, slashPos As Long
    dotPos = InStrRev(filePath, ".")
    slashPos = InStrRev(filePath, "\")
    If dotPos > slashPos Then
        CoverageReportPath = Left$(filePath, dotPos - 1) & "_coverage.txt"
    Else
        CoverageReportPath = filePath & "_coverage.txt"
    End If
End Function

' APROPOSPLUS: one line per loaded rule, "pattern -> template", built
' from mPatTexts/mPatForms (the same parallel-Collection pair
' EnglishListRefSlots already reads) - VLA.VlaWriteForm resolves each
' rule's template to its final call text, so this is the ENGLISH
' SENTENCE half of apropos, carried down to VLA.bas the same way
' vocabulary macro text already is. Pushed at the end of every load,
' cleared at every reset - see the two VlaAproposCarry call sites.
Private Function AproposRulesBlob() As String
    Dim r As String
    Dim i As Long
    For i = 1 To mPatTexts.Count
        Dim forms As Collection
        Set forms = mPatForms.Item(i)
        Dim tmplText As String
        tmplText = ""
        Dim f As Variant
        For Each f In forms
            If Len(tmplText) > 0 Then tmplText = tmplText & " "
            tmplText = tmplText & VLA.VlaWriteForm(f)
        Next
        r = r & mPatTexts.Item(i) & "  ->  " & tmplText & vbCrLf
    Next
    AproposRulesBlob = r
End Function

' LISTOPS-PROVENANCE: sourceName+startLine names the CALL SITE only -
' fine when one directive produces one rule, not fine when a generator
' (today, one written by hand; eventually a LISTOPS table-walker)
' produces many from one call, since every spliced child shares that
' one line. ProvLoc adds the one thing a call site can never carry on
' its own: which of the several a failure came from, when the
' generator (or a human, directly - at-row is not LISTOPS-only) said
' so via at-row. Empty rowTag reproduces the exact old wording,
' unchanged - a genuine typo/failure with no at-row in play looks
' exactly as it always did.
Private Function ProvLoc(ByVal sourceName As String, ByVal lineNo As Long, ByVal rowTag As String) As String
    ProvLoc = sourceName & " line " & lineNo
    If Len(rowTag) > 0 Then ProvLoc = ProvLoc & " (" & rowTag & ")"
End Function

Private Sub DispatchVocabForm(fl As Variant, ByVal startLine As Long, ByVal rowTag As String, ByVal sourceName As String, _
                               ByRef count As Long, testSents As Collection, testExps As Collection, _
                               testLines As Collection, testKinds As Collection, testRowTags As Collection, _
                               expTexts As Collection, expTags As Collection, _
                               ByVal allowExpansion As Boolean, ByVal depth As Long)
    If depth > 20 Then
        VLA_Messages.RaiseMsg "english-vocab-nesting-too-deep", "loc", ProvLoc(sourceName, startLine, rowTag)
    End If
    ' G0: every real directive is a parenthesized form (a list) - a
    ' bare atom (a typo, or old-format text fed to this reader by
    ' mistake) is not one, and `Set flc = fl` on a non-object value
    ' raises VBA's own opaque "Object required" (424) - a real crash,
    ' not a teaching refusal. Caught here instead, with the line and
    ' the atom - reachable both for a raw top-level form and for
    ' whatever a generator's own (begin ...) spliced in.
    If Not IsObject(fl) Then
        VLA_Messages.RaiseMsg "english-vocab-expected-directive", "loc", ProvLoc(sourceName, startLine, rowTag), "word", CStr(fl)
    End If
    Dim flc As Collection
    Set flc = fl
    Dim head As String
    head = CStr(flc.Item(1))

    ' A generator's own expansion may be (begin dir1 dir2 ...) - splice
    ' its children back in as separate directives, recursively (one
    ' generator calling another nests this way). Runs regardless of
    ' allowExpansion - a spliced child is already fully expanded, so
    ' checking for "begin" here is free and handles a generator body
    ' that itself contains a nested (begin ...). rowTag passes through
    ' UNCHANGED to every sibling - begin does not itself distinguish
    ' rows; at-row does.
    If head = "begin" Then
        Dim j As Long
        For j = 2 To flc.Count
            DispatchVocabForm flc.Item(j), startLine, rowTag, sourceName, count, testSents, testExps, testLines, testKinds, testRowTags, expTexts, expTags, False, depth + 1
        Next
        Exit Sub
    End If

    ' LISTOPS-PROVENANCE: (at-row label form) - a transparent wrapper,
    ' checked alongside "begin" for the same reason (both are
    ' structural, not directives themselves). label is whatever the
    ' caller wrote (a bare number, a bare word, or a quoted string -
    ' StripQuoteSigil handles either); form is dispatched exactly as
    ' if unwrapped, EXCEPT every failure inside it (including nested
    ' ones - a defmacro's own arity error, a test-success proof
    ' failure run later) now names label too. Not LISTOPS-only: usable
    ' today, by hand, anywhere a human wants failures inside a large
    ' hand-written block distinguished from one another.
    If head = "at-row" Then
        If flc.Count <> 3 Then
            VLA_Messages.RaiseMsg "english-vocab-at-row-arity", "loc", ProvLoc(sourceName, startLine, rowTag)
        End If
        Dim newRowTag As String
        newRowTag = StripQuoteSigil(CStr(flc.Item(2)))
        DispatchVocabForm flc.Item(3), startLine, newRowTag, sourceName, count, testSents, testExps, testLines, testKinds, testRowTags, expTexts, expTags, allowExpansion, depth + 1
        Exit Sub
    End If

    If Len(head) > 13 And Right$(head, 13) = "-vla-override" Then
        ' G3: an explicit, earned replacement - see AddPhraseRule.
        RecordExpandedForm expTexts, expTags, flc, rowTag   ' GEXPANDER.0
        AddPhraseRule StripQuoteSigil(CStr(flc.Item(2))), JoinFormsText(flc, 3), True
        count = count + 1
    ElseIf Len(head) > 4 And Right$(head, 4) = "-vla" Then
        ' F.13: the source language rides in the head on purpose -
        ' see BETA_ROADMAP.md's F.13 entry - and every one of them
        ' is a phrase-rule directive, no per-language dispatch
        ' needed. Polyglot files are deliberate, not tolerated:
        ' this never checks which prefix a file already committed
        ' to, because it never commits to one.
        RecordExpandedForm expTexts, expTags, flc, rowTag   ' GEXPANDER.0
        AddPhraseRule StripQuoteSigil(CStr(flc.Item(2))), JoinFormsText(flc, 3)
        count = count + 1
    ElseIf head = "test-success" Then
        ' Collected now, run after every rule in the file exists -
        ' rowTag captured NOW, at dispatch time, into its own parallel
        ' collection (testRowTags), because the proof itself doesn't
        ' run until well after this whole dispatch pass unwinds (see
        ' EnglishLoadVocabularyText's second loop) - by then rowTag
        ' the PARAMETER is long gone, same reason startLine is already
        ' captured into testLines instead of read live.
        RecordExpandedForm expTexts, expTags, flc, rowTag   ' GEXPANDER.0
        testSents.Add StripQuoteSigil(CStr(flc.Item(2)))
        testExps.Add JoinFormsText(flc, 3)
        testLines.Add startLine
        testKinds.Add "t"
        testRowTags.Add rowTag
    ElseIf head = "test-fail" Then
        ' G4: the negative proof. The sentence must REFUSE, and the
        ' refusal must contain the fragment - so a phrasebook
        ' protects its refusal decisions the way test-success
        ' protects its translations.
        Dim tfFragment As String
        tfFragment = StripQuoteSigil(CStr(flc.Item(3)))
        If Len(tfFragment) = 0 Then
            VLA_Messages.RaiseMsg "english-vocab-test-fail-empty-fragment", "loc", ProvLoc(sourceName, startLine, rowTag)
        End If
        RecordExpandedForm expTexts, expTags, flc, rowTag   ' GEXPANDER.0
        testSents.Add StripQuoteSigil(CStr(flc.Item(2)))
        testExps.Add tfFragment
        testLines.Add startLine
        testKinds.Add "f"
        testRowTags.Add rowTag
    ElseIf head = "defmacro" Then
        ' L4: phrasebooks carry (defmacro ...) forms, appended to
        ' every translation (bottom, source-map neutral - L4.2) so
        ' the transpiler's own pass-1 collection registers them -
        ' the macro system and the vocabulary system compose. See
        ' RegisterVocabMacro for the name-uniqueness and probe-
        ' transpile validation, unchanged from the old loader.
        RecordExpandedForm expTexts, expTags, flc, rowTag   ' GEXPANDER.0
        RegisterVocabMacro flc, sourceName, startLine, rowTag
    ElseIf Len(head) > 9 And Right$(head, 9) = "-function" Then
        ' function-word target is a bare atom, never quoted - it
        ' names a VLA/VBA function, the same role a template
        ' argument plays elsewhere, never English prose (the
        ' owner's own catch during this migration's build).
        RecordExpandedForm expTexts, expTags, flc, rowTag   ' GEXPANDER.0
        RegisterFunctionWord StripQuoteSigil(CStr(flc.Item(2))), CStr(flc.Item(3)), _
                             ProvLoc(sourceName, startLine, rowTag)
    ElseIf head = "keyword-alias" Then
        ' LX5.1: (keyword-alias "si" "if") - a language file's own
        ' control-flow vocabulary, the same directive shape as every
        ' other real directive here, so a second language never touches
        ' a .bas file to gain if/repeat/while/etc. See
        ' CanonicalizeStructuralWords for the consumer.
        RecordExpandedForm expTexts, expTags, flc, rowTag   ' GEXPANDER.0
        RegisterKeywordAlias StripQuoteSigil(CStr(flc.Item(2))), StripQuoteSigil(CStr(flc.Item(3)))
    ElseIf Len(head) > 9 And Left$(head, 9) = "requires-" Then
        ' F.10: the tag was already read, checked and (if unmet)
        ' refused by the lexical pre-pass, long before this dispatch
        ' pass started - see this module's F.10 section. Reaching it
        ' here at top level therefore means it PASSED, and this arm
        ' exists only so a legitimately-placed declaration isn't
        ' reported as an unrecognized directive.
        '
        ' depth > 0 means it arrived spliced out of a (begin ...) or
        ' produced by a generator's expansion - which the pre-pass,
        ' reading raw source text, could not have seen and therefore
        ' never checked. Honouring it here would be honouring an
        ' unchecked requirement; ignoring it silently would let a
        ' generator emit a declaration that does nothing while
        ' looking, to a reader, exactly like one that does. Refused
        ' instead, naming why.
        If depth > 0 Then
            VLA_Messages.RaiseMsg "english-vocab-requires-from-expansion", _
                "loc", ProvLoc(sourceName, startLine, rowTag), "head", head
        End If
        RecordExpandedForm expTexts, expTags, flc, rowTag   ' GEXPANDER.0
    ElseIf allowExpansion Then
        ExpandVocabMacroCall flc, head, sourceName, startLine, rowTag, count, testSents, testExps, testLines, testKinds, testRowTags, expTexts, expTags, depth
    Else
        VLA_Messages.RaiseMsg "english-vocab-expansion-not-directive", "loc", ProvLoc(sourceName, startLine, rowTag), "head", head
    End If
End Sub

' GEXPANDER.0: appends one real directive's own literal text (already
' fully macro-expanded, and already flattened out of any "begin"/
' "at-row" wrapper by DispatchVocabForm's own recursion above) to the
' running expanded-form record, in the exact order this dispatch pass
' reaches each real directive - called from every branch above that IS
' one, never for the two structural wrappers themselves.
Private Sub RecordExpandedForm(expTexts As Collection, expTags As Collection, fl As Variant, ByVal rowTag As String)
    expTexts.Add VLA.VlaWriteForm(fl)
    expTags.Add rowTag
End Sub

' METAVOCAB: the one real attempt. flc's head named nothing this
' module recognizes as a directive - try it as a call to a macro this
' FILE has itself carried so far (mVocabMacros, the exact accumulated
' text RegisterVocabMacro's own probe already re-transpiles on every
' defmacro) plus the prelude (VLA.VlaExpandText always includes it).
' toFixpoint:=True means one call expands any nested macro calls too -
' a generator calling another generator needs no help from here, only
' DispatchVocabForm's own (begin ...) splicing above does.
' Unrecognized-and-unexpandable (a genuine typo, or a name this file
' never carries) comes back byte-identical to what went in - VLA's own
' engine makes no other promise about a call it can't expand - so an
' unchanged round-trip is the signal to raise the ORIGINAL message,
' unchanged, not a new one: this pass changes nothing about what a
' plain mistake looks like.
Private Sub ExpandVocabMacroCall(flc As Collection, ByVal head As String, ByVal sourceName As String, _
                                  ByVal startLine As Long, ByVal rowTag As String, ByRef count As Long, _
                                  testSents As Collection, testExps As Collection, _
                                  testLines As Collection, testKinds As Collection, testRowTags As Collection, _
                                  expTexts As Collection, expTags As Collection, ByVal depth As Long)
    Dim callText As String
    callText = VLA.VlaWriteForm(flc)

    Dim expandedText As String
    On Error GoTo expandFailed
    Dim fired As Long
    expandedText = VLA.VlaExpandText(mVocabMacros & vbLf & callText, True, fired)
    On Error GoTo 0

    Dim expandedForms As Collection
    Set expandedForms = VLA.VlaReadForms(expandedText)
    If expandedForms.Count <> 1 Then GoTo unrecognized
    Dim ef As Variant
    AssignVar ef, expandedForms.Item(1)
    If Not IsObject(ef) Then GoTo unrecognized
    If VLA.VlaWriteForm(ef) = callText Then GoTo unrecognized   ' no macro matched - unchanged

    DispatchVocabForm ef, startLine, rowTag, sourceName, count, testSents, testExps, testLines, testKinds, testRowTags, expTexts, expTags, False, depth + 1
    Exit Sub

unrecognized:
    VLA_Messages.RaiseMsg "english-vocab-unrecognized-directive", "loc", ProvLoc(sourceName, startLine, rowTag), "head", head
    Exit Sub
expandFailed:
    ' A real macro DID match, but its own expansion raised (wrong
    ' arity, a reserved name...) - captured to a local before use, per
    ' this codebase's own G0 doctrine on live Err properties (VLA's
    ' own VlaReadForms/VlaReadFormsWithLines fail: handlers already
    ' guard this the same way). Re-raised with THIS file's real line -
    ' VlaExpandText's own line numbers describe a synthetic blob
    ' (mVocabMacros & callText), not this file, so passing its message
    ' through raw would name the wrong place.
    Dim errDesc As String
    errDesc = Err.Description
    VLA_Messages.RaiseMsg "english-vocab-macro-expansion-failed", "loc", ProvLoc(sourceName, startLine, rowTag), "head", head, "detail", errDesc
End Sub

' =====================================================================
'  SEC.2: raw behind explicit, per-phrasebook consent. THREAT_MODEL.md
'  SS1.4: a (raw "...") form splices its own string content verbatim
'  into the emitted VBA module - full VBA privilege, a gap "no static
'  analysis can ever bound... by design," so it is gated by consent
'  instead.
'
'  Gated at EnglishLoadVocabulary (the FILE-path loader, below) rather
'  than the shared EnglishLoadVocabularyText primitive both this file's
'  file-based loaders AND VLA_Browser.bas's host-free translate API
'  call - checked, not assumed, and it matters: VLA_Browser.bas
'  (PORT.1, already shipped) documents EnglishTranslateTextToVla/ToVba
'  as callers that "never touch a file, never show a MsgBox," and
'  tools/check_translate_purity.ps1 (PORT.2) pins EnglishLoadVocabularyText
'  itself as one of the tracked host-free functions. A first draft of
'  this gate lived inside EnglishLoadVocabularyText and only APPEARED
'  to pass that ratchet because the MsgBox sat in a sibling function
'  outside the line range the script scans - a loophole, not real
'  compliance, and a real regression against VLA_Browser.bas's own
'  documented contract (it calls EnglishLoadVocabularyText directly,
'  with phrasebook TEXT it already holds in memory - a raw-bearing
'  phrasebook passed through it would have popped a live dialog from
'  a function explicitly promised never to show one). EnglishLoadVocabulary
'  already does file I/O (Dir$/Open, both already "host" operations by
'  this codebase's own definition) and is NOT one of PORT.2's tracked
'  functions, so a consent prompt belongs there instead - it also
'  means only FILE-based loads are gated (the real "org phrasebook" /
'  "community phrasebook" tiers THREAT_MODEL.md SS2 describes are
'  files on disk), while the shipped base corpus's embedded-sheet path
'  (VLA_IDE.IdeLoadVocab's embedded-chain fallback) and VLA_Browser.bas's
'  text-based path stay exactly as host-free as documented - correct,
'  not merely convenient: embedded/shipped content is THREAT_MODEL's
'  own tier-1 "trusted by construction" tier, and a host-free caller
'  has no dialog to show in the first place.
'
'  Declining means the WHOLE phrasebook does not load - not just its
'  raw-bearing rules - so "before its rules become reachable" (SEC.2's
'  own roadmap wording) is exactly true: a partial load would leave a
'  program author unable to tell, from one missing rule alone, whether
'  that was a security refusal or a bug.
'
'  Consent is remembered per exact CONTENT, not per path - keyed by
'  EnglishSourceHash(filePath) (already-audited, already used for the
'  Expanded Phrasebook's own staleness stamp) rather than a new hash
'  function, since a real file path is always available here. A path
'  can be reused across different content, and THREAT_MODEL.md SS2
'  already found no provenance tag distinguishes phrasebooks today, so
'  editing the phrasebook at all - even a comment - invalidates prior
'  consent and re-prompts, deliberately stricter than a path-keyed
'  record.
'
'  SEC.11, 0.5.3 - the hash that keying rests on is now SHA-256.
'  Reusing EnglishSourceHash was right; what was wrong was that
'  EnglishSourceHash was a 32-bit polynomial, which made "keyed per
'  exact content" a claim the key could not actually support. A second
'  preimage against h*31+b is constructible by hand, so "editing the
'  phrasebook at all invalidates prior consent" held only against an
'  editor who was not trying to defeat it. It now holds against one who
'  is. Nothing above this paragraph changed shape - the same function is
'  called at the same place for the same reason; only its strength did.
'
'  TWO REMEMBER SCOPES, an owner-requested refinement over a single
'  blanket "remember" - offered as an explicit, disclosed choice, not
'  a silent default:
'    - WORKBOOK scope (CustomDocumentProperties on the captured
'      ActiveWorkbook, keyed by the same content hash): forging this
'      approval needs write access to that one specific file. An
'      attacker who already has that could just as easily inline the
'      payload directly - so this buys real containment, not
'      obscurity.
'    - DEVICE scope (SaveSetting/GetSetting, "Frazaro"/RAW_CONSENT_SECTION,
'      unchanged from the mechanism's first draft): ANY code able to
'      run as this Windows user can read or WRITE this exact registry
'      location - it is scoped only by the app-name string "Frazaro",
'      which is public (this project is open source), and the content
'      hash is a public, documented, easily-reproduced algorithm. A
'      forged device-scope grant does not require ever having run
'      Frazaro, still less this phrasebook - only some unrelated
'      foothold on the same Windows account, staged separately from
'      the phrasebook itself. This is a REAL, named limitation, not
'      hidden from the person choosing it - VlaAskRawConsent's own
'      prompt text says exactly this, in plain language, so "device"
'      is something a user opts INTO informed, not something they
'      silently inherit as the only option.
'  Neither scope defends against an attacker who ALREADY has arbitrary
'  code execution as this Windows user, or against raw VBA that has
'  ALREADY run once with full privilege writing a future consent
'  record itself - no persisted secret written by same-process VBA can
'  be made unforgeable by more privileged same-process VBA; nothing
'  here claims otherwise. What workbook-scope changes is the SHAPE of
'  the remaining risk: a forged grant then affects one artifact a
'  person can watch (this file), not every Frazaro use by this Windows
'  account forever.
'
'  ActiveWorkbook, never ThisWorkbook, for the workbook-scoped record -
'  see EnglishLoadVocabulary's own comment at the capture site.
'
'  Deliberately NO test-bypass toggle anywhere in this mechanism - a
'  settable "skip the real prompt" surface, however narrowly scoped,
'  is itself a standing backdoor (any other code in the project, or a
'  forgotten debug leftover, could set it and silently defeat consent
'  for the rest of the session with no audit trail). The "already
'  consented" path is instead tested by pre-seeding the exact state a
'  real click would have left, for EACH scope (see this section's own
'  tests), which exercises the real code path with no shortcut built
'  in. The live prompt-and-decline/accept interaction itself is
'  owner-verified manually, the same precedent DI.1's own trust dialog
'  already set in this codebase (BETA_ROADMAP1.md DI.1: "Owner-verified
'  live, all three trust states in sequence" - never folded into the
'  automated VlaSelfTestsAll suite).
'
'  TWO CHAINED STOCK MsgBox DIALOGS, not one custom UserForm - VBA's
'  MsgBox cannot relabel its own buttons ("Approve for this workbook"
'  is not an available caption for any MsgBox button set), and this
'  codebase's one existing custom UserForm (frmCLI.frm) carries its own
'  documented dev-reload fragility (a prior "Can't move focus" VBIDE
'  error traced to corrupted import/reload state, not a code bug) -
'  not worth risking, or worth an artifact this text-based house style
'  cannot review as plainly as everything else here, for a two-question
'  flow two ordinary dialogs already answer cleanly. MsgBox is
'  otherwise used only in VLA_IDE.bas; it appears here as a deliberate,
'  narrow exception - a callback indirection into the IDE layer would
'  only add a second settable bypass surface for no real safety gain.
' =====================================================================

' Does this vocabulary TEXT carry a (raw ...) form anywhere - top-level
' (VLA.bas's EmitTop) or nested inside a statement body (VLA.bas's
' EmitStmt) both dispatch "raw" verbatim into the emitted module, so
' detection walks every paren's own head symbol rather than trusting a
' substring search, which would false-positive on a rule whose own
' English sentence text or a nearby comment merely contains the word
' "raw" (this file has several). Comment-aware (";" to end of line, the
' real .vla comment rule - VLA.bas's own reader, ~line 2725), string-
' literal-aware (backslash-escaped, same as every other reader in this
' codebase) character walk - the same shape EnglishResolveCheck already
' uses for the identical reason, over the phrasebook's raw source text
' rather than an already-parsed form tree (Nth/IsList are Private to
' VLA.bas - not reachable from this module).
Private Function VocabTextHasRawForm(ByVal text As String) As Boolean
    Dim n As Long
    n = Len(text)
    Dim inLit As Boolean
    Dim i As Long
    i = 1
    Do While i <= n
        Dim c As String
        c = Mid$(text, i, 1)
        If inLit Then
            If c = "\" Then
                i = i + 2
            ElseIf c = """" Then
                inLit = False
                i = i + 1
            Else
                i = i + 1
            End If
        ElseIf c = """" Then
            inLit = True
            i = i + 1
        ElseIf c = ";" Then
            Do While i <= n
                If Mid$(text, i, 1) = vbCr Or Mid$(text, i, 1) = vbLf Then Exit Do
                i = i + 1
            Loop
        ElseIf c = "(" Then
            Dim j As Long
            j = i + 1
            Do While j <= n And Mid$(text, j, 1) = " "
                j = j + 1
            Loop
            Dim hEnd As Long
            hEnd = j
            Do While hEnd <= n
                If InStr(" ()" & vbCr & vbLf & vbTab, Mid$(text, hEnd, 1)) > 0 Then Exit Do
                hEnd = hEnd + 1
            Loop
            If VLA_Identity.Fold(Mid$(text, j, hEnd - j)) = "raw" Then
                VocabTextHasRawForm = True
                Exit Function
            End If
            i = hEnd
        Else
            i = i + 1
        End If
    Loop
End Function

' hostWb may be Nothing (no active workbook at the moment of a raw-
' bearing load - rare, but not impossible, e.g. a dev-rig context with
' nothing open) - workbook-scope simply has nowhere to check in that
' case, so only device-scope can ever answer True for it.
Private Function VlaRawConsentGranted(ByVal contentHash As String, ByVal hostWb As Workbook) As Boolean
    If GetSetting("Frazaro", RAW_CONSENT_SECTION, contentHash, "") = "granted" Then
        VlaRawConsentGranted = True
        Exit Function
    End If
    If hostWb Is Nothing Then Exit Function
    Dim v As String
    On Error Resume Next
    v = hostWb.CustomDocumentProperties(RawConsentPropName(contentHash)).Value
    On Error GoTo 0
    VlaRawConsentGranted = (v = "granted")
End Function

Private Function RawConsentPropName(ByVal contentHash As String) As String
    RawConsentPropName = RAW_CONSENT_SECTION & " " & contentHash
End Function

Private Sub VlaRawConsentRecordDevice(ByVal contentHash As String)
    SaveSetting "Frazaro", RAW_CONSENT_SECTION, contentHash, "granted"
End Sub

' No workbook to scope to (hostWb Is Nothing) is not fatal - this
' load already succeeded via the live prompt; only a FUTURE load would
' have benefited from the record, so it silently no-ops and the next
' load simply asks again.
Private Sub VlaRawConsentRecordWorkbook(ByVal contentHash As String, ByVal hostWb As Workbook)
    If hostWb Is Nothing Then Exit Sub
    Dim propName As String
    propName = RawConsentPropName(contentHash)
    On Error Resume Next
    hostWb.CustomDocumentProperties(propName).Value = "granted"
    If Err.Number <> 0 Then
        Err.Clear
        hostWb.CustomDocumentProperties.Add Name:=propName, LinkToContent:=False, Type:=4, Value:="granted"   ' 4 = msoPropertyTypeString - the literal, not the named Office constant, so this compiles with no dependency on the Office Object Library being a checked reference
    End If
    On Error GoTo 0
End Sub

' The DI.1-shaped prompt SEC.2 asks for, as two chained stock dialogs
' (see this section's own header for why not one custom dialog): names
' the specific phrasebook, states the actual capability in plain
' language rather than a generic "trust this file" (DI.1's own finding
' was that a generic trust rubber-stamp is not what a reviewer should
' be asked to click), then - only if the first answer was yes - asks
' ONE further, purely escalating question: also remember this for
' every workbook, not just this one. Owner correction to this pin,
' this session: an earlier wording asked "remember this approval?"
' with Yes=workbook/No=device, which read as the second dialog merely
' re-confirming the first "yes" (both answered "Yes") rather than as a
' distinct choice - a user skimming two Yes/No dialogs in a row could
' misread which scope they landed on. The load itself is ALREADY
' settled by the first dialog; workbook-scope is what "yes, load it"
' already implies at minimum, so the second dialog now asks only
' whether to go further, with Yes=the broader, riskier scope and
' No=the default, narrower one - the answer that matches "no further
' action" also matches the safer outcome, so skimming past this dialog
' can no longer land on the riskier choice by mis-reading Yes/Yes as
' agreement twice. Fails toward the SAFER, narrower scope on any
' answer that is not unambiguously "yes, go wider" - never toward the
' broader one.
Private Function VlaAskRawConsent(ByVal sourceName As String) As Long
    Dim loadPrompt As String
    loadPrompt = "'" & sourceName & "' wants to load a rule that runs unrestricted VBA - file access, other applications, anything VBA itself can do, not just Excel actions." & vbCrLf & vbCrLf & _
                 "Only allow this if you trust where this phrasebook came from." & vbCrLf & vbCrLf & _
                 "Load it anyway?"
    If MsgBox(loadPrompt, vbYesNoCancel Or vbExclamation Or vbDefaultButton2, "Frazaro - phrasebook needs raw VBA access") <> vbYes Then
        VlaAskRawConsent = RawConsentDeclined
        Exit Function
    End If

    Dim scopePrompt As String
    scopePrompt = "This loads into THIS workbook, and that approval is remembered here either way." & vbCrLf & vbCrLf & _
                  "Also approve it for EVERY workbook on this device, so you are never asked again anywhere?" & vbCrLf & vbCrLf & _
                  "Yes - more convenient, but a wider target: any other code able to run as you on this computer could forge this approval for a phrasebook of its own choosing." & vbCrLf & vbCrLf & _
                  "No - keep it to this workbook only. Safer: forging this approval would require tampering with this specific file."
    If MsgBox(scopePrompt, vbYesNo Or vbQuestion Or vbDefaultButton2, "Frazaro - approve for every workbook on this device?") = vbYes Then
        VlaAskRawConsent = RawConsentDevice
    Else
        VlaAskRawConsent = RawConsentWorkbook
    End If
End Function

' =====================================================================
'  F.10: (requires-<namespace> "<value>") - a phrasebook's own
'  declared preconditions, parsed and checked BEFORE a single rule
'  from the file registers.
'
'  SYNTAX. The namespace rides in the HEAD, not in an argument:
'      (requires-version "0.5.2")
'      (requires-capability "sendmail")
'      (requires-form "paint cell")
'  BETA_ROADMAP.md's own F.10 entry spells this "requires:
'  <namespace>:<value>", which is prose predating F.13's migration to
'  an all-forms file format - there is no line-oriented "key: value"
'  syntax left in a .vla file for it to be. Head-carried namespace is
'  this format's dominant pattern already (english-vla/espanol-vla -
'  "the source language rides in the head on purpose"; english-
'  function; -vla-override), and it is the only spelling that reads
'  correctly: a BARE atom in argument position names a FUNCTION here
'  ((english-function "keys of" vladictkeys) - that directive's own
'  comment: "a bare atom, never quoted - it names a VLA/VBA function"),
'  so (requires version "0.5.2") would claim `version` is a function.
'  The Lisp answer, (requires 'version "0.5.2"), is unavailable: this
'  dialect has no quote sigil at all (StripQuoteSigil strips a leading
'  double-quote and nothing else), and inventing one for a single
'  directive is a language change, not a tag.
'
'  WHY A LEXICAL PRE-PASS, not a DispatchVocabForm arm. Dispatch runs
'  in FILE ORDER, so a (requires-...) sitting below a rule would be
'  read only after that rule had already registered - SEC.2's gate,
'  the precedent this follows, fires "BEFORE a single rule from this
'  file registers." So the check walks the raw source text, in
'  VocabTextHasRawForm's exact shape (comment-aware, string-literal-
'  aware) and for the same reason: a substring search would false-
'  positive on prose. Being lexical and PRE-expansion also closes a
'  real hole - a phrasebook cannot disguise a requirement by defining
'  a macro of the same name, which ExpandVocabMacroCall would
'  otherwise happily expand into something else entirely. The
'  pre-pass is authoritative; DispatchVocabForm's own requires- arm
'  below is a no-op that exists only so a legitimately-placed tag
'  isn't reported as an unrecognized directive, and refuses outright
'  if one arrives from a generator's expansion (which the pre-pass,
'  reading raw text, could never have seen).
'
'  WHERE EACH NAMESPACE IS CHECKED - one parser, deliberately TWO
'  enforcement sites, because refusing and asking are not the same
'  act:
'    - version / form / unknown are pure yes-or-no refusals. They
'      touch no host object, so they belong in
'      EnglishLoadVocabularyText - which is also the only way
'      VLA_Browser.bas's host-free translate path gets gated at all.
'      A phrasebook that genuinely needs 0.7.0 is genuinely unusable
'      there too, and saying so raises a message; it opens no dialog.
'    - capability has to ASK (SEC.7's own two-scope consent UX), and
'      a MsgBox in EnglishLoadVocabularyText is precisely the trap
'      SEC.2 was corrected for: VLA_Browser.bas documents its own
'      callers as functions that "never touch a file, never show a
'      MsgBox," and check_translate_purity.ps1 tracks that function
'      by name. So capability is checked one level up, in
'      EnglishLoadVocabulary, beside SEC.2's raw gate.
'  This is NOT the F.10/CO.3 near-duplication that had to be
'  corrected: that was two SYNTAXES for one idea. Here one syntax and
'  one parser feed two sites that genuinely differ.
'
'  VERSION-FIRST ORDERING (owner decision, with the unknown-namespace
'  policy below). Every requires-version tag is evaluated before any
'  other tag, regardless of where it sits in the file. This is what
'  makes deny-by-default humane: a phrasebook written for a newer
'  build that declares its version gets "needs 0.7.0, this is 0.5.1"
'  - the precise, actionable message - instead of a vague complaint
'  about whatever unknown tag happened to come first.
'
'  UNKNOWN NAMESPACE: REFUSE (owner decision). Deny-by-default is
'  SD-15's posture, but the deciding argument is narrower and does
'  not depend on it. First, the head word settles it semantically:
'  every requires- tag is a PRECONDITION by construction - there is
'  no advisory one - so a requirement this build cannot understand is
'  a requirement it cannot confirm was met, and ignoring it is
'  loading a phrasebook whose stated precondition was never checked.
'  Second, and decisively, the forward-compatibility cost is ALREADY
'  PAID everywhere else in this format: DispatchVocabForm already
'  refuses any unrecognized top-level directive head, so an 0.5 build
'  already refuses an 0.7 phrasebook that uses any new directive.
'  Ignoring unknown namespaces would make requires- uniquely more
'  permissive than every other directive in the language - failing
'  open in the one construct whose whole job is gating. The namespace
'  set is closed and grows only by release.
'
'  ONE CORNER, STATED RATHER THAN HIDDEN: the pre-pass walks every
'  paren, not only top-level ones, so a (requires-...) written inside
'  a defmacro BODY is seen and checked even though it is not a
'  top-level declaration. That is conservative in the right direction
'  and never silently wrong: if its requirement holds, the dispatch
'  arm below still refuses the form the moment the generator is
'  actually called; if it does not hold, the file refuses earlier,
'  with the unmet message rather than the from-expansion one. Either
'  way a requirement inside a macro body is an error, and either way
'  it is reported - only the wording differs by which check reaches it
'  first.
' =====================================================================

' One parsed (requires-...) tag per entry, in file order. Returns the
' count. Namespaces come back folded; values come back with their
' string escapes resolved, exactly as the reader would have.
Private Function VocabRequiresScan(ByVal text As String, ByRef nsOut As Collection, _
                                    ByRef valOut As Collection, ByRef posOut As Collection) As Long
    Set nsOut = New Collection
    Set valOut = New Collection
    Set posOut = New Collection

    Dim n As Long
    n = Len(text)
    Dim inLit As Boolean
    Dim i As Long
    i = 1
    Do While i <= n
        Dim c As String
        c = Mid$(text, i, 1)
        If inLit Then
            If c = "\" Then
                i = i + 2
            ElseIf c = """" Then
                inLit = False
                i = i + 1
            Else
                i = i + 1
            End If
        ElseIf c = """" Then
            inLit = True
            i = i + 1
        ElseIf c = ";" Then
            Do While i <= n
                If Mid$(text, i, 1) = vbCr Or Mid$(text, i, 1) = vbLf Then Exit Do
                i = i + 1
            Loop
        ElseIf c = "(" Then
            Dim headStart As Long
            headStart = i + 1
            Do While headStart <= n And Mid$(text, headStart, 1) = " "
                headStart = headStart + 1
            Loop
            Dim hEnd As Long
            hEnd = headStart
            Do While hEnd <= n
                If InStr(" ()" & vbCr & vbLf & vbTab, Mid$(text, hEnd, 1)) > 0 Then Exit Do
                hEnd = hEnd + 1
            Loop
            Dim head As String
            head = VLA_Identity.Fold(Mid$(text, headStart, hEnd - headStart))
            If Len(head) > 9 And Left$(head, 9) = "requires-" Then
                nsOut.Add Mid$(head, 10)
                valOut.Add RequiresReadValue(text, hEnd)
                posOut.Add i
            End If
            i = hEnd
        Else
            i = i + 1
        End If
    Loop
    VocabRequiresScan = nsOut.Count
End Function

' The quoted value directly after a requires- head. Returns "" when
' there isn't one (a bare (requires-version) or (requires-version
' 0.5.2) with the quotes forgotten) - the caller turns that into the
' missing-value refusal, so this stays a pure reader and never raises.
Private Function RequiresReadValue(ByVal text As String, ByVal fromPos As Long) As String
    Dim n As Long
    n = Len(text)
    Dim i As Long
    i = fromPos
    Do While i <= n
        Dim c As String
        c = Mid$(text, i, 1)
        If c = """" Then Exit Do
        ' Anything but whitespace before the opening quote means there
        ' is no quoted value to read - stop rather than skipping over
        ' the rest of the form hunting for an unrelated string.
        If InStr(" " & vbCr & vbLf & vbTab, c) = 0 Then Exit Function
        i = i + 1
    Loop
    If i > n Then Exit Function

    i = i + 1
    Dim r As String
    Do While i <= n
        Dim d As String
        d = Mid$(text, i, 1)
        If d = "\" Then
            r = r & Mid$(text, i + 1, 1)
            i = i + 2
        ElseIf d = """" Then
            RequiresReadValue = r
            Exit Function
        Else
            r = r & d
            i = i + 1
        End If
    Loop
End Function

' 1-based line number of a character position, counted the way every
' other message in this module counts them. Called only when a tag
' actually needs to be named in a refusal, so walking the prefix is
' free in practice.
' lineNo, never `line`: Line is a VBA keyword (Line Input #, and the
' Line method on a drawing surface), and this codebase has already
' been bitten once by a local whose name collided with a VBA
' intrinsic.
Private Function RequiresLineOf(ByVal text As String, ByVal pos As Long) As Long
    Dim lineNo As Long
    lineNo = 1
    Dim i As Long
    For i = 1 To pos - 1
        If Mid$(text, i, 1) = vbLf Then lineNo = lineNo + 1
    Next
    RequiresLineOf = lineNo
End Function

' The host-free half: version, form, and the unknown-namespace
' refusal. Safe to call from EnglishLoadVocabularyText (and therefore
' from VLA_Browser.bas's host-free translate path) because every exit
' is a RaiseMsg, never a dialog. Idempotent and cheap, which is why
' the file-path loader calls it too rather than reordering the gates.
Private Sub VocabRequiresCheckPure(ByVal text As String, ByVal sourceName As String)
    Dim ns As Collection, vals As Collection, poss As Collection
    If VocabRequiresScan(text, ns, vals, poss) = 0 Then Exit Sub

    ' Pass 1 - every version tag, before anything else in the file is
    ' considered. See VERSION-FIRST ORDERING above.
    Dim i As Long
    For i = 1 To ns.Count
        If CStr(ns.Item(i)) = "version" Then
            RequiresCheckVersion CStr(vals.Item(i)), text, CLng(poss.Item(i)), sourceName
        End If
    Next

    ' Pass 2 - everything else this site owns, in file order.
    For i = 1 To ns.Count
        Dim nm As String
        nm = CStr(ns.Item(i))
        Select Case nm
            Case "version", "capability"
                ' version done above; capability belongs to the
                ' consent-shaped site, not this one.
            Case "form"
                ' CO.6 built docs/GRAMMAR_SINCE.md, which is what a
                ' per-form gate would have to read, but the ledger is
                ' a repo document and nothing carries it into a
                ' running add-in yet. The namespace is RECOGNIZED
                ' here on purpose - so the syntax was designed
                ' against it rather than bent to fit later - and
                ' refuses in words as not-yet-implemented, which is
                ' honest, rather than passing silently.
                VLA_Messages.RaiseMsg "english-vocab-requires-form-unsupported", _
                    "loc", ProvLoc(sourceName, RequiresLineOf(text, CLng(poss.Item(i))), ""), _
                    "form", CStr(vals.Item(i))
            Case Else
                VLA_Messages.RaiseMsg "english-vocab-requires-unknown-namespace", _
                    "loc", ProvLoc(sourceName, RequiresLineOf(text, CLng(poss.Item(i))), ""), _
                    "namespace", nm, "have", VLA.VLA_RELEASE_VERSION
        End Select
    Next
End Sub

Private Sub RequiresCheckVersion(ByVal wanted As String, ByVal text As String, _
                                  ByVal pos As Long, ByVal sourceName As String)
    Dim loc As String
    loc = ProvLoc(sourceName, RequiresLineOf(text, pos), "")
    If Len(wanted) = 0 Then
        VLA_Messages.RaiseMsg "english-vocab-requires-missing-value", "loc", loc, "namespace", "version"
    End If
    ' CO.4 built VlaVersionParse as a pure predicate that never raises
    ' precisely so this line can validate user-typed text with no
    ' error handler around it.
    ' vMaj/vMin/vPat, matching CO.4's own aMaj/aMin/aPat naming rather
    ' than the bare min/max spellings a reader could mistake for
    ' intrinsics.
    Dim vMaj As Long, vMin As Long, vPat As Long
    If Not VLA.VlaVersionParse(wanted, vMaj, vMin, vPat) Then
        VLA_Messages.RaiseMsg "english-vocab-requires-version-malformed", "loc", loc, "wanted", wanted
    End If
    If Not VLA.VlaVersionAtLeast(wanted) Then
        VLA_Messages.RaiseMsg "english-vocab-requires-version-unmet", "loc", loc, _
            "wanted", wanted, "have", VLA.VLA_RELEASE_VERSION
    End If
End Sub

' The consent-shaped half. Today every capability refuses: SEC.7 owns
' the grant side (its own two-scope, hash-keyed consent UX, reused
' from SEC.2 rather than DI.1's blanket per-publisher trust), and
' until it exists nothing can grant anything - so "refuse, naming the
' capability" is the correct and complete answer, not a placeholder.
' SEC.7 replaces the refusal below with its ask; the parse, the
' syntax, and this call site all stay exactly as they are.
Private Sub VocabRequiresCheckCapability(ByVal text As String, ByVal sourceName As String)
    Dim ns As Collection, vals As Collection, poss As Collection
    If VocabRequiresScan(text, ns, vals, poss) = 0 Then Exit Sub
    Dim i As Long
    For i = 1 To ns.Count
        If CStr(ns.Item(i)) = "capability" Then
            Dim loc As String
            loc = ProvLoc(sourceName, RequiresLineOf(text, CLng(poss.Item(i))), "")
            If Len(CStr(vals.Item(i))) = 0 Then
                VLA_Messages.RaiseMsg "english-vocab-requires-missing-value", "loc", loc, "namespace", "capability"
            End If
            VLA_Messages.RaiseMsg "english-vocab-requires-capability-ungranted", _
                "loc", loc, "capability", CStr(vals.Item(i))
        End If
    Next
End Sub

' F.13: load rules from a string of real VLA - read whole via
' VLA.VlaReadFormsWithLines (so every directive keeps a precise source
' line for its own error messages, the same way the old line-oriented
' scanner did), then dispatched by each top-level form's own head
' symbol. No custom line/paren scanning left here at all - forms are
' the only path, same as everywhere else since F.2's dual-path
' retirement. Returns rules added (matches the old contract exactly;
' EnglishLoadVocabulary and every other caller are unaffected).
Public Function EnglishLoadVocabularyText(ByVal text As String, _
                                          Optional ByVal sourceName As String = "vocabulary") As Long
    EnsureInit

    ' F.10: the phrasebook's own declared preconditions, checked
    ' before a single rule registers - the host-free namespaces only
    ' (see this module's F.10 section for why capability cannot be
    ' checked from here). This is the gate VLA_Browser.bas's host-free
    ' translate path gets, and it opens no dialog to get it.
    VocabRequiresCheckPure text, sourceName

    mLoadSource = sourceName   ' G3: provenance for every rule this load registers

    Dim forms As Collection, formLines As Collection
    Set forms = VLA.VlaReadFormsWithLines(text, formLines)

    Dim count As Long
    Dim testSents As New Collection
    Dim testExps As New Collection
    Dim testLines As New Collection
    Dim testKinds As New Collection   ' G4: "t" = test-success, "f" = test-fail
    Dim testRowTags As New Collection ' LISTOPS-PROVENANCE: "" unless at-row wrapped it
    Dim expTexts As New Collection    ' GEXPANDER.0: one entry per real directive, in dispatch order
    Dim expTags As New Collection     ' GEXPANDER.0: parallel - "" unless at-row wrapped it

    Dim k As Long
    k = 0
    Dim f As Variant
    For Each f In forms
        k = k + 1
        Dim startLine As Long
        startLine = CLng(formLines.Item(k))
        DispatchVocabForm f, startLine, "", sourceName, count, testSents, testExps, testLines, testKinds, testRowTags, expTexts, expTags, True, 0
    Next

    ' Run the file's proofs against the fully loaded grammar, in
    ' file order - positives (test-success) and negatives (test-fail)
    ' alike.
    Dim i As Long
    For i = 1 To testSents.Count
        If CStr(testKinds.Item(i)) = "f" Then
            mFailsRun = mFailsRun + 1
            RunVocabFailTest CStr(testSents.Item(i)), CStr(testExps.Item(i)), _
                             sourceName, CLng(testLines.Item(i)), CStr(testRowTags.Item(i))
        Else
            mTestsRun = mTestsRun + 1
            RunVocabTest CStr(testSents.Item(i)), CStr(testExps.Item(i)), _
                         sourceName, CLng(testLines.Item(i)), CStr(testRowTags.Item(i))
        End If
    Next

    mLoadSource = ""
    mLastLoadSource = sourceName

    ' P-PROBE round 2, correctness fix: VlaMacroDoc's own documented
    ' LIFETIME contract (VLA.bas, right above its declaration) says
    ' vocabulary-carried macros stay visible there "because their
    ' probe-transpile and every subsequent parse carry them through
    ' this same table" - a real, load-bearing side effect of the OLD
    ' per-macro probe (it called VlaTranspile directly, un-pushed, so
    ' mMacros was left holding prelude + every accumulated macro after
    ' the LAST registration). RegisterVocabMacro's own probe is now
    ' context-pushed/popped (VLA.VlaProbeMacroForm) specifically so it
    ' does NOT leak into mMacros per call - which silently broke this
    ' contract until a live test caught it (L17: "a vocabulary-carried
    ' docstring survives the probe and reads back"). Fixed by doing,
    ' ONCE per load instead of once per macro, exactly what the old
    ' per-macro probe did as a side effect: a real Pass-1-only compile
    ' (VLA.VlaCompileToForms, no context push - deliberately left to
    ' leak into mMacros, same as the old probe) over the FULL
    ' accumulated mVocabMacros, republishing prelude + every vocab
    ' macro into mMacros for VlaMacroDoc/apropos/GetMacro to read
    ' post-load. O(1) calls per load, not O(macros) - the quadratic
    ' cost this whole item exists to remove never comes back.
    VLA.VlaCompileToForms mVocabMacros

    ' L11.1: push the carried vocabulary down so apropos answers
    ' state-independently (see VlaAproposCarry). APROPOSPLUS: the
    ' second argument carries every loaded rule's own pattern text
    ' alongside its macro text, so apropos can answer with the English
    ' sentence too, not just the macro it reaches.
    VlaAproposCarry mVocabMacros, AproposRulesBlob()

    mLastExpandedBlob = BuildExpandedBlob(expTexts, expTags)   ' GEXPANDER.0

    EnglishLoadVocabularyText = count
End Function

' GEXPANDER.0: one line per real directive, in expTexts' own order
' (registration order - see RecordExpandedForm); an at-row label
' (expTags's parallel entry, "" when none) prints as its own comment
' line directly above the form it labels. AproposRulesBlob's own
' precedent (walk a loaded-forms collection, VLA.VlaWriteForm each
' entry, join) applied to the whole directive stream, not just
' phrase-rule templates.
Private Function BuildExpandedBlob(expTexts As Collection, expTags As Collection) As String
    Dim r As String
    Dim i As Long
    For i = 1 To expTexts.Count
        Dim tag As String
        tag = CStr(expTags.Item(i))
        If Len(tag) > 0 Then r = r & "; row: " & tag & vbCrLf
        r = r & CStr(expTexts.Item(i)) & vbCrLf
    Next
    BuildExpandedBlob = r
End Function

' F.13: joins VLA.VlaWriteForm of fl's elements from fromIdx onward,
' space-separated - reconstructs "template text" (or "expected VLA
' text") from a directive's own trailing forms, so AddPhraseRule/
' RunVocabTest/RunVocabFailTest (all unchanged, still text-based)
' never need to know their caller now holds real forms instead of
' hand-scanned text. A directive with more than one trailing form (a
' multi-statement template) joins all of them, in order - unexercised
' by any real rule today (see G-RENDER's own note on this), but the
' old format supported it and this preserves that.
Private Function JoinFormsText(fl As Collection, ByVal fromIdx As Long) As String
    Dim r As String
    Dim idx As Long
    For idx = fromIdx To fl.Count
        If Len(r) > 0 Then r = r & " "
        r = r & VLA.VlaWriteForm(fl.Item(idx))
    Next
    JoinFormsText = r
End Function

' =====================================================================
'  VOCABDIFF: structural diff between two vocabulary sources - see the
'  header note at this module's own version stamp for the design.
' =====================================================================

' One record per top-level form, in file order: (kind, key, canonical
' text, source line). Does NOT call EnglishLoadVocabularyText - a diff
' must not register rules, run macro probes, or touch whatever grammar
' is currently loaded; it only reads and classifies, mirroring that
' function's own head-symbol dispatch without any of its side effects.
' An unrecognized or malformed top-level form still gets a record
' (kind "other") instead of raising - this tool reports differences,
' it does not validate a vocabulary, so it is deliberately more
' tolerant than a real load.
Private Function DiffCollectItems(ByVal text As String) As Collection
    Dim forms As Collection, formLines As Collection
    Set forms = VLA.VlaReadFormsWithLines(text, formLines)

    Dim out As New Collection
    Dim k As Long
    k = 0
    Dim f As Variant
    For Each f In forms
        k = k + 1
        Dim ln As Long
        ln = CLng(formLines.Item(k))

        Dim kind As String, key As String, canon As String
        If Not IsObject(f) Then
            kind = "other"
            key = CStr(f) & " #" & k
            canon = CStr(f)
        Else
            Dim fl As Collection
            Set fl = f
            Dim head As String
            head = CStr(fl.Item(1))
            canon = VLA.VlaWriteForm(fl)

            If Len(head) > 13 And Right$(head, 13) = "-vla-override" Then
                kind = "override"
                key = StripQuoteSigil(CStr(fl.Item(2)))
            ElseIf Len(head) > 4 And Right$(head, 4) = "-vla" Then
                kind = "rule"
                key = StripQuoteSigil(CStr(fl.Item(2)))
            ElseIf head = "test-success" Then
                kind = "test-success"
                key = StripQuoteSigil(CStr(fl.Item(2)))
            ElseIf head = "test-fail" Then
                kind = "test-fail"
                key = StripQuoteSigil(CStr(fl.Item(2)))
            ElseIf head = "defmacro" Then
                kind = "macro"
                If fl.Count >= 2 And IsObject(fl.Item(2)) Then
                    Dim nf As Collection
                    Set nf = fl.Item(2)
                    key = CStr(nf.Item(1))
                Else
                    key = canon   ' malformed defmacro - fall back to itself
                End If
            ElseIf Len(head) > 9 And Right$(head, 9) = "-function" Then
                kind = "function"
                key = StripQuoteSigil(CStr(fl.Item(2)))
            Else
                kind = "other"
                key = head & " #" & k
            End If
        End If

        ' NOT "As New" - a New-typed Dim inside a loop auto-instantiates
        ' once and every later iteration reuses that SAME instance
        ' (Dim is hoisted; only the first Nothing access triggers New),
        ' so every record would end up pointing at one shared,
        ' ever-growing Collection instead of getting its own. Set
        ' fresh, explicitly, every iteration.
        Dim rec As Collection
        Set rec = New Collection
        rec.Add kind
        rec.Add key
        rec.Add canon
        rec.Add ln
        out.Add rec
    Next
    Set DiffCollectItems = out
End Function

' A record's own (kind, key) folded into one lookup key - kind keeps a
' rule and a same-named macro from colliding; folding keeps the match
' case/whitespace-blind, matching CollGet's own convention elsewhere
' in this module.
Private Function DiffItemKey(ByVal kind As String, ByVal key As String) As String
    DiffItemKey = VLA_Identity.Fold(kind & "|" & key)
End Function

Private Function DiffCounts(items As Collection) As String
    Dim rules As Long, overrides As Long, tsucc As Long, tfail As Long
    Dim macros As Long, funcs As Long, other As Long
    Dim it As Variant
    For Each it In items
        Dim rec As Collection
        Set rec = it
        Select Case CStr(rec.Item(1))
            Case "rule": rules = rules + 1
            Case "override": overrides = overrides + 1
            Case "test-success": tsucc = tsucc + 1
            Case "test-fail": tfail = tfail + 1
            Case "macro": macros = macros + 1
            Case "function": funcs = funcs + 1
            Case Else: other = other + 1
        End Select
    Next
    DiffCounts = rules & " rules, " & overrides & " overrides, " & macros & " macros, " & _
                 tsucc & " tests (" & tfail & " expected fails), " & funcs & " functions" & _
                 IIf(other > 0, ", " & other & " unrecognized", "")
End Function

' Keyed by DiffItemKey so a rule/override/test/macro/function-word
' each occupy their own namespace. A key present on both sides is
' UNCHANGED or CHANGED by comparing VLA.VlaWriteForm's canonical text
' (immune to reflow/reindent/reordering); present only on one side is
' REMOVED or ADDED. Report order: A's own order for CHANGED/REMOVED,
' then B's own order for ADDED - each side keeps its real order, this
' never invents one merged ordering neither file actually has.
Private Function DiffReport(itemsA As Collection, itemsB As Collection, _
                             ByVal labelA As String, ByVal labelB As String) As String
    Dim lookupB As New Collection
    Dim it As Variant
    For Each it In itemsB
        Dim rb0 As Collection
        Set rb0 = it
        Dim dkb As String
        dkb = DiffItemKey(CStr(rb0.Item(1)), CStr(rb0.Item(2)))
        If Not CollHasKey(lookupB, dkb) Then lookupB.Add rb0, dkb   ' first occurrence wins
    Next

    Dim matchedB As New Collection
    Dim added As Long, removed As Long, changed As Long, sameCount As Long
    Dim body As String

    For Each it In itemsA
        Dim ra As Collection
        Set ra = it
        Dim kindA As String, keyA As String, canonA As String, lineA As Long
        kindA = CStr(ra.Item(1))
        keyA = CStr(ra.Item(2))
        canonA = CStr(ra.Item(3))
        lineA = CLng(ra.Item(4))
        Dim dk As String
        dk = DiffItemKey(kindA, keyA)

        Dim found As Boolean
        Dim matchB As Variant
        AssignVar matchB, CollGet(lookupB, dk, found)
        If Not found Then
            removed = removed + 1
            body = body & "  REMOVED " & kindA & " " & labelA & ":" & lineA & " """ & keyA & """" & vbCrLf
        Else
            Dim rbm As Collection
            Set rbm = matchB
            If Not CollHasKey(matchedB, dk) Then matchedB.Add True, dk
            If CStr(rbm.Item(3)) = canonA Then
                sameCount = sameCount + 1
            Else
                changed = changed + 1
                body = body & "  CHANGED " & kindA & " """ & keyA & """" & vbCrLf
                body = body & "    " & labelA & ":" & lineA & "  " & canonA & vbCrLf
                body = body & "    " & labelB & ":" & CLng(rbm.Item(4)) & "  " & CStr(rbm.Item(3)) & vbCrLf
            End If
        End If
    Next

    For Each it In itemsB
        Dim rc As Collection
        Set rc = it
        Dim kindC As String, keyC As String, lineC As Long
        kindC = CStr(rc.Item(1))
        keyC = CStr(rc.Item(2))
        lineC = CLng(rc.Item(4))
        If Not CollHasKey(matchedB, DiffItemKey(kindC, keyC)) Then
            added = added + 1
            body = body & "  ADDED " & kindC & " " & labelB & ":" & lineC & " """ & keyC & """" & vbCrLf
        End If
    Next

    Dim r As String
    r = "VocabDiff: " & labelA & " (" & DiffCounts(itemsA) & ")" & vbCrLf
    r = r & "        vs " & labelB & " (" & DiffCounts(itemsB) & ")" & vbCrLf
    r = r & body
    r = r & changed & " changed, " & added & " added, " & removed & " removed, " & sameCount & " unchanged" & vbCrLf
    DiffReport = r
End Function

' Compares two vocabulary sources given as text. See this module's own
' VOCABDIFF.0 version-stamp note for the full design and its stated
' scope (two revisions of one vocabulary, not cross-language matching).
Public Function EnglishDiffVocabularyText(ByVal textA As String, ByVal textB As String, _
                                           Optional ByVal labelA As String = "A", _
                                           Optional ByVal labelB As String = "B") As String
    EnglishDiffVocabularyText = DiffReport(DiffCollectItems(textA), DiffCollectItems(textB), labelA, labelB)
End Function

' File-path convenience wrapper, the same pairing EnglishLoadVocabulary
' takes over EnglishLoadVocabularyText.
Public Function EnglishDiffVocabulary(ByVal pathA As String, ByVal pathB As String) As String
    EnglishDiffVocabulary = EnglishDiffVocabularyText(VocabReadFile(pathA), VocabReadFile(pathB), pathA, pathB)
End Function

' LISTOPS-PROVENANCE: mirrors DefineMacro's own L17 docstring-index
' rule (VLA.bas) exactly - RegisterVocabMacro needs to know where the
' TEMPLATE half of a defmacro form starts, to wrap ONLY that half in
' (gen-row ...), leaving the signature and any docstring untouched.
' Small, stable, and duplicated rather than shared because DefineMacro
' is Private to a different module - the rule itself (L17: a string
' literal at form 3 is documentation only when at least one more form
' follows) is narrow and has not changed since it shipped.
Private Function MacroBodyStartIdx(macForm As Collection) As Long
    MacroBodyStartIdx = 3
    If macForm.Count >= 4 Then
        Dim d3 As Variant
        AssignVar d3, macForm.Item(3)
        If Not IsObject(d3) Then
            If Left$(CStr(d3), 1) = Chr$(34) Then MacroBodyStartIdx = 4
        End If
    End If
End Function

' F.13: registers one (defmacro ...) form - the same validation the
' old macro: handling did (name uniqueness, probe-transpile through
' the engine's own gates), but the name comes from the form's own
' structure (Item(2).Item(1), the "(name params...)" list's own first
' element) instead of VocabMacroName's hand-rolled text scan, and
' there is no "carries (defmacro ...) forms only" check left to make -
' the dispatch above only ever calls this when the form's own head
' symbol already proved that structurally.
Private Sub RegisterVocabMacro(macForm As Collection, ByVal sourceName As String, ByVal startLine As Long, _
                                Optional ByVal rowTag As String = "")
    ' LISTOPS-PROVENANCE: rowTag<>"" (this defmacro came from an
    ' (at-row label (defmacro ...)) directive, hand-written or
    ' generator-emitted) rewraps the TEMPLATE half only - signature and
    ' any docstring pass through untouched - in (gen-row "label" ...),
    ' VLA.bas's own EmitStmt case, so EVERY future call to this macro,
    ' in any program, carries the row it was generated from. rowTag=""
    ' (every macro before this pass, and any macro not at-row-wrapped)
    ' skips this entirely - effForm is macForm itself, byte-for-byte
    ' the old behavior.
    Dim effForm As Collection
    Set effForm = macForm
    If Len(rowTag) > 0 Then
        Dim bodyStart As Long
        bodyStart = MacroBodyStartIdx(macForm)
        Dim wrapped As New Collection
        wrapped.Add "gen-row"
        wrapped.Add Chr$(34) & rowTag
        Dim bi As Long
        For bi = bodyStart To macForm.Count
            wrapped.Add macForm.Item(bi)
        Next
        Dim rebuilt As New Collection
        rebuilt.Add "defmacro"
        rebuilt.Add macForm.Item(2)
        If bodyStart = 4 Then rebuilt.Add macForm.Item(3)   ' the docstring, untouched
        rebuilt.Add wrapped
        Set effForm = rebuilt
    End If

    Dim macText As String
    macText = VLA.VlaWriteForm(effForm)

    Dim nameForm As Collection
    Set nameForm = macForm.Item(2)
    Dim macName As String
    macName = CStr(nameForm.Item(1))

    Dim mnFound As Boolean
    mnFound = False
    Dim mnPrev As Variant
    mnPrev = CollGet(mVocabMacroNames, VLA_Identity.Fold(macName), mnFound)
    If mnFound Then
        ' U.9: the guard's teaching matches the case. Same path
        ' re-carried = a raw re-load without a reset (the loader is
        ' ADDITIVE by design for L4 composition - the resetting
        ' wrappers are what made weeks of reloads clean) - so say the
        ' fix. Different paths = the cross-file collision the guard
        ' was built for - the pinned wording stands.
        If CStr(mnPrev) = sourceName Then
            VLA_Messages.RaiseMsg "english-vocab-file-already-carried-same", "loc", ProvLoc(sourceName, startLine, rowTag)
        End If
        VLA_Messages.RaiseMsg "english-vocab-macro-name-collision", "loc", ProvLoc(sourceName, startLine, rowTag), "name", macName, "prev", CStr(mnPrev)
    End If
    ' P-PROBE: this used to re-transpile mVocabMacros (every macro
    ' registered so far) & macText on EVERY new registration -
    ' O(macros^2) in the number of vocab macros, confirmed the
    ' dominant cost of a full corpus load (BETA_ROADMAP1.md's own
    ' P-PROBE entry). Read against VlaTranspile's actual Pass 1/Pass 2
    ' split (VLA.bas) before changing it: Pass 1 (DefineMacro) collects
    ' every top-level defmacro form into a FRESH mMacros Collection on
    ' every call regardless of what's fed in, and allows silent
    ' redefinition (`mMacros.Remove` then `Add`, no collision check);
    ' Pass 2 only expands/emits non-defmacro top forms, and the probe's
    ' old trailing "(sub vla-macro-check ())" called nothing - so no
    ' macro in the corpus is ever invoked or cross-checked against
    ' another here. The only thing this probe actually proves is that
    ' THIS macro's own form is well-formed (DefineMacro's reserved-
    ' name/param/template checks) and that its VlaWriteForm text
    ' round-trips through the reader - both entirely local to macText,
    ' independent of every other macro already carried. Re-feeding the
    ' whole accumulated corpus each time was pure waste; the name-
    ' collision case is already caught above (mVocabMacroNames), before
    ' this probe ever runs.
    ' Second round, found live: dropping the accumulated corpus above
    ' alone barely moved a real load's own wall time - VlaTranspile
    ' ALSO splices in and re-parses PreludeMacros() (prelude.vla, 34 KB
    ' / 46 macros) and re-registers all 46 of them via DefineMacro on
    ' EVERY call, a fixed ~8-15 ms cost (Timer-confirmed) paid by every
    ' registration regardless of the corpus fix above - the real
    ' dominant cost at today's corpus size. VLA.VlaProbeMacroForm
    ' (VLA.bas, same F5.0 push/pop-context discipline as VlaReadForms)
    ' runs exactly DefineMacro's own check with no prelude splice and
    ' no Pass 2 at all - the "(sub vla-macro-check ())" wrapper is gone
    ' too, since Pass 2 never validated anything macro-specific to
    ' begin with.
    On Error GoTo macroProbeFailed
    VLA.VlaProbeMacroForm macText
    On Error GoTo 0
    mVocabMacros = mVocabMacros & IIf(Len(mVocabMacros) > 0, vbCrLf, "") & macText
    mVocabMacroNames.Add sourceName, VLA_Identity.Fold(macName)
    Exit Sub

macroProbeFailed:
    Dim macErr As String
    macErr = Err.Description
    On Error GoTo 0
    VLA_Messages.RaiseMsg "english-vocab-macro-form-invalid", "loc", ProvLoc(sourceName, startLine, rowTag), "detail", macErr
End Sub

' U.10: the load-report counters line - one honest copy instead of
' an assembled memory. Counts are cumulative since the last reset
' (multi-file composition adds up, exactly like the rule store);
' the source names the LAST load.
Public Function EnglishVocabStats() As String
    EnsureInit
    Dim nr As Long
    nr = mPatItems.Count - mPreludeCount
    Dim r As String
    r = "loaded: " & nr & " rule" & IIf(nr = 1, "", "s") & ", " & _
        mVocabMacroNames.Count & " macro" & IIf(mVocabMacroNames.Count = 1, "", "s") & ", " & _
        mTestsRun & " test" & IIf(mTestsRun = 1, "", "s") & _
        " (" & mFailsRun & " expected fail" & IIf(mFailsRun = 1, "", "s") & ")"
    If Len(mLastLoadSource) > 0 Then r = r & " from " & mLastLoadSource
    EnglishVocabStats = r
End Function

' GO.6: one line per distinct phrasebook source currently registered
' (the built-in prelude excluded), in first-load order, with a rule
' count - the "can a user see what they've loaded" visibility GO.6's
' own open question #3 asked for, alongside the ADD-not-replace
' semantics EnglishIdeLoadPhrasebook (VLA_IDE.bas) actually uses.
' AddPhraseRule's own override provenance format is "newSrc (overrides
' oldSrc)" - only the part before " (overrides" names who currently
' OWNS the rule, so that is the bucket key here; the full string still
' prints per rule wherever Explain already shows it (RuleSourceOf) -
' this report only answers "which files are active", not "who
' overrode whom". Linear same-name scan, not a keyed lookup: the
' number of distinct sources is always tiny (a base corpus plus a
' handful of loaded phrasebooks, per PersistPhrasebookPath's own
' 8-slot ceiling), so an O(n^2) scan over n sources costs nothing a
' human would notice - matching EnglishLintReport's own simplicity.
Public Function EnglishLoadedSourcesReport() As String
    EnsureInit
    Dim names As New Collection
    Dim counts() As Long
    ReDim counts(1 To mPatSources.Count + 1)
    Dim distinctCount As Long
    distinctCount = 0
    Dim i As Long
    For i = mPreludeCount + 1 To mPatSources.Count
        Dim src As String
        src = CStr(mPatSources.Item(i))
        Dim cut As Long
        cut = InStr(src, " (overrides ")
        If cut > 0 Then src = Left$(src, cut - 1)
        Dim foundAt As Long
        foundAt = 0
        Dim j As Long
        For j = 1 To distinctCount
            If CStr(names.Item(j)) = src Then
                foundAt = j
                Exit For
            End If
        Next j
        If foundAt = 0 Then
            distinctCount = distinctCount + 1
            names.Add src
            foundAt = distinctCount
            counts(foundAt) = 0
        End If
        counts(foundAt) = counts(foundAt) + 1
    Next i
    If distinctCount = 0 Then
        EnglishLoadedSourcesReport = "No phrasebook loaded beyond the built-in grammar."
        Exit Function
    End If
    Dim r As String
    For j = 1 To distinctCount
        r = r & CStr(names.Item(j)) & " - " & counts(j) & " rule" & IIf(counts(j) = 1, "", "s") & vbCrLf
    Next j
    EnglishLoadedSourcesReport = r
End Function

' Translate one test sentence and compare against the expected VLA
' (whitespace-normalized). Failures raise loudly with file and line:
' a vocabulary does not load unless its proofs hold.
Private Sub RunVocabTest(ByVal sentence As String, ByVal expected As String, _
                         ByVal sourceName As String, ByVal lineNo As Long, Optional ByVal rowTag As String = "")
    Dim savedD As Collection, savedA As Collection
    Set savedD = mDeclared
    Set savedA = mAssigned
    Set mDeclared = New Collection
    Set mAssigned = New Collection
    ' S2: a proof firing is evidence of loading, not of usage - the
    ' counters measure real programs, so every rule with a test: line
    ' does not arrive pre-counted at every load. Restored on BOTH
    ' exit paths.
    Dim savedSusp As Boolean
    savedSusp = mUsageSuspended
    mUsageSuspended = True

    Dim toks() As String
    Dim pos As Long
    Dim got As String
    Dim d As String

    On Error GoTo failed
    toks = CanonicalizeStructuralWords(EnTokenize(sentence))   ' LX5.1
    pos = 1
    mLastRuleIdx = 0
    got = ParseStmt(toks, pos, 0)
    Do While TokAt(toks, pos) = PARA_TOK
        pos = pos + 1
    Loop
    If pos <= UBound(toks) Then
        VLA_Messages.RaiseMsg "english-extra-words-after-statement"
    End If
    On Error GoTo 0
    ' The passing sentence becomes the rule's worked example.
    If mLastRuleIdx > 0 Then
        On Error Resume Next
        mRuleExamples.Add sentence, "r" & mLastRuleIdx
        On Error GoTo 0
        BumpRuleTestCount mLastRuleIdx   ' AS.1: the real firing count
    End If
    Set mDeclared = savedD
    Set mAssigned = savedA
    mUsageSuspended = savedSusp

    If NormalizeWs(got) <> NormalizeWs(expected) Then
        VLA_Messages.RaiseMsg "english-test-failed", "loc", ProvLoc(sourceName, lineNo, rowTag), "sentence", sentence, "expected", NormalizeWs(expected), "got", NormalizeWs(got)
    End If
    Exit Sub

failed:
    d = Err.Description
    Set mDeclared = savedD
    Set mAssigned = savedA
    mUsageSuspended = savedSusp
    VLA_Messages.RaiseMsg "english-test-failed-to-translate", "loc", ProvLoc(sourceName, lineNo, rowTag), "sentence", sentence, "error", d
End Sub

' AS.1: the counting primitive for mRuleTestCounts - same "Collections
' cannot update in place, remove and re-add under the same key" idiom
' BumpUsage already established (see there for the fuller reasoning),
' minus BumpUsage's own case-folding (a rule INDEX, not an author-typed
' key, is already canonical) and minus its mUsageSuspended gate (a
' proof firing is never suspended - see this file's own field comment
' on mRuleTestCounts).
Private Sub BumpRuleTestCount(ByVal ruleIdx As Long)
    If mRuleTestCounts Is Nothing Then Set mRuleTestCounts = New Collection
    Dim key As String
    key = "r" & ruleIdx
    Dim found As Boolean
    Dim cur As Variant
    cur = CollGet(mRuleTestCounts, key, found)
    If found Then
        mRuleTestCounts.Remove key
        mRuleTestCounts.Add CLng(cur) + 1, key
    Else
        mRuleTestCounts.Add 1&, key
    End If
End Sub

' G4: the negative proof. Translate the sentence with the same
' machinery as RunVocabTest (state saved, usage suspended - a proof
' firing is evidence of loading, not usage, negatives included).
' If it TRANSLATES, the proof fails: the vocabulary was supposed to
' refuse this. If it refuses but the message lacks the fragment
' (whitespace-normalized, case-blind - the suite's Norm philosophy),
' the wording drifted and the proof fails naming both texts. Either
' failure refuses the whole load, exactly like test:.
Private Sub RunVocabFailTest(ByVal sentence As String, ByVal fragment As String, _
                             ByVal sourceName As String, ByVal lineNo As Long, Optional ByVal rowTag As String = "")
    Dim savedD As Collection, savedA As Collection
    Set savedD = mDeclared
    Set savedA = mAssigned
    Set mDeclared = New Collection
    Set mAssigned = New Collection
    Dim savedSusp As Boolean
    savedSusp = mUsageSuspended
    mUsageSuspended = True

    Dim toks() As String
    Dim pos As Long
    Dim got As String
    Dim d As String

    On Error GoTo refused
    toks = CanonicalizeStructuralWords(EnTokenize(sentence))   ' LX5.1
    pos = 1
    got = ParseStmt(toks, pos, 0)
    Do While TokAt(toks, pos) = PARA_TOK
        pos = pos + 1
    Loop
    If pos <= UBound(toks) Then
        VLA_Messages.RaiseMsg "english-extra-words-after-statement"
    End If
    On Error GoTo 0
    Set mDeclared = savedD
    Set mAssigned = savedA
    mUsageSuspended = savedSusp
    VLA_Messages.RaiseMsg "english-failtest-translated", "loc", ProvLoc(sourceName, lineNo, rowTag), "sentence", sentence, "became", NormalizeWs(got)
    Exit Sub

refused:
    d = Err.Description
    Set mDeclared = savedD
    Set mAssigned = savedA
    mUsageSuspended = savedSusp
    If InStr(1, NormalizeWs(d), NormalizeWs(fragment), vbTextCompare) = 0 Then
        VLA_Messages.RaiseMsg "english-failtest-message-drifted", "loc", ProvLoc(sourceName, lineNo, rowTag), "sentence", sentence, "wanted", fragment, "message", d
    End If
End Sub

' Collapse all whitespace runs to single spaces for test comparison.
Private Function NormalizeWs(ByVal s As String) As String
    s = Replace(s, vbCrLf, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, vbTab, " ")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    NormalizeWs = Trim$(s)
End Function

' =====================================================================
'  S2: rule-usage profiling - what does the grammar ACTUALLY do?
'  The parse-failure log captures what users couldn't say; these
'  counters capture what the loaded grammar really does: which phrase
'  rules fire (every firing, nested included), which structural form
'  claimed each top-level sentence, and - just as important - which
'  loaded rules NEVER fire. Counters are in-memory and session-scoped,
'  keyed by pattern text so they survive grammar resets and reloads;
'  vocab test: proofs are deliberately not counted (a proof firing is
'  evidence of loading, not of usage). Consumers on the roadmap: G1
'  merges the rule families that fire and reviews the dead ones for
'  deletion; U2 pins the most-used sentences on evidence; latency work
'  names its dead weight. Persistence across sessions is deliberately
'  deferred - the pilot's corpus pipeline (M2a) is the durable home
'  for usage data; this instrument is for the dev loop first.
'
'    Debug.Print EnglishRuleUsage()      the report, sorted by count
'    EnglishRuleUsageCount("rule: set {v:var} to {e:expr}")
'    EnglishRuleUsageReset               start a fresh measurement
' =====================================================================

' =====================================================================
'  S4.2: the resolve check - S4's re-scope after the probe's empirical
'  death. We generate the code and we know every name it may call, so
'  helper resolution needs no VBA compiler: every called name that
'  starts with "vla" must be either defined in the program's own VLA
'  (its subs/functions, pass 1) or provided by the runtime
'  (VlaHelperManifest, cross-module - the S3 precedent). Anything
'  else is the V5.3 class - a dialect template's typo'd helper, a
'  program expecting a newer Frazaro - and refuses at CHECK, naming
'  the helper and its line, before anything is injected. Names not
'  vla-prefixed pass: VBA builtins and user-word calls belong to
'  VBA's own judgement (the recorded residual). String literals are
'  skipped; (at-line N ...) markers thread the source line to the
'  refusal - and source line IS sheet row since S3.3. An empty
'  manifest means "cannot verify" and passes: the check is auxiliary
'  machinery and fails softer than the Run it guards.
' =====================================================================
Public Function EnglishResolveCheck(ByVal vlaText As String, _
                                    Optional ByRef lineOut As Long) As String
    lineOut = 0
    Dim manifest As String
    ' V8: one manifest read per session (see mManifestCache).
    If Len(mManifestCache) = 0 Then
        On Error Resume Next
        mManifestCache = VlaHelperManifest()
        On Error GoTo 0
    End If
    manifest = mManifestCache
    If Len(Trim$(manifest)) = 0 Then Exit Function

    Dim defined As String
    defined = " "
    Dim n As Long
    n = Len(vlaText)
    Dim pass As Long
    For pass = 1 To 2
        Dim inLit As Boolean
        inLit = False
        Dim curLine As Long
        curLine = 0
        Dim i As Long
        i = 1
        Do While i <= n
            Dim c As String
            c = Mid$(vlaText, i, 1)
            If inLit Then
                If c = "\" Then
                    i = i + 2
                ElseIf c = """" Then
                    inLit = False
                    i = i + 1
                Else
                    i = i + 1
                End If
            ElseIf c = """" Then
                inLit = True
                i = i + 1
            ElseIf c = "(" Then
                Dim j As Long
                j = i + 1
                Do While j <= n And Mid$(vlaText, j, 1) = " "
                    j = j + 1
                Loop
                Dim h As Long
                h = j
                Do While h <= n
                    If InStr(" ()" & vbCr & vbLf & vbTab, Mid$(vlaText, h, 1)) > 0 Then Exit Do
                    h = h + 1
                Loop
                Dim head As String
                head = Mid$(vlaText, j, h - j)
                Dim keyName As String
                keyName = VLA_Identity.Fold(Replace(Replace(head, "-", ""), "_", ""))
                If pass = 1 Then
                    If keyName = "sub" Or keyName = "function" Then
                        ' the next symbol is a defined name
                        Dim j2 As Long
                        j2 = h
                        Do While j2 <= n And Mid$(vlaText, j2, 1) = " "
                            j2 = j2 + 1
                        Loop
                        Dim h2 As Long
                        h2 = j2
                        Do While h2 <= n
                            If InStr(" ()" & vbCr & vbLf & vbTab, Mid$(vlaText, h2, 1)) > 0 Then Exit Do
                            h2 = h2 + 1
                        Loop
                        If h2 > j2 Then
                            defined = defined & VLA_Identity.Fold(Replace(Replace(Mid$(vlaText, j2, h2 - j2), "-", ""), "_", "")) & " "
                        End If
                    ElseIf keyName = "defmacro" Then
                        ' L4.1 (maiden-run incident, owner-caught): a
                        ' carried macro DEFINES its name for resolution
                        ' - the call expands away at transpile and can
                        ' never miss at runtime - but pass 1 predated
                        ' L4 and only read sub/function, so the first
                        ' vocabulary macro drew the missing-helper
                        ' dialog at Reload. This engine's defmacro puts
                        ' the name INSIDE the signature list -
                        ' (defmacro (name params...) ...) - the L2.1
                        ' dialect lesson, third appearance, so skip to
                        ' the inner "(" before reading the name. Also
                        ' covers a raw-row defmacro a program carries
                        ' itself (L0), for free.
                        Dim j3 As Long
                        j3 = h
                        Do While j3 <= n And InStr(" " & vbCr & vbLf & vbTab, Mid$(vlaText, j3, 1)) > 0
                            j3 = j3 + 1
                        Loop
                        If Mid$(vlaText, j3, 1) = "(" Then
                            j3 = j3 + 1
                            Do While j3 <= n And InStr(" " & vbCr & vbLf & vbTab, Mid$(vlaText, j3, 1)) > 0
                                j3 = j3 + 1
                            Loop
                            Dim h3 As Long
                            h3 = j3
                            Do While h3 <= n
                                If InStr(" ()" & vbCr & vbLf & vbTab, Mid$(vlaText, h3, 1)) > 0 Then Exit Do
                                h3 = h3 + 1
                            Loop
                            If h3 > j3 Then
                                defined = defined & VLA_Identity.Fold(Replace(Replace(Mid$(vlaText, j3, h3 - j3), "-", ""), "_", "")) & " "
                            End If
                        End If
                    End If
                Else
                    If keyName = "atline" Then
                        curLine = CLng(Val(Mid$(vlaText, h)))
                    ElseIf Left$(keyName, 3) = "vla" Then
                        If InStr(defined, " " & keyName & " ") = 0 And _
                           InStr(manifest, " " & keyName & " ") = 0 Then
                            lineOut = curLine
                            EnglishResolveCheck = head
                            Exit Function
                        End If
                    End If
                End If
                If h > i Then i = h Else i = i + 1
            Else
                i = i + 1
            End If
        Loop
    Next
End Function

Public Function EnglishRuleUsage() As String
    EnsureInit
    Dim r As String
    r = "===== RULE USAGE (this session) =====" & vbCrLf
    Dim n As Long
    If Not mUsageKeys Is Nothing Then n = mUsageKeys.Count
    If n = 0 Then
        r = r & "  (no sentences translated yet this session)" & vbCrLf
    Else
        ' Copy to arrays and selection-sort descending; ties keep
        ' first-seen order (strict > leaves earlier entries in place).
        Dim keys() As String
        Dim cnts() As Long
        ReDim keys(1 To n)
        ReDim cnts(1 To n)
        Dim i As Long, j As Long
        Dim found As Boolean
        For i = 1 To n
            keys(i) = mUsageKeys.Item(i)
            cnts(i) = CLng(CollGet(mUsageCounts, VLA_Identity.Fold(keys(i)), found))
        Next
        Dim best As Long, tk As String, tc As Long
        For i = 1 To n - 1
            best = i
            For j = i + 1 To n
                If cnts(j) > cnts(best) Then best = j
            Next
            If best <> i Then
                tk = keys(i): keys(i) = keys(best): keys(best) = tk
                tc = cnts(i): cnts(i) = cnts(best): cnts(best) = tc
            End If
        Next
        r = r & "fired:" & vbCrLf
        For i = 1 To n
            r = r & Right$("      " & cnts(i), 6) & "  " & keys(i) & vbCrLf
        Next
    End If

    ' Dead-rule view: every currently loaded phrase rule with no
    ' count. (Counts for rules no longer loaded still show above -
    ' text-keyed counters outlive the rules they measured.)
    Dim deadCount As Long
    Dim dead As String
    Dim k As Long
    For k = 1 To mPatTexts.Count
        If Not UsageHasKey("rule: " & mPatTexts.Item(k)) Then
            dead = dead & "        " & mPatTexts.Item(k) & vbCrLf
            deadCount = deadCount + 1
        End If
    Next
    If deadCount > 0 Then
        r = r & "never fired (" & deadCount & " of " & mPatTexts.Count & _
                " loaded phrase rules):" & vbCrLf & dead
    Else
        r = r & "  (every loaded phrase rule has fired)" & vbCrLf
    End If
    r = r & "  (vocab test: proofs are not counted; counters survive" & vbCrLf & _
            "   Reload - EnglishRuleUsageReset starts a fresh measurement)" & vbCrLf
    EnglishRuleUsage = r
End Function

' Programmatic accessor (the self-test's window, and any future tool's):
' the full display key, e.g. "rule: set {v:var} to {e:expr}" or
' "form: the Create declaration". Unknown keys read 0.
Public Function EnglishRuleUsageCount(ByVal key As String) As Long
    If mUsageCounts Is Nothing Then Exit Function
    Dim found As Boolean
    Dim v As Variant
    v = CollGet(mUsageCounts, VLA_Identity.Fold(key), found)
    If found Then EnglishRuleUsageCount = CLng(v)
End Function

Public Sub EnglishRuleUsageReset()
    Set mUsageCounts = Nothing
    Set mUsageKeys = Nothing
End Sub

' The counting primitive. Keys are lowercased (every text comparison
' in this project is case-insensitive; patterns arrive in the case
' their author wrote). Collections cannot update in place - remove
' and re-add under the same key.
Private Sub BumpUsage(ByVal key As String)
    If mUsageSuspended Then Exit Sub
    If mUsageCounts Is Nothing Then
        Set mUsageCounts = New Collection
        Set mUsageKeys = New Collection
    End If
    Dim disp As String
    disp = key                       ' the report shows author's case
    key = VLA_Identity.Fold(key)                ' the count key is case-blind
    Dim found As Boolean
    Dim cur As Variant
    cur = CollGet(mUsageCounts, key, found)
    If found Then
        mUsageCounts.Remove key
        mUsageCounts.Add CLng(cur) + 1, key
    Else
        mUsageCounts.Add 1&, key
        mUsageKeys.Add disp
    End If
End Sub

Private Function UsageHasKey(ByVal key As String) As Boolean
    If mUsageCounts Is Nothing Then Exit Function
    UsageHasKey = CollHasKey(mUsageCounts, VLA_Identity.Fold(key))
End Function

' All accumulated grammar warnings from rule registration, or a clean
' bill of health. Print after loading vocabularies:
'   Debug.Print EnglishLintReport()
Public Function EnglishLintReport() As String
    EnsureInit
    If mLintWarnings.Count = 0 Then
        EnglishLintReport = "No grammar warnings."
        Exit Function
    End If
    Dim r As String
    Dim e As Variant
    For Each e In mLintWarnings
        r = r & "WARNING: " & e & vbCrLf
    Next
    EnglishLintReport = r
End Function

' Remove all loaded/added rules, keeping only the built-in prelude.
' Call before reloading vocabulary files so reloads are idempotent.
Public Sub EnglishResetGrammar()
    EnsureInit
    Do While mPatItems.Count > mPreludeCount
        mPatItems.Remove mPatItems.Count
        mPatForms.Remove mPatForms.Count
        mPatTexts.Remove mPatTexts.Count
        mPatSigs.Remove mPatSigs.Count
        mPatSources.Remove mPatSources.Count
    Loop
    mLoadSource = ""
    mVocabMacros = ""
    Set mVocabMacroNames = New Collection
    Set VLA_English.mKeywordAlias = New Collection     ' LX5.1: no aliases survive a reset
    mTestsRun = 0                         ' U.10
    mFailsRun = 0
    mLastLoadSource = ""
    VlaAproposCarry "", ""                ' L11.1/APROPOSPLUS: apropos back to prelude-only
    ' (Overrides never reach the prelude zone - the built-in core is
    ' not overridable, so popping to mPreludeCount restores exactly
    ' the built-ins, provenance included.)
    Set mLintWarnings = New Collection
    Set mRuleExamples = New Collection    ' indices shift on reset
    Set mRuleTestCounts = New Collection  ' AS.1: same reason - indices shift
    RegisterBuiltinFuncWords              ' drop user-added function words
    mDspValid = False                     ' V7: indices shifted here too
    RebuildSigOwner                       ' G3 perf: same reason - popping
                                           ' back to mPreludeCount leaves
                                           ' mSigOwner's own entries for
                                           ' every removed rule stale (live
                                           ' bug, caught this way: a LATER
                                           ' registration's duplicate check
                                           ' looked up a stale index and
                                           ' mPatTexts.Item(that index)
                                           ' subscripted out of range,
                                           ' since the freshly-reset table
                                           ' doesn't have that many rows
                                           ' yet) - cheap even called this
                                           ' often, since it only ever
                                           ' walks back down to the small,
                                           ' fixed prelude
End Sub

Public Function EnglishRuleCount() As Long
    EnsureInit
    EnglishRuleCount = mPatItems.Count
End Function

' Read a text file as UTF-8 with an ANSI fallback. (Kept private and
' duplicated from VLA_Loader so this module stays self-contained.)
Private Function VocabReadFile(ByVal filePath As String) As String
    On Error GoTo ansiFallback
    Dim st As Object
    ' V8: reuse the session's stream (see mVocabStream).
    If mVocabStream Is Nothing Then Set mVocabStream = CreateObject("ADODB.Stream")
    Set st = mVocabStream
    st.Type = 2                              ' adTypeText
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile filePath
    VocabReadFile = st.ReadText(-1)
    st.Close
    Exit Function

ansiFallback:
    ' A failed UTF-8 read may leave the stream open or wedged - drop
    ' the cached object entirely; the next call recreates it.
    On Error Resume Next
    mVocabStream.Close
    On Error GoTo 0
    Set mVocabStream = Nothing
    On Error GoTo 0
    Dim f As Integer
    Dim b() As Byte
    f = FreeFile
    Open filePath For Binary Access Read As #f
    If LOF(f) = 0 Then
        Close #f
        VocabReadFile = ""
        Exit Function
    End If
    ReDim b(0 To LOF(f) - 1)
    Get #f, , b
    Close #f
    VocabReadFile = StrConv(b, vbUnicode)
End Function

' =====================================================================
'  Diagnostics: best-partial-match error messages
' =====================================================================

' Record a rule-attempt failure at token position p, keeping the
' attempt that got furthest. The earliest-registered rule to reach
' mBestProgress leads (mBestExpect, the grammar's own precedence
' order, unchanged) - LE.2: ruleIdx (the mPatItems/mPatTexts index of
' the rule that was attempting the match) rides alongside in
' mBestTieIdx, growing past that one leader whenever a LATER rule
' reaches the exact same furthest position (a genuine tie: same
' prefix, different next word), so DidYouMean can offer every rule
' that got equally far, not just whichever was registered first.
Private Sub NoteFail(ByVal p As Long, ByVal expected As String, ByVal ruleIdx As Long)
    If p > mBestProgress Then
        mBestProgress = p
        mBestExpect = expected
        Set mBestTieIdx = New Collection
        mBestTieIdx.Add ruleIdx
    ElseIf p = mBestProgress And mBestTieIdx.Count < 3 Then
        Dim tv As Variant
        For Each tv In mBestTieIdx
            If CLng(tv) = ruleIdx Then Exit Sub
        Next
        mBestTieIdx.Add ruleIdx
    End If
End Sub

' Human description of a slot category, for error messages.
' G1: expected-token description for an alternation slot -
' "one of 'into'/'in'" - used by the near-miss reporting the same
' way SlotDesc serves the category slots.
Private Function AltDesc(ByVal cat As String) As String
    Dim alts() As String
    Dim ai As Long
    Dim r As String
    Dim sf As Collection
    Dim f As Variant
    alts = Split(cat, "|")
    For ai = LBound(alts) To UBound(alts)
        ' G1.1: a stem/suffix branch lists BOTH surfaces - the reader
        ' of a near-miss sees words, never the notation.
        Set sf = SurfaceForms(VLA_Identity.Fold(alts(ai)))
        For Each f In sf
            If Len(r) > 0 Then r = r & "/"
            r = r & "'" & CStr(f) & "'"
        Next
    Next
    AltDesc = "one of " & r
End Function

' =====================================================================
'  G2: reference shape validators. Check-time can judge SHAPE, never
'  existence - a well-formed range may still be missing at runtime,
'  and that stays runtime's step-numbered job. Quoted references
'  never arrive here (quotes assert at the binding site). A bang
'  splits an optional sheet qualifier (data!b2, and the chained
'  data!a1:b10) from the reference; 'Q1 Data'!A1 forms are string-
'  tagged by the tokenizer and thus assert as quoted. The honest
'  limits, recorded: a 1-3 letter word passes the column shape
'  ("cat" is a legal-looking column), and :sheet validates nothing -
'  any word can name a sheet, so that category exists for the touch
'  data below, not for refusals.
' =====================================================================
Private Function RefShapeOk(ByVal cat As String, ByVal tok As String) As Boolean
    Dim tail As String
    Dim bp As Long
    tail = tok
    If cat = "cell" Or cat = "range" Then
        bp = InStrRev(tail, "!")
        If bp > 0 Then
            If bp = 1 Or bp = Len(tail) Then Exit Function
            tail = Mid$(tail, bp + 1)
        End If
    End If
    Select Case cat
        Case "cell": RefShapeOk = IsCellPart(tail)
        Case "range": RefShapeOk = IsRangePart(tail)
        Case "column": RefShapeOk = IsColLetters(tail)
        Case "sheet": RefShapeOk = True
        Case "color": RefShapeOk = IsColorWord(tail)
    End Select
End Function

Private Function IsColLetters(ByVal s As String) As Boolean
    Dim i As Long
    If Len(s) < 1 Or Len(s) > 3 Then Exit Function
    For i = 1 To Len(s)
        If Not Mid$(s, i, 1) Like "[a-z]" Then Exit Function
    Next
    IsColLetters = True
End Function

Private Function IsRowDigits(ByVal s As String) As Boolean
    Dim i As Long
    If Len(s) < 1 Or Len(s) > 7 Then Exit Function
    For i = 1 To Len(s)
        If Not Mid$(s, i, 1) Like "[0-9]" Then Exit Function
    Next
    IsRowDigits = True
End Function

' Letters then digits - A1, aa100 - both parts present and sized.
Private Function IsCellPart(ByVal s As String) As Boolean
    Dim i As Long
    i = 1
    Do While i <= Len(s)
        If Mid$(s, i, 1) Like "[a-z]" Then i = i + 1 Else Exit Do
    Loop
    If i = 1 Or i > Len(s) Then Exit Function
    IsCellPart = IsColLetters(Left$(s, i - 1)) And IsRowDigits(Mid$(s, i))
End Function

' A cell IS a range; a span pairs two parts of ONE kind - cells
' (A1:C50), column letters (A:C), or row numbers (1:5).
Private Function IsRangePart(ByVal s As String) As Boolean
    Dim cp As Long
    cp = InStr(s, ":")
    If cp = 0 Then
        IsRangePart = IsCellPart(s)
        Exit Function
    End If
    Dim a As String, b As String
    a = Left$(s, cp - 1)
    b = Mid$(s, cp + 1)
    If InStr(b, ":") > 0 Then Exit Function
    IsRangePart = (IsCellPart(a) And IsCellPart(b)) _
               Or (IsColLetters(a) And IsColLetters(b)) _
               Or (IsRowDigits(a) And IsRowDigits(b))
End Function

' G2: the machine-readable "what does this rule touch" surface,
' pre-registered for U1's Rehearse (one band later). One line per
' loaded rule carrying typed reference slots: the pattern, then its
' slotname:category pairs. Shape-blind rules are omitted - the
' report answers "which rules touch the workbook, and through what".
'   ?EnglishListRefSlots   in the Immediate window prints it.
Public Function EnglishListRefSlots() As String
    EnsureInit
    Dim r As String
    Dim i As Long
    Dim items As Collection
    Dim it As Variant
    Dim t As String, sn As String, ct As String
    Dim lineTxt As String
    For i = 1 To mPatItems.Count
        Set items = mPatItems.Item(i)
        lineTxt = ""
        For Each it In items
            t = it
            If IsSlotTok(t, sn, ct) Then
                Select Case ct
                    Case "range", "cell", "column", "sheet", "color", "path"
                        If Len(lineTxt) > 0 Then lineTxt = lineTxt & ", "
                        lineTxt = lineTxt & sn & ":" & ct
                End Select
            End If
        Next
        If Len(lineTxt) > 0 Then
            r = r & mPatTexts.Item(i) & "  ->  " & lineTxt & vbCrLf
        End If
    Next
    EnglishListRefSlots = r
End Function

' G3: which file a rule came from - "(built-in)", "(added directly)",
' or the phrasebook's source name, with " (overrides <file>)" when an
' override: replaced an earlier rule. Explain prints it, so "which
' file won this sentence" is one Explain away - the load-order
' question answered per sentence instead of by archaeology.
' G3: the provenance of the rule that matched the last translated
' sentence - the machine-readable half of Explain's "[from ...]",
' for pins and future tooling (the phrasebook manager's "which file
' won this rule" column is this function walked over a program).
Public Function EnglishRuleSource() As String
    If mLastRuleIdx > 0 Then EnglishRuleSource = RuleSourceOf(mLastRuleIdx)
End Function

Private Function RuleSourceOf(ByVal idx As Long) As String
    If idx >= 1 And idx <= mPatSources.Count Then
        RuleSourceOf = mPatSources.Item(idx)
    Else
        RuleSourceOf = "(unknown)"
    End If
End Function

' Render one token back into user-facing form.
Private Function RenderTok(ByVal t As String) As String
    If IsStrTok(t) Then
        RenderTok = """" & Mid$(t, 2) & """"
    Else
        RenderTok = t
    End If
End Function

' Render tokens first..last as readable text.
Private Function RenderTokens(toks() As String, ByVal first As Long, ByVal last As Long) As String
    Dim r As String
    Dim i As Long
    For i = first To last
        If Len(r) > 0 Then r = r & " "
        r = r & RenderTok(TokAt(toks, i))
    Next
    RenderTokens = r
End Function

' Build the failure message for an ununderstood sentence, using the
' best partial match if any rule made real progress, and suggesting
' registered phrases that start with the same first word.
Private Function BuildParseError(toks() As String, ByVal pos As Long) As String
    Dim msg As String
    Dim found As String

    If mBestProgress > pos Then
        found = TokAt(toks, mBestProgress)
        If Len(found) = 0 Or found = "." Then
            found = "the end of the sentence"
        Else
            found = "'" & RenderTok(found) & "'"
        End If
        msg = "I understood '" & RenderTokens(toks, pos, mBestProgress - 1) & _
              "' - then I expected " & mBestExpect & " but found " & found & "."
    Else
        msg = "Don't understand: '" & SentenceContext(toks, pos) & "'"
    End If

    Dim suggest As String
    Dim bestTies As Collection
    If mBestProgress > pos Then Set bestTies = mBestTieIdx
    suggest = DidYouMean(PeekWord(toks, pos), bestTies)
    If Len(suggest) > 0 Then
        msg = msg & " Did you mean: " & suggest
    ElseIf mPatItems.Count <= mPreludeCount Then
        ' No sentence starts with this word AND nothing beyond the
        ' prelude is registered: almost always a missing/lost
        ' EnglishLoadVocabulary (VBA state resets clear the grammar).
        msg = msg & " Only the " & mPatItems.Count & " built-in phrases are loaded - was EnglishLoadVocabulary run in this session?"
    Else
        msg = msg & " No loaded sentence starts with '" & PeekWord(toks, pos) & _
              "' - click the 'Known Sentences' button to see all " & mPatItems.Count & _
              " sentences this program understands (or, in VBA, print the list with ?EnglishListPhrases in the Immediate window)."
    End If
    Dim ep As Long
    ep = pos
    If mBestProgress > pos Then ep = mBestProgress
    msg = msg & LineTag(ep)
    BuildParseError = msg
End Function

' Up to three registered phrase patterns to suggest. LE.2: bestTies
' (every rule NoteFail recorded as tied for the furthest progress into
' THIS sentence, earliest-registered first, wired through from
' BuildParseError) leads when present - a relevance-scored near-miss
' list, not a guess, and no longer a single pick: two rules that
' genuinely diverge at the exact same token (same prefix, different
' next word) both surface, instead of only the earlier-registered one
' winning silently. The first-word scan that used to be the whole
' function now only fills any slots ties left open, skipping every
' tied index so none is ever repeated.
Private Function DidYouMean(ByVal firstWord As String, bestTies As Collection) As String
    Dim r As String
    Dim i As Long, hits As Long
    Dim items As Collection
    Dim tv As Variant

    If Not bestTies Is Nothing Then
        For Each tv In bestTies
            If hits > 0 Then r = r & "  |  "
            r = r & "'" & mPatTexts.Item(CLng(tv)) & "'"
            hits = hits + 1
        Next
    End If

    If Len(firstWord) > 0 And hits < 3 Then
        For i = 1 To mPatItems.Count
            If Not TiedAt(bestTies, i) Then
                Set items = mPatItems.Item(i)
                If items.Item(1) = firstWord Then
                    If hits > 0 Then r = r & "  |  "
                    r = r & "'" & mPatTexts.Item(i) & "'"
                    hits = hits + 1
                    If hits = 3 Then Exit For
                End If
            End If
        Next
    End If

    DidYouMean = r
End Function

' Whether rule idx is already in DidYouMean's tie list, so the
' first-word fallback never repeats a suggestion the tie already gave.
Private Function TiedAt(ties As Collection, ByVal idx As Long) As Boolean
    If ties Is Nothing Then Exit Function
    Dim tv As Variant
    For Each tv In ties
        If CLng(tv) = idx Then
            TiedAt = True
            Exit Function
        End If
    Next
End Function

' The structural shapes: they live in the parser, not the rule list,
' so they are written out by hand here - shared by EnglishListPhrases
' (the flat cheat-sheet text) and EnglishPhraseRows (the structured
' Template/Example table, LE.1's thin slice) so the two never drift
' apart. Kept honest by a self-test that asserts each one appears in
' EnglishListPhrases's own output, so adding a structural form without
' updating this list fails a named test rather than silently hiding
' the form from users.
Private Function PhraseBuiltinShapes() As Variant
    PhraseBuiltinShapes = Array( _
        "Create a number/text/value/list/lookup called <name>.", _
        "If <condition>, <sentence>.", _
        "If <condition>: ...  Otherwise, if <condition>: ...  Otherwise: ...", _
        "When <value> is <case>: ...  When it is <case>: ...  Otherwise: ...", _
        "Try: ...  If that fails: ...", _
        "Repeat <n> times: ...        (the word ""counter"" runs 1 to N)", _
        "Repeat until <condition>: ...", _
        "While <condition>: ...", _
        "Count <name> from <a> to <b>: ...   (also ""down from"")", _
        "For each <x> in <collection>: ...", _
        "Stop the loop.   Done.", _
        "To <name>: ...  then  <name>.       (your own actions; ""with"" adds parameters)", _
        "To <name> of <param>: ...  Give back <value>.   (then ""<name> of x"" works anywhere a value goes)", _
        "To [get] <name> using <p> of <default> and <q>: ...   (named parameters; call: <name> using <p> of 5, defaults fill the rest)", _
        "To get <name>: ...  Give back <value>.   (no parameters; then <name> alone is a value anywhere)", _
        "Define <name> as <fixed value>.")
End Function

' A "what can I say" cheat sheet: every registered phrase (prelude +
' loaded), grouped by first word in first-encounter order. Generated
' from the live grammar, so it is never out of date.
Public Function EnglishListPhrases() As String
    EnsureInit
    Dim r As String
    r = "built into the language (every vocabulary shares these):" & vbCrLf
    Dim shapes As Variant, si As Long
    shapes = PhraseBuiltinShapes()
    For si = LBound(shapes) To UBound(shapes)
        r = r & "    " & shapes(si) & vbCrLf
    Next
    r = r & vbCrLf
    Dim firstWords As New Collection
    Dim items As Collection
    Dim i As Long
    Dim f As String
    ' Collect distinct first words in registration order.
    For i = 1 To mPatItems.Count
        Set items = mPatItems.Item(i)
        f = items.Item(1)
        On Error Resume Next
        firstWords.Add f, f
        On Error GoTo 0
    Next
    ' Emit each group.
    Dim fw As Variant
    Dim exv As Variant
    Dim exf As Boolean
    For Each fw In firstWords
        r = r & CStr(fw) & ":" & vbCrLf
        For i = 1 To mPatItems.Count
            Set items = mPatItems.Item(i)
            If items.Item(1) = CStr(fw) Then
                r = r & "    " & mPatTexts.Item(i) & vbCrLf
                exf = False
                exv = CollGet(mRuleExamples, "r" & i, exf)
                If exf Then r = r & "        e.g. " & CStr(exv) & vbCrLf
            End If
        Next
    Next
    r = r & "values (usable anywhere a value goes):" & vbCrLf & EnglishListFunctionWords()
    EnglishListPhrases = r
End Function

' LE.1 (thin slice): the same walk EnglishListPhrases performs, over
' the same live grammar state, handed back as parallel rows - template,
' worked example (from a rule's own passing test: sentence, "" when
' none), and whether the row is a section/group header - instead of
' one preformatted text blob, so a caller can render a real two-column
' table (VLA_IDE.bas's "What can I say?" button) instead of re-parsing
' text back apart. Deliberately its own pass rather than a refactor of
' EnglishListPhrases into a shared row-then-format helper: that walk is
' small, stable, and self-test-pinned (VLA_Tests.bas's "listing covers
' structural forms" fragment check) - duplicating it here is lower-risk
' than restructuring a working, pinned function.
Public Sub EnglishPhraseRows(ByRef outTemplates As Collection, ByRef outExamples As Collection, _
                              ByRef outIsHeader As Collection)
    EnsureInit
    Set outTemplates = New Collection
    Set outExamples = New Collection
    Set outIsHeader = New Collection

    AddPhraseRow outTemplates, outExamples, outIsHeader, _
        "built into the language (every vocabulary shares these)", "", True
    Dim shapes As Variant, si As Long
    shapes = PhraseBuiltinShapes()
    For si = LBound(shapes) To UBound(shapes)
        AddPhraseRow outTemplates, outExamples, outIsHeader, CStr(shapes(si)), "", False
    Next

    Dim firstWords As New Collection
    Dim items As Collection
    Dim i As Long
    Dim f As String
    For i = 1 To mPatItems.Count
        Set items = mPatItems.Item(i)
        f = items.Item(1)
        On Error Resume Next
        firstWords.Add f, f
        On Error GoTo 0
    Next

    Dim fw As Variant
    Dim exv As Variant
    Dim exf As Boolean
    For Each fw In firstWords
        AddPhraseRow outTemplates, outExamples, outIsHeader, CStr(fw), "", True
        For i = 1 To mPatItems.Count
            Set items = mPatItems.Item(i)
            If items.Item(1) = CStr(fw) Then
                exf = False
                exv = CollGet(mRuleExamples, "r" & i, exf)
                AddPhraseRow outTemplates, outExamples, outIsHeader, _
                             CStr(mPatTexts.Item(i)), IIf(exf, CStr(exv), ""), False
            End If
        Next
    Next

    AddPhraseRow outTemplates, outExamples, outIsHeader, _
        "values (usable anywhere a value goes)", "", True
    Dim fnLines() As String
    fnLines = Split(EnglishListFunctionWords(), vbCrLf)
    Dim vi As Long
    For vi = LBound(fnLines) To UBound(fnLines)
        If Len(Trim$(fnLines(vi))) > 0 Then _
            AddPhraseRow outTemplates, outExamples, outIsHeader, Trim$(fnLines(vi)), "", False
    Next
End Sub

Private Sub AddPhraseRow(templates As Collection, examples As Collection, isHeader As Collection, _
                          ByVal tmpl As String, ByVal example As String, ByVal hdr As Boolean)
    templates.Add tmpl
    examples.Add example
    isHeader.Add hdr
End Sub

' =====================================================================
'  Translation-time call checking. Named-argument mismatches and
'  missing required parameters are VBA COMPILE errors - they happen
'  before any code runs, so the runtime step handler can never catch
'  them. The English layer sees definitions and calls together, so it
'  refuses the translation instead, with the sentence quoted back.
' =====================================================================

Private Sub RecordCall(ByVal name As String, args As Collection, ByVal sentence As String, ByVal lineNo As Long)
    If mCallNames Is Nothing Then Exit Sub   ' e.g. vocab tests outside a translation
    mCallNames.Add VLA_Identity.Fold(name)
    mCallArgs.Add args
    mCallTexts.Add sentence
    mCallLines.Add lineNo
End Sub

Private Function FindAction(ByVal name As String) As Long
    Dim i As Long
    For i = 1 To mActNames.Count
        If mActNames.Item(i) = name Then
            FindAction = i
            Exit Function
        End If
    Next
End Function

Private Function InNameList(col As Collection, ByVal name As String) As Boolean
    Dim e As Variant
    For Each e In col
        If CStr(e) = name Then
            InNameList = True
            Exit Function
        End If
    Next
End Function

Private Function JoinNames(col As Collection) As String
    Dim e As Variant, r As String
    For Each e In col
        If Len(r) > 0 Then r = r & ", "
        r = r & e
    Next
    If Len(r) = 0 Then r = "(none)"
    JoinNames = r
End Function

Private Sub ValidateActionCalls()
    Dim i As Long, j As Long, idx As Long
    Dim params As Collection, reqs As Collection, args As Collection
    Dim a As Variant
    For i = 1 To mCallNames.Count
        idx = FindAction(CStr(mCallNames.Item(i)))
        If idx = 0 Then
            ' Not defined in this file: might live in another module,
            ' so warn rather than refuse.
            mLintWarnings.Add "'" & mCallTexts.Item(i) & "' calls '" & mCallNames.Item(i) & _
                "', which is not defined in this file - if it is not defined elsewhere, running will fail"
        Else
            Set params = mActParams.Item(idx)
            Set reqs = mActReq.Item(idx)
            Set args = mCallArgs.Item(i)
            For Each a In args
                If Not InNameList(params, CStr(a)) Then
                    mErrLine = CLng(mCallLines.Item(i))
                    VLA_Messages.RaiseMsg "english-call-unknown-param", "call", mCallTexts.Item(i), "action", mCallNames.Item(i), "param", a, "list", JoinNames(params), "loc", LineSuf(mErrLine)
                End If
            Next
            For j = 1 To params.Count
                If reqs.Item(j) = "1" Then
                    If Not InNameList(args, CStr(params.Item(j))) Then
                        mErrLine = CLng(mCallLines.Item(i))
                        VLA_Messages.RaiseMsg "english-call-missing-param", "call", mCallTexts.Item(i), "action", mCallNames.Item(i), "param", params.Item(j), "loc", LineSuf(mErrLine)
                    End If
                End If
            Next
        End If
    Next
End Sub

' =====================================================================
'  The safe front door for buttons: reset, load vocabularies, read,
'  translate, compile, run - with every failure presented as a plain
'  message box (which has no Debug button, so no path into the VBE).
'  Wire worksheet buttons to a one-line sub calling this.
' =====================================================================
Public Sub EnglishRunProgram(ByVal programPath As String, _
                             ByVal moduleName As String, _
                             ParamArray vocabPaths() As Variant)
    On Error GoTo failed
    EnglishResetGrammar
    Dim i As Long
    For i = LBound(vocabPaths) To UBound(vocabPaths)
        EnglishLoadVocabulary CStr(vocabPaths(i))
    Next
    ' Lint warnings go to the Immediate window: visible to a developer,
    ' invisible to a button user.
    Debug.Print EnglishLintReport()

    If Len(Dir$(programPath)) = 0 Then
        VLA_Messages.RaiseMsg "english-program-file-not-found", "path", programPath
    End If
    EnglishCompileToModule VocabReadFile(programPath), moduleName
    Application.Run "'" & ActiveWorkbook.Name & "'!" & moduleName & ".main"
    Exit Sub

failed:
    ' S3: through the message seam (VLA_Runtime ships inside the
    ' add-in, so the call resolves in every world this Sub runs in).
    VlaShowError Err.Description
End Sub

' =====================================================================
'  Translate to VLA / Translate to VBA - the same auditability
'  G-EXPANDER gave vocabulary files, for FILE-BASED PROGRAMS. Today the
'  only way to see what a program actually compiles to is
'  EnglishRunProgram itself - which compiles AND RUNS, leaving no
'  readable trace behind. These two are EnglishRunProgram's own
'  sibling, same shape, minus the compile-and-run tail: reset, load
'  vocab, translate, WRITE - never execute. Targets English-prose
'  program files (instructions.txt-shaped) - not scripts/alonzo.vla,
'  which is a raw-VLA (include ...) library with no main, not a
'  standalone runnable program.
'  Named "Translate", not "Export" (owner correction, renamed from
'  EnglishExportVla/EnglishExportVba): this pipeline never touches
'  workbook content at all - a picked SOURCE file becomes an OUTPUT
'  file, purely external. "Export" implies something moving OUT of the
'  workbook (which EnglishExpandedVocabularyText's own export genuinely
'  does - it exports the currently-loaded LIVE grammar state); this is
'  a pure file-to-file transformation, the same shape as translating a
'  document from one language to another.
' =====================================================================

' <folder>\<base><ext> - instructions.txt -> instructions.vla, no infix
' (GoldenPathFor's own "_golden" infix is that function's own concern,
' not this one's - a translation is the primary artifact, not a
' golden). Public: also the SUGGESTED filename for VLA_IDE.bas's own
' Save-As dialog (TranslateViaRibbon) - a genuine cross-module caller
' now, not just this module's own default-path computation.
Public Function TranslatedPathFor(ByVal programPath As String, ByVal ext As String) As String
    Dim sep As String
    sep = Application.PathSeparator
    Dim i As Long
    i = InStrRev(programPath, sep)
    Dim folder As String, base As String
    folder = Left$(programPath, i)
    base = Mid$(programPath, i + 1)
    Dim d As Long
    d = InStrRev(base, ".")
    If d > 0 Then base = Left$(base, d - 1)
    TranslatedPathFor = folder & base & ext
End Function

' Rule-12 duplicate of the Open.../Print#/Close writer (VLA_Build.bas's
' WriteTextFile is the model, including its Kill-if-exists guard) - its
' own small copy rather than reusing WriteExpandedSibling's inline
' writer, whose own failure message is tailored to its own caller.
Private Sub WriteTextFileVla(ByVal filePath As String, ByVal content As String)
    If Len(Dir$(filePath)) > 0 Then Kill filePath
    Dim f As Integer
    f = FreeFile
    Open filePath For Output As #f
    Print #f, content
    Close #f
End Sub

' Both return True on success - a workbook's own button, wired the same
' way EnglishRunProgram's own callers are, can ignore the return value
' entirely (self-contained: shows its own message box on failure,
' matching EnglishRunProgram's "safe front door" contract exactly). The
' return value exists for a caller that ALSO wants its own success
' confirmation (VLA_IDE.bas's TranslateViaRibbon) - without it, that
' caller could not tell an internally-swallowed failure (which already
' showed its own error box) from a real success, and would show a
' bogus "Translated" confirmation on top of the error.
' outPathOverride (owner-reported gap): pass "" for a workbook's own
' button - gets the auto-derived <base>.vla/.vba sibling, same as
' always. TranslateViaRibbon passes an explicit path here - the one a
' human just chose via its own Save-As dialog - because the ORIGINAL
' ribbon flow only ever showed an OPEN dialog (to pick the SOURCE) and
' then silently auto-wrote the output with no dialog at all; typing a
' not-yet-existing name into that Open dialog hit Windows' own native
' "file not found" refusal - an Open dialog enforces existence, a Save
' dialog does not. Required, not Optional: VBA does not allow Optional
' parameters before a ParamArray at all (confirmed live - a compile
' error, not a style choice) - ParamArray must stay last, so the
' override sits between programPath and vocabPaths, not appended
' after, and every caller must pass it explicitly.
Public Function EnglishTranslateToVla(ByVal programPath As String, ByVal outPathOverride As String, _
                                  ParamArray vocabPaths() As Variant) As Boolean
    On Error GoTo failed
    EnglishResetGrammar
    Dim i As Long
    For i = LBound(vocabPaths) To UBound(vocabPaths)
        EnglishLoadVocabulary CStr(vocabPaths(i))
    Next
    Debug.Print EnglishLintReport()

    If Len(Dir$(programPath)) = 0 Then
        VLA_Messages.RaiseMsg "english-program-file-not-found", "path", programPath
    End If
    Dim outPath As String
    outPath = outPathOverride
    If Len(outPath) = 0 Then outPath = TranslatedPathFor(programPath, ".vla")
    If StrComp(outPath, programPath, vbTextCompare) = 0 Then
        VLA_Messages.RaiseMsg "english-translate-vla-overwrite", "path", programPath
    End If

    Dim vla As String
    vla = EnglishToVla(VocabReadFile(programPath))
    WriteTextFileVla outPath, VLA_Lint.VlaLintFormat(vla)
    EnglishTranslateToVla = True
    Exit Function

failed:
    VlaShowError Err.Description
End Function

' outPathOverride: same contract as EnglishTranslateToVla's own (above) -
' required, not Optional, for the identical reason (VBA disallows
' Optional before a ParamArray).
Public Function EnglishTranslateToVba(ByVal programPath As String, ByVal outPathOverride As String, _
                                  ParamArray vocabPaths() As Variant) As Boolean
    On Error GoTo failed
    EnglishResetGrammar
    Dim i As Long
    For i = LBound(vocabPaths) To UBound(vocabPaths)
        EnglishLoadVocabulary CStr(vocabPaths(i))
    Next
    Debug.Print EnglishLintReport()

    If Len(Dir$(programPath)) = 0 Then
        VLA_Messages.RaiseMsg "english-program-file-not-found", "path", programPath
    End If
    Dim outPath As String
    outPath = outPathOverride
    If Len(outPath) = 0 Then outPath = TranslatedPathFor(programPath, ".vba")
    If StrComp(outPath, programPath, vbTextCompare) = 0 Then
        VLA_Messages.RaiseMsg "english-translate-vba-overwrite", "path", programPath
    End If

    ' No lint pass - VLA_Lint understands .vla S-expression text only,
    ' not generated VBA; VlaTranspile's own output is already
    ' deterministic (TestGoldens's own pin).
    Dim code As String
    code = EnglishToVba(VocabReadFile(programPath))
    WriteTextFileVla outPath, code
    EnglishTranslateToVba = True
    Exit Function

failed:
    VlaShowError Err.Description
End Function

' =====================================================================
'  "Define <name> as <fixed value>." - program-wide named constants.
'  Emits a module-level Const, so the name exists before any code
'  runs, is visible in every action, and cannot be assigned.
'  Values must be fixed: quoted text, a number, or an earlier alias.
' =====================================================================
Private Sub ParseDefine(toks() As String, ByRef pos As Long)
    pos = pos + 1                              ' past "define"
    Dim name As String
    name = ExpectWord(toks, pos, "a name after 'Define'")
    CheckName name
    ExpectWordIs toks, pos, "as"
    Dim t As String, val As String
    t = TokAt(toks, pos)
    If IsStrTok(t) Then
        val = VlaStringLit(Mid$(t, 2))
        pos = pos + 1
    ElseIf IsNumTok(t) Then
        val = t
        pos = pos + 1
    ElseIf Len(WordAt(toks, pos)) > 0 And CollHasKey(mAliasNames, t) Then
        val = t                                ' an earlier alias
        pos = pos + 1
    Else
        VLA_Messages.RaiseMsg "english-define-needs-fixed-value", "context", SentenceContext(toks, pos)
    End If
    ExpectTok toks, pos, ".", "'.' at the end of the sentence"
    AddKeyed mAliasNames, name
    mAliasDefs = mAliasDefs & "(const " & name & " " & val & ")" & vbCrLf
End Sub

' V5.3: the runtime helper zoo (VlaColor, VlaCount, VlaItem, the
' VlaDict family, VlaFindRow, VlaSendMail, VlaEnsureSheet) moved to
' module VLA_Runtime - a generated program calls them from the
' user's workbook, where an add-in Private cannot reach, so they
' must ship as their own injectable module. See VLA_Runtime.bas.

' =====================================================================
'  Function-word registry: the extension API for expressions, the way
'  EnglishAddPhrase is for sentences. One line per value-word:
'      EnglishAddFunctionWord "largest", "application.worksheetfunction.max"
'      EnglishAddFunctionWord "today", "date", True
'  or in a vocab file:
'      function: largest of => application.worksheetfunction.max
'      function: today => date
' =====================================================================

Public Sub EnglishAddFunctionWord(ByVal word As String, ByVal target As String, _
                                  Optional ByVal nullary As Boolean = False)
    EnsureInit
    If nullary Then
        AddFnEntry mFnNullary, word, target
        mFnDisplay.Add VLA_Identity.Fold(word)
    Else
        AddFnEntry mFnOf, word, target
        mFnDisplay.Add VLA_Identity.Fold(word) & " of ..."
    End If
End Sub

Private Sub RegisterFunctionWord(ByVal lhs As String, ByVal target As String, ByVal context As String)
    Dim parts() As String
    parts = Split(Trim$(lhs), " ")
    If UBound(parts) - LBound(parts) = 1 And VLA_Identity.Fold(parts(UBound(parts))) = "of" Then
        EnglishAddFunctionWord parts(LBound(parts)), target, False
    ElseIf UBound(parts) = LBound(parts) Then
        EnglishAddFunctionWord parts(LBound(parts)), target, True
    Else
        VLA_Messages.RaiseMsg "english-function-word-not-one-word", "context", context
    End If
End Sub

Private Sub RegisterBuiltinFuncWords()
    Set mFnOf = New Collection
    Set mFnNullary = New Collection
    Set mFnDisplay = New Collection
    AddFnEntry mFnOf, "length", "len"
    AddFnEntry mFnOf, "uppercase", "ucase"
    AddFnEntry mFnOf, "lowercase", "lcase"
    AddFnEntry mFnOf, "absolute", "abs"
    AddFnEntry mFnOf, "month", "month"
    AddFnEntry mFnOf, "year", "year"
    AddFnEntry mFnOf, "day", "day"
    AddFnEntry mFnOf, "hour", "hour"
    AddFnEntry mFnOf, "minute", "minute"
    ' count is core since B1 (lists live in the core language), backed
    ' by VlaCount so one word answers lists and ranges alike. Vocab
    ' files must NOT redefine it back to worksheetfunction.count, or
    ' loading them would silently break "count of" over lists.
    AddFnEntry mFnOf, "count", "vlacount"
    ' B7: element access over lists and ranges.
    AddFnEntry mFnOf, "first", "vlafirst"
    AddFnEntry mFnOf, "last", "vlalast"
    ' V3.1: reading a lookup entry inside "For each pair in <lookup>".
    ' Engine-seeded, not vocabulary: the pair walk is STRUCTURAL, and
    ' a structural feature must not emit calls whose reading depends
    ' on a vocab file being loaded. Safe as of-words: "value" only
    ' takes this reading when "of" follows ("value in cell B2" checks
    ' for "in", the Create kind is its own parser, with/using
    ' parameters named value are parameter-name-driven, and of-words
    ' are legal as plain names everywhere else - CheckName refuses
    ' only nullary words).
    AddFnEntry mFnOf, "key", "vlapairkey"
    AddFnEntry mFnOf, "value", "vlapairvalue"
    AddFnEntry mFnNullary, "today", "date"
    AddFnEntry mFnNullary, "now", "now"
    Dim w As Variant
    For Each w In Array("length of ...", "uppercase of ...", "lowercase of ...", _
                        "absolute of ...", "month of ...", "year of ...", "day of ...", _
                        "hour of ...", "minute of ...", "count of ...", _
                        "first of ...", "last of ...", "item <n> of <list>", _
                        "key of <pair>", "value of <pair>", "today", "now")
        mFnDisplay.Add CStr(w)
    Next
End Sub

Private Sub AddFnEntry(col As Collection, ByVal word As String, ByVal target As String)
    On Error Resume Next
    col.Remove VLA_Identity.Fold(word)               ' allow redefinition
    On Error GoTo 0
    col.Add target, VLA_Identity.Fold(word)
End Sub

Private Function IsFnWord(col As Collection, ByVal word As String) As Boolean
    If col Is Nothing Then Exit Function
    Dim f As Boolean, v As Variant
    v = CollGet(col, word, f)
    IsFnWord = f
End Function

Private Function FnTarget(col As Collection, ByVal word As String) As String
    Dim f As Boolean
    FnTarget = CStr(CollGet(col, word, f))
End Function

' The function words, for the "Known Sentences" listing.
Public Function EnglishListFunctionWords() As String
    EnsureInit
    Dim r As String
    Dim w As Variant
    For Each w In mFnDisplay
        r = r & "    " & w & vbCrLf
    Next
    EnglishListFunctionWords = r
End Function


' =====================================================================
'  EnglishExplain - the "why did it do that?" tool. Prints the tokens,
'  then: the construct that claimed the sentence (a structural form,
'  a value call, an action call) or the matched phrase rule - plus
'  the inner rule when both apply (a one-line If body) - then the
'  VLA, the macro expansion when one fired (L1: VlaExpandText to
'  fixpoint - what the emitter sees), and the VBA. On failure, a
'  CLAIMED sentence shows who claimed
'  it and what it refused; a true no-match shows the failure with
'  first-word candidates. For developers and phrasebook authors.
'  Program-scoped words (To-actions, using-calls) are those of the
'  most recent translation - run Check Sentences first, then Explain.
' =====================================================================
' =====================================================================
'  G5: the rule scratchpad. One Immediate-window call auditions a
'  candidate rule against the FINISHED authoring surface -
'  alternations, surface forms, typed slots, overrides, carried
'  macros - without editing a file or reloading two hundred proofs:
'
'    EnglishTryRule "zap cell {r:cell}", "(debug-print {r})", "Zap cell B2."
'    EnglishTryRule "warm cell {r:cell}", "(debug-print 2)", "Warm cell B2.", True   ' audition an override
'
'  The candidate registers into the live grammar, the probe runs
'  through EnglishExplain (tokens, claim, matched rule with its
'  [from ...] provenance, VLA, macro panel), a VERDICT line says
'  whether the CANDIDATE fired or an earlier rule shadowed it (the
'  audition fact that matters most), any lint findings the candidate
'  drew are printed - and then everything is discarded: the five
'  rule collections are snapshot before and restored after, so
'  append and in-place override replacement restore uniformly, lint
'  warnings roll back, usage counters and declared-name state are
'  suspended and restored (a try is not usage), and the grammar is
'  byte-for-byte as before. A registration the load would refuse
'  (an unmarked duplicate, an unearned override, a bad shape) is
'  REPORTED as the audition finding it is, never a crash. In a
'  fresh workbook, call it through the add-in:
'    Application.Run "Frazaro.xlam!EnglishTryRule", "...", "...", "..."
'  Macros audition through VlaExpand/VlaTry, not here - this bench
'  is for sentences.
' =====================================================================
Public Sub EnglishTryRule(ByVal pattern As String, ByVal template As String, _
                          ByVal probe As String, Optional ByVal asOverride As Boolean = False)
    EnsureInit
    Debug.Print "=== EnglishTryRule"
    Debug.Print "  candidate: " & IIf(asOverride, "override: ", "") & pattern

    Dim sIt() As Variant, sTx() As Variant
    Dim sSg() As Variant, sSr() As Variant, sFo() As Variant
    SnapshotRules sIt, sTx, sSg, sSr, sFo
    Dim lintBefore As Long
    lintBefore = mLintWarnings.Count
    Dim savedD As Collection, savedA As Collection
    Set savedD = mDeclared
    Set savedA = mAssigned
    Set mDeclared = New Collection
    Set mAssigned = New Collection
    Dim savedSusp As Boolean
    savedSusp = mUsageSuspended
    mUsageSuspended = True
    Dim savedSource As String
    savedSource = mLoadSource
    mLoadSource = "(try)"

    On Error GoTo regFailed
    AddPhraseRule pattern, template, asOverride
    On Error GoTo 0

    Dim li As Long
    For li = lintBefore + 1 To mLintWarnings.Count
        Debug.Print "  lint: " & mLintWarnings.Item(li)
    Next

    Debug.Print "  probe: " & probe
    mLastRuleIdx = 0
    Dim probeErr As String
    On Error Resume Next
    EnglishExplain probe
    If Err.Number <> 0 Then probeErr = Err.Description
    On Error GoTo 0
    If Len(probeErr) > 0 Then Debug.Print "  probe refused: " & probeErr

    If mLastRuleIdx > 0 Then
        If mPatTexts.Item(mLastRuleIdx) = pattern Then
            Debug.Print "  verdict: the CANDIDATE matched the probe"
        Else
            Debug.Print "  verdict: an EARLIER rule matched - '" & mPatTexts.Item(mLastRuleIdx) & _
                        "' [from " & RuleSourceOf(mLastRuleIdx) & "] - the candidate never fired"
        End If
    Else
        Debug.Print "  verdict: no phrase rule fired (a structural form claimed it, or it refused - see above)"
    End If
    GoTo cleanup

regFailed:
    Debug.Print "  registration would refuse the load: " & Err.Description
    On Error GoTo 0

cleanup:
    RestoreRules sIt, sTx, sSg, sSr, sFo
    Do While mLintWarnings.Count > lintBefore
        mLintWarnings.Remove mLintWarnings.Count
    Loop
    Set mDeclared = savedD
    Set mAssigned = savedA
    mUsageSuspended = savedSusp
    mLoadSource = savedSource
    Debug.Print "  (discarded - the grammar is exactly as before)"
End Sub

' G5: full snapshot of the five parallel rule collections. Uniform
' restore whether the try appended or replaced in place. EnsureInit
' guarantees the prelude exists, so the count is never zero.
Private Sub SnapshotRules(ByRef it() As Variant, _
                          ByRef tx() As Variant, ByRef sg() As Variant, ByRef sr() As Variant, _
                          ByRef fo() As Variant)
    Dim n As Long
    n = mPatItems.Count
    ReDim it(1 To n)
    ReDim tx(1 To n)
    ReDim sg(1 To n)
    ReDim sr(1 To n)
    ReDim fo(1 To n)
    Dim i As Long
    For i = 1 To n
        Set it(i) = mPatItems.Item(i)
        tx(i) = mPatTexts.Item(i)
        sg(i) = mPatSigs.Item(i)
        sr(i) = mPatSources.Item(i)
        Set fo(i) = mPatForms.Item(i)
    Next
End Sub

Private Sub RestoreRules(ByRef it() As Variant, _
                         ByRef tx() As Variant, ByRef sg() As Variant, ByRef sr() As Variant, _
                         ByRef fo() As Variant)
    Set mPatItems = New Collection
    Set mPatTexts = New Collection
    Set mPatSigs = New Collection
    Set mPatSources = New Collection
    Set mPatForms = New Collection
    Dim i As Long
    For i = LBound(it) To UBound(it)
        mPatItems.Add it(i)
        mPatTexts.Add tx(i)
        mPatSigs.Add sg(i)
        mPatSources.Add sr(i)
        mPatForms.Add fo(i)
    Next
    mDspValid = False   ' V7: a try's audition indexed the candidate;
                        ' the restored store must never consult that
                        ' index (a stale entry can point past the end)
    RebuildSigOwner     ' G3 perf: same reasoning - the candidate's own
                        ' AddPhraseRule call already added ITS entries
                        ' into mSigOwner; re-derive from the restored
                        ' table rather than trying to patch them back
                        ' out (this is EnglishTryRule's own one-rule-
                        ' at-a-time dev bench, not a hot path - an
                        ' O(corpus size) rebuild here costs nothing
                        ' that matters)
End Sub

Public Sub EnglishExplain(ByVal sentence As String)
    EnsureInit
    Dim savedD As Collection, savedA As Collection
    Set savedD = mDeclared
    Set savedA = mAssigned
    Set mDeclared = New Collection
    Set mAssigned = New Collection

    Debug.Print "=== Explain: " & sentence
    Dim toks() As String
    Dim t As Variant
    Dim r As String
    On Error GoTo restore
    toks = CanonicalizeStructuralWords(EnTokenize(sentence))   ' LX5.1
    For Each t In toks
        If IsStrTok(CStr(t)) Then
            r = r & " [""" & Mid$(CStr(t), 2) & """]"
        Else
            r = r & " [" & t & "]"
        End If
    Next
    Debug.Print "  tokens:" & r

    Dim pos As Long
    Dim got As String
    Dim failMsg As String
    mLastRuleIdx = 0
    mClaim = ""
    pos = 1
    On Error Resume Next
    got = ParseStmt(toks, pos, 0)
    If Err.Number <> 0 Then failMsg = Err.Description
    On Error GoTo restore

    If Len(failMsg) > 0 Then
        ' A4: two different failures wear different traces. When a
        ' construct CLAIMED the sentence (a structural case, a value
        ' call, an action call) and then stopped, the sentence was
        ' recognized - the useful trace is who claimed it and what it
        ' refused, and the first-word candidate list would be noise
        ' (structural words own no phrase rules, so it used to print
        ' an empty header here). Only a true no-match - nothing
        ' claimed, no rule consumed the whole sentence - gets the
        ' candidate listing.
        If Len(mClaim) > 0 Then
            Debug.Print "  claimed by: " & mClaim
            Debug.Print "  refused: " & failMsg
            If EnglishLastErrorLine() > 0 Then Debug.Print "  error line: " & EnglishLastErrorLine()
            GoTo restore
        End If
        Debug.Print "  NO MATCH: " & failMsg
        If EnglishLastErrorLine() > 0 Then Debug.Print "  error line: " & EnglishLastErrorLine()
        ' Candidates sharing the first word, with examples.
        Dim fw As String
        fw = WordAt(toks, 1)
        If Len(fw) > 0 Then
            Dim i As Long
            Dim items As Collection
            Dim exv As Variant
            Dim exf As Boolean
            Debug.Print "  rules starting with '" & fw & "':"
            For i = 1 To mPatItems.Count
                Set items = mPatItems.Item(i)
                If items.Item(1) = fw Then
                    Debug.Print "    " & mPatTexts.Item(i)
                    exf = False
                    exv = CollGet(mRuleExamples, "r" & i, exf)
                    If exf Then Debug.Print "        e.g. " & exv
                End If
            Next
        End If
    Else
        If pos <= UBound(toks) Then Debug.Print "  (note: sentence continued past the first statement)"
        ' A4: the headline is the OUTERMOST construct. A structural
        ' claim wins it; a phrase rule that fired inside (a one-line
        ' If body, say) is still shown, as the inner rule. A plain
        ' phrase sentence keeps its old line exactly.
        If Len(mClaim) > 0 Then
            Debug.Print "  matched: " & mClaim
            If mLastRuleIdx > 0 Then Debug.Print "  inner rule: " & mPatTexts.Item(mLastRuleIdx) & "   [from " & RuleSourceOf(mLastRuleIdx) & "]"
        ElseIf mLastRuleIdx > 0 Then
            Debug.Print "  matched rule: " & mPatTexts.Item(mLastRuleIdx) & "   [from " & RuleSourceOf(mLastRuleIdx) & "]"
        Else
            Debug.Print "  matched: a core sentence form"
        End If
        Debug.Print "  VLA: " & got
        ' L1: the macro panel. English emits macro calls in everyday
        ' sentences (add! for Increase, dotimes for Repeat), so when
        ' any macro fires in this sentence's VLA, print the fixpoint
        ' expansion - the middle layer becomes readable without
        ' inferring it back from the VBA below. Soft-failure doctrine:
        ' an expansion problem never kills the Explain - the panel is
        ' simply skipped and the VBA panel tells the story slower.
        Dim expFired As Long
        Dim expText As String
        On Error Resume Next
        expText = VlaExpandText(got, True, expFired)
        On Error GoTo restore
        If expFired > 0 And Len(expText) > 0 Then
            Debug.Print "  expands to (" & expFired & " macro application" & _
                        IIf(expFired = 1, "", "s") & "):"
            Dim xl As Variant
            For Each xl In Split(Replace(expText, vbCrLf, vbLf), vbLf)
                If Len(Trim$(CStr(xl))) > 0 Then Debug.Print "    " & xl
            Next
        End If
        Dim vba As String
        On Error Resume Next
        ' V4: the probe rides the same (at-line ...) wrapper a Run
        ' would give this sentence, so the VBA panel shows the real
        ' three-layer map tags (the probe's sentence line is 1).
        vba = VlaTranspile("(sub explain-probe ()" & vbCrLf & _
                           "(at-line 1" & vbCrLf & got & vbCrLf & "))")
        If Err.Number <> 0 Then vba = "(transpile error: " & Err.Description & ")"
        On Error GoTo restore
        Debug.Print "  VBA:"
        Dim ln As Variant
        For Each ln In Split(Replace(vba, vbCrLf, vbLf), vbLf)
            If Len(Trim$(CStr(ln))) > 0 Then Debug.Print "    " & ln
        Next
        If InStr(1, vba, "src:") > 0 Then
            Debug.Print "  (map: ' vla:N src:M = vla line N of the generated middle layer," & _
                        " sentence line M of the program - the probe is line 1)"
        End If
    End If

restore:
    If Err.Number <> 0 Then Debug.Print "  (explain error: " & Err.Description & ")"
    Set mDeclared = savedD
    Set mAssigned = savedA
End Sub

' =====================================================================
'  The phrasebook audit, shipped. Loads the given phrasebook(s) into
'  a fresh grammar and returns "" when clean, or the registration
'  warnings (duplicate signatures, operator-words-after-expression-
'  slots, noise-word-before-slot - F.4 shape 2) when not, PLUS one
'  post-load pass over the finished corpus for cross-rule shadowing
'  (F.4 shape 1, AuditCrossRuleShadow - audit-time only, deliberately
'  not a registration-time hook: see its own header for why). Loading
'  also runs every test: line, so a failing proof raises before the
'  audit even reports - three gates in one call. Leaves the audited
'  phrasebook loaded.
' =====================================================================

' EDITIONMANIFEST.1: also accepts a single pre-built array (VLA_Build's
' own edition phrasebook chains) - VBA's ParamArray has no unpacking
' syntax, so a caller holding a Variant array built at runtime (not a
' fixed, hand-typed list of literal paths) cannot spread it across
' positional arguments the way every existing caller here already
' does. Detected, not a second entry point: exactly one argument that
' is itself an array means "here is the list," the same way a caller
' passing several bare paths already meant "here they are, in order."
Public Function EnglishAuditPhrasebook(ParamArray filePaths() As Variant) As String
    EnglishResetGrammar
    Dim paths As Variant
    If UBound(filePaths) = LBound(filePaths) And IsArray(filePaths(LBound(filePaths))) Then
        paths = filePaths(LBound(filePaths))
    Else
        paths = filePaths
    End If
    Dim i As Long
    For i = LBound(paths) To UBound(paths)
        EnglishLoadVocabulary CStr(paths(i))
    Next
    AuditCrossRuleShadow   ' F.4 shape 1: audit-time only, see its own header
    EnglishAuditPhrasebook = AuditWarnings()
End Function

' Same audit from a string - used by the self-test.
Public Function EnglishAuditText(ByVal text As String, Optional ByVal sourceName As String = "phrasebook") As String
    ' G3: audit mode - unmarked same-shape duplicates REFUSE a real
    ' load, but the audit's contract is to report everything a load
    ' would refuse, so duplicates stay recorded warnings here and the
    ' walk continues. The flag restores on every exit, error included.
    EnglishResetGrammar
    mAuditMode = True
    On Error GoTo cleanup
    EnglishLoadVocabularyText text, sourceName
    mAuditMode = False
    AuditCrossRuleShadow   ' F.4 shape 1: audit-time only, see its own header
    EnglishAuditText = AuditWarnings()
    Exit Function
cleanup:
    mAuditMode = False
    ' G0: capture Err's properties to locals BEFORE re-raising - a
    ' documented VBA fragility, the same hazard VlaReadForms's/
    ' VlaReadFormsWithLines's own fail: handlers already guard against
    ' (VLA.bas): Err.Raise Err.Number/Err.Source/Err.Description
    ' passed live can misbehave, since Err.Raise's own invocation can
    ' disturb Err's properties before all three arguments finish
    ' evaluating. This exact line crashed with an opaque "Object
    ' required" instead of relaying the real underlying error, caught
    ' live during the F.13 migration.
    Dim errNum As Long, errDesc As String, errSrc As String
    errNum = Err.Number: errDesc = Err.Description: errSrc = Err.Source
    Err.Raise errNum, errSrc, errDesc
End Function

Private Function AuditWarnings() As String
    If mLintWarnings.Count = 0 Then Exit Function
    Dim r As String
    Dim e As Variant
    For Each e In mLintWarnings
        r = r & "AUDIT: " & e & vbCrLf
    Next
    AuditWarnings = r
End Function
