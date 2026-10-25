Attribute VB_Name = "VLA"
Option Explicit
' DI.3a: the product's release version - orthogonal to every VLA_xxx_VERSION
' constant in this project (including VLA_CORE_VERSION just below), which
' each name the roadmap item that last touched THAT module, not a release
' number. This one constant is what "what build is this" actually means;
' bump it only at an actual release, by hand, same discipline every other
' VLA_xxx_VERSION already uses. Lives here (not VLA_Identity.bas, despite
' DI.3's own scoping note proposing that) because VLA_Identity's own header
' is explicit that nothing beyond its two comparison verbs belongs there
' (REBUILD.md SS4, "no VLA_Utilities") - read fresh before writing this,
' not assumed from the scoping pass. VLA.bas ships in every configuration
' and is the natural place a reader already looks for "the version".
' VLA_Build.bas's VlaBuildAddin writes this value into
' installer\version.iss at build time, so Frazaro.iss's AppVersion can
' never again silently rot the way it did at a hardcoded "1.0" (DI.3's own
' finding). VLA_IDE.bas surfaces it in "Copy Feedback" and the ribbon.
' Starts at 0.5.0, not 0.1.0 - the owner's own choice, both a beta-status
' signal (see SD-14 for what each of the three numbers means from here)
' and a deliberate homage to John McCarthy's LISP 1.5 Programmer's Manual,
' this project's own most direct ancestor in spirit.
Public Const VLA_RELEASE_VERSION As String = "0.11.0"
Public Const VLA_CORE_VERSION As String = "VLA.6"
' VLA.6: A NAME AT EXPAND TIME, AND A REFERENCE'S TWIN. (define name form)
' binds a word to a form from the next form of the source on: the walk
' passes a define by as written (DefinePassedBy), and once the form of the
' source it stands in has expanded, ExpandSourceForm - the one place the
' three entry points expand a form of the source - takes the defines from
' that form's statements, checks each name, refuses a name used in its own
' form and a ring of names, and registers them (DefinesTake), so that
' neither the trampoline's order nor a begin's shape decides what a word
' means. The walk's atom case puts a define's form where its name stands
' as an expression and looks at it again, a link of the chain; a list's
' head is passed by, never a name. (last X) is a head of the formula syntax
' beside ref and error, X's twin, its sheet's name with .last after it
' (FsLastForm); (sheet-expand X) and (range-expand X) are forms of the walk
' beside formula-expand, the two texts ref takes, so that a verb writes a
' row at a name (ExpandReferencePart); the three read one reference on a
' sheet alike (FsSheetReference). Found while scoping and fixed: every
' return of the trampoline but an atom's dropped the forms it had set aside
' - (defmacro (m) (debug-print 1) (quote (a b))) expanded to (quote (a b))
' - and each now keeps them (ExpandMacros' finish). vla-lang's expand.rs
' and formula_syntax.rs mirror each; the expand golden holds the two.
' VLA.5: A LOOP AT EXPAND TIME, AND A PIECE BY ITS HEAD. (each (name list)
' form ...) substitutes its forms once for each element of the list, the
' name bound to the element under the template's own bindings, and
' (each-index (index name list start) form ...) binds a count beside it,
' from the start up by one (EachFormResults). A loop is a form of the
' walks, beside formula-expand, and no LISTOPS primitive: Substitute takes
' it at a list's head, standing alone, as one begin (EachFormBegin), as a
' list's element its results spliced in (EachFormInto), and where a value
' is read - ResolveListopsArg, unquote, unquote-splicing - as the list of
' its results (SubstituteValue), the rest parameter's own rule with a
' statement's begin; ExpandMacros takes it at a form's head, its begin
' handed back to the trampoline so its rows expand, and among a list's
' elements spliced (ExpandEachInto). A loop's scope is made for each
' element from the template's bindings its forms read (EachReadBindings),
' since a Collection cannot list its keys. A name a loop binds is refused
' when it is an expand-time name or one the template already binds, and a
' list written as a macro's call is refused, never walked as its symbols.
' (assoc-expand head entries) is the twenty-eighth primitive
' (EvalAssocExpand): the first entry whose head is head, compared as
' equal? compares, or the empty list, Scheme's assoc. The expand-time
' names are one table now, ExpandReservedDesc, which DefineMacro and a
' loop's names both ask. Found on the way and fixed, since a loop's begin
' meets it: a begin a macro made whose last form is an atom kept the atom
' alone and dropped the forms set aside before it, so (defmacro (m)
' (debug-print 1) x) expanded to x; ExpandMacros keeps them now, in order
' (PendingBegin). vla-lang's expand.rs mirrors each; the expand golden
' holds the two to one text.
' KERNEL.8, SLICE F: ARRAYS IN THE FORMULA SYNTAX. FormulaSyntaxWrite
' writes (quote (1 2 3)) as the array constant {1,2,3} and
' (quote ((1 2) (3 4))) as {1,2;3,4}, each value a numeral, a text, a
' truth value or (error "#N/A"), every row as long as the first, and
' refuses by name an empty one, one mixing values and rows, rows of two
' lengths and any other value (FsQuote); and (: a b) as a:b, the range
' operator between two expressions that answer references, bound above
' every other operator (FsBinaryLevel 8, an atom now 9). vla-lang's
' formula_syntax holds the same writer and reads both back; the expand
' golden holds the two writers to one text.
' VLA.4: THE FORMULA SYNTAX. FormulaSyntaxWrite writes an expression as a
' formula's canonical text in Excel's invariant spelling - no space
' outside a text, the fewest parentheses Excel's precedence allows, upper-
' case names, a sheet quoted only where RefersQuoteSheet says, a number
' and a text as written - and refuses by name what no formula spells: a
' symbol RefersScan does not read whole as one reference or name
' (tax-rate, never renamed), VBA's numerals (&HFF), a head the head table
' refuses. (formula-expand <expression>) is it at expand time, an
' operation of ExpandMacros' walk (ExpandFormulaCall) that expands its
' argument first, reserved in DefineMacro. EmitFormula, deflambda's
' dialect, keeps its spelling, and its Case Else now refuses a head the
' table refuses (F.7's table: no row undeclared), where it wrote new(x).
' vla-lang's formula_syntax mirrors the writer and adds the reader;
' scripts/expand_golden.txt holds the two writers to one text. An alias
' reaches its form by the head table itself (FsCanonicalHead), never by a
' compile's mHeadAliases, which the expander does not set: the owner's
' pass found (si ...) written SI(...) in a session no compile had run in.
' VLA.3: THE EXPAND-TIME OBJECT MODEL. Ten primitives join the seventeen,
' each named where those are (DefineMacro's reserved names, ExpandMacros'
' standalone Select, Substitute's inline Select, EvalListopsPrim), each a
' name ending in expand: -expand, *expand, quotient-expand and mod-expand
' in EvalArithExpand, the last two Excel's QUOTIENT and MOD over whole
' numbers, a zero divisor refused; &expand (EvalConcatExpand), atoms
' joined into a string literal, so that a macro writes an address or a
' formula with no leading-quote trick; characters-expand
' (EvalCharactersExpand), a text split into one-character strings, a
' surrogate pair kept whole; address-expand, row-expand, column-expand and
' column-letters-expand (EvalAddressExpand, EvalAddressPartExpand,
' EvalColumnLettersExpand), a row and a column written as an address and
' read back by one written grammar (ParseAddressText), the column by
' VLA_Refers' rule. vla-lang's expand.rs mirrors each, and
' scripts/expand_golden.txt (VlaWriteExpandGolden) holds the two to one
' text. Alonzo's cartridge sources asked for them (its SPEC.md section 12).
' VLA.1: TEXT A STRANGER WROTE IS REFUSED BY NAME, NEVER A CRASH. ParseForm
' recursed once a list with no limit, so a text of thousands of nested
' lists ran VBA out of stack where a refusal was owed, and the expander's
' walk did the same for a list a macro nested deeply; vla-lang's reader and
' expander, ports of these, overflowed the same way (a cartridge or a
' program of 20 KB). All of them now stop at VLA_FORM_DEPTH_LIMIT, 64, with
' vla-form-too-deep naming the line: ParseForm takes listDepth, the depth of
' the list it would open, and ExpandMacros and ExpandOnePass an Optional
' walkDepth, their refusal RefuseFormTooDeep's, so each recursive frame
' gains one Long and no local (trap 8). The same lists, the same id, in
' both implementations (SD-18). The corpus nests 10 deep at most. A macro
' that grows its argument at every application, (defmacro (deeper x)
' (deeper (x))) or (defmacro (dup x) (dup (x x))), meets the 5000 chain
' limit here, since Substitute shares an argument by reference and builds
' no copies; vla-lang copies them, so it weighs each application's result
' and refuses the two sooner (vla-form-too-deep, and
' vla-macro-expansion-too-large past 1,000,000 forms in one application).
' A weighing here, tried and withdrawn the same day, counted the shared
' structure as copies and made a long walk quadratic.
' SEC.15: THE FORMULA SINK IS THE RUNTIME'S. EmitStmt's "set!" arm writes
' Call VlaSetFormula(obj, text) for exactly (set! (. obj formula) v), where
' it wrote obj.Formula2 = text: the helper (VLA_Runtime.bas) refuses a
' formula that reaches outside the workbook on its own (WEBSERVICE, a DDE
' link, ...) by name before writing, and writes through Formula2 as this
' arm did. The interpreter's Case "formula" calls the same helper, so both
' backends refuse the same text with the same words. The compile golden's
' eighteen formula lines change and nothing else; the core's emit/stmt.rs
' mirrors the arm.
' F.7 (the two latent bugs): A LONE OPERAND KEEPS ITS SIGN IN A FORMULA, AND
' A QUOTE MARK INSIDE A FORMULA'S TEXT IS DOUBLED. Found 2026-09-27 while
' Contemplation 9's collapse 1 was scoped; fixed 2026-10-01, ahead of
' PORT.5, so the port never reproduces a bug to match a golden. EmitFormula's
' operator arm wrote the operator only between operands, so (- x) became
' (x) and the sign was lost; a lone operand under - now writes (-x), the
' shape EmitExpr's own - arm already had. EmitFormula's string arm and
' FormulaQuote's wrapped a text as it stood, so a text holding a quote mark
' gave Excel a formula it could not read; FormulaText, new, doubles each
' quote mark inside the text, Excel's own spelling, and EmitDeflambda's
' VBA-escape of the whole formula at the Add line then doubles those again,
' as it doubles every other quote. Pinned by TestF7 (VLA_Tests_Grammar.bas);
' the corpus holds no deflambda, so the goldens do not move. The rest of
' F.7 - the operator table and the test that fails when the emitter gains a
' case the dialect neither supports nor refuses - stays open.
' TER-8: TOKENIZE READS TEXT WITH NO TOKEN. Found by TER-7, which fixed the
' English tokenizer's twin. Tokenize ended with ReDim toksArr(1 To
' outc.Count), and for text holding no token at all - empty, only
' whitespace, only ; comments - that is ReDim (1 To 0), which VBA refuses
' with error 9, "Subscript out of range", before a form is read. The P-TOK
' comment above it said the zero case was handled; it was not. Zero tokens
' now give (0 To 0) for both arrays, the empty list every reader here
' already expects: tokens live at 1..UBound, ParseAll and ParseForm loop
' while pos <= UBound, TokRawLine returns 0 past UBound, and nothing reads
' LBound or walks the arrays with For Each. Reached only by the readers that
' do not prepend the prelude - VlaReadForms, VlaReadFormsWithLines, VlaFormat
' and VlaProbeMacroForm - each of which now returns no forms. Their callers
' already say so in words (a template, a slot value or a rendered form must
' be exactly one form; Lint's reformatter refuses a count other than one),
' DATALOG and PROLOG now refuse empty rules text by name
' (datalog-rules-empty, prolog-clauses-empty), and an empty or comments-only
' phrasebook loads no rules instead of raising.
'
' LINTERPOLATE.0: (interpolate tpl :key val ...) - EmitExpr's own Case
' "interpolate", below, next to "array" - the compile-time twin of
' VLA_Interpreter.bas's EvalExpr Case "interpolate" (that module's own
' LINTERPOLATE.0 note has the full design: named holes, why the
' template must be a literal, the "&" coercion rule). Named
' "interpolate", not CL's "format" - this codebase's own
' format-as-currency/-percent/-date (english.vla) already own "format"
' for visual cell styling; a second, unrelated "format" primitive doing
' string substitution would collide in READING even with no symbol
' collision in the code (owner's own call, this session - the roadmap
' item itself was renamed L-FORMAT -> L-INTERPOLATE alongside this).
' tpl must be a literal string node so its hole names are known at
' COMPILE time - there is no runtime here to hand a computed template
' to and ask it what its own holes are. Emits exactly the "&"-chain
' shape the two DATALOG macros this primitive replaces already
' hand-spliced (datalog-filter-place/datalog-chain-place, english.vla):
' each hole's own EmitExpr(...) text spliced between the template's
' literal runs, "&"-joined, one splice per OCCURRENCE (a hole reused
' three times in one template gets its value's EmitExpr text emitted
' three times, not cached into a temp variable - correct and simplest
' for BETA_ROADMAP.md's own "not a second grammar" scoping, since
' EmitExpr is pure text generation with no side effect to worry about
' repeating). New messages (VLA_Messages.bas): vla-interpolate-
' template-must-be-literal, vla-interpolate-expected-keyword-arg,
' vla-interpolate-unclosed-hole, vla-interpolate-empty-hole-name,
' vla-interpolate-unknown-key, vla-interpolate-unused-argument (a
' dangling ":key" with no value reuses the existing
' vla-keyword-arg-missing-value).
' PF4C.0: (for-each-row (row range) body...) - PRODUCT + PERFORMANCE's
' PF.4c, the last of PF.4's three sub-items, landing the array-slab
' language feature itself on top of PF.4a's array-element set!/read
' parity and PF.4b's VlaSlabRead/VlaSlabWrite. Compiled emission
' (EmitForEachRow, this file): bulk-read the range once into `arr`,
' loop over its rows by index. `row` is not a real emitted variable -
' (row i)/(set! (row i) v) compile DIRECTLY to arr(iVar, i) (EmitExpr's
' own dispatch, checked ahead of its normal Select Case, driven by
' mSlabRowVar/mSlabArrVar/mSlabIVar - dynamic extent via save/restore
' in EmitForEachRow, mAtLine/mGenRow's own established precedent). One
' bulk write-back after the loop. Body may only read/write the CURRENT
' row's own columns - no raw Cells/Range calls inside it - the owner's
' own explicit adjudication after the automatic-loop-rewrite
' alternative was rejected on determinism/auditability grounds
' (docs/BETA_REARVIEW.md's own PF.4 entry has the full history): a
' fixed, unconditional emission template is what keeps this auditable,
' the same shape every other VLA form already has.
' All-or-nothing commit falls out of the emission shape itself, not
' special-cased: VlaSlabWrite is the only place anything reaches the
' live sheet, called exactly once, after the loop, and a bare exit-for
' needs no special-casing of its own to cooperate with that - because
' (row i) writes straight into arr with nothing deferred, whatever's
' mutated up to and including the break's own row is already sitting
' in arr the moment exit-for fires. return does not commit, though:
' EmitReturn emits a bare Exit Sub/Exit Function, unconditionally
' leaving the WHOLE procedure and skipping the write-back call no
' matter how deep the nesting - the interpreter's own ExecForEachRow
' (VLA_Interpreter.bas) checks mProcReturn/mGotoLabel FIRST and skips
' its own write-back to match, a real cross-backend asymmetry caught
' and fixed during scoping, not assumed symmetric from the shape alone.
' "a raised error discards everything" holds by the same construction
' (VlaSlabWrite unreached whenever anything above it fails to return
' normally, on either backend) - argued from the shape at first, then
' confirmed live by the owner (2026-09-03) rather than left resting on
' the argument alone: a genuine out-of-bounds error mid-loop (row 1 and
' row 2 already mutated in memory, row 3 raising "Subscript out of
' range" before ever reaching it) left the range fully unchanged after
' both Run and Interpret, on a real workbook, not just the self-test
' harness.
' Three real bugs caught live across two owner-run passes, not reasoned
' through in advance - none obvious from reading the code in isolation:
' (1) Debug > Compile, first pass: one of the hidden bookkeeping names
' was `cVar`, which collides with `CVar`, a real VBA intrinsic
' (Convert-to-Variant) - VBA identifiers are case-insensitive, so `Dim
' cVar As String` alone, nothing else on the line, was a genuine
' "Syntax error", isolated by noticing only THAT one of several
' structurally-identical sibling Dims was ever flagged. (2) Chasing (1),
' this function's own string-building was reshaped from one long multi-
' line continuation into separate `r = r & ...` appends, matching
' EmitFor/EmitForEach/EmitIf/EmitSelect's own established style - not
' confirmed as a second real cause, but the right shape regardless.
' (3) VlaSelfTestsAll, second pass, the real design flaw: the original
' fix for (1)/(2) still sliced each row into its OWN small array
' (`row(1 To cols)`), ran the body, then copied it BACK into `arr` in a
' SEPARATE loop placed AFTER the body - so a bare exit-for inside the
' body skipped straight past that trailing copy-back, losing the
' mutation on the exact row where the break fired (measured live:
' `20 + 20 + 30 = 70`, not the expected `20 + 40 + 30 = 90` - row 2's
' own doubling never reached `arr` before the break). Not a small
' patch: fixed by removing the separate row array entirely in favor of
' the direct-substitution design this entry now describes, which
' deletes the copy-back step (and the whole bug class with it) rather
' than special-casing exit-for to route through it correctly.
' PF7.0: VlaEmbeddedText's O(n^2) string-build (`s = s & CStr(cell.Value)
' & vbCrLf` inside a `For i = 1 To n` cell-by-cell loop) fixed - PRODUCT +
' PERFORMANCE's PF.7, BETA_REARVIEW.md/BETA_ROADMAP.md, scoped before any
' code changed per the owner's own P-TOK-precedent request. One bulk
' Range.Value read (one COM round-trip regardless of n, replacing n
' separate sh.Cells(i, 1) calls) flattened into a 1-D String array, then
' one Join - genuinely O(n), not another positionally-indexed structure
' wearing a new shape (P-TOK's own mid-session lesson: a materialized
' ARRAY is O(1) indexed, unlike a Collection's Item(k)). n=1 handled as
' its own branch, not an oversight: Range.Value on a single-cell range
' returns a bare scalar, not a 2-D array - not hypothetical, it's
' IdeVocabFileName's own VLAe_Name shape, hit on every Check in a built
' edition. Scoping pass found the whole loop body unreachable from the
' dev workbook/self-test suite as it stood - PreludeMacros' and
' IdeVocabPath's own external-file lookups both win there, so VLAp_Source/
' VLAe_Source are never read via this path in dev; only a built .xlam's
' embedded sheets ever reach it. Test coverage added alongside the fix,
' not left to the existing 2-row case alone: n=1 (VLA_Tests_Host.bas,
' TestEmbeddedTextFallback) and n=1-with-a-never-written cell (End(xlUp)
' returns row 1 on a wholly blank column, never row 0 - the real shape
' "no content" takes), plus an mRunScaleTests-gated round-trip at
' english.vla's own real order of magnitude (2,849 rows) for a
' repeatable, dev-side number, matching TestVocabMacroProbe's own scale-
' case discipline (elapsed time printed, not pinned). PF.8 (RegisterVocab-
' Macro's own mVocabMacros string-build, VLA_SentenceEngine.bas) checked
' during scoping and left separate, not folded in: same antipattern
' family, but mVocabMacros accumulates across many separate calls over a
' vocabulary load rather than reading n rows known upfront in one call,
' and is itself re-read (as raw text, re-parsed) on every metavocab macro
' call during translation - a different mechanical shape wanting its own
' fix, not this one repeated.
' PNTH.0: cdr/cddr's O(n^2) copy replaced with O(1) structural sharing -
' MACHINE + OPTIMIZATION's P-NTH, BETA_ROADMAP.md, the "full fix,"
' owner's own call after P-DICT's live before/after ruled out GetMacro
' and confirmed this family (cdr/ListTail) as the real cost. New class
' VlaSlice.cls: a read-only, O(1) VIEW over the tail of an existing
' Collection - Backing + Offset, never copies, never nests (a slice of
' a slice flattens to point at the SAME root Backing with an adjusted
' Offset, so .Item/.Count stay O(1) regardless of chain length).
' Deliberately a flat VIEW, not a general cons-cell: cdr/cddr (repeated
' tail-walking) is the evidenced hot path; cons (prepending) is not,
' so EvalCons still copies, only widened to accept a VlaSlice tail
' without crashing - a real cons-cell chain for THAT would be
' disproportionate for an unmeasured cost. S3.1 checked deliberately,
' not assumed: the stale-ObjPtr hazard that made cdr/cddr COPY in the
' first place was about a FREED address being reused by an unrelated
' later list, never about sharing a live, still-referenced object - a
' fresh VlaSlice is still constructed on every call, still TagLine'd 0,
' the same discipline that already protected the old Collection-copy
' result, so nothing about that hazard reopens here.
' Touch points, ALL traced by reading actual call chains, not assumed
' safe by category: Nth and TagLine widened Collection -> Object (zero-
' risk, every existing caller keeps working - a Collection passed where
' Object is expected always does); ListTail rewritten to build a
' VlaSlice; the entire LISTOPS dispatch chain a VlaSlice can flow
' through - EvalCar/EvalCdr/EvalCddr/EvalCons/EvalList/EvalNullQ/EvalEq/
' EvalEqual/EvalQuoteIf/EvalArithExpand/EvalCond/EvalCondFrom,
' EvalListopsPrim, FuseSymbol, ResolveListopsArg, ListHeadIs, DeepEqual,
' NthList, ExpandOne - all widened on their own "the form/value being
' read" parameter; ExpandMacros' and ExpandOnePass' own trampolines
' widened the same way, since a macro whose ENTIRE body is a bare cdr/
' cddr call can hand a VlaSlice back to the dispatch loop itself, not
' just to a listops sub-call, however unusual that shape is - not
' fixed on suspicion, traced concretely, request by request, until each
' one either resolved to "always a genuine Collection here" (parsing/
' emission - WriteDatum/WritePretty excepted, both debug/apropos
' rendering paths a computed value CAN reach) or got widened. Several
' For Each loops (Substitute's own unquote-splicing splice, WriteDatum,
' the two trampolines' own final rebuild) converted to indexed access -
' VlaSlice has no enumerator, deliberately: building one is the classic
' VBA NewEnum/IEnumVARIANT trick, real complexity bought for nothing,
' since every consumer already only ever needed .Count/.Item(i).
' PDICT.0: mMacros (the macro registry) swapped from a Collection to a
' late-bound Scripting.Dictionary (no project reference) - MACHINE +
' OPTIMIZATION's P-DICT, BETA_ROADMAP.md, gated on and directly measured
' by P-PROF's own live run: GetMacro's Collection.Item(key) lookup was a
' linear scan (Collection has no hash table for keyed access), the
' dominant contributor to "expand" reading 71.3% of compile-side transpile
' time against the scale-included suite. GetMacro now uses .Exists (throw-
' free, hashed) instead of On Error Resume Next around a possible miss;
' DefineMacro's own registration is a direct Dictionary-item overwrite,
' not the old Remove-then-Add, removing a guaranteed exception-throw on
' every ordinary FIRST-time registration too (Remove always missed there -
' redefinition is the rare case, not the common one). Six touch points,
' fully enumerated during scoping, all in this file plus VlaFrame.cls's
' own typed Macros field (the F5.0 push/pop snapshot): the declaration,
' four independent reset sites (VlaTranspile/VlaCompileToForms/
' VlaProbeMacroForm/VlaExpandText), DefineMacro's write, GetMacro's read,
' and the Apropos listing's iteration (mMacros.Items now, not a bare
' For Each - a Dictionary walks keys by default, not values). GetMacro's
' own contract (returns the macro's record Collection, or Nothing) is
' unchanged, so its four callers (ExpandMacros/ExpandOnePass/VlaMacroDoc/
' EmitDoc) needed no changes. Deliberately NOT touched: cdr/ListTail's own
' independent O(n^2) cost (P-NTH's own territory, filed separately,
' scoped next) and every other Collection-as-registry site in this file
' (Substitute/FormLine's own CollGet lookups) - neither is evidenced hot
' by P-PROF's own numbers the way GetMacro was, and this tranche's own
' rule is real cost before a fix, not suspicion.
' PPROF.0: per-phase timing behind one switch (MACHINE + OPTIMIZATION's
' P-PROF, BETA_ROADMAP.md), the instrument the rest of that tranche
' (P-DICT/P-NTH/P-TOK) is gated on rather than built on suspicion.
' VlaTranspile's own existing Pass 1/Pass 2 structure is the seam - four
' compile-side buckets (tokenize/parse/expand/emit), summed across every
' call while the new mProfileOn switch is on, mirroring
' VLA_Tests.mRunScaleTests's own bare-Public-Boolean shape. expand/emit
' are timed per top form inside the Pass 2 loop, with a profPhase flag
' so emitfail's own handler can still credit whichever phase was running
' when a call fails partway (the 5500-deep TestListopsBudget case, which
' is SUPPOSED to fail once ExpandMacros' trampoline cap fires, still
' contributes its own real cost to the expand bucket rather than vanishing
' from the count) - existing error text, Err.Number/Source, and every
' other emitfail behavior are untouched; the accumulation happens before
' emitfail's own logic runs, never inside it. Zero cost when mProfileOn
' is False (the default, and every ordinary Check/Compile/Run path):
' Timer is never read at all, only the boolean is checked - the
' golden-diff-empty invariant this module's own IN.0.5 comment already
' calls "a real, proven-fragile thing to risk" stays untouched. See
' VLA_SentenceEngine.bas (translate-side split, same switch) and
' VLA_DevRig.bas's new VlaProfileAll (the dial that turns it on).
' LX2.0: this module's 134 of 141 raw Err.Raise refusal sites now route
' through VLA_Messages.RaiseMsg with a stable id - SD-2/LX.2, the last
' and largest file in the migration. The 7 that stay raw on purpose
' (~563/565/568/570, VlaTranspile's own emitfail; ~688 and ~1084, two
' VlaReadForms-family fail: handlers; ~2185, VlaExpandStepText's
' restoreBudget) all re-raise an already-caught error - emitfail
' additionally appends a location suffix to Err.Description before
' re-raising, but Err.Number/Source and the substance of Description
' still come from whatever inner call actually failed, often already
' carrying its own id - so this is propagation/annotation, not
' origination, the same exception every other migrated file already
' has. "vla-defmacro-reserved-name" alone covers 22 call sites (one id,
' two named slots, per reserved primitive); "vla-keyword-misuse" covers
' 4 sites that just relay KwMisuseMsg's own return value.
' "vla-interpreter-only-handler" preserves VLA_ERR_INTERPRETER_ONLY (this
' module's own named non-5 error number) via VLA_Messages.bas's one
' documented cross-module reference. Rendered text and Err.Number/
' Err.Source are unchanged (verified by manual trace at migration time
' for every site; the full 134-to-99-unique-id mapping is cross-checked
' against a fresh grep count, not just counted by hand).
' COND.0: the tenth engine primitive - (cond (test1 form1) ... (else
' formN)), generalizing quote-if's own strict two-branch shape to N
' clauses. Scoped across several passes in BETA_REARVIEW.md's own COND
' entry, with one real correction along the way worth keeping on
' record: a first pass split it into two SEPARATE proposals under two
' roadmap items (COND, a prelude.vla macro deferring every test to
' runtime; LISTOPS-COND, an engine primitive requiring every test to
' fold at expand time) - only caught, before either was built, that
' both wanted the name `cond`, and `DefineMacro` refuses shadowing
' outright: whichever loaded second would have silently never fired,
' the exact `map`/`reduce` collision class this project has already
' been burned by twice, this time self-inflicted rather than found by
' a live run. Resolved by unifying into ONE primitive rather than
' renaming either: `EvalCond` (VLA.bas) tries expand-time folding per
' clause first (`ResolveListopsArg`, quote-if's own path), and for any
' test that doesn't resolve to literal true/false, defers - splicing
' the raw test into a native (if test (then form) (else <rest>)) for
' RUNTIME evaluation instead of raising. Every clause gets this choice
' independently, so a later clause that DOES fold true is hoisted
' straight into the enclosing else slot rather than wrapped in another
' redundant if. `else` is a syntactic keyword (IsSym/Fold, matching
' EmitIf's own then/elseif/else), must be last or raises; no match and
' no else expands to (begin), matching walk-rows's own base-case
' convention, not an error. Reserved-name refusal (DefineMacro), both
' recognition sites (ExpandMacros/Substitute), and EvalListopsPrim's
' dispatcher all extended the same way LISTOPS.0 wired its own nine in
' originally - seventeen primitives now. Shared, not new, limitation
' documented at EvalCond's own header: a test that is itself a bare
' bound template parameter can't be re-resolved through that
' indirection, the same property quote-if's own test already has,
' simply never previously exercised.
' LISTOPSEXPAND.0: seven new expand-time primitives - +expand/=expand/
' <>expand/>expand/<expand/>=expand/<=expand - the arithmetic/comparison
' counterpart to LISTOPS.0's own nine structural primitives, scoped and
' adjudicated against LISTOPS-PURITY/the anti-gensym auditability
' doctrine in BETA_REARVIEW.md's own LISTOPS-EXPAND entry before a line
' of this was written. Verdict there: safe, under three guardrails -
' operands must already be resolved literal numbers within the SAME
' Substitute pass (never force further expansion of an unresolved call
' to get one, filter's own forbidden move), non-numeric operands RAISE
' rather than coerce, and no and/or-style combinators ride along (the
' actual slope toward eval). Unblocks LISTOPS-STDLIB's own `length`
' (still a runtime sum of 1s today, unchanged by this pass - no caller
' rewritten yet) and gives a future expand-time `cond` (LISTOPS-COND,
' also in BETA_REARVIEW.md) a comparison vocabulary beyond null?/eq?/
' equal?. Checked, not assumed, before writing IsNumericLiteralText/
' EvalArithExpand (VLA.bas): every atom in this system - numbers
' included - is stored as a plain string with no separate reader-time
' number type (IsSym's own only distinction is "does it start with a
' quote character"), so VBA's IsNumeric/CDbl/CStr were never safe to
' reuse here - all three respect Application.International/regional
' Windows settings, which would let identical source text fold to a
' different literal result on two machines. Val()/Str$() instead
' (Microsoft-documented locale-invariant, always a period decimal
' separator) plus a hand-rolled strict literal check, corrects last
' pass's own scoping text, which had recommended reusing "the reader's
' numeric classification" - no such classification exists to reuse.
' Reserved-name refusal (DefineMacro), both recognition sites
' (ExpandMacros's standalone Select Case, Substitute's own inline
' dispatch), and EvalListopsPrim's dispatcher all extended the same way
' LISTOPS.0 wired its own nine in originally - sixteen primitives now,
' one shared EvalArithExpand switching on the operator rather than
' seven near-identical functions. No TagLine concern (LISTOPS-PURITY's
' own binding note): these return a SCALAR, never a new Collection.
' `length` itself deliberately left unrewritten this pass - LISTOPS-
' EXPAND was scoped and built as infrastructure, not bundled with a
' caller-facing change, so the two can be reviewed/tested separately.
' G6.0: `array` - a new runtime expression primitive (EmitExpr here,
' EvalExpr/VLA_Interpreter.bas symmetrically), the lowering target for
' VLA_English.bas's new list-valued slots ({name:cat-list}, G6). Given
' an explicit case in both backends rather than relying on the compile
' side's generic Case Else function-call fallback (which would happen
' to also work, since VBA is case-insensitive and "array"/"Array" are
' the same identifier to it) - AS.8's own operator-gap finding is
' exactly the cost of leaning on an implicit fallthrough for something
' that actually needs to be a real, recognized primitive. Deliberately
' NOT a LISTOPS primitive - see LISTOPS.0's own note directly below for
' the expand-time/runtime line this stays on the correct side of.
' LISTOPS.0: the complete expand-time data bundle named in the roadmap's
' own METAMETAMACRO LINE entry - car/cdr/cddr/cons/list/null?/eq?/equal?
' - plus one primitive the entry's own bundle didn't name: quote-if.
' Finding, surfaced and owner-confirmed before any code: none of the
' eight named primitives can make a self-recursive macro STOP, because
' nothing in this codebase has ever had an expand-time conditional -
' EmitStmt's "if" (-> EmitIf) and EmitFormula's "if" (-> Excel IF()) are
' both emit-time, reached only after expansion finishes, and mMacros/
' GetMacro/DefineMacro key macros by name only, no arity-based dispatch
' to branch on argument shape either. quote-if is a NEW special form for
' exactly this reason - like quasiquote/unquote, not a plain eager
' primitive: only the SELECTED branch's form is ever resolved via
' Substitute; the other branch is never touched, which is what lets
' (quote-if (null? lst) (quote (begin)) (begin ... (walk (cdr lst))))
' actually terminate instead of eagerly evaluating the recursive call at
' the base case too. All nine share one dispatcher, EvalListopsPrim,
' called from both recognition sites - Substitute's own per-template
' walk (with the enclosing macro's real bindings/restName/restItems,
' same machinery car/cdr/cons resolve their own arguments through) and
' ExpandMacros's standalone walk (empty tables) - same dual-site
' treatment `symbol` already got under QUASIQUOTE. Unlike `symbol`,
' none of the nine are recognized while inQuasi is True: a shielded
' (car x) is ordinary shielded list data like anything else under a
' quasiquote, reachable only via an explicit (unquote (car x)) - `symbol`
' was exempted because an unbound argument is always a safe passthrough;
' `car` on arbitrary shielded template data is not. Honest caveat found
' while scoping the tests: this gate is local to Substitute's OWN walk -
' a shielded-but-never-unquoted (car x) that becomes a macro's own final
' template OUTPUT is still subject to ExpandMacros's ordinary post-
' expansion re-walk, same as any macro result already gets, unconditional
' and pre-existing. Since none of the nine have any EmitExpr/EmitStmt
' case at all, a shielded call left unquoted was always a dead end
' either way - the only choice is which of two errors fires, never a
' silent wrong result. The real, intended pattern is always an explicit
' (unquote (car ...)) - computing the value inside the shield, not
' leaving the bare form for later.
' null?/eq?/equal? return the SYMBOL "true" or "false", never a raw VBA
' Boolean - a Boolean would slip past IsSym's own object/string-literal
' check as if it were a stray identifier the moment anything but
' quote-if consumed it. eq? compares two ATOMS only (symbols via
' VLA_Identity.Fold, string literals by exact content, a symbol is
' never equal? to same-text string literal); equal? is eq?'s same
' atom-level comparison generalized recursively over list structure
' (DeepEqual). cdr/cddr/cons/list all build NEW Collections and all
' call TagLine ..., 0 on every one - LISTOPS-PURITY's own binding note,
' so the S3.1 stale-ObjPtr-tag bug can never reopen through one of
' these. Depth guard (ExpandMacros): raised 200 -> 1000 on the first
' pass, per LISTOPS-BUDGET's own finding that a naive one-row-per-frame
' recursive walker hit 200 easily on an ordinary-sized table - REVERTED
' back to 200 on a second pass, once a live host run disproved the
' raise: VBA's real native stack exhausts between 150 (safe) and 250
' (crashes with "Out of stack space"), meaning 200 was already at the
' real ceiling with no headroom, and 1000 just replaced a clean "too
' deep" error with a raw crash for the pre-existing 250-deep
' TestListopsBudget pin. LISTOPS-BUDGET's depth-chain problem is
' therefore NOT solved by this pass - a naive walker has real headroom
' for roughly 150-190 rows, below "a 250-row pricing sheet is entirely
' ordinary". The real fix (steering LISTOPS's own walking convention
' away from one long chain) is left open.
' P-PROBE's own O(macros^2) registration cost (separately
' documented, MACHINE - OPTIMIZATION) is NOT addressed here - the
' acceptance test below stays a small, fixed row count specifically so
' it doesn't need to be. Reserved names: all nine refused in
' DefineMacro, same treatment quote/quasiquote/unquote/unquote-splicing/
' symbol already get. Naming verified collision-free against the
' runtime list accessors (vlafirst/vlalast/"first of", VLA_Runtime.bas/
' VLA_English.bas) - different namespace entirely, expand-time
' primitives that vanish during compilation, never emitted, never
' reachable from a runtime English sentence. `TestListops` (38
' assertions) and `TestListopsDepthSafety` (2 assertions) wired into
' VlaSelfTest right after TestQuasiquote.
' **First live run: 768/3 failed** - the pre-existing 250-deep
' TestListopsBudget pin and both of TestListopsDepthSafety's own new
' assertions, all three for the identical reason (the depth-guard raise
' above), not three separate bugs. Fixed by reverting the guard to 200
' and rewriting TestListopsDepthSafety to stop asserting specific large
' row-counts are physically safe without live data backing the claim -
' see its own comment in VLA_Tests.bas. **Second round, predicted:**
' self-test 733 -> 771, 0 failures; host self-test unchanged at 114/0
' (compiler-only surface, no host-side change).
' QUASIQUOTE.0: the full bundle - (quasiquote ...), (unquote ...),
' (unquote-splicing ...), (symbol ...) - closes METAVOCAB's own stated
' limit (a generator could substitute a bound value only as a whole
' symbol, never glue part of one into a new identifier). Substitute
' gains a new Optional inQuasi parameter (default False, so every prior
' call site is untouched byte-for-byte) and a new FuseSymbol helper;
' ExpandMacros gains matching recognition/guards for the standalone
' case. Owner-gated before any code was written (this bundle reads as
' richer substitution vocabulary, not procedural, under L-TIER3) and
' scoped with two deliberate deferrals kept on the record: nested
' quasiquote (a data-structure mismatch, not a stack-safety concern -
' a Boolean inQuasi can express "shielded" but not "how deep", and
' shipping the depth-tracked version with no real caller to verify it
' against would be speculative generality) and unquote-splicing's
' operand auto-unwrapping a (quote ...)-wrapped argument (nothing in
' the corpus passes one with splice-intent; revisit once LISTOPS makes
' computed lists real). `symbol` alone is recognized standalone
' (outside any defmacro), not just inside Substitute's own walk -
' unlike its three siblings, it needs no ambient bindings environment,
' since a bare unbound argument already passes through as itself.
' `TestQuasiquote` (VLA_Tests.bas, 20 assertions) wired into
' VlaSelfTest right after TestListopsConfluence.
' APROPOSPLUS.0: apropos's carry mechanism (L11.1) gains a second
' channel - VlaAproposCarry's own new optional argument, mAproposRulesCarry
' beside mAproposCarry, VlaFrame's own new AproposRulesCarry field riding
' VlaPushContext/VlaPopContext alongside AproposCarry. VLA_English.bas
' pushes one "pattern -> template" line per loaded rule down the same
' call that already pushes macro text, so a search now surfaces the
' ENGLISH SENTENCE that reaches a macro, not just the macro itself -
' whichever half of the stack the search term names. See
' VLA_English.bas's own APROPOSPLUS note (AproposRulesBlob) for the
' other half.
' IN.7: a distinct error NUMBER (not just wording, like every other
' emitter refusal here uses) for "this form is interpreter-native and
' cannot be compiled" - EmitProc's colon-name guard raises it for any
' handler-declaration shape compiled parity does not yet cover (today,
' only "on:sheet-change"; "on:click:<slug>" and EmitStmt's own
' "make-button" case both compile for real now, button-click's own
' compiled-parity follow-up). The one caller that needs to tell this
' refusal apart from a REAL build problem is VLA_IDE.bas's DoCheck,
' which transpiles every translation just to validate it builds
' (P.L6's own doc-population side effect) - without a distinct
' number, that validation would fail CHECK (not just Compile) for
' every program using "When the sheet changes:", a real regression
' for the already-shipped, live-verified sheet-change half.
' vbObjectError's own +512 offset is VBA's documented range for
' custom error numbers.
Public Const VLA_ERR_INTERPRETER_ONLY As Long = vbObjectError + 7001
' IN.7 (button-click half): the compiled module's own button-creation
' helper (EmitStmt's "make-button" case Calls it; VlaTranspile appends
' its body once, only when something actually called it) - reserved
' like "vla_step"/"vla_fail"/"vla_tco_N" (EmitProc's own comment on
' that convention), never a name English or hand-authored VLA could
' collide with by accident.
Private Const VLA_MAKEBUTTON_SUB As String = "VlaCompiledMakeButton"
' F5.0: a compilation context, scoped to this module's own eighteen
' module-level variables (VLA_English.bas's 63 are a separate, later
' item). New VlaFrame.cls (a plain data bag - VBA classes cannot read
' another module's Private variables, so the copying lives here) plus
' three new Public entry points: VlaPushContext snapshots the current
' eighteen fields onto a new module-level stack (mContextStack) without
' resetting them - a nested VlaTranspile/VlaCompileToForms call already
' resets what it needs at its own top (mMacros/mHeadAliases/mSubDocs via
' "Set x = New Collection", confirmed by reading both entry points
' before writing this, not assumed), so Push does not need to reproduce
' that per-function reset logic, only save what was there before it ran.
' VlaPopContext restores the top frame and discards it; VlaContextDepth
' reports the stack's current size, for a caller (a test, or a future
' nested-compile site) to assert balance. Pure save/restore, callable
' around any nested operation - no existing function's signature or
' body changed, so every current caller is unaffected. On error, the
' caller is responsible for popping (On Error GoTo cleanup: VlaPopContext),
' the same discipline prelude.vla's with-fast-excel already keeps for
' Excel's own save/restore state; nothing here can wrap arbitrary
' caller code the way VBA has no try/finally to do that with.
' (A0.1 -> C1.0 repays a drift: the (new ...) expression form landed
' during B1 without a bump. From here the stamp moves with the module.)
' LX3.0: every identity-deciding LCase$ site now folds through
' VLA_Identity.Fold (invariant ASCII, never locale-dependent) - R6/SD-8.
' F6.0: SrcLineTag/OpenedAtTag now share one LabeledLine core instead
' of two copies of the same IncLookup branch (the "fourth reporting
' surface" comment below flagged this itself); emitfail's 3x-repeated
' at-line suffix is one AtLineSuffix() call now. Output unchanged byte
' for byte - this is deduplication, not a wording change.
' IO1.0: VlaCompileToModule refuses to overwrite a module it did not
' create (a foreign module/macro sharing the target name) instead of
' silently destroying it. BEHAVIOR CHANGE, not dedup: every module this
' function writes now carries a leading marker comment, and every
' compile after this pass checks for it (or the pre-existing ' vla:N
' source-map tag, for modules compiled before this marker existed).
' Message corrected same session: it no longer suggests "use a
' different program name" - not actually available for the default
' program, whose module name (Frazaro_EN_Sheet) Frazaro chooses, not the user.
' LX6.0: SymName transliterates non-ASCII letters to their closest
' ASCII equivalent (TransliterateChar/TransliterateToAscii) instead of
' passing them straight into generated VBA source, where they are not
' legal identifier characters. Refuses in words if nothing legal
' survives. ASCII input (100% of the existing corpus) is untouched -
' this should be an empty golden diff.
' F1.1 (reverted by F1.2, same session): tried tagging a macro call's
' single-statement expansion with its own source line, on the theory
' that it 1:1-replaces what the user wrote. Wrong: TWO existing pins
' (dotimes' template-clean/args-keep-lines split, inc!'s "deliberately
' untagged" S3.1 regression proof) already tested the opposite, and
' the theory was never checked against them before acting on it. All
' macro expansions are untagged, uniformly, by design - see
' ExpandMacros' own comment. F1.2: reverted; the original 8-line
' observation this pass chased was correct, expected behavior, not a
' bug - the known cost of routing a rule through a macro at all, paid
' by every macro-based rule since long before F.1, newly visible only
' because F.1 put macros in bare-statement position for the first time.
' IN0.5: added VlaCompileToForms, a new Public entry point beside
' VlaTranspile - same Tokenize/ParseAll/DefineMacro/ExpandMacros
' pipeline, stopped one step short of EmitTop so a second backend can
' consume the expanded forms directly. Built for VLA_Interpreter.bas,
' the walking-skeleton interpreter (dev-rig only). VlaTranspile itself
' is unchanged; the new function duplicates its setup rather than
' factor it out, on purpose - see VlaCompileToForms' own comment.
' LX4.0: the alias chokepoint. New ResolveHeadAlias, consulted at the
' top of all four Select Case dispatchers (EmitTop/EmitStmt/EmitExpr/
' EmitFormula) instead of the raw folded head - so (fijar! x 5) reaches
' the same Case "set!" arm (set! x 5) would. Backed by mHeadAliases, a
' new module-level Collection set fresh at the top of VlaTranspile and
' VlaCompileToForms from VLA_HeadTable.VlaHeadTableAliasMap - built
' once per compile, O(1) per lookup, not VlaHeadTableRow's per-call
' linear scan. VLA_HeadTable.bas carries the actual alias data (three
' rows today); this module only consumes it.
' F3.0: PreludeMacros() was a 322-line VBA string builder (every quote
' in every docstring/body doubled to survive a VBA string literal -
' the escape-stack bug class this move deletes); it is now a ten-line
' file read. The 33 macros themselves moved verbatim to scripts/
' prelude.vla, with their surrounding design-rationale comments
' preserved as VLA ;-comments - nothing here changed their expansion.
' New PreludeVlaPath resolves the file ThisWorkbook-relative (the
' engine's own location), not ActiveWorkbook-relative the way P.L7's
' (include ...) resolves a user's own library file - the prelude is
' core language infrastructure, not something scoped per user
' workbook, so it does not reuse SpliceIncludes/ReadIncludeFile's path
' resolution, only the same "read a file's text and splice it in"
' shape their existence proved out. mLineOffset (already computed as
' CountLf(prelude) + 1 at every call site, unchanged) absorbs whatever
' line count the file turns out to have, so the ' vla:N tags in
' generated VBA are unaffected by prelude.vla's size - confirmed via
' the golden diff, not assumed.

' =====================================================================
'  VLA - "Visual Lisp for Applications"
'  An s-expression front-end for VBA, written in VBA.
'
'  Design rule: VLA is VBA wearing parentheses. Every form maps 1:1
'  to a VBA construct; VBA semantics are preserved exactly. The one
'  genuinely Lispy addition is a template macro system (defmacro),
'  which expands before code generation.
'
'  Public API:
'    VlaTranspile(source)                 -> VBA source text
'    VlaCompileToModule(source, modName)  -> injects code into a module
'                                            (needs "Trust access to the
'                                            VBA project object model")
'    VlaExpand source                     -> L1: prints the macro
'                                            expansion (one pass, then
'                                            fixpoint) to the Immediate
'                                            window - "show me what this
'                                            form becomes"
'    VlaExpandText(source, toFixpoint)    -> the programmatic half
'                                            (Explain's macro panel,
'                                            the self-test)
'    VlaPushContext / VlaPopContext       -> F.5: save/restore this
'                                            module's compile-time state
'                                            around a nested or isolated
'                                            compile. VlaContextDepth
'                                            reports the stack's size.
'
'  Datum representation (reader output):
'    list   = Collection
'    atom   = String
'             - string literals carry a leading quote char as a marker
'             - everything else (symbols, numbers) is raw text
' =====================================================================

Private mMacros As Object          ' PDICT.0: late-bound Scripting.Dictionary
                                   ' (no project reference), macro records
                                   ' keyed by lowercase name - was a
                                   ' Collection; GetMacro's own keyed lookup
                                   ' is a linear scan on a Collection, O(n)
                                   ' per call, confirmed the dominant compile-
                                   ' side cost (P-PROF's own live "expand"
                                   ' number) once a chain resolves through a
                                   ' growing table. See GetMacro/DefineMacro,
                                   ' below, for the read/write sides.
Private mHeadAliases As Collection ' LX.4: folded alias -> canonical head
                                   ' symbol (VLA_HeadTable.VlaHeadTableAliasMap,
                                   ' fresh per compile, same reason mMacros is -
                                   ' no module-level cache survives a recompile,
                                   ' LESSONS.md XXXII)
Private mSubDocs As Collection     ' P.L6: procedure-docstring records
                                   ' (name-as-written, doc, kind), keyed
                                   ' by lowercase name. Reset by
                                   ' VlaTranspile ONLY - never by
                                   ' VlaExpandText, so apropos's own
                                   ' prelude+carry refresh cannot wipe
                                   ' the program it is about to list.
                                   ' Reflects the LAST program
                                   ' transpiled; a state wipe empties
                                   ' it like everything else (S5.1).
Private mAproposCarry As String    ' L11.1: the loaded vocabulary's
                                   ' defmacro text, pushed down by
                                   ' VLA_English at load - apropos
                                   ' re-parses it per ask, so its
                                   ' answer never depends on which
                                   ' machinery parsed last
Private mAproposRulesCarry As String  ' APROPOSPLUS: one line per loaded
                                   ' rule ("pattern -> template"),
                                   ' pushed down alongside mAproposCarry
                                   ' by the same call - already-resolved
                                   ' text (VLA.VlaWriteForm's own output
                                   ' at carry time), so unlike the macro
                                   ' text above it needs no per-ask
                                   ' re-expansion, only a per-ask match
Private mFuncName As String        ' name of function currently being emitted
Private mInFunction As Boolean     ' inside Function (vs Sub)?

' C2: the source map. The tokenizer records a raw line per token; the
' parser keys each parsed list (by object identity, ObjPtr) to its
' USER-source line - raw line minus the prepended prelude's offset, so
' prelude forms map to nothing. Macro expansion rebuilds every list,
' and the rebuild path copies the tag; macro-TEMPLATE forms carry none
' (deliberate: the map marks the user's statements, and a template
' statement has no user line to name). The emitter appends ' vla:N to
' the first line of every mapped statement and remembers the last
' mapped line it saw, so an emitter error can name a neighborhood.
' V4: the three-layer join. An (at-line N stmt ...) wrapper is a
' zero-runtime annotation carrying the line in the ORIGINATING
' source - for English-generated VLA, the sentence's line - and the
' map tag becomes ' vla:N src:M for everything emitted inside it
' (dynamic extent: begin bodies and If arms inherit; a nested
' wrapper overrides within its own extent). The transpiler still
' knows nothing about English: "the source one layer up" is a
' layer-neutral idea, and hand-written VLA simply never writes the
' wrapper. Emitter failures name both coordinates when both are
' known ("near vla line N, source line M").
Private mTokLines() As Long        ' P-TOK: raw line number per token,
                                   ' 1-based array parallel to a
                                   ' Tokenize call's own returned token
                                   ' array - O(1) positional access,
                                   ' materialized once after the scan
                                   ' (see Tokenize/TokRawLine)
Private mLineOffset As Long        ' raw lines occupied before user line 1
Private mFormLines As Collection   ' CStr(ObjPtr(list)) -> CStr(source line)
Private mTcoName As String         ' L18: the current function's
                                   ' SymName when TCO is armed ("" =
                                   ' off); armed by EmitProc only for
                                   ' an all-ByVal function whose body
                                   ' contains a (return (self ...)).
Private mTcoParams As Collection   ' L18: the VBA parameter names, in
                                   ' order, for the rebind.
Private mProcEmitted As Boolean    ' G12.1: a procedure has been
                                   ' emitted this transpile - module-
                                   ' level declarations after this
                                   ' point refuse (see TopoGuard).
Private mIncMap As Collection      ' P.L7: expanded-source line k ->
                                   ' "label|fileLine|anchor" at item k.
                                   ' Nothing when the program has no
                                   ' includes (the splicer's fast path),
                                   ' so every consult site falls through
                                   ' to today's wording byte-for-byte.
                                   ' label "" = the main program; anchor
                                   ' = the main-file line an included
                                   ' line reports at RUNTIME (its
                                   ' include statement's line), while
                                   ' transpile-time errors use
                                   ' label+fileLine for precision.
Private mEmitLine As Long          ' last mapped statement line emitted
Private mAtLine As Long            ' V4: innermost enclosing at-line value
                                   ' (0 = none); dynamic extent via
                                   ' save/restore in the at-line case
Private mGenRow As String         ' LISTOPS-PROVENANCE: innermost enclosing
                                   ' gen-row label ("" = none); dynamic
                                   ' extent via save/restore in the
                                   ' gen-row case, same shape as mAtLine -
                                   ' safe here (unlike VLA_English.bas's
                                   ' own rowTag) because emission is one
                                   ' immediate recursive walk, never
                                   ' deferred the way test-success proofs
                                   ' are at the vocabulary-load layer.
Private mSlabCounter As Long       ' PF.4c: for-each-row's own hidden
                                   ' bookkeeping variables (the bulk-
                                   ' read array, the captured range, the
                                   ' row/column indices) need names the
                                   ' user never supplied - the no-gensym
                                   ' doctrine's own answer (LESSONS.md
                                   ' XVII: "every label and temporary is
                                   ' a caller-supplied parameter... VBA's
                                   ' own compile step is the collision
                                   ' detector") is a FIXED name, not a
                                   ' random one, so a second use in the
                                   ' same Sub still needs a distinct
                                   ' suffix or it would collide with the
                                   ' first's own Dim. Deterministic, not
                                   ' gensym: a monotonic per-transpile
                                   ' counter, one increment per for-each-
                                   ' row form actually emitted - the SAME
                                   ' program always produces the SAME
                                   ' suffix sequence, so a reader can
                                   ' predict "the Nth for-each-row in
                                   ' this file gets suffix N" by hand,
                                   ' exactly the traceability bar this
                                   ' project's own doctrine sets, not an
                                   ' opaque freshly-minted symbol.
Private mSlabRowVar As String      ' PF.4c: a real bug caught live
                                   ' (owner-run VlaSelfTestsAll) forced a
                                   ' redesign - the first EmitForEachRow
                                   ' sliced each row into its own small
                                   ' array, ran the body, then copied it
                                   ' back into the bulk array in a
                                   ' SEPARATE loop AFTER the body. A bare
                                   ' exit-for inside the body skips
                                   ' straight past that trailing copy-
                                   ' back, so the row WHERE the break
                                   ' happens lost its own mutation - not
                                   ' "up to and including the break," up
                                   ' to but NOT including it. Fixed by
                                   ' removing the separate row array
                                   ' entirely: (row i)/(set! (row i) v)
                                   ' now compile directly to arr(iVar, i)
                                   ' (EmitExpr's own dispatch, below),
                                   ' so nothing is ever deferred for a
                                   ' bare Exit For to skip - exit-for
                                   ' needs no special-casing of its own
                                   ' at all. "" = no for-each-row
                                   ' currently active; dynamic extent via
                                   ' save/restore in EmitForEachRow, the
                                   ' same shape mAtLine/mGenRow already
                                   ' use above - a plain nested for/for-
                                   ' each/while/do-until inside the body
                                   ' never touches this, so (row i)
                                   ' keeps resolving to the enclosing
                                   ' for-each-row's own row through any
                                   ' such nesting, unchanged.
Private mSlabArrVar As String      ' PF.4c: the bulk array name (row i)
                                   ' substitutes into, paired with
                                   ' mSlabRowVar - see that declaration.
Private mSlabIVar As String        ' PF.4c: the row-index variable name
                                   ' (row i) substitutes into, paired
                                   ' with mSlabRowVar - see that
                                   ' declaration.
Private mBudgetOn As Boolean       ' L15: when True, applications
Private mExpandBudget As Long      '   beyond the budget do not fire -
                                   '   the stepper's whole mechanism.
                                   '   Boolean gate because a Long's
                                   '   default 0 would silently budget
                                   '   EVERY parse to zero.
Private mExpandFired As Long       ' L1: macro applications since the
                                   ' current public entry reset it -
                                   ' ExpandMacros and ExpandOnePass each
                                   ' count one per application, so
                                   ' VlaExpandText can report "did any
                                   ' macro fire?" without diffing text
Private mFormulaExpandOpen As Boolean ' VLA.4: the walk stands inside a
                                   ' formula-expand's argument, where
                                   ' another one is refused; cleared by
                                   ' ExpandFormulaCall on every way out
Private mDefines As Object         ' VLA.6: the words a source's defines
                                   ' bind, a late-bound Scripting.Dictionary
                                   ' keyed by the folded word, each its form
                                   ' as written; Nothing until a form's
                                   ' defines register, so that the walk asks
                                   ' one Is Nothing of an atom in a source
                                   ' with none. Reset beside mMacros by each
                                   ' entry point, carried by VlaFrame
Private mDefineMet As Boolean      ' VLA.6: the walk passed a define by in
                                   ' the form ExpandSourceForm expands
Private mFormulaTexts As Collection ' VLA.6: the texts formula-expand wrote
                                   ' in that form, read for a word used in
                                   ' the form that defines it; Nothing for
                                   ' none

' F5.0: the push/pop stack itself - NOT one of the eighteen fields a
' VlaFrame snapshots (it is the stack that HOLDS those snapshots).
' Nothing until the first VlaPushContext call (Is Nothing, not an
' empty Collection, so VlaContextDepth can tell "never pushed" from
' "pushed and popped back to zero" without needing to - both report 0).
Private mContextStack As Collection

' PORT.1: VBA's real rule, found live (compile error: "Only comments
' may appear after End Sub, End Function, or End Property") - a bare
' module-level variable declaration is NOT legal scattered between
' procedures the way a Sub/Function is; every module-level Private/Dim
' in a standard module must sit in the one contiguous declarations
' region before the file's first procedure (ProfMs, immediately below
' PPROF.0's own accumulators - VlaTranspile itself follows right after).
' The two Public Subs that USE these (VlaSetPreludeOverride/
' VlaClearPreludeOverride) stay near PreludeMacros, where they're
' actually used - only the bare declarations had to move.
Private mPreludeOverrideSet As Boolean
Private mPreludeOverrideText As String

' PPROF.0: the one switch (bare Public Boolean, mRunScaleTests' own
' shape) and the four compile-side accumulators it gates - see the
' header comment above and VlaTranspile's own body for where each is
' timed. mProfCalls counts every VlaTranspile call attempted while the
' switch is on, success or failure alike (incremented up front, not at
' the end, so a call that ultimately raises still counts as profiled).
Public mProfileOn As Boolean
Public mProfMsTokenize As Double
Public mProfMsParse As Double
Public mProfMsExpand As Double
Public mProfMsEmit As Double
Public mProfCalls As Long

' IO.1: every module VlaCompileToModule writes carries this as its
' first line, so a later compile can tell "ours, safe to overwrite"
' from "something else's, refuse" - see VlaCompileToModule. A leading
' comment before Option Explicit is already shipping precedent
' (VlaInjectRuntime prepends three lines the same way), so this adds
' no new VBA-legality question, only a fourth line where there were
' three.
Private Const FRAZARO_MODULE_MARKER As String = _
    "' Frazaro-managed module - rewritten on every Check/Run/compile; edits here do not persist."

' VLA.1: the deepest a list may nest, the top-level list being 1. ParseForm
' refuses a list opened deeper, and so do the expander's walks, so that no
' stage after them recurses past it: VBA's stack gives out between 150 and
' 250 frames of ExpandMacros (LISTOPS-BUDGET's measurement). vla-lang's
' reader::FORM_DEPTH_LIMIT is the same number.
Private Const VLA_FORM_DEPTH_LIMIT As Long = 64

' =====================================================================
'  Public API
' =====================================================================

' =====================================================================
'  CO.4 - grammar semantic versioning. THE DECISION FIRST, because the
'  code below is small only because of it: the grammar's compatibility
'  version IS VLA_RELEASE_VERSION, under SD-14's already-decided
'  triggers. No second version axis was invented. Scoped by reading
'  what actually depends on this item - CO.3's compiled-in version
'  stamp, F.10's `requires: version:X` - before designing anything.
'
'  WHY REUSE IS SOUND, not merely convenient. The one thing that would
'  break it is a release that changes grammar-visible behavior without
'  the version number moving enough to notice. SD-14 already
'  forecloses that, by its own definitions:
'    PATCH is DEFINED as "no observable change to what an existing
'      pilot's sentences do" - so a release that changed
'      grammar-visible behavior is not a PATCH at all; it has already
'      become MINOR or MAJOR before this question gets asked.
'    MINOR is "adds something a user can now do or say that they
'      couldn't before", and its own first worked example is "a new
'      grammar rule or section" - precisely the event a phrasebook's
'      `requires: version:` exists to test for.
'    MAJOR is SD-4 invoked, OR "a comparable architectural break where
'      an existing pilot workbook's program could behave differently
'      after upgrading" - broader than SD-4 alone, so a break in the
'      phrasebook DSL itself is covered even when no shipped
'      instructions.txt sentence changed meaning.
'  A separate grammar number would restate all three rules and then
'  need its own doctrine for when the two disagree. This project has
'  already been burned once by inventing a second scheme where one
'  covered the need (F.10/CO.3's near-duplication, corrected the same
'  session this item was scoped); not repeated here.
'
'  WHAT VLA_CORE_VERSION IS NOT: it is not this, and a
'  `requires: version:` check must never read it. It and the 23 other
'  VLA_xxx_VERSION constants each name the ROADMAP ITEM that last
'  touched THAT module - this file's own header, line 3, says so
'  outright. Three modules carry "LINTERPOLATE.0" simultaneously right
'  now (VLA, VLA_Interpreter, VLA_Messages) because one item touched
'  all three, which is not something a version axis can do. It stays
'  exactly as it is: an opaque, human-readable, per-module changelog
'  tag with no ordering contract, serving a different purpose.
'
'  ORDERING is plain numeric MAJOR.MINOR.PATCH, left to right - every
'  value this project has ever shipped, and no more. Deliberately NOT
'  full semver: pre-release suffixes have never been needed here, and
'  one would flow straight into the `v<version>` git tag that
'  tools/release.ps1 builds, into Frazaro.iss's AppVersion, and into
'  Windows' own DisplayVersion in Add/Remove Programs - none of which
'  has ever seen a suffix. A malformed version REFUSES IN WORDS (LX.8)
'  rather than comparing as anything at all: "banana" must never
'  quietly satisfy a requirement, and must never quietly fail one
'  either, since both hide the typo that caused it.
'
'  Consumers to keep in step - a full-repo census, because this
'  constant's VALUE is parsed and embedded well outside this module:
'  tools/release.ps1 (a regex on the exact `Public Const ... As String
'  = "..."` line shape, then exact string equality against -Version and
'  the v<version> tag), VLA_Build.bas's VlaWriteInstallerVersion ->
'  installer\version.iss -> Frazaro.iss's AppVersion, and VLA_IDE.bas's
'  understands-sheet and Copy Feedback. Nothing below changes that
'  line's shape or its value's format, on purpose.
' =====================================================================

' Parse "MAJOR.MINOR.PATCH" into its three numbers. A PURE PREDICATE:
' returns False on anything malformed and never raises, so a caller
' validating text a user typed (F.10's `requires:` line) can ask the
' question without an error handler around it. VlaVersionCompare is
' the raising verb; this is the asking verb. On False the three
' out-params are left at 0 rather than half-filled - a partial parse
' is not a version.
Public Function VlaVersionParse(ByVal s As String, ByRef major As Long, _
                                ByRef minor As Long, ByRef patch As Long) As Boolean
    major = 0
    minor = 0
    patch = 0
    Dim parts() As String
    ' Guarded before Split rather than relying on UBound(-1) over the
    ' zero-length array Split("") returns: that is safe VBA, but this
    ' is the one path a caller reaches with genuinely empty text, and
    ' an explicit exit says so without depending on the subtlety.
    If Len(Trim$(s)) = 0 Then Exit Function
    parts = Split(Trim$(s), ".")
    If UBound(parts) <> 2 Then Exit Function
    Dim n(0 To 2) As Long
    Dim i As Long, j As Long, p As String, c As Long
    For i = 0 To 2
        p = parts(i)
        ' Len > 9 cannot overflow a Long (max 2147483647, 10 digits);
        ' an empty part ("1..3", "1.2.") is not a number at all.
        If Len(p) = 0 Or Len(p) > 9 Then Exit Function
        For j = 1 To Len(p)
            c = AscW(Mid$(p, j, 1))
            ' Digits only - this also rejects a leading "+"/"-" and any
            ' pre-release suffix, which is the point: neither has a
            ' defined order here, so neither may parse.
            If c < 48 Or c > 57 Then Exit Function
        Next j
        n(i) = CLng(p)
    Next i
    major = n(0)
    minor = n(1)
    patch = n(2)
    VlaVersionParse = True
End Function

' -1 / 0 / +1 as `a` is older than / the same as / newer than `b`.
' Refuses in words if either side is malformed: a version that cannot
' be read has no place in an ordering, and silently answering 0
' ("same") would let a `requires:` typo pass review unnoticed.
Public Function VlaVersionCompare(ByVal a As String, ByVal b As String) As Long
    Dim aMaj As Long, aMin As Long, aPat As Long
    Dim bMaj As Long, bMin As Long, bPat As Long
    If Not VlaVersionParse(a, aMaj, aMin, aPat) Then
        VLA_Messages.RaiseMsg "vla-version-malformed", "value", a
    End If
    If Not VlaVersionParse(b, bMaj, bMin, bPat) Then
        VLA_Messages.RaiseMsg "vla-version-malformed", "value", b
    End If
    If aMaj <> bMaj Then
        VlaVersionCompare = Sgn(aMaj - bMaj)
    ElseIf aMin <> bMin Then
        VlaVersionCompare = Sgn(aMin - bMin)
    ElseIf aPat <> bPat Then
        VlaVersionCompare = Sgn(aPat - bPat)
    End If
End Function

' The verb F.10's `requires: version:X` and CO.3's compiled-in stamp
' both actually want: is the running build at least `minimum`?
' Compares against VLA_RELEASE_VERSION - DI.3a's own "this one
' constant is what 'what build is this' actually means". Raises
' through VlaVersionCompare if `minimum` is malformed, which is the
' refusal a phrasebook with a typo'd requires: line should get.
Public Function VlaVersionAtLeast(ByVal minimum As String) As Boolean
    VlaVersionAtLeast = (VlaVersionCompare(VLA_RELEASE_VERSION, minimum) >= 0)
End Function

' PPROF.0: Rule-12 duplicate of VLA_DevRig.bas's own Private DevMs (not
' visible across modules) - same midnight-wrap handling, kept beside the
' phase vars it feeds rather than promoting DevMs to Public for one
' extra caller.
Private Function ProfMs(ByVal t0 As Double) As Double
    Dim d As Double
    d = Timer - t0
    If d < 0 Then d = d + 86400#
    ProfMs = d * 1000#
End Function

Public Function VlaTranspile(ByVal source As String) As String
    ' P.L7: reader-level include splice, BEFORE tokenization - the
    ' prelude offset below is untouched by it, and the per-line map
    ' it builds (or leaves at Nothing) feeds SrcLineTag, the
    ' emitfail formatter, and MapTag.
    source = SpliceIncludes(source)
    Dim prelude As String
    prelude = PreludeMacros()
    mLineOffset = CountLf(prelude) + 1   ' prelude lines + the joining break
    Set mFormLines = New Collection
    mEmitLine = 0
    mAtLine = 0
    mGenRow = ""
    mSlabCounter = 0
    mSlabRowVar = ""
    If mProfileOn Then mProfCalls = mProfCalls + 1

    Dim pt0 As Double
    If mProfileOn Then pt0 = Timer
    Dim toks() As String
    toks = Tokenize(prelude & vbCrLf & source)
    If mProfileOn Then mProfMsTokenize = mProfMsTokenize + ProfMs(pt0)

    If mProfileOn Then pt0 = Timer
    Dim forms As Collection
    Set forms = ParseAll(toks)

    Set mMacros = CreateObject("Scripting.Dictionary")   ' PDICT.0
    Set mDefines = Nothing          ' VLA.6: no word is defined yet
    Set mHeadAliases = VLA_HeadTable.VlaHeadTableAliasMap()   ' LX.4
    Set mSubDocs = New Collection   ' P.L6: per-transpile, see declaration
    mProcEmitted = False            ' G12.1: per-transpile topology
    mFuncName = ""
    mInFunction = False

    ' Pass 1: collect macro definitions (order-independent). PPROF.0:
    ' timed together with ParseAll above, one "parse" bucket - P-PROF's
    ' own four named phases (tokenize/parse/expand/emit, ALPHA6_ROADMAP's
    ' own record) never carved out a fifth for macro collection, and
    ' DefineMacro's own cost (table insertion) is a different question
    ' from GetMacro's (Pass 2's own per-lookup cost, P-DICT's actual
    ' target) - folding them together here would blur exactly the number
    ' P-DICT needs clean.
    Dim bodyForms As New Collection
    Dim f As Variant
    For Each f In forms
        If ListHeadIs(f, "defmacro") Then
            Dim cl As Collection
            Set cl = f
            DefineMacro cl
        Else
            bodyForms.Add f
        End If
    Next
    If mProfileOn Then mProfMsParse = mProfMsParse + ProfMs(pt0)

    ' Pass 2: expand macros, then emit. On failure, name the last
    ' mapped statement line the emitter reached - approximate on
    ' purpose (macro-generated statements carry no line), which is
    ' why it says "near". Raising from the handler propagates to the
    ' caller with the augmented text; nothing is swallowed.
    Dim sb As String, sbU As Long
    ' V8: builder - the old r = r & EmitTop(...) copied the whole
    ' output once per top form.
    SbAdd sb, sbU, "Option Explicit" & vbCrLf & vbCrLf
    ' PPROF.0: profPhase names whichever of expand/emit is mid-flight,
    ' so emitfail below can still credit it on a raise (e.g. a
    ' 5000+-deep chain is SUPPOSED to fail once ExpandMacros' trampoline
    ' cap fires - that real cost belongs in the expand bucket, not lost
    ' silently). Cleared right after each successful add so a later,
    ' unrelated failure elsewhere in this function can never double
    ' count a phase that already finished.
    Dim profPhase As String
    Dim expanded As Variant
    On Error GoTo emitfail
    For Each f In bodyForms
        If mProfileOn Then
            pt0 = Timer
            profPhase = "expand"
        End If
        AssignVar expanded, ExpandSourceForm(f)   ' VLA.6
        If mProfileOn Then
            mProfMsExpand = mProfMsExpand + ProfMs(pt0)
            profPhase = ""
            pt0 = Timer
            profPhase = "emit"
        End If
        SbAdd sb, sbU, EmitTop(expanded)
        If mProfileOn Then
            mProfMsEmit = mProfMsEmit + ProfMs(pt0)
            profPhase = ""
        End If
        SbAdd sb, sbU, vbCrLf
    Next
    VlaTranspile = SbText(sb, sbU)
    ' IN.7 (button-click half): the shared button-creation helper
    ' (EmitStmt's "make-button" case comment has the full reasoning)
    ' rides at the bottom, appended once, only when this program
    ' actually used make-button - L4.2's own "phrasebook macros ride
    ' at the bottom, appended only when used" precedent (VLA_English.
    ' bas), applied here instead of always emitting dead code into
    ' every compiled module.
    If InStr(1, VlaTranspile, "Call " & VLA_MAKEBUTTON_SUB & "(") > 0 Then
        VlaTranspile = VlaTranspile & vbCrLf & MakeButtonHelperText()
    End If
    Exit Function
emitfail:
    ' PPROF.0: credit whichever of expand/emit was mid-flight when this
    ' fired, BEFORE any of the existing error-augmentation logic below -
    ' reading Err.Number/Source/Description here changes nothing about
    ' them, so this cannot alter what gets raised. profPhase is "" (and
    ' this is a no-op) for any failure outside the Pass 2 loop, and for
    ' the ordinary mProfileOn=False path it is skipped entirely.
    If mProfileOn And Len(profPhase) > 0 Then
        If profPhase = "expand" Then
            mProfMsExpand = mProfMsExpand + ProfMs(pt0)
        ElseIf profPhase = "emit" Then
            mProfMsEmit = mProfMsEmit + ProfMs(pt0)
        End If
        profPhase = ""
    End If
    If mEmitLine > 0 And InStr(1, Err.Description, "vla line") = 0 Then
        ' P.L7: an emit-time error inside an included file names the
        ' file and its LOCAL line; main-file errors keep today's
        ' wording byte-for-byte (IncLookup is False on the fast path).
        Dim lbE As String, flE As Long, anE As Long
        If IncLookup(mEmitLine, lbE, flE, anE) Then
            ' Mapped: included lines name their file; main-file lines
            ' after an include use the MAIN file's own numbering
            ' (flE), never the expanded line.
            If Len(lbE) > 0 Then
                Err.Raise Err.Number, Err.Source, Err.Description & " (near " & lbE & " line " & flE & AtLineSuffix() & ")"
            Else
                Err.Raise Err.Number, Err.Source, Err.Description & " (near vla line " & flE & AtLineSuffix() & ")"
            End If
        End If
        Err.Raise Err.Number, Err.Source, Err.Description & " (near vla line " & mEmitLine & AtLineSuffix() & ")"
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

' F.6: the "at-line" source-annotation suffix, formatted once instead
' of three times inline (emitfail above repeated the identical IIf on
' every one of its three raise paths). "" when no (at-line ...) form
' encloses the point of failure.
Private Function AtLineSuffix() As String
    If mAtLine > 0 Then AtLineSuffix = ", source line " & mAtLine
End Function

' LX.4: the emitter's one alias-resolution chokepoint. Every Emit*
' Select Case (EmitTop/EmitStmt/EmitExpr/EmitFormula) dispatches on
' this instead of the raw folded head, so (fijar! ...) reaches the
' same Case "set!" arm (set! ...) would - one lookup before the
' Select Case, not a second dispatch table beside it. Reads
' mHeadAliases (set fresh per compile by VlaTranspile/VlaCompileToForms
' from VLA_HeadTable.VlaHeadTableAliasMap) rather than calling
' VLA_HeadTable.VlaHeadTableRow directly: that function's alias match
' is a linear scan over 66 rows, fine for the occasional catalog
' lookup it was built for, wrong for a call on every node of every
' emit - this is a plain keyed lookup instead, O(1) per call, built
' once per compile rather than once per node.
Private Function ResolveHeadAlias(ByVal foldedHead As String) As String
    ResolveHeadAlias = foldedHead
    If mHeadAliases Is Nothing Then Exit Function
    On Error Resume Next
    Dim canon As String
    canon = CStr(mHeadAliases.Item(foldedHead))
    On Error GoTo 0
    If Len(canon) > 0 Then ResolveHeadAlias = canon
End Function

' IN.0.5: the same Tokenize/ParseAll/macro-definition/ExpandMacros
' pipeline VlaTranspile runs, stopped one step short - it hands back
' the expanded top-level FORMS themselves instead of handing them to
' EmitTop, so a second backend (VLA_Interpreter.bas) can consume the
' exact same forms the VBA emitter does, not a re-parse. Still NOT
' factored out of VlaTranspile into a shared helper both call, though
' this comment's original excuse for that is gone: IN.2 (the real
' evaluator) is built, and VLA_Interpreter stopped being a dev-only
' caller at IN.6 (it ships, and is the DEFAULT runtime a real Run
' reaches first). What still argues against touching it today is
' narrower than "not yet warranted" - VlaTranspile's golden-diff-empty
' invariant is a real, proven-fragile thing to risk (F.9/F.3's own
' history), so factoring the shared prefix out is a legitimate,
' still-open cleanup, not a closed question - just one that deserves
' its own scoped pass with its own golden-diff verification, not a
' rider on whatever else happened to be touching this file.
Public Function VlaCompileToForms(ByVal source As String) As Collection
    source = SpliceIncludes(source)
    Dim prelude As String
    prelude = PreludeMacros()
    mLineOffset = CountLf(prelude) + 1
    Set mFormLines = New Collection
    mEmitLine = 0
    mAtLine = 0
    mGenRow = ""
    mSlabCounter = 0
    mSlabRowVar = ""
    Dim toks() As String
    toks = Tokenize(prelude & vbCrLf & source)
    Dim forms As Collection
    Set forms = ParseAll(toks)

    Set mMacros = CreateObject("Scripting.Dictionary")   ' PDICT.0
    Set mDefines = Nothing          ' VLA.6: no word is defined yet
    Set mHeadAliases = VLA_HeadTable.VlaHeadTableAliasMap()   ' LX.4 - unused
                        ' by this function (it never calls Emit*), set
                        ' anyway to match VlaTranspile so neither entry
                        ' point leaves mHeadAliases stale from a prior call
    Set mSubDocs = New Collection
    mProcEmitted = False
    mFuncName = ""
    mInFunction = False

    Dim bodyForms As New Collection
    Dim f As Variant
    For Each f In forms
        If ListHeadIs(f, "defmacro") Then
            Dim cl As Collection
            Set cl = f
            DefineMacro cl
        Else
            bodyForms.Add f
        End If
    Next

    Dim expanded As New Collection
    For Each f In bodyForms
        expanded.Add ExpandSourceForm(f)   ' VLA.6
    Next
    Set VlaCompileToForms = expanded
End Function

' F.2: the bare reader half of VlaTranspile/VlaCompileToForms's shared
' prefix, stopped even earlier - tokenize + parse, nothing else. No
' prelude splice (a short fragment is not a program), no macro
' collection/expansion, no golden-diff-sensitive line-tag bookkeeping a
' caller reading a fragment in isolation has no use for. Exists so a
' caller that wants to BUILD forms directly - VLA_English.bas's phrase-
' rule templates, today spliced together as text and handed back
' through this same reader a second time (Tier 1's "double round-trip",
' BETA_REARVIEW.md's F.2) - can turn a VLA-syntax fragment into real
' forms once, at rule-load time, instead of every match. Context-
' pushed/popped (F5.0) so a call mid-compile can never disturb
' mTokLines/mFormLines/mLineOffset the enclosing compile still needs.
Public Function VlaReadForms(ByVal text As String) As Collection
    VlaPushContext
    On Error GoTo fail
    mLineOffset = 0
    Set mFormLines = New Collection
    Dim toks() As String
    toks = Tokenize(text)
    Set VlaReadForms = ParseAll(toks)
    VlaPopContext
    Exit Function
fail:
    Dim errNum As Long, errDesc As String, errSrc As String
    errNum = Err.Number: errDesc = Err.Description: errSrc = Err.Source
    VlaPopContext
    Err.Raise errNum, errSrc, errDesc
End Function

' P-PROBE: RegisterVocabMacro's own probe (VLA_English.bas) needs to
' know whether ONE already-canonicalized defmacro form's own text is
' well-formed - DefineMacro's own reserved-name/rest-param/docstring-
' template checks - and that the text round-trips through the reader.
' Nothing about that needs PreludeMacros() spliced in (a bare defmacro
' form never references a prelude macro) or Pass 2/EmitTop (the old
' probe's own trailing "(sub vla-macro-check ())" called nothing, so
' Pass 2 never validated anything macro-specific - confirmed by
' reading VlaTranspile's own Pass 1/Pass 2 split, not assumed). Going
' through VlaTranspile anyway paid PreludeMacros()'s own disk read
' plus a full Tokenize/ParseAll of prelude.vla (34 KB, 46 macros as of
' this writing) AND DefineMacro on all 46 of them, on EVERY probe
' call - a real, Timer-confirmed ~8-15 ms fixed cost per call,
' independent of the accumulated-vocab-corpus cost the rest of
' P-PROBE's own fix (RegisterVocabMacro) already removed, and NOT
' fixed by it: a live corpus load (134 macros) showed barely any
' change until this was found too. Same F.2/F5.0 shape as
' VlaReadForms just above - context pushed/popped so a call mid-
' compile can never disturb the enclosing compile's own state - just
' without that function's caller handing back the parsed forms (this
' one's whole point is the DefineMacro side effect + its errors, not
' the forms themselves).
Public Sub VlaProbeMacroForm(ByVal macroText As String)
    VlaPushContext
    On Error GoTo fail
    mLineOffset = 0
    Set mFormLines = New Collection
    Dim toks() As String
    toks = Tokenize(macroText)
    Dim forms As Collection
    Set forms = ParseAll(toks)
    Set mMacros = CreateObject("Scripting.Dictionary")   ' PDICT.0
    Dim f As Variant
    For Each f In forms
        If ListHeadIs(f, "defmacro") Then
            Dim cl As Collection
            Set cl = f
            DefineMacro cl
        End If
    Next
    VlaPopContext
    Exit Sub
fail:
    Dim errNum As Long, errDesc As String, errSrc As String
    errNum = Err.Number: errDesc = Err.Description: errSrc = Err.Source
    VlaPopContext
    Err.Raise errNum, errSrc, errDesc
End Sub

' F5.0: snapshots this module's nineteen compile-time fields onto a new
' stack frame, WITHOUT resetting them - see VLA_CORE_VERSION's own F5.0
' comment for why that is safe and correct (VlaTranspile/
' VlaCompileToForms already reset what each needs at its own top). A
' caller wanting a fully independent nested compile just runs one of
' those two entry points after pushing; a caller wanting only a
' checkpoint to roll back to may mutate the module state directly and
' pop to undo it. Object/Collection-typed fields are captured by
' reference (Set), which is safe here specifically because every reset
' site in this module reassigns a FRESH object (a new Collection for
' most fields; PDICT.0 changed mMacros's own reset from "New Collection"
' to a fresh "CreateObject(""Scripting.Dictionary"")", same reassign-
' not-clear shape) rather than clearing one in place - the snapshot's
' own object and the module's are never the same reference after a
' reset, so popping an outer frame can never see a nested compile's
' mutations.
Public Sub VlaPushContext()
    Dim f As New VlaFrame
    Set f.Macros = mMacros
    Set f.Defines = mDefines        ' VLA.6
    Set f.HeadAliases = mHeadAliases
    Set f.SubDocs = mSubDocs
    f.AproposCarry = mAproposCarry
    f.AproposRulesCarry = mAproposRulesCarry
    f.FuncName = mFuncName
    f.InFunction = mInFunction
    f.TokLines = mTokLines
    f.LineOffset = mLineOffset
    Set f.FormLines = mFormLines
    f.TcoName = mTcoName
    Set f.TcoParams = mTcoParams
    f.ProcEmitted = mProcEmitted
    Set f.IncMap = mIncMap
    f.EmitLine = mEmitLine
    f.AtLine = mAtLine
    f.GenRow = mGenRow
    f.BudgetOn = mBudgetOn
    f.ExpandBudget = mExpandBudget
    f.ExpandFired = mExpandFired
    If mContextStack Is Nothing Then Set mContextStack = New Collection
    mContextStack.Add f
End Sub

' F5.0: restores the most recently pushed frame and discards it. Raises
' (does not silently no-op) on an unbalanced pop - a bare stack
' underflow is a programming error in the caller, the same class of
' mistake Nth/HeadSym already raise loudly on elsewhere in this module,
' not a user-facing refusal SD-2 governs.
Public Sub VlaPopContext()
    If mContextStack Is Nothing Then
        VLA_Messages.RaiseMsg "vla-pop-context-empty"
    ElseIf mContextStack.Count = 0 Then
        VLA_Messages.RaiseMsg "vla-pop-context-empty"
    End If
    Dim f As VlaFrame
    Set f = mContextStack.Item(mContextStack.Count)
    mContextStack.Remove mContextStack.Count
    Set mMacros = f.Macros
    Set mDefines = f.Defines        ' VLA.6
    Set mHeadAliases = f.HeadAliases
    Set mSubDocs = f.SubDocs
    mAproposCarry = f.AproposCarry
    mAproposRulesCarry = f.AproposRulesCarry
    mFuncName = f.FuncName
    mInFunction = f.InFunction
    mTokLines = f.TokLines
    mLineOffset = f.LineOffset
    Set mFormLines = f.FormLines
    mTcoName = f.TcoName
    Set mTcoParams = f.TcoParams
    mProcEmitted = f.ProcEmitted
    Set mIncMap = f.IncMap
    mEmitLine = f.EmitLine
    mAtLine = f.AtLine
    mGenRow = f.GenRow
    mBudgetOn = f.BudgetOn
    mExpandBudget = f.ExpandBudget
    mExpandFired = f.ExpandFired
End Sub

' F5.0: 0 both before the first push and after every push has been
' popped back off - see mContextStack's own declaration comment for why
' that is deliberate. A test (or a future nested-compile caller) can
' assert balance with this instead of reaching into module state.
Public Function VlaContextDepth() As Long
    If Not mContextStack Is Nothing Then VlaContextDepth = mContextStack.Count
End Function

' Compiles VLA source straight into a standard module of the active
' VBA project. Requires the host trust setting:
'   File > Options > Trust Center > Trust Center Settings >
'   Macro Settings > "Trust access to the VBA project object model"
Public Sub VlaCompileToModule(ByVal vlaSource As String, ByVal moduleName As String, _
                              Optional ByVal targetWb As Workbook)
    Dim code As String
    code = VlaTranspile(vlaSource)

    Dim proj As Object, comp As Object
    ' Target the workbook the user is working in - correct both when
    ' VLA runs from the same workbook and when it runs from an add-in
    ' (where ActiveVBProject may point at the add-in itself). D1: a
    ' caller that has already captured its host passes it explicitly,
    ' so a focus change between capture and this VBE call cannot
    ' redirect the injection; ActiveWorkbook remains the default for
    ' standalone use. (No "= Nothing" default - object optionals may
    ' not have one, hazard 11's Optional-constant rule; omission IS
    ' the Nothing default.)
    If targetWb Is Nothing Then Set targetWb = ActiveWorkbook
    Set proj = targetWb.VBProject
    On Error Resume Next
    Set comp = proj.VBComponents(moduleName)
    On Error GoTo 0
    If comp Is Nothing Then
        Set comp = proj.VBComponents.Add(1)   ' 1 = vbext_ct_StdModule
        comp.Name = moduleName
    Else
        ' IO.1: refuse to silently destroy a module we did not create.
        ' Three cases reach here: empty (nothing to lose - a user can
        ' add a blank module by hand and never touch it), ours (marker
        ' present, OR the source-map's ' vla:N tag anywhere in the
        ' module - a module compiled by a Frazaro build from BEFORE
        ' this marker existed carries no marker line but is soaked in
        ' that tag on every mapped statement, so treating it as ours
        ' too is what keeps this guard from refusing every workbook's
        ' very next Run the day this pass ships), or a foreign
        ' module/macro that happens to share this name (neither
        ' signal, real content) - only the third refuses. A pilot
        ' prerequisite: this is "the first thing that breaks on a real
        ' enterprise machine," per the roadmap entry this implements.
        If comp.CodeModule.CountOfLines > 0 Then
            Dim existingCode As String
            existingCode = comp.CodeModule.Lines(1, comp.CodeModule.CountOfLines)
            If InStr(1, existingCode, FRAZARO_MODULE_MARKER, vbBinaryCompare) = 0 And _
               InStr(1, existingCode, "' vla:", vbBinaryCompare) = 0 Then
                VLA_Messages.RaiseMsg "vla-module-not-ours", "name", moduleName
            End If
        End If
    End If
    If comp.CodeModule.CountOfLines > 0 Then
        comp.CodeModule.DeleteLines 1, comp.CodeModule.CountOfLines
    End If
    comp.CodeModule.AddFromString FRAZARO_MODULE_MARKER & vbCrLf & code
End Sub

' =====================================================================
'  L1: macroexpand - "show me what this form becomes". Pure tooling,
'  zero compiler territory: parse, expand, print - the emitter is
'  never involved, so any form the reader accepts can be inspected,
'  bare statements included. The prelude's macros are in force and
'  (defmacro ...) forms in the source are collected first, exactly as
'  a transpile would; the definitions themselves are not echoed, only
'  the expanded body forms.
'
'  From the Immediate window:
'    VlaExpand "(with-fast-excel done (set! x 1))"
'  prints two views. ONE PASS replaces each OUTERMOST macro call with
'  its template substitution and leaves the result untouched - what
'  YOUR template produces, nested calls still visible. FIXPOINT runs
'  expansion to the end - what the emitter sees. When no macro fires
'  the form is already core VLA and one normalized printing suffices.
'
'  VlaExpandText is the programmatic half (Explain's macro panel, the
'  self-test): returns the expanded VLA as text; firedCount reports
'  macro applications (0 = core VLA already). Reader errors raise as
'  usual - the printing Sub catches and prints them instead, because
'  an Immediate-window tool should answer in the Immediate window.
'
'  The printed text is a normalized re-serialization of the datums:
'  source-map tags and comments do not survive (this is a debugging
'  view, not a formatter), and string literals re-escape \" and \\ so
'  the output re-reads identically.
' =====================================================================

Public Sub VlaExpand(ByVal source As String)
    Dim onePass As String, full As String
    Dim fired1 As Long, firedN As Long
    On Error GoTo oops
    onePass = VlaExpandText(source, False, fired1)
    full = VlaExpandText(source, True, firedN)
    Debug.Print "=== VlaExpand"
    If firedN = 0 Then
        Debug.Print "  (no macro fired - this is already core VLA)"
        PrintIndented onePass
        Exit Sub
    End If
    Debug.Print "  one pass (each outermost macro call substituted once):"
    PrintIndented onePass
    If full = onePass Then
        Debug.Print "  fixpoint: identical (" & firedN & " macro application" & PluralS(firedN) & ")"
    Else
        Debug.Print "  fixpoint (what the emitter sees; " & firedN & " application" & PluralS(firedN) & "):"
        PrintIndented full
    End If
    Exit Sub
oops:
    Debug.Print "=== VlaExpand"
    Debug.Print "  (could not expand: " & Err.Description & ")"
End Sub

Public Function VlaExpandText(ByVal source As String, ByVal toFixpoint As Boolean, _
                              Optional ByRef firedCount As Long) As String
    ' firedCount is a typed ByRef Optional: pass a Long variable or
    ' omit it (hazard 11 - a Variant may not ride into it).
    Dim prelude As String
    prelude = PreludeMacros()
    mLineOffset = CountLf(prelude) + 1
    Set mFormLines = New Collection
    mEmitLine = 0
    mAtLine = 0
    mGenRow = ""
    mSlabCounter = 0
    mSlabRowVar = ""
    mExpandFired = 0

    Dim toks() As String
    toks = Tokenize(prelude & vbCrLf & source)
    Dim forms As Collection
    Set forms = ParseAll(toks)

    ' Same two-pass shape as VlaTranspile: collect every defmacro
    ' (prelude's and the source's own), expand only the body forms.
    Set mMacros = CreateObject("Scripting.Dictionary")   ' PDICT.0
    Set mDefines = Nothing          ' VLA.6: no word is defined yet
    Dim bodyForms As New Collection
    Dim f As Variant
    For Each f In forms
        If ListHeadIs(f, "defmacro") Then
            Dim cl As Collection
            Set cl = f
            DefineMacro cl
        Else
            bodyForms.Add f
        End If
    Next

    Dim sb As String, sbU As Long
    Dim ex As Variant
    For Each f In bodyForms
        If toFixpoint Then
            AssignVar ex, ExpandSourceForm(f)   ' VLA.6
        Else
            AssignVar ex, ExpandOnePass(f)
        End If
        SbAdd sb, sbU, WritePretty(ex, 0)
        SbAdd sb, sbU, vbCrLf
    Next
    firedCount = mExpandFired
    VlaExpandText = SbText(sb, sbU)
End Function

' One pass: top-down; a list whose head is a macro is replaced by its
' template substitution and NOT walked further - so the substitution
' is shown verbatim, nested macro calls inside it still spelled as
' calls. Lists with non-macro heads recurse into their elements, so
' sibling macro calls each get their one application.
Private Function ExpandOnePass(d As Variant, Optional ByVal walkDepth As Long = 1) As Variant
    If Not IsList(d) Then
        ExpandOnePass = d
        Exit Function
    End If
    If walkDepth > VLA_FORM_DEPTH_LIMIT Then RefuseFormTooDeep d   ' VLA.1
    ' PNTH.0: Object, not Collection - d (and so lst) may be a VlaSlice,
    ' same reasoning as ExpandMacros' own twin widening just above.
    Dim lst As Object
    Set lst = d
    If lst.Count > 0 Then
        If IsSym(lst.Item(1)) Then
            ' P.L5: quote skip - the fixpoint walk's twin, same
            ' reason, same exactness (see ExpandMacros). VLA.6: a define is
            ' left whole as written too, its word and its form, since it
            ' belongs to the fixpoint walk; one pass replaces no word.
            If VLA_Identity.Fold(CStr(lst.Item(1))) = "quote" Or IsDefineForm(lst) Then
                Set ExpandOnePass = lst
                Exit Function
            End If
            Dim rec As Collection
            Set rec = GetMacro(CStr(lst.Item(1)))
            If Not rec Is Nothing Then
                If mBudgetOn And mExpandFired >= mExpandBudget Then GoTo budgetSpent1  ' L15
                Dim r As Variant
                AssignVar r, ExpandOne(lst, rec)
                mExpandFired = mExpandFired + 1
                If IsObject(r) Then Set ExpandOnePass = r Else ExpandOnePass = r
                Exit Function
            End If
budgetSpent1:
        End If
    End If
    ' PNTH.0: indexed, not For Each - lst may be a VlaSlice, no
    ' enumerator to walk with For Each.
    Dim outc As New Collection
    Dim epi As Long
    For epi = 1 To lst.Count
        outc.Add ExpandOnePass(lst.Item(epi), walkDepth + 1)
    Next
    TagLine outc, FormLine(lst)   ' S3.1: same explicit-tag hygiene as
                                  ' the fixpoint rebuild above
    Set ExpandOnePass = outc
End Function

' VLA.1: the expanders' refusal of a list nested past VLA_FORM_DEPTH_LIMIT,
' whether a text or a macro nested it so, at the list's own line, or at
' none for a list a macro made or the prelude holds. An atom past the limit
' is only a deep list's element, so it passes. A Sub of its own, so that
' the recursive walks gain no local (trap 8).
Private Sub RefuseFormTooDeep(d As Variant)
    If Not IsList(d) Then Exit Sub
    Dim loc As String
    Dim ln As Long
    ln = FormLine(d)
    If ln > 0 Then loc = SrcLineTag(ln + mLineOffset)
    VLA_Messages.RaiseMsg "vla-form-too-deep", "limit", CStr(VLA_FORM_DEPTH_LIMIT), "loc", loc
End Sub

' Flat datum writer: symbols and numbers print raw; string literals
' (leading-quote-marked, escapes already processed at tokenize) are
' re-escaped - backslashes FIRST, then quotes, or the backslash that
' quote-escaping introduces would be doubled.
' =====================================================================
'  V8: the string builder - a pre-sized buffer grown by doubling and
'  written with Mid$ assignment, behind two tiny helpers. VBA's
'  r = r & piece reallocates and copies ALL of r per append, so a
'  loop that grows its result is O(n^2) in the result's length -
'  invisible at 100 sentences, dominant at 5,000. The builder makes
'  the same loop linear amortized. Callers hold the state as a local
'  pair (buf, used); SbText truncates the padding off. Behavior-
'  invariant by construction: the output bytes are identical, only
'  the copying is gone. (Rule 12: a twin pair lives in VLA_English -
'  Privates are invisible cross-module, and a hot helper should not
'  pay a cross-module call anyway.)
' =====================================================================
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

Private Function WriteDatum(v As Variant) As String
    If Not IsList(v) Then
        Dim s As String
        s = CStr(v)
        If Left$(s, 1) = Chr$(34) Then
            WriteDatum = """" & Replace(Replace(Mid$(s, 2), "\", "\\"), """", "\""") & """"
        Else
            WriteDatum = s
        End If
        Exit Function
    End If
    ' PNTH.0: Object, indexed not For Each - v may be a VlaSlice (a
    ' cdr/cddr result), which has no enumerator to walk with For Each.
    ' PROLOG.28: every item but the LAST written by the recursive call it
    ' always was; a last item that is itself a list is opened in the same
    ' buffer and walked in a loop, its closing paren owed and paid once at
    ' the end. The text is the same byte for byte - "(a (b c))" either way,
    ' proven over random terms by the item's transliteration (diff28) -
    ' but a cons chain, whose tail IS its last item, no longer nests one
    ' frame per cell. PROLOG's rendering reaches here with whatever a query
    ' built, and an improper chain thousands long (append over an unbound
    ' tail) used to run VBA's stack out.
    Dim lst As Object
    Dim sb As String, sbU As Long
    Dim wi As Long, closers As Long
    Set lst = v
    Do
        SbAdd sb, sbU, "("
        closers = closers + 1
        If lst.Count = 0 Then Exit Do
        For wi = 1 To lst.Count - 1
            If wi > 1 Then SbAdd sb, sbU, " "
            SbAdd sb, sbU, WriteDatum(lst.Item(wi))
        Next
        If lst.Count > 1 Then SbAdd sb, sbU, " "
        If IsList(lst.Item(lst.Count)) Then
            Set lst = lst.Item(lst.Count)
        Else
            SbAdd sb, sbU, WriteDatum(lst.Item(lst.Count))
            Exit Do
        End If
    Loop
    SbAdd sb, sbU, String$(closers, ")")
    WriteDatum = SbText(sb, sbU)
End Function

' F.2: public, single-line wrapper over WriteDatum - the flat half of
' VlaFormat's own pretty-writer, exposed so a caller that BUILT a form
' (rather than parsed one from text) can render it back to VLA source
' text without duplicating the atom/string/list-write rules here a
' second time. Always flat (VlaFormat's width-wrapped WritePretty is
' for human-facing reformatting; a caller splicing this into a larger
' generated line wants one line, not a paragraph).
Public Function VlaWriteForm(v As Variant) As String
    VlaWriteForm = WriteDatum(v)
End Function

' F.13: VlaReadForms, plus each top-level form's own source line -
' captured DURING the read, not after. A first attempt at this exposed
' a bare VlaFormLine(v) wrapper over the existing Private FormLine and
' shipped it broken: VlaReadForms's own VlaPushContext/VlaPopContext
' discipline (F5.0) builds a FRESH mFormLines for the read, then
' restores mFormLines to whatever it held BEFORE the call the instant
' VlaReadForms returns - by design, the same discipline that makes
' nested compiles safe - so the tags this read just computed are
' already gone before any caller can look one up. Live-tested,
' confirmed broken (VlaFormLine returned 0 for everything), root-caused
' immediately. The fix isn't patching the lookup; it's never handing
' control back before the tags are captured - this duplicates
' VlaReadForms's own body (deliberately, rather than touching a
' function every existing caller already trusts) and reads every
' returned form's line while mFormLines is still the fresh one, before
' VlaPopContext restores the old one out from under it.
Public Function VlaReadFormsWithLines(ByVal text As String, ByRef outLines As Collection) As Collection
    VlaPushContext
    On Error GoTo fail
    mLineOffset = 0
    Set mFormLines = New Collection
    Dim toks() As String
    toks = Tokenize(text)
    Dim forms As Collection
    Set forms = ParseAll(toks)
    Set outLines = New Collection
    Dim f As Variant
    For Each f In forms
        outLines.Add FormLine(f)
    Next
    Set VlaReadFormsWithLines = forms
    VlaPopContext
    Exit Function
fail:
    Dim errNum As Long, errDesc As String, errSrc As String
    errNum = Err.Number: errDesc = Err.Description: errSrc = Err.Source
    VlaPopContext
    Err.Raise errNum, errSrc, errDesc
End Function

' Pretty writer: a form that fits on one line at its indent prints
' flat; a wider one opens with its head and gives every remaining
' element its own indented line, closing on the last. Width-based,
' no head knowledge - robust over forms no macro has invented yet.
Private Function WritePretty(v As Variant, ByVal indent As Long) As String
    Dim pad As String
    pad = String$(indent, " ")
    Dim flat As String
    flat = WriteDatum(v)
    If Not IsList(v) Then
        WritePretty = pad & flat
        Exit Function
    End If
    If indent + Len(flat) <= 90 Then
        WritePretty = pad & flat
        Exit Function
    End If
    ' PNTH.0: Object, not Collection - v may be a VlaSlice.
    Dim lst As Object
    Set lst = v
    If lst.Count = 0 Then
        WritePretty = pad & "()"
        Exit Function
    End If
    Dim sb As String, sbU As Long
    SbAdd sb, sbU, pad & "(" & WriteDatum(Nth(lst, 1))
    Dim i As Long
    For i = 2 To lst.Count
        SbAdd sb, sbU, vbCrLf
        SbAdd sb, sbU, WritePretty(Nth(lst, i), indent + 4)
    Next
    SbAdd sb, sbU, ")"
    WritePretty = SbText(sb, sbU)
End Function

Private Function PluralS(ByVal n As Long) As String
    If n <> 1 Then PluralS = "s"
End Function

Private Sub PrintIndented(ByVal text As String)
    Dim ln As Variant
    For Each ln In Split(Replace(text, vbCrLf, vbLf), vbLf)
        If Len(CStr(ln)) > 0 Then Debug.Print "    " & ln
    Next
End Sub

' PORT.1: the one host touch on the translate path (VlaTranspile calls
' this; nothing else in the English->VLA->VBA chain reaches the host at
' all - confirmed by grep before this seam was added, not assumed).
' VlaSetPreludeOverride lets a caller supply the prelude as text and
' skip PreludeVlaPath/VlaEmbeddedText entirely - the seam a host-free
' caller (VLA_Browser.bas; eventually a non-VBA port) needs, added without
' touching a single existing caller: the override defaults to unset, so
' every call site that has never heard of this stays byte-identical.
' Deliberately a global toggle, not a VlaTranspile parameter - matches
' this codebase's own idiom for an opt-in alternate mode (VlaMessageCapture/
' mMsgCaptureOn, VLA_Runtime.bas) rather than threading a new parameter
' through VlaTranspile's many call sites for one caller's benefit.
' (mPreludeOverrideSet/mPreludeOverrideText themselves are declared up
' in the module's declarations region, above VlaTranspile - see that
' declaration's own comment for why a bare Private cannot live here.)
Public Sub VlaSetPreludeOverride(ByVal preludeText As String)
    mPreludeOverrideText = preludeText
    mPreludeOverrideSet = True
End Sub

Public Sub VlaClearPreludeOverride()
    mPreludeOverrideSet = False
    mPreludeOverrideText = ""
End Sub

' F.3: prelude.vla, not a VBA string - see the file's own header for
' the reasons (deletes the escape-stack bug class every doubled quote
' below used to risk, gains lint coverage, is editable by anyone who
' does not read VBA). Read fresh on every compile - no module-level
' cache, same reason nothing else per-compile is cached here
' (LESSONS.md XXXII: a recompile drops module state) - and the file
' read costs nothing a human would notice next to everything else one
' compile already does.
' DI2.1: an external prelude.vla still wins when present - unchanged
' from F.3, so a dev checkout or a user's own override always takes
' priority. Only when PreludeVlaPath finds neither candidate does this
' fall back to the add-in's own embedded copy (VLA_Build's
' EmbedTextAsSheet, written into "VLAp_Source" at build time) - the
' shipped add-in ships every VLA compile needs, not just the file.
Private Function PreludeMacros() As String
    ' PORT.1: checked first, before any host touch below - so a caller
    ' that has set the override never reaches PreludeVlaPath/Dir$/
    ' ThisWorkbook/VlaEmbeddedText at all, not even to fail past them.
    If mPreludeOverrideSet Then
        PreludeMacros = mPreludeOverrideText
        Exit Function
    End If
    Dim p As String
    p = PreludeVlaPath()
    If Len(p) > 0 Then
        PreludeMacros = VLA_Loader.VlaReadFile(p)
    Else
        PreludeMacros = VlaEmbeddedText("VLAp_Source")
        If Len(PreludeMacros) = 0 Then
            VLA_Messages.RaiseMsg "vla-prelude-not-found"
        End If
    End If
End Function

' prelude.vla's path - ThisWorkbook-relative (the engine's own
' location: the dev workbook during development, the built add-in
' once deployed), NOT ActiveWorkbook/HostBook-relative like a user's
' own vocab file (VLA_IDE.bas's IdeVocabPath) - the prelude is core
' language infrastructure, not something scoped per user workbook.
' Same two-candidate shape VLA_DevRig.bas's DevFindFile and
' VLA_Tests.bas's FindDevFile already use for other engine-shipped
' dev files - this is the one production caller of that shape.
' DI2.1: returns "" rather than raising when neither candidate exists -
' PreludeMacros decides what "" means (the embedded fallback, or a
' refusal if that's empty too), the same way IdeVocabPath already
' returns its best-guess candidate for VLA_IDE's own caller to judge.
Private Function PreludeVlaPath() As String
    Dim sep As String
    sep = Application.PathSeparator
    Dim c1 As String, c2 As String
    c1 = ThisWorkbook.Path & sep & "prelude.vla"
    c2 = ThisWorkbook.Path & sep & "scripts" & sep & "prelude.vla"
    If Len(Dir$(c1)) > 0 Then
        PreludeVlaPath = c1
    ElseIf Len(Dir$(c2)) > 0 Then
        PreludeVlaPath = c2
    End If
End Function

' DI2.1: reads a plain-text asset back from a very-hidden sheet -
' VLA_Build.EmbedTextAsSheet's own reader half. Deliberately Public: a
' second embedded asset (english.vla, via VLA_IDE.IdeLoadVocab) needs
' the identical logic, and R7's "duplicate small helpers" discipline is
' scoped to the VLAX_ injection boundary, not to every module pair -
' this one has no such boundary between it and VLA_IDE.bas. Returns ""
' if the sheet doesn't exist; callers decide whether that's fatal.
' PF.7: used to build its result with `s = s & CStr(cell.Value) & vbCrLf`
' inside a `For i = 1 To n` loop reading one cell at a time - O(n^2) in
' row count from the repeated concatenation (each `&` reallocates the
' whole growing string), plus n separate COM round-trips
' (`sh.Cells(i, 1)`) stacked on top. Confirmed hot for the two
' multi-hundred/multi-thousand-row callers (VLAp_Source/prelude.vla,
' VLAe_Source/english.vla via IdeLoadVocab, itself inside DoCheck's own
' Timer-measured "vocabulary" leg) - unmeasurable from the dev
' workbook/self-test suite as it stood, since both readers always find
' an external file first there and this loop body never ran; only a
' built .xlam's embedded sheets ever reached it. Fixed: one bulk
' Range.Value read (one COM round-trip regardless of n) flattened into
' a 1-D array, then one Join (genuinely O(n) - unlike a Collection's
' own positional Item(k), a materialized array is O(1) indexed, the
' distinction P-TOK's own mid-session lesson turned on). `Join(...)`
' alone drops the trailing terminator TestEmbeddedTextFallback pins as
' contractual (every row, including the last, ends in vbCrLf) - the
' explicit `& vbCrLf` after Join restores it.
' n=1 is handled separately because `Range.Value` on a single-cell
' range returns a bare scalar, not a 2-D array - not a hypothetical
' edge case: IdeVocabFileName's own VLAe_Name sheet is exactly one row
' (a bare filename) and hits this on every Check in a built edition.
' A bulk read buys that call site nothing (O(1) either way); it still
' needs the branch to stay correct.
Public Function VlaEmbeddedText(ByVal sheetName As String) As String
    On Error Resume Next
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If sh Is Nothing Then Exit Function
    Dim n As Long
    n = sh.Cells(sh.Rows.Count, 1).End(xlUp).Row
    If n = 1 Then
        VlaEmbeddedText = CStr(sh.Cells(1, 1).Value) & vbCrLf
        Exit Function
    End If
    Dim vals As Variant
    vals = sh.Range(sh.Cells(1, 1), sh.Cells(n, 1)).Value
    Dim lines() As String
    ReDim lines(1 To n)
    Dim i As Long
    For i = 1 To n
        lines(i) = CStr(vals(i, 1))
    Next
    VlaEmbeddedText = Join(lines, vbCrLf) & vbCrLf
End Function

' =====================================================================
'  Reader: tokenizer + parser
' =====================================================================

Private Function Tokenize(ByVal s As String) As String()
    Dim outc As New Collection
    Dim outLn As New Collection      ' C2: raw line per token, parallel
    Dim i As Long, n As Long
    Dim c As String, d As String, buf As String
    Dim start As Long
    Dim ln As Long, startLn As Long
    ln = 1
    n = Len(s)
    i = 1
    Do While i <= n
        c = Mid$(s, i, 1)
        Select Case c
            Case vbLf
                ln = ln + 1
                i = i + 1
            Case " ", vbTab, vbCr
                i = i + 1
            Case "("
                outc.Add "(": outLn.Add ln
                i = i + 1
            Case ")"
                outc.Add ")": outLn.Add ln
                i = i + 1
            Case ";"          ' comment runs to end of line (the line
                              ' break itself is left for the arm above,
                              ' which counts it)
                Do While i <= n
                    c = Mid$(s, i, 1)
                    If c = vbCr Or c = vbLf Then Exit Do
                    i = i + 1
                Loop
            Case """"         ' string literal; \" and \\ are escapes
                startLn = ln
                i = i + 1
                buf = ""
                Do While i <= n
                    c = Mid$(s, i, 1)
                    If c = "\" Then
                        d = Mid$(s, i + 1, 1)
                        If d = """" Then
                            buf = buf & """"
                            i = i + 2
                        ElseIf d = "\" Then
                            buf = buf & "\"
                            i = i + 2
                        Else
                            buf = buf & c
                            i = i + 1
                        End If
                    ElseIf c = """" Then
                        i = i + 1
                        Exit Do
                    Else
                        If c = vbLf Then ln = ln + 1
                        buf = buf & c
                        i = i + 1
                    End If
                Loop
                outc.Add Chr$(34) & buf   ' leading quote char marks "string literal"
                outLn.Add startLn
            Case Else         ' symbol or number
                start = i
                Do While i <= n
                    c = Mid$(s, i, 1)
                    If c = "(" Or c = ")" Or c = " " Or c = vbTab _
                       Or c = vbCr Or c = vbLf Or c = ";" Or c = """" Then Exit Do
                    i = i + 1
                Loop
                outc.Add Mid$(s, start, i - start)
                outLn.Add ln
        End Select
    Loop
    ' P-TOK: materialize once into arrays for O(1) positional access -
    ' the scan above is untouched, still building via cheap sequential
    ' Collection.Add calls; only the destination changes.
    ' TER-8: this comment used to say an empty program was ReDim'd "to
    ' size 0". It was not: ReDim (1 To 0) is refused by VBA with error 9,
    ' so text with no token raised before a form was read. Zero tokens now
    ' give (0 To 0) - allocated, UBound 0 - the empty list every reader
    ' expects (TER-7 made the same repair to EnTokenize).
    ' MUST be For Each, not "For k = 1 To .Count: x = .Item(k)" - a
    ' Collection's own Item(k) is itself an O(k) positional walk, so an
    ' indexed copy loop is O(n^2), the exact cost this item exists to
    ' remove (caught live: it relocated parse's own slowdown into
    ' Tokenize instead of eliminating it). For Each uses the
    ' Collection's real enumerator - true O(n) for the whole copy.
    Dim toksArr() As String
    If outc.Count = 0 Then
        ReDim toksArr(0 To 0)
        ReDim mTokLines(0 To 0)
    Else
        ReDim toksArr(1 To outc.Count)
        ReDim mTokLines(1 To outLn.Count)
    End If
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
    Tokenize = toksArr
End Function

' C2 source-map helpers ------------------------------------------------

Private Function CountLf(ByVal s As String) As Long
    Dim i As Long, k As Long
    i = 1
    Do
        i = InStr(i, s, vbLf)
        If i = 0 Then Exit Do
        k = k + 1
        i = i + 1
    Loop
    CountLf = k
End Function

Private Function TokRawLine(ByVal pos As Long) As Long
    On Error GoTo notReady   ' mTokLines never allocated - no Tokenize
                             ' has run yet in this frame (mirrors the
                             ' old "Is Nothing" guard for a Collection)
    If pos < 1 Or pos > UBound(mTokLines) Then Exit Function
    TokRawLine = mTokLines(pos)
    Exit Function
notReady:
End Function

' F.6: the shared core of every "which line, in which file" tag -
' "vla line N" or "lib.vla line N" - with no wrapping phrase of its
' own, so each of the reporting surfaces below (and the emitfail
' handler in VlaTranspile) can wrap it in their own wording without
' re-deriving the IncLookup branch. ok = False means rawLn maps to
' prelude territory (no line to report at all).
Private Function LabeledLine(ByVal rawLn As Long, ByRef ok As Boolean) As String
    Dim lc As Long
    lc = rawLn - mLineOffset
    ok = (lc >= 1)
    If Not ok Then Exit Function
    Dim lb As String, fl As Long, an As Long
    If IncLookup(lc, lb, fl, an) Then
        ' Mapped: fl is the line WITHIN its own file - for included
        ' lines that file is named; for main-file lines after an
        ' include, fl is the main file's own numbering (the expanded
        ' line lc would be wrong there).
        If Len(lb) > 0 Then
            LabeledLine = lb & " line " & fl
        Else
            LabeledLine = "vla line " & fl
        End If
        Exit Function
    End If
    LabeledLine = "vla line " & lc
End Function

' " (vla line N)" for reader messages, or "" for prelude territory.
' P.L7: when the line falls inside an included file the tag names
' that file and its LOCAL line instead - "(lib.vla line 12)" - and
' the no-include fast path (mIncMap Is Nothing) keeps today's
' wording byte-for-byte.
Private Function SrcLineTag(ByVal rawLn As Long) As String
    Dim ok As Boolean, ll As String
    ll = LabeledLine(rawLn, ok)
    If Not ok Then Exit Function
    SrcLineTag = " (" & ll & ")"
End Function

' PL7.1: the unbalanced-parens raise formats its own line phrase
' ("for the list opened at ...") inline - a FOURTH reporting
' surface the P.L7 design's "three chokepoints" claim missed; the
' owner's failing-pin paste named it. F.6 routes it through the same
' LabeledLine core as SrcLineTag now, closing that gap for real:
' included files get their name and local line, main-file lines keep
' the existing phrase with the main file's own numbering, and the
' no-include fast path keeps today's wording byte-for-byte.
Private Function OpenedAtTag(ByVal rawLn As Long) As String
    Dim ok As Boolean, ll As String
    ll = LabeledLine(rawLn, ok)
    If Not ok Then Exit Function
    OpenedAtTag = " for the list opened at " & ll
End Function

' =====================================================================
'  P.L7 (P-INCLUDE): (include "lib.vla") - a reader-level textual
'  splice. The line must stand ALONE (leading whitespace allowed):
'  the splicer is line-based and runs BEFORE the tokenizer, so the
'  file name between its quotes is raw text - no escapes, and a
'  name containing a quote is simply not an include line (it then
'  meets the emitter's worded refusal like any other misplaced
'  include form). Relative names resolve against the active
'  workbook's folder; absolute paths pass through. Includes nest;
'  depth caps at 16 with a cycle-suspicion message. Stated limit:
'  a line INSIDE a multi-line string literal that spells an
'  include line will be spliced (the splicer cannot know it is
'  inside a string) - do not spell includes inside multi-line
'  strings.
' =====================================================================
Private Function SpliceIncludes(ByVal src As String) As String
    If InStr(1, src, "(include", vbTextCompare) = 0 Then
        Set mIncMap = Nothing        ' fast path: no includes, no map
        SpliceIncludes = src
        Exit Function
    End If
    Set mIncMap = New Collection
    Dim sb As String, sbU As Long
    SpliceWalk src, "", 0, 0, sb, sbU
    SpliceIncludes = SbText(sb, sbU)
End Function

' label "" = the main program (each line anchors itself); an
' included block inherits the ANCHOR of the include statement that
' pulled it in, however deep the nesting.
Private Sub SpliceWalk(ByVal text As String, ByVal label As String, ByVal depth As Long, _
                       ByVal anchor As Long, ByRef sb As String, ByRef sbU As Long)
    If depth > 16 Then VLA_Messages.RaiseMsg "vla-include-too-deep", "label", label
    Dim lines() As String
    lines = Split(Replace(text, vbCrLf, vbLf), vbLf)
    Dim i As Long
    Dim fn As String
    Dim ownAnchor As Long
    For i = 0 To UBound(lines)
        ownAnchor = anchor
        If ownAnchor = 0 Then ownAnchor = i + 1
        If IsIncludeLine(lines(i), fn) Then
            SpliceWalk ReadIncludeFile(fn), fn, depth + 1, ownAnchor, sb, sbU
        Else
            ' The map count IS the number of lines already emitted,
            ' so break-before keeps text and map aligned by
            ' construction, recursion included.
            If mIncMap.Count > 0 Then SbAdd sb, sbU, vbCrLf
            SbAdd sb, sbU, lines(i)
            mIncMap.Add label & "|" & (i + 1) & "|" & ownAnchor
        End If
    Next
End Sub

Private Function IsIncludeLine(ByVal lineText As String, ByRef fileName As String) As Boolean
    Dim t As String
    t = Trim$(lineText)
    If VLA_Identity.Fold(Left$(t, 9)) <> "(include " Then Exit Function
    If Right$(t, 1) <> ")" Then Exit Function
    Dim inner As String
    inner = Trim$(Mid$(t, 10, Len(t) - 10))
    If Len(inner) < 2 Then Exit Function
    If Left$(inner, 1) <> """" Or Right$(inner, 1) <> """" Then Exit Function
    fileName = Mid$(inner, 2, Len(inner) - 2)
    If Len(fileName) = 0 Then Exit Function
    If InStr(1, fileName, """") > 0 Then Exit Function
    IsIncludeLine = True
End Function

Private Function ReadIncludeFile(ByVal name As String) As String
    Dim path As String
    path = name
    If InStr(1, path, ":") = 0 And Left$(path, 1) <> "\" And Left$(path, 1) <> "/" Then
        On Error Resume Next
        path = ActiveWorkbook.path & Application.PathSeparator & name
        On Error GoTo 0
    End If
    Dim fnum As Integer
    fnum = FreeFile
    On Error GoTo cantread
    Open path For Input As #fnum
    If LOF(fnum) > 0 Then ReadIncludeFile = Input$(LOF(fnum), #fnum)
    Close #fnum
    Exit Function
cantread:
    On Error Resume Next
    Close #fnum
    On Error GoTo 0
    VLA_Messages.RaiseMsg "vla-include-cannot-read", "name", name, "path", path
End Function

Private Function IncLookup(ByVal localLine As Long, ByRef label As String, _
                           ByRef fileLine As Long, ByRef anchor As Long) As Boolean
    If mIncMap Is Nothing Then Exit Function
    If localLine < 1 Or localLine > mIncMap.Count Then Exit Function
    Dim parts() As String
    parts = Split(mIncMap.Item(localLine), "|")
    label = parts(0)
    fileLine = CLng(parts(1))
    anchor = CLng(parts(2))
    IncLookup = True
End Function

' Key a parsed list to its user-source line. Object identity (ObjPtr)
' is the key; a freed list's address can be reused within one
' transpile, so last-write-wins rather than raising on a duplicate.
' PNTH.0: widened from Collection to Object - ObjPtr and the rest of
' this Sub's own body never assumed anything Collection-specific, so a
' VlaSlice (cdr/cddr's own result, VlaSlice.cls) tags exactly like any
' other freshly-built list value always has.
Private Sub TagLine(f As Object, ByVal srcLine As Long)
    If mFormLines Is Nothing Then Exit Sub
    ' S3.1 (first-run incident, caught by the determinism pin):
    ' "deliberately untagged" is now an EXPLICIT stored zero, never an
    ' absent key. ObjPtr keys are memory addresses, and the allocator
    ' recycles them - a list born at a dead tagged form's address, if
    ' left keyless, READ the dead form's line: a phantom map tag that
    ' came and went with heap history (transpiles of identical text
    ' differed on cold sessions and agreed on warm ones). Storing 0
    ' means every list this module creates OVERWRITES whatever stale
    ' entry its address inherited; FormLine then never reads the dead.
    If srcLine < 1 Then srcLine = 0
    Dim k As String
    k = CStr(ObjPtr(f))
    On Error Resume Next
    mFormLines.Remove k
    On Error GoTo 0
    mFormLines.Add CStr(srcLine), k
End Sub

Private Function FormLine(v As Variant) As Long
    If mFormLines Is Nothing Then Exit Function
    If Not IsList(v) Then Exit Function
    Dim found As Boolean
    Dim got As Variant
    got = CollGet(mFormLines, CStr(ObjPtr(v)), found)
    If found Then FormLine = CLng(got)
End Function

' Append the map comment to the FIRST emitted line of a statement -
' block statements tag their opener, and the comment is legal VBA
' after any of them (assignments, If ... Then, Do While, labels).
Private Function MapTag(ByVal emitted As String, ByVal srcLine As Long) As String
    ' V4: inside an (at-line N ...) wrapper the tag carries both
    ' coordinates - the vla line AND the originating-source line.
    ' P.L7: an included line's RUNTIME tag anchors to the main-file
    ' line of the include statement that pulled it in, so the tag
    ' format and the runtime map's main-file coherence are both
    ' untouched; transpile-time errors carry the precise per-file
    ' coordinates instead (SrcLineTag / emitfail).
    Dim tag As String
    ' LISTOPS-PROVENANCE: srcLine <= 0 reaches here ONLY when mGenRow is
    ' set (the caller's own gate skips this call otherwise, preserving
    ' S3.1's "macro-body statements carry no vla: tag at all" exactly) -
    ' so there is no real vla:N to print here; "vla:0" would be a lie,
    ' not a fallback, so the line-number half of the tag is omitted
    ' entirely rather than printed as a placeholder.
    If srcLine > 0 Then
        Dim lbT As String, flT As Long, anT As Long
        Dim useLine As Long
        useLine = srcLine
        If IncLookup(srcLine, lbT, flT, anT) Then
            If Len(lbT) > 0 Then useLine = anT
        End If
        tag = " ' vla:" & useLine
        If mAtLine > 0 Then tag = tag & " src:" & mAtLine
    Else
        tag = " '"
    End If
    ' gen-row's own label - a distinct, additional field, never
    ' combined into vla:N's own number (see (gen-row ...) in EmitStmt:
    ' this is a generator's OWN row provenance, not a line number, and
    ' the two happen to never coincide in practice since a macro's
    ' template body is always Substitute-copied, hence always
    ' srcLine<=0 whenever mGenRow is active - but neither branch above
    ' assumes that; either can carry the other if it ever changes).
    If Len(mGenRow) > 0 Then tag = tag & " vla-row:" & mGenRow
    Dim i As Long
    i = InStr(1, emitted, vbCrLf)
    If i = 0 Then
        MapTag = emitted & tag
    Else
        MapTag = Left$(emitted, i - 1) & tag & Mid$(emitted, i)
    End If
End Function

Private Function ParseAll(toks() As String) As Collection
    Dim pos As Long
    pos = 1
    Dim outc As New Collection
    Do While pos <= UBound(toks)
        outc.Add ParseForm(toks, pos, 1)
    Loop
    Set ParseAll = outc
End Function

' listDepth: the depth of the list a "(" here would open (VLA.1).
Private Function ParseForm(toks() As String, ByRef pos As Long, ByVal listDepth As Long) As Variant
    Dim t As String
    Dim lnRaw As Long
    lnRaw = TokRawLine(pos)              ' C2: this token's raw line
    t = toks(pos)
    pos = pos + 1
    If t = "(" Then
        If listDepth > VLA_FORM_DEPTH_LIMIT Then VLA_Messages.RaiseMsg "vla-form-too-deep", "limit", CStr(VLA_FORM_DEPTH_LIMIT), "loc", SrcLineTag(lnRaw)
        Dim lst As New Collection
        TagLine lst, lnRaw - mLineOffset ' C2: a list is born on the
                                         ' line of its opening paren
        Do While pos <= UBound(toks)
            If toks(pos) = ")" Then
                pos = pos + 1
                Set ParseForm = lst
                Exit Function
            End If
            lst.Add ParseForm(toks, pos, listDepth + 1)
        Loop
        VLA_Messages.RaiseMsg "vla-unbalanced-parens", "loc", OpenedAtTag(lnRaw)
    ElseIf t = ")" Then
        VLA_Messages.RaiseMsg "vla-unexpected-close-paren", "loc", SrcLineTag(lnRaw)
    Else
        ParseForm = t
    End If
End Function

' =====================================================================
'  Datum helpers
' =====================================================================

Private Function IsList(v As Variant) As Boolean
    IsList = IsObject(v)
End Function

Private Function IsSym(v As Variant) As Boolean
    If IsObject(v) Then Exit Function
    Dim s As String
    s = CStr(v)
    If Len(s) = 0 Then Exit Function
    IsSym = (Left$(s, 1) <> Chr$(34))
End Function

Private Function IsKeywordArg(v As Variant) As Boolean
    If Not IsSym(v) Then Exit Function
    Dim s As String
    s = CStr(v)
    IsKeywordArg = (Left$(s, 1) = ":" And Len(s) > 1)
End Function

' Fetch element i of a list, handling the object/value Set-vs-Let split.
' PNTH.0: widened from Collection to Object - safe and non-breaking for
' every existing caller (a Collection passed where Object is expected
' always works), and it is what lets this same function keep serving
' the handful of listops call sites where lst may now hold a VlaSlice
' (cdr/cddr's own result) instead of a Collection, with no separate
' Nth-for-slices twin to keep in sync.
Private Function Nth(lst As Object, ByVal i As Long) As Variant
    If i > lst.Count Then VLA_Messages.RaiseMsg "vla-form-missing-element", "n", i
    If IsObject(lst.Item(i)) Then
        Set Nth = lst.Item(i)
    Else
        Nth = lst.Item(i)
    End If
End Function

' PNTH.0: lst (the container being indexed) widened Collection ->
' Object - EvalCondFrom's own lst (below) can reach here as a VlaSlice
' if cond's own dispatch was itself re-entered on a computed value.
' Return type deliberately untouched (still Collection): the ELEMENT
' this asserts and returns list-shaped is always genuinely-parsed
' syntax at every call site (a cond clause, a macro signature, a
' function header...), never itself a listops-computed value.
Private Function NthList(lst As Object, ByVal i As Long) As Collection
    If i > lst.Count Then VLA_Messages.RaiseMsg "vla-form-missing-element", "n", i
    If Not IsObject(lst.Item(i)) Then
        VLA_Messages.RaiseMsg "vla-expected-list-at-position", "n", i, "value", CStr(lst.Item(i))
    End If
    Set NthList = lst.Item(i)
End Function

Private Function SymText(v As Variant) As String
    If IsObject(v) Then VLA_Messages.RaiseMsg "vla-expected-symbol-got-list"
    Dim s As String
    s = CStr(v)
    If Left$(s, 1) = Chr$(34) Then VLA_Messages.RaiseMsg "vla-expected-symbol-got-string"
    SymText = s
End Function

Private Function HeadSym(lst As Collection) As String
    If lst.Count = 0 Then VLA_Messages.RaiseMsg "vla-empty-form"
    HeadSym = SymText(Nth(lst, 1))
End Function

Private Function ListHeadIs(v As Variant, ByVal name As String) As Boolean
    If Not IsList(v) Then Exit Function
    ' PNTH.0: Object, not Collection - v may be a VlaSlice (a computed
    ' cdr/cddr result flowing back through, e.g., ResolveListopsArg's
    ' own quote-unwrap check on a nested listops call's own result).
    Dim lst As Object
    Set lst = v
    If lst.Count = 0 Then Exit Function
    If IsObject(lst.Item(1)) Then Exit Function
    Dim t As String
    t = CStr(lst.Item(1))
    If Left$(t, 1) = Chr$(34) Then Exit Function
    ListHeadIs = (VLA_Identity.Fold(t) = VLA_Identity.Fold(name))
End Function

Private Function StrLitContent(v As Variant) As String
    If IsObject(v) Then VLA_Messages.RaiseMsg "vla-expected-string-got-list"
    Dim s As String
    s = CStr(v)
    If Left$(s, 1) <> Chr$(34) Then VLA_Messages.RaiseMsg "vla-expected-string-got-other", "value", s
    StrLitContent = Mid$(s, 2)
End Function

' Assign a Variant that may hold an object or a value.
Private Sub AssignVar(ByRef target As Variant, ByVal v As Variant)
    If IsObject(v) Then
        Set target = v
    Else
        target = v
    End If
End Sub

Private Function CopyList(src As Collection) As Collection
    Dim outc As New Collection
    Dim e As Variant
    For Each e In src
        outc.Add e
    Next
    Set CopyList = outc
End Function

' Collection lookup by key without dying on a missing key.
' V8 audit note (the checklist's CollGet item): the trap INSTALL is
' near-free; the cost is the THROW, paid only on a missing key. The
' hot miss-dominant paths are Substitute (template symbols outside
' the bindings - most symbols of most templates) and FormLine (forms
' without a recorded line - every macro-generated form). A throw-free
' lookup means Scripting.Dictionary, a structural swap this pass does
' not buy blind: the transpile dial after the string conversion is
' the evidence that decides whether it is ever worth buying. Verdict
' recorded on the Alpha 5 ledger; the code stands.
Private Function CollGet(col As Collection, ByVal key As String, ByRef found As Boolean) As Variant
    On Error GoTo missing
    If IsObject(col.Item(key)) Then
        Set CollGet = col.Item(key)
    Else
        CollGet = col.Item(key)
    End If
    found = True
    Exit Function
missing:
    found = False
End Function

' =====================================================================
'  Macro system: (defmacro (name p1 p2 & rest) template...)
'
'  Template-substitution macros: parameters are replaced by the forms
'  passed at the call site; the rest parameter splices its forms into
'  whatever list it appears in. Multiple template forms are wrapped in
'  an implicit (begin ...). No gensym, no compile-time computation -
'  by design (see README).
' =====================================================================

Private Sub DefineMacro(f As Collection)
    Dim sig As Collection
    Set sig = NthList(f, 2)
    Dim mname As String
    mname = HeadSym(sig)
    ' P.L5, QUASIQUOTE, LISTOPS, VLA.3, VLA.4, VLA.5, VLA.6: the walks
    ' recognize every expand-time name as a form of their own - quote's data
    ' skip, the shielding bundle, the primitives, formula-expand, the loop,
    ' define and the reference's two parts -
    ' so a macro defined under one could never fire; refuse the name
    ' loudly instead of shadowing it silently, saying what it already is.
    ' The names and their words are one table, ExpandReservedDesc, which a
    ' loop's names are refused by too (VLA.5).
    Dim reservedDesc As String
    reservedDesc = ExpandReservedDesc(VLA_Identity.Fold(mname))
    If Len(reservedDesc) > 0 Then VLA_Messages.RaiseMsg "vla-defmacro-reserved-name", "word", VLA_Identity.Fold(mname), "desc", reservedDesc

    Dim params As New Collection
    Dim restName As String
    restName = ""
    Dim j As Long
    j = 2
    Do While j <= sig.Count
        Dim t As String
        t = SymText(Nth(sig, j))
        If t = "&" Then
            If j + 1 > sig.Count Then VLA_Messages.RaiseMsg "vla-defmacro-rest-param-missing", "name", mname
            restName = SymText(Nth(sig, j + 1))
            If j + 1 <> sig.Count Then VLA_Messages.RaiseMsg "vla-defmacro-rest-param-not-last", "name", mname
            j = sig.Count + 1
        Else
            params.Add t
            j = j + 1
        End If
    Loop

    ' L17: an optional docstring. A string literal at form 3 is
    ' documentation ONLY when at least one more form follows - a lone
    ' string IS the template (CL's own rule, so string-valued macros
    ' stay expressible). One docstring at most; a second string after
    ' it belongs to the template and the engine judges it as it
    ' always has. The doc is never part of the template, so expansion
    ' ignores it by construction. Escapes stay as written (the
    ' stored literal's content, verbatim) - L11's (doc ...) can
    ' prettify when it arrives; this pass only carries the text.
    Dim mdoc As String
    Dim docIdx As Long
    mdoc = ""
    docIdx = 3
    If f.Count >= 4 Then
        Dim d3 As Variant
        AssignVar d3, Nth(f, 3)
        If Not IsList(d3) Then
            If Left$(CStr(d3), 1) = Chr$(34) Then
                mdoc = StrLitContent(d3)
                docIdx = 4
            End If
        End If
    End If

    Dim tmpl As Variant
    If f.Count = docIdx Then
        AssignVar tmpl, Nth(f, docIdx)
    ElseIf f.Count > docIdx Then
        Dim b As New Collection
        b.Add "begin"
        For j = docIdx To f.Count
            b.Add Nth(f, j)
        Next
        Set tmpl = b
    Else
        VLA_Messages.RaiseMsg "vla-defmacro-missing-template", "name", mname
    End If

    Dim rec As New Collection
    rec.Add params
    rec.Add restName
    rec.Add tmpl
    rec.Add mdoc                   ' L17: item 4, "" when undocumented
    rec.Add mname                  ' L11: item 5, the name as written -
                                   ' Collections cannot enumerate their
                                   ' keys, so apropos reads it here (rec
                                   ' is still a Collection - only the
                                   ' outer mMacros table itself, PDICT.0,
                                   ' became a Dictionary)
    ' PDICT.0: direct overwrite, not Remove-then-Add - a Dictionary's own
    ' keyed Item-assignment already replaces an existing entry (or inserts
    ' a new one) with no error either way, so "allow redefinition" needs
    ' no separate step. This is a deliberate choice over an .Exists guard:
    ' redefinition is the RARE case in practice (a growing phrasebook is
    ' overwhelmingly new names), and the old Remove-then-Add shape paid a
    ' real VBA exception-throw cost on every ordinary FIRST-time
    ' registration (Remove always missed), not just actual redefinitions -
    ' worse the larger a corpus gets, exactly the direction phrasebooks
    ' and corpora are headed, not a one-time saving.
    Set mMacros.Item(VLA_Identity.Fold(mname)) = rec
End Sub

' VLA.5: the expand-time names, each with what it is, "" for any other
' name: the one table DefineMacro refuses a macro's name by and a loop
' refuses a name it binds by, vla-lang's RESERVED in its order. The walks
' recognize each name as a form of their own (P.L5's quote, QUASIQUOTE's
' bundle, the twenty-eight primitives, VLA.4's formula-expand, VLA.5's
' loop, VLA.6's define and the reference's two parts), so a macro of the
' name could never fire, and a loop's name so spelled would stand where the
' walks read the form. A define's word is refused by it too (VLA.6).
Private Function ExpandReservedDesc(ByVal wordFold As String) As String
    Select Case wordFold
        Case "quote"
            ExpandReservedDesc = "data-literal form"
        Case "quasiquote"
            ExpandReservedDesc = "template-shielding form"
        Case "unquote"
            ExpandReservedDesc = "quasiquote escape form"
        Case "unquote-splicing"
            ExpandReservedDesc = "quasiquote splice form"
        Case "symbol"
            ExpandReservedDesc = "identifier-fusion primitive"
        Case "car"
            ExpandReservedDesc = "list-head primitive"
        Case "cdr"
            ExpandReservedDesc = "list-tail primitive"
        Case "cddr"
            ExpandReservedDesc = "list-tail primitive"
        Case "cons"
            ExpandReservedDesc = "list-construction primitive"
        Case "list"
            ExpandReservedDesc = "list-construction primitive"
        Case "null?"
            ExpandReservedDesc = "empty-list primitive"
        Case "eq?"
            ExpandReservedDesc = "atom-comparison primitive"
        Case "equal?"
            ExpandReservedDesc = "structural-comparison primitive"
        Case "quote-if"
            ExpandReservedDesc = "expand-time conditional"
        Case "cond"
            ExpandReservedDesc = "multi-clause conditional (folds at expand time when possible, defers to runtime otherwise)"
        Case "+expand"
            ExpandReservedDesc = "expand-time addition primitive"
        Case "=expand"
            ExpandReservedDesc = "expand-time equality primitive"
        Case "<>expand"
            ExpandReservedDesc = "expand-time inequality primitive"
        Case ">expand"
            ExpandReservedDesc = "expand-time comparison primitive"
        Case "<expand"
            ExpandReservedDesc = "expand-time comparison primitive"
        Case ">=expand"
            ExpandReservedDesc = "expand-time comparison primitive"
        Case "<=expand"
            ExpandReservedDesc = "expand-time comparison primitive"
        Case "-expand"
            ExpandReservedDesc = "expand-time subtraction primitive"
        Case "*expand"
            ExpandReservedDesc = "expand-time multiplication primitive"
        Case "quotient-expand"
            ExpandReservedDesc = "expand-time integer quotient primitive"
        Case "mod-expand"
            ExpandReservedDesc = "expand-time remainder primitive"
        Case "&expand"
            ExpandReservedDesc = "expand-time text-joining primitive"
        Case "characters-expand"
            ExpandReservedDesc = "expand-time text-splitting primitive"
        Case "address-expand"
            ExpandReservedDesc = "expand-time address primitive"
        Case "row-expand"
            ExpandReservedDesc = "expand-time row-number primitive"
        Case "column-expand"
            ExpandReservedDesc = "expand-time column-number primitive"
        Case "column-letters-expand"
            ExpandReservedDesc = "expand-time column-letters primitive"
        Case "formula-expand"
            ExpandReservedDesc = "expand-time formula-syntax operation"
        Case "assoc-expand"
            ExpandReservedDesc = "expand-time entry-lookup primitive"
        Case "each"
            ExpandReservedDesc = "expand-time loop form"
        Case "each-index"
            ExpandReservedDesc = "expand-time counting loop form"
        Case "define"
            ExpandReservedDesc = "expand-time naming form"
        Case "sheet-expand"
            ExpandReservedDesc = "expand-time reference-sheet operation"
        Case "range-expand"
            ExpandReservedDesc = "expand-time reference-range operation"
    End Select
End Function

' L17: read a macro's docstring - "" for an undocumented or unknown
' name. LIFETIME, stated plainly: the macro table rebuilds on every
' parse (VlaTranspile / VlaExpandText both reset it), so this answers
' as of the LAST parse - the honest query flow is parse-then-ask,
' which is exactly the seam L11's (doc ...) will ride. Vocabulary-
' carried macros are visible here after any load or translation,
' because their probe-transpile and every subsequent parse carry
' them through this same table.
Public Function VlaMacroDoc(ByVal name As String) As String
    Dim rec As Collection
    Set rec = GetMacro(name)
    If rec Is Nothing Then Exit Function
    If rec.Count >= 4 Then VlaMacroDoc = CStr(rec.Item(4))
End Function

' P.L6: the procedure-docstring table's writer. Every emitted sub
' and function records here (name AS WRITTEN, doc or "", kind) -
' last-write-wins on a duplicate name, since VBA's own compile step
' is the loud judge of the duplication itself.
Private Sub RecordSubDoc(ByVal asWritten As String, ByVal doc As String, ByVal kind As String)
    ' L14: kind widened from a Boolean to the kind string itself -
    ' "sub" / "function" / "lambda" - so deflambda registers on the
    ' same table (the bench consuming P.L6, as the session order
    ' promised). Apropos renders "lambda" as "(worksheet function)".
    If mSubDocs Is Nothing Then Set mSubDocs = New Collection
    Dim k As String
    k = VLA_Identity.Fold(asWritten)
    On Error Resume Next
    mSubDocs.Remove k
    On Error GoTo 0
    Dim rec As New Collection
    rec.Add asWritten
    rec.Add doc
    rec.Add kind
    mSubDocs.Add rec, k
End Sub

' P.L6: read a procedure's docstring - "" for an undocumented or
' unknown name, case-insensitive like the table. LIFETIME, stated
' plainly (L17's flow, one table over): the table is reset by
' VlaTranspile and populated at emission, so it reflects the LAST
' program transpiled - after a Check that is the user's program
' (the translate-once transpile runs last); after a vocabulary
' load it is empty (the macro probes transpile defmacro text, no
' procedures); after any Run or scratch the S5.1 wipe empties it
' with everything else. Parse-then-ask.
Public Function VlaSubDoc(ByVal name As String) As String
    ' Object items must be fetched with Set (the brief's Set-vs-Let
    ' hazard; CollGet Let-copies and is for values) - GetMacro's own
    ' pattern, mirrored.
    If mSubDocs Is Nothing Then Exit Function
    Dim rec As Collection
    On Error Resume Next
    Set rec = mSubDocs.Item(VLA_Identity.Fold(name))
    On Error GoTo 0
    If rec Is Nothing Then Exit Function
    VlaSubDoc = CStr(rec.Item(2))
End Function

' IN.7 (button-click half): does a proc by this AS-WRITTEN name exist
' in what THIS transpile has emitted so far? mSubDocs's own existence
' (not its doc text, which VlaSubDoc above already reports "" for
' either an empty doc or an unknown name - indistinguishable by
' design there) is what EmitStmt's "make-button" case needs, to leave
' a button's OnAction unset rather than reference a Sub that was
' never declared (a decorative button with no handler is legal under
' the interpreter too - VlaDispatchButtonClick's own "unregistered =
' no-op" precedent). Reachable without an ordering requirement
' because VLA_English.bas always appends "(sub main ...)" LAST, after
' every handler declaration, regardless of where the user wrote the
' two sentences relative to each other - so by the time EmitTop
' reaches a make-button call inside sub main, any matching handler
' is already recorded here. (A make-button call inside a named
' "To ...:" action declared textually BEFORE its own handler is the
' one case this can still miss - a real, narrow limitation, not
' solved here.)
Private Function SubExists(ByVal asWrittenName As String) As Boolean
    If mSubDocs Is Nothing Then Exit Function
    Dim rec As Collection
    On Error Resume Next
    Set rec = mSubDocs.Item(VLA_Identity.Fold(asWrittenName))
    On Error GoTo 0
    SubExists = Not (rec Is Nothing)
End Function

' L11: (doc <macro-name>) - a statement form resolved at TRANSPILE
' time (the emitter's privilege, like at-line): the doc is looked up
' in this parse's macro table and baked into a Debug.Print, so the
' running program simply speaks it. Parse-scoped by the same honesty
' as VlaMacroDoc: a raw scratch (VlaTry) carries prelude + its own
' macros, not the vocabulary's - and the unknown-name message says
' exactly that instead of pretending.
Private Function EmitDoc(lst As Collection, ByVal pad As String) As String
    If lst.Count < 2 Then VLA_Messages.RaiseMsg "vla-doc-arity"
    Dim nm As String
    nm = SymText(Nth(lst, 2))
    Dim rec As Collection
    Set rec = GetMacro(nm)
    Dim msg As String
    If rec Is Nothing Then
        msg = nm & ": (not a known macro in this parse - vocabulary macros ride loads and translations, not raw scratches)"
    ElseIf Len(CStr(rec.Item(4))) = 0 Then
        msg = nm & ": (no documentation)"
    Else
        msg = nm & ": " & CStr(rec.Item(4))
    End If
    EmitDoc = pad & "Debug.Print " & Chr$(34) & Replace(msg, Chr$(34), Chr$(34) & Chr$(34)) & Chr$(34) & vbCrLf
End Function

' L11: apropos - name-match over the macro table (names AND
' docstrings) plus the runtime helper manifest, per the definition.
' Definition order, macros first: the table walks as defined
' (prelude, then whatever the last parse carried), then helpers as
' the manifest lists them. Parse-then-ask lifetime as ever - and a
' vocabulary load ENDS with every carried macro parsed, so
' load-then-ask sees the whole bench. The manifest read is fresh
' per call (a human asks once; no cache) and SOFT: where it cannot
' be read (no VBE trust), apropos still answers macros - the
' resolve check's own soft-pass doctrine. An empty pattern lists
' everything.
' L11.1: English pushes the carried vocabulary down at load and
' clears it at reset - the correct dependency direction (English
' already depends on core; core never reads English).
' APROPOSPLUS: rulesSrc is the optional second half of that same push -
' one "pattern -> template" line per loaded rule (VLA_English's own
' AproposRulesBlob), so apropos can answer with the ENGLISH SENTENCE
' that reaches a macro, not just the macro itself. Optional and
' defaulted so this stays source-compatible with any caller that only
' ever pushed macro text.
Public Sub VlaAproposCarry(ByVal src As String, Optional ByVal rulesSrc As String = "")
    mAproposCarry = src
    mAproposRulesCarry = rulesSrc
End Sub

Public Function VlaAproposText(Optional ByVal pattern As String = "") As String
    ' L11.1: state-independence. The maiden smoke taught that
    ' parse-then-ask makes a human-facing tool feel broken - the
    ' answer depended on which machinery parsed last. So apropos
    ' re-parses the bench itself on every ask: prelude plus the
    ' carried vocabulary. A human asks at human speed; the parse is
    ' cheap. ((doc ...) stays parse-scoped on purpose - it is an
    ' emitter form, and its teaching message says so.)
    Dim refreshFired As Long
    Dim refreshText As String
    refreshText = VlaExpandText(mAproposCarry, True, refreshFired)
    Dim r As String
    Dim rec As Variant
    Dim nm As String
    Dim dc As String
    If Not mMacros Is Nothing Then
        ' PDICT.0: Dictionary's own For Each walks KEYS, not values (a
        ' Collection's own walked values) - .Items returns the values
        ' array directly, the closer match to the old behavior.
        For Each rec In mMacros.Items
            If rec.Count >= 5 Then
                nm = CStr(rec.Item(5))
                dc = CStr(rec.Item(4))
                If Len(pattern) = 0 Or InStr(1, nm, pattern, vbTextCompare) > 0 _
                   Or InStr(1, dc, pattern, vbTextCompare) > 0 Then
                    r = r & nm & " - " & IIf(Len(dc) > 0, dc, "(no documentation)") & vbCrLf
                End If
            End If
        Next
    End If
    ' APROPOSPLUS: the English-sentence tier - one line per loaded rule,
    ' "pattern -> template", carried down by VLA_English.bas the same
    ' call that carries macro text above (VlaAproposCarry's own second
    ' parameter). Matches against the WHOLE line, so a search hits
    ' whether the word lives in the sentence ("pivot") or in what it
    ' calls ("pivot-create") - the point of this tier is exactly that a
    ' search for a macro name should also surface the English sentence
    ' that reaches it, and vice versa. Already-resolved text (built via
    ' VLA.VlaWriteForm at carry time, in VLA_English.bas), so unlike
    ' the macro tier above this needs no per-ask re-expansion - a wipe
    ' cannot desync it from mMacros because it never depended on
    ' mMacros to begin with.
    If Len(mAproposRulesCarry) > 0 Then
        Dim ruleLines() As String
        ruleLines = Split(mAproposRulesCarry, vbCrLf)
        Dim li As Long
        For li = LBound(ruleLines) To UBound(ruleLines)
            Dim oneLine As String
            oneLine = ruleLines(li)
            If Len(Trim$(oneLine)) > 0 Then
                If Len(pattern) = 0 Or InStr(1, oneLine, pattern, vbTextCompare) > 0 Then
                    r = r & oneLine & "   (sentence)" & vbCrLf
                End If
            End If
        Next
    End If

    ' P.L6: the program-procedure tier - the last transpiled
    ' program's subs and functions, name-and-doc matched like the
    ' macros above. This lists AFTER a Check (the translate-once
    ' transpile runs last, so the table holds the user's program)
    ' and empties on any wipe with everything else; the prelude
    ' refresh above deliberately cannot touch it (VlaExpandText
    ' never resets the table - the declaration says why).
    If Not mSubDocs Is Nothing Then
        Dim srec As Variant
        For Each srec In mSubDocs
            nm = CStr(srec.Item(1))
            dc = CStr(srec.Item(2))
            If Len(pattern) = 0 Or InStr(1, nm, pattern, vbTextCompare) > 0 _
               Or InStr(1, dc, pattern, vbTextCompare) > 0 Then
                If CStr(srec.Item(3)) = "lambda" Then
                    r = r & nm & " - " & IIf(Len(dc) > 0, dc, "(no documentation)") & " (worksheet function)" & vbCrLf
                Else
                    r = r & nm & " - " & IIf(Len(dc) > 0, dc, "(no documentation)") & " (program " & CStr(srec.Item(3)) & ")" & vbCrLf
                End If
            End If
        Next
    End If
    Dim man As String
    On Error Resume Next
    man = VlaHelperManifest()
    On Error GoTo 0
    If Len(man) > 0 Then
        Dim names() As String
        Dim i As Long
        ' L11.2: the manifest is SPACE-joined - one line, every name.
        ' The owner's smoke paste showed it plainly and the pin caught
        ' the line-break assumption. Normalize all whitespace to one
        ' space and split on that; each helper then lists on its own
        ' line, which is also simply the better report.
        man = Replace(Replace(Replace(Replace(man, vbCrLf, " "), vbCr, " "), vbLf, " "), vbTab, " ")
        names = Split(man, " ")
        For i = LBound(names) To UBound(names)
            nm = Trim$(names(i))
            If Len(nm) > 0 Then
                If Len(pattern) = 0 Or InStr(1, nm, pattern, vbTextCompare) > 0 Then
                    r = r & nm & " - (runtime helper)" & vbCrLf
                End If
            End If
        Next
    End If
    If Len(r) = 0 Then
        r = "(nothing matches '" & pattern & "' in this parse)" & vbCrLf
        ' L11.3: distinguish "no match" from "nothing loaded" - after
        ' any Run or scratch injection, module memory is wiped (S5.1,
        ' the same fact that sent L6's history into workbook Names),
        ' so the carried vocabulary is gone until the next load. Say
        ' so instead of shrugging.
        If Len(mAproposCarry) = 0 Then
            r = r & "(no vocabulary is carried right now - a Run or scratch injection wipes module memory; Reload, then ask again)" & vbCrLf
        End If
    End If
    VlaAproposText = r
End Function

Public Sub VlaApropos(Optional ByVal pattern As String = "")
    Debug.Print VlaAproposText(pattern)
End Sub

' L12: the L1 writer as the CANONICAL FORMATTER. Tokenize, parse,
' pretty-print - no prelude, no expansion (formatting must never
' expand: what you wrote is what gets formatted; (when ...) stays
' (when ...)). Parse errors propagate with their hinted messages
' (C2's "missing ')' for the list opened at vla line N"), so a
' malformed source formats into a teaching refusal, not silence.
Public Function VlaFormat(ByVal source As String) As String
    mLineOffset = 0                  ' no prelude here: line 1 is the
                                     ' user's line 1, so the parser's
                                     ' "opened at vla line N" hints
                                     ' number the text being formatted
                                     ' (the offset is module state a
                                     ' prior transpile leaves behind)
    Dim toks() As String
    toks = Tokenize(source)
    Dim forms As Collection
    Set forms = ParseAll(toks)
    Dim sb As String, sbU As Long
    Dim f As Variant
    For Each f In forms
        SbAdd sb, sbU, WritePretty(f, 0)
        SbAdd sb, sbU, vbCrLf
    Next
    VlaFormat = SbText(sb, sbU)
End Function

' L12: the balance count as words - the L0 scanner's arithmetic
' surfaced as a hint. String-aware WITH the tokenizer's own escape
' rules (\" and \\ inside strings - a naive quote-toggle miscounts
' them) and ;-comment-aware. Returns "" for balanced text; otherwise
' the definition's own sentence shape. Pure and pinned.
Public Function VlaBalanceHint(ByVal t As String) As String
    Dim i As Long
    Dim c As String
    Dim depth As Long
    Dim inQuote As Boolean
    i = 1
    Do While i <= Len(t)
        c = Mid$(t, i, 1)
        If inQuote Then
            If c = "\" Then
                i = i + 2
            Else
                If c = """" Then inQuote = False
                i = i + 1
            End If
        ElseIf c = """" Then
            inQuote = True
            i = i + 1
        ElseIf c = ";" Then
            Do While i <= Len(t)
                c = Mid$(t, i, 1)
                If c = vbCr Or c = vbLf Then Exit Do
                i = i + 1
            Loop
        Else
            If c = "(" Then depth = depth + 1
            If c = ")" Then depth = depth - 1
            i = i + 1
        End If
    Loop
    If depth > 0 Then
        VlaBalanceHint = "this row opens " & depth & " form" & IIf(depth = 1, "", "s") & " it never closes"
    ElseIf depth < 0 Then
        VlaBalanceHint = "this row closes " & (-depth) & " form" & IIf(depth = -1, "", "s") & " it never opened"
    End If
End Function

' L15: the macro-stepper - step-N on the odometer. The text after
' exactly stepN applications, with totalApps reporting the fixpoint
' count so callers can say "application 3 of 7". Deterministic by
' construction: the expansion walk is a fixed order, so the first N
' applications of any two runs are the SAME N - a budget prefix is
' THE prefix. stepN 0 is the parse as written; stepN at or past the
' total is the fixpoint. The budget gate restores on every exit,
' errors included - a stepper must never leave the engine stingy.
Public Function VlaExpandStepText(ByVal source As String, ByVal stepN As Long, _
                                  ByRef totalApps As Long) As String
    Dim fired As Long
    Dim full As String
    full = VlaExpandText(source, True, fired)    ' learn the total
    totalApps = fired
    If stepN < 0 Or stepN >= totalApps Then
        VlaExpandStepText = full
        Exit Function
    End If
    mBudgetOn = True
    mExpandBudget = stepN
    On Error GoTo restoreBudget
    VlaExpandStepText = VlaExpandText(source, True, fired)
    mBudgetOn = False
    Exit Function
restoreBudget:
    mBudgetOn = False
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

' The narrator: no argument walks every application from "as written"
' to the fixpoint; a stepN prints that one frame. Teaching material
' by design - this printout IS the macro chapter of VLA_GUIDE.
Public Sub VlaExpandStep(ByVal source As String, Optional ByVal stepN As Long = -1)
    Dim total As Long
    Dim i As Long
    Dim t As String
    If stepN >= 0 Then
        t = VlaExpandStepText(source, stepN, total)
        Debug.Print "application " & stepN & " of " & total & ":"
        Debug.Print t
    Else
        t = VlaExpandStepText(source, 0, total)
        Debug.Print "step 0 of " & total & " (as written):"
        Debug.Print t
        For i = 1 To total
            t = VlaExpandStepText(source, i, total)
            Debug.Print "application " & i & " of " & total & ":"
            Debug.Print t
        Next
    End If
End Sub

' PDICT.0: was On Error Resume Next around a Collection.Item lookup - a
' Collection has no hash table for keyed access, so that Item call was a
' linear scan over every registered macro, the dominant cost P-PROF's own
' live "expand" number (71.3% of compile-side transpile time, scale-
' included suite) traced back to. mMacros is now a Dictionary; .Exists is
' throw-free and O(1)-hashed, so the miss case (the common one at the end
' of every macro chain) costs a lookup, not an exception.
Private Function GetMacro(ByVal name As String) As Collection
    Dim key As String
    key = VLA_Identity.Fold(name)
    If mMacros.Exists(key) Then Set GetMacro = mMacros.Item(key)
End Function

Private Function ExpandMacros(d As Variant, ByVal depth As Long, Optional ByVal walkDepth As Long = 1) As Variant
    ' LISTOPS-BUDGET / LISTOPS.0: LIVE FINDING, this pass's own "raised
    ' 200 -> 1000, a measured 5x, not reckless" claim was WRONG -
    ' disproven by a real host run, not a hypothetical. A raise to 1000
    ' turned a clean "too deep" error into a raw "Out of stack space"
    ' crash for the pre-existing 250-deep TestListopsBudget pin, which
    ' had passed at 200 before this session touched it: VBA's actual
    ' native stack exhausts somewhere between 150 (still safe) and 250
    ' (crashes), meaning 200 was ALREADY sitting at the real ceiling
    ' with essentially no headroom to raise into - the linear-scaling
    ' argument this comment's own prior revision made was right about
    ' the SHAPE of the risk and wrong that there was room to move it.
    ' Reverted to 200, its prior, already-validated value.
    '
    ' TABLESPEC-SCALE / TABLESPECSCALE.0: the actual fix, found once the
    ' owner asked directly whether tail-call optimization could sidestep
    ' this whole discussion. VBA has none, at any layer of this stack -
    ' but the chain below (macro expands to a list still headed by
    ' another macro name; substitute; check again; repeat) is a textbook
    ' tail call by hand: nothing happens to a chain link's own result
    ' except returning it. That made it hand-trampolineable - a Do loop
    ' reassigning ONE local (cur), not a recursive call.
    '
    ' FIRST VERSION of this fix was WRONG - caught by re-tracing before
    ' any live run, not by a failure. It only trampolined a macro whose
    ' template expands DIRECTLY to another macro call at its own head
    ' position (quote-map/reverse/append's shape - genuinely fixed by
    ' it). walk-rows and tablespec - the actual motivating case - don't
    ' do that: their own else-branch is `(begin <side content> (walk-
    ' rows (cdr lst)))` - the self-call sits ONE LEVEL INSIDE a `begin`,
    ' which is not itself a macro, so the loop exited after one
    ' substitution and fell through to the ORIGINAL, un-trampolined
    ' `rebuild children` step below - which is where walk-rows/
    ' tablespec's real stack cost was ALWAYS coming from, not the head-
    ' chain this comment originally (and wrongly) described. Worse than
    ' a no-op: `depth` no longer incremented for this shape either,
    ' silently disarming the safety net that used to turn this into a
    ' clean "too deep" error - a large table would have hit a raw,
    ' uncontrolled "Out of stack space" crash instead, with no warning.
    '
    ' THE FIX: `begin` is now a transparent TAIL-POSITION wrapper inside
    ' this loop too, the same way a real tail-call-optimizing evaluator
    ' treats `(begin e1 ... en)` - only en is in tail position. Reaching
    ' `(begin e1 ... en)` here stashes e1..e(n-1) (ordinary, bounded work
    ' - walk-rows's own debug-print call, tablespec-row's own call -
    ' never further chained) into pendingHead and loops again on en
    ' alone. Once the loop exits for a real reason, pendingHead's own
    ' stashed elements are combined with the final rebuild into one
    ' result. No new VBA stack frame is pushed per chain link either
    ' way, so chain length (a table walker's own row count, LISTOPS-
    ' BUDGET's original worry) is fully uncoupled from native stack risk
    ' - it was ONLY ever the chain shape that cost stack, never the
    ' shape of the data being built (LISTOPS-BUDGET's own "many
    ' independent, never-chained calls" finding already said as much;
    ' this closes the chained half, begin-wrapped or not). `depth`
    ' (untouched by this loop, still ByVal, still checked once per call
    ' exactly as before) keeps meaning exactly what it always meant -
    ' genuine VBA call-stack nesting from resolving pendingHead's own
    ' stashed elements and the final rebuild below, each independent and
    ' shallow (never nested in each other, matching the already-proven-
    ' safe "many independent siblings" shape) - which LISTOPS-CONFLUENCE
    ' already established is bounded by a template's own WRITTEN
    ' structure, not by data size. `chainLen` (new, local to this call,
    ' never threaded anywhere) is the chain's own separate counter -
    ' incremented for both a macro substitution AND a begin-unwrap, so
    ' either kind of non-terminating loop is caught - capped generously
    ' (5000 - twenty times "a 250-row pricing sheet is entirely
    ' ordinary," LISTOPS-BUDGET's own illustrative case) purely as a
    ' logical guard against a genuinely infinite self-expanding macro
    ' (one with no base case) - it no longer maps to any hardware risk,
    ' so unlike 200 there was real headroom to be generous here.
    '
    ' SECOND correction, also caught before any live run: treating EVERY
    ' `begin` as a transparent tail wrapper would have re-tagged an
    ' ORDINARY, hand-written top-level `(begin ...)` statement's real
    ' source line as 0 for no reason - only a macro's own SUBSTITUTED
    ' output (already 0-tagged, S3.1's own convention) should ever take
    ' this path. Gated on `FormLine(cur) = 0` below, so a hand-written
    ' begin falls straight through to the unchanged rebuild step, real
    ' line intact.
    If walkDepth > VLA_FORM_DEPTH_LIMIT Then RefuseFormTooDeep d   ' VLA.1
    If depth > 200 Then VLA_Messages.RaiseMsg "vla-macro-expansion-too-deep"
    Dim cur As Variant
    AssignVar cur, d
    Dim chainLen As Long
    chainLen = 0
    Dim pendingHead As New Collection
    ' PNTH.0: Object, not Collection - cur (below) can legitimately BE
    ' a VlaSlice mid-loop (a macro whose own body is a bare cdr/cddr
    ' call, however unusual a shape that is) - this only needs to
    ' re-check headFold against it, never .Add/For Each it directly.
    Dim lst As Object
    Dim headFold As String
    Dim lprimTop As Variant
    Dim rec As Collection
    Dim r As Variant
    Dim bi As Long
    Do
        If Not IsList(cur) Then
            ' VLA.6: a word a define binds, its form put where the word stands
            ' and looked at again, one link of the chain. A source with no
            ' define pays the one Is Nothing.
            If Not mDefines Is Nothing Then
                If DefineNamed(cur) Then
                    chainLen = chainLen + 1
                    If chainLen > 5000 Then VLA_Messages.RaiseMsg "vla-macro-self-expansion-limit"
                    AssignVar cur, mDefines.Item(VLA_Identity.Fold(CStr(cur)))
                    GoTo nextLink
                End If
            End If
            ' VLA.5: an atom with forms set aside before it is the last form
            ' of a begin a macro made, combined with them after the loop.
            If pendingHead.Count > 0 Then Exit Do
            ExpandMacros = cur
            Exit Function
        End If
        Set lst = cur
        If lst.Count = 0 Then Exit Do
        If Not IsSym(lst.Item(1)) Then Exit Do
        headFold = VLA_Identity.Fold(CStr(lst.Item(1)))
        ' P.L5: quote is a DATA context - expansion never enters.
        ' A quoted word list may legally contain macro names
        ' ((quote (when unless if)) is data, not code), so the
        ' walk returns the form untouched, children included.
        ' DefineMacro refuses "quote" as a macro name, so the
        ' head can never be a macro - the skip is exact, not
        ' heuristic. Deterministic and budget-neutral (L15):
        ' skipping fires nothing. VLA.6: every return of a form of its
        ' own, this one and those below, goes by finish, which keeps the
        ' forms set aside before it.
        If headFold = "quote" Then
            Set lprimTop = lst
            GoTo finish
        End If
        ' QUASIQUOTE: (symbol ...) is the one primitive of the
        ' bundle recognized standalone, not just inside a defmacro
        ' template's own Substitute walk - its arguments don't
        ' require the ambient bindings table Substitute threads
        ' (a bare unbound symbol already passes through as itself),
        ' so it resolves the same way here with empty tables.
        ' Recognized before the macro lookup below since "symbol"
        ' can never itself be a defined macro name (DefineMacro
        ' refuses it, same as quote).
        If headFold = "symbol" Then
            lprimTop = FuseSymbol(lst, New Collection, "", New Collection)
            GoTo finish
        End If
        ' VLA.4: (formula-expand <expression>), the formula syntax at
        ' expand time - an operation of this walk and of no other site.
        ' Substitute leaves it alone, so a template's parameters are
        ' filled in inside it; ExpandFormulaCall expands its argument to
        ' its fixpoint before writing it, so an author's macros inside the
        ' expression fire and a macro's call is never written as a
        ' function's. A helper, so this recursive procedure gains no
        ' local (trap 8).
        If headFold = "formula-expand" Then
            lprimTop = ExpandFormulaCall(lst, depth, walkDepth)
            GoTo finish
        End If
        ' VLA.6: a define, a form of the walk that writes nothing here: passed
        ' by as written, its word and its form never walked, for the form of
        ' the source it stands in to take once that has expanded
        ' (ExpandSourceForm); refused inside a formula's expression.
        If headFold = "define" Then
            Set lprimTop = DefinePassedBy(lst)
            GoTo finish
        End If
        ' VLA.6: a reference's sheet and its cell or range as texts, forms of
        ' the walk beside formula-expand, their argument expanded first so
        ' that a word a define binds is replaced (ExpandReferencePart).
        If headFold = "sheet-expand" Or headFold = "range-expand" Then
            lprimTop = ExpandReferencePart(headFold, lst, depth, walkDepth)
            GoTo finish
        End If
        ' QUASIQUOTE: unlike symbol, quasiquote/unquote/unquote-
        ' splicing DO require the ambient bindings table Substitute
        ' threads - unquote's entire job is looking a name up in
        ' it, and no such table exists outside one macro's own
        ' expansion. Refuse loudly with words here, at the general
        ' walk, rather than let these fall through unrecognized to
        ' a confusing emit-time error.
        Select Case headFold
            Case "quasiquote", "unquote", "unquote-splicing"
                VLA_Messages.RaiseMsg "vla-quasiquote-outside-template", "head", headFold
        End Select
        ' LISTOPS: the twenty-eight list/arithmetic/conditional primitives
        ' (the original nine, LISTOPS-EXPAND's own seven, cond, VLA.3's ten
        ' and VLA.5's assoc-expand) are recognized standalone too, same
        ' treatment symbol gets - none need Substitute's ambient bindings
        ' table intrinsically, since a bound-name reference inside an
        ' argument can't exist outside an enclosing macro in the first place
        ' (same reasoning FuseSymbol's own comment gives for symbol).
        Select Case headFold
            Case "car", "cdr", "cddr", "cons", "list", "null?", "eq?", "equal?", "quote-if", "cond", _
                 "+expand", "=expand", "<>expand", ">expand", "<expand", ">=expand", "<=expand", _
                 "-expand", "*expand", "quotient-expand", "mod-expand", "&expand", "characters-expand", _
                 "address-expand", "row-expand", "column-expand", "column-letters-expand", "assoc-expand"
                AssignVar lprimTop, EvalListopsPrim(headFold, lst, New Collection, "", New Collection)
                DefineNoteChosen lprimTop   ' VLA.6
                GoTo finish
        End Select
        ' `FormLine(cur) = 0` gates this to template-copy output ONLY -
        ' a hand-written top-level `(begin ...)` statement, parsed
        ' straight from source, carries its own real line (tagged
        ' during parsing, well before this function ever sees it) and
        ' must keep it; only a macro's own SUBSTITUTED result (already
        ' 0-tagged by Substitute/ExpandOne per LISTOPS-PURITY's own
        ' binding note - never a real user line to begin with) is safe
        ' to reassemble and re-tag 0 down in the pendingHead/finalOut
        ' combine below. Without this gate, an ordinary hand-written
        ' `(begin (foo x) (bar y))` reached directly would lose its own
        ' real line tag for no reason - the whole point here is chains
        ' that were ALREADY untagged, not a new class of untagged form.
        ' VLA.5: a loop at a form's head, outside any template, is a
        ' statement: its results as one begin (tagged 0), handed back to this
        ' loop so that its rows expand - a primitive's result above is
        ' returned as it is, and a loop's rows would then never expand. One
        ' link of the chain, as a begin unwrapped is. A helper, so this
        ' recursive procedure gains no local (trap 8).
        If headFold = "each" Or headFold = "each-index" Then
            chainLen = chainLen + 1
            If chainLen > 5000 Then VLA_Messages.RaiseMsg "vla-macro-self-expansion-limit"
            Set cur = EachFormBegin(lst, New Collection, "", New Collection)
        ElseIf headFold = "begin" And lst.Count > 1 And FormLine(cur) = 0 Then
            For bi = 2 To lst.Count - 1
                pendingHead.Add Nth(lst, bi)
            Next
            chainLen = chainLen + 1
            If chainLen > 5000 Then VLA_Messages.RaiseMsg "vla-macro-self-expansion-limit"
            AssignVar cur, Nth(lst, lst.Count)
        ElseIf headFold = "begin" Then
            Exit Do
        Else
            Set rec = GetMacro(CStr(lst.Item(1)))
            If rec Is Nothing Then Exit Do
            If mBudgetOn And mExpandFired >= mExpandBudget Then Exit Do   ' L15
            chainLen = chainLen + 1
            If chainLen > 5000 Then VLA_Messages.RaiseMsg "vla-macro-self-expansion-limit"
            AssignVar r, ExpandOne(lst, rec)
            mExpandFired = mExpandFired + 1   ' L1: one per application
            AssignVar cur, r
        End If
nextLink:
    Loop
    ' VLA.5: FOUND WHILE BUILDING THE LOOP, AND FIXED. A begin a macro made
    ' whose last form is an atom - (defmacro (m) (debug-print 1) x), or a
    ' loop's results that end in one - reaches here with its other forms set
    ' aside, where the atom alone used to be returned and they were dropped:
    ' (m) expanded to x. They are kept now, in order, the atom last.
    If Not IsList(cur) Then
        Set ExpandMacros = PendingBegin(pendingHead, cur, depth, walkDepth)
        Exit Function
    End If
    ' No macro at the head: rebuild with expanded elements. C2: the
    ' rebuild copies the source-line tag - expansion rebuilds EVERY
    ' list, so without this hop the whole map would die here. The
    ' macro path above deliberately does NOT tag its result: template
    ' statements have no user line, while the user's argument forms
    ' arrive as elements and keep their tags through this very path.
    ' (S3.1: "no user line" rides through as an explicit zero now -
    ' FormLine of a template copy reads 0, TagLine stores 0, and the
    ' rebuilt list is zero-tagged in turn, never keyless.)
    ' (F.1 postscript: a same-session attempt to tag a macro call's
    ' single-statement expansion with its own line broke two pins in
    ' VLA_Tests.bas that predate it - dotimes' template-clean/args-
    ' keep-lines split, and inc!'s "deliberately untagged" S3.1
    ' regression proof - and neither was checked before the change was
    ' made. Reverted; those two pins are the record of why, and both
    ' pass again now that this reads exactly as it did before.)
    ' PNTH.0: indexed, not For Each - cur can reach here as a bare
    ' VlaSlice (a macro whose own body is a bare cdr/cddr call, with no
    ' begin/if wrapping it, so nothing else consumed it first), and
    ' VlaSlice has no enumerator to walk with For Each. Object-typed and
    ' indexed works identically for either a Collection or a VlaSlice,
    ' no copy needed either way - cheaper than materializing first.
    Dim curObj As Object
    Set curObj = cur
    Dim outc As New Collection
    Dim ei As Long
    ' VLA.6: a list's head is never a word a define binds, so an atom there
    ' is passed by, as the walk would have returned it.
    ei = 1
    If curObj.Count > 0 Then
        If Not IsObject(curObj.Item(1)) Then
            outc.Add curObj.Item(1)
            ei = 2
        End If
    End If
    For ei = ei To curObj.Count
        ' VLA.5: a loop among the elements, its results spliced in, each
        ' expanded where the loop stood (ExpandEachInto).
        If IsEachForm(curObj.Item(ei)) Then
            ExpandEachInto outc, curObj.Item(ei), depth, walkDepth + 1
        Else
            outc.Add ExpandMacros(curObj.Item(ei), depth, walkDepth + 1)
        End If
    Next
    TagLine outc, FormLine(cur)
    ' TABLESPEC-SCALE / TABLESPECSCALE.0: pendingHead is empty for every
    ' macro shape that isn't begin-wrapped-tail-recursive - outc alone,
    ' returned exactly as before, is the whole story for those (LISTOPS,
    ' quote-map, ordinary macros, everything this file's own tests
    ' already pinned). Non-empty only for a walk-rows/tablespec-shaped
    ' chain - its own stashed elements are resolved now (independent,
    ' shallow calls, never nested in each other or in outc's own
    ' resolution) and combined into one begin alongside outc, mirroring
    ' the flat structure the fully-unwound pre-trampoline recursion used
    ' to build, just assembled by a loop instead of by stack depth.
    ' Untagged (TagLine ..., 0), same convention the macro path above
    ' already holds - this wrapper is assembled template output, not a
    ' form with a user line of its own (PendingBegin, VLA.5's helper, so
    ' that the atom's case above and this one assemble it alike).
    If pendingHead.Count = 0 Then
        Set ExpandMacros = outc
        Exit Function
    End If
    Set ExpandMacros = PendingBegin(pendingHead, outc, depth, walkDepth)
    Exit Function
    ' VLA.6: FOUND WHILE SCOPING, AND FIXED. A form of its own the loop above
    ' returns as it is - a quote, symbol's atom, formula-expand's text, a
    ' primitive's result, a define passed by, a reference's part - used to be
    ' returned alone, and the forms a begin a macro made had set aside before
    ' it were dropped: (defmacro (m) (debug-print 1) (quote (a b))) expanded
    ' to (quote (a b)). Each comes here in lprimTop, and those forms are kept
    ' before it, as an atom's are since VLA.5 (PendingBegin).
finish:
    If pendingHead.Count > 0 Then Set lprimTop = PendingBegin(pendingHead, lprimTop, depth, walkDepth)
    If IsObject(lprimTop) Then Set ExpandMacros = lprimTop Else ExpandMacros = lprimTop
End Function

' TABLESPEC-SCALE / VLA.5: the forms ExpandMacros' trampoline set aside,
' each expanded as a chain of its own (a loop among them spliced, as a
' loop among a list's elements is), then the last form, already expanded:
' one begin, tagged 0. PNTH.0: pendingHead is always a genuine Collection
' (built by pendingHead.Add Nth(lst, bi) - a parsed sub-form, never a
' VlaSlice), so For Each is safe here.
Private Function PendingBegin(pendingHead As Collection, lastForm As Variant, ByVal depth As Long, ByVal walkDepth As Long) As Collection
    Dim finalOut As New Collection
    finalOut.Add "begin"
    Dim e As Variant
    For Each e In pendingHead
        If IsEachForm(e) Then
            ExpandEachInto finalOut, e, depth, walkDepth + 1
        Else
            finalOut.Add ExpandMacros(e, depth, walkDepth + 1)
        End If
    Next
    finalOut.Add lastForm
    TagLine finalOut, 0
    Set PendingBegin = finalOut
End Function

' VLA.4: (formula-expand <expression>) - vla-lang's
' Expander::formula_expand. The one argument is expanded to its fixpoint,
' then written by the formula syntax (FormulaSyntaxWrite) as one string
' atom, the formula's text with its = (the leading quote character is the
' reader's string marker). Arity is refused first, then a formula-expand
' inside another's argument, whose text would be a text there.
' mFormulaExpandOpen is cleared on every way out, a refusal's included,
' and a refusal raised again from locals (trap 1).
Private Function ExpandFormulaCall(lst As Object, ByVal depth As Long, ByVal walkDepth As Long) As String
    If lst.Count <> 2 Then VLA_Messages.RaiseMsg "vla-formula-expand-arity", "n", (lst.Count - 1)
    If mFormulaExpandOpen Then VLA_Messages.RaiseMsg "vla-formula-expand-nested"
    Dim expr As Variant
    Dim errNum As Long, errSrc As String, errDesc As String
    Dim syntaxText As String
    mFormulaExpandOpen = True
    On Error GoTo failed
    AssignVar expr, ExpandMacros(lst.Item(2), depth, walkDepth + 1)
    mFormulaExpandOpen = False
    On Error GoTo 0
    syntaxText = FormulaSyntaxWrite(expr)
    FormulaTextKept syntaxText   ' VLA.6
    ExpandFormulaCall = Chr$(34) & syntaxText
    Exit Function
failed:
    errNum = Err.Number
    errSrc = Err.Source
    errDesc = Err.Description
    mFormulaExpandOpen = False
    On Error GoTo 0
    Err.Raise errNum, errSrc, errDesc
End Function

' VLA.5: a loop met among a list's elements by ExpandMacros' walk, outside
' any template: its results, the loop evaluated with no bindings but its
' own names, added to outc in order, each expanded at the depth the loop
' stood at - spliced, as Substitute splices a loop that is a template's
' element.
Private Sub ExpandEachInto(outc As Collection, eachForm As Variant, ByVal depth As Long, ByVal walkDepth As Long)
    Dim results As Collection
    Set results = EachFormResults(eachForm, New Collection, "", New Collection)
    Dim r As Variant
    For Each r In results
        outc.Add ExpandMacros(r, depth, walkDepth)
    Next
End Sub

' PNTH.0: callForm widened Collection -> Object - both ExpandMacros'
' own trampoline and ExpandOnePass can call this with a VlaSlice (a
' prior cdr/cddr's own result standing in as "the call form" when a
' macro's body was itself a bare listops expression). Only ever reads
' .Count/.Item(j), both VlaSlice-safe.
Private Function ExpandOne(callForm As Object, rec As Collection) As Variant
    Dim params As Collection
    Set params = rec.Item(1)
    Dim restName As String
    restName = rec.Item(2)
    Dim tmpl As Variant
    AssignVar tmpl, Nth(rec, 3)

    Dim argc As Long
    argc = callForm.Count - 1
    If Len(restName) = 0 Then
        If argc <> params.Count Then
            VLA_Messages.RaiseMsg "vla-macro-arity", "name", CStr(callForm.Item(1)), "n", params.Count, "given", argc
        End If
    Else
        If argc < params.Count Then
            VLA_Messages.RaiseMsg "vla-macro-arity-min", "name", CStr(callForm.Item(1)), "n", params.Count, "given", argc
        End If
    End If

    Dim bindings As New Collection
    Dim j As Long
    For j = 1 To params.Count
        bindings.Add callForm.Item(j + 1), VLA_Identity.Fold(CStr(params.Item(j)))
    Next
    Dim restItems As New Collection
    For j = params.Count + 2 To callForm.Count
        restItems.Add callForm.Item(j)
    Next

    Dim res As Variant
    AssignVar res, Substitute(tmpl, bindings, VLA_Identity.Fold(restName), restItems)
    If IsObject(res) Then Set ExpandOne = res Else ExpandOne = res
End Function

' QUASIQUOTE: inQuasi defaults False so every existing call site (just
' ExpandOne below) is untouched, byte-for-byte, unless it opts in via
' an explicit (quasiquote ...) in its own template. True means "inside
' an active quasiquote shield" - bare-symbol/restName auto-substitution
' below is suppressed (data by default), and only an explicit
' (unquote ...)/(unquote-splicing ...) hole punches back through to
' ordinary substitution.
Private Function Substitute(t As Variant, bindings As Collection, ByVal restName As String, restItems As Collection, Optional ByVal inQuasi As Boolean = False) As Variant
    If Not IsList(t) Then
        If Not inQuasi Then
            If IsSym(t) Then
                Dim key As String
                key = VLA_Identity.Fold(CStr(t))
                If Len(restName) > 0 Then
                    If key = restName Then
                        ' Rest param used standalone: substitute the whole list.
                        Set Substitute = CopyList(restItems)
                        Exit Function
                    End If
                End If
                Dim found As Boolean
                Dim v As Variant
                AssignVar v, CollGet(bindings, key, found)
                If found Then
                    If IsObject(v) Then Set Substitute = v Else Substitute = v
                    Exit Function
                End If
            End If
        End If
        Substitute = t
        Exit Function
    End If

    Dim lst As Collection
    Set lst = t

    ' QUASIQUOTE: (symbol ...) consumes the whole form and returns one
    ' fused atom - checked before the ordinary per-element walk below,
    ' same as ExpandMacros's own quote skip-check is checked before its
    ' macro lookup. Recognized regardless of inQuasi (see FuseSymbol's
    ' own comment for why it alone needs no shielding awareness).
    ' quasiquote/unquote/unquote-splicing likewise consume their whole
    ' form and return early.
    If lst.Count > 0 Then
        If IsSym(lst.Item(1)) Then
            Dim h As String
            h = VLA_Identity.Fold(CStr(lst.Item(1)))
            Select Case h
                Case "symbol"
                    Substitute = FuseSymbol(lst, bindings, restName, restItems)
                    Exit Function
                Case "quasiquote"
                    If inQuasi Then VLA_Messages.RaiseMsg "vla-quasiquote-nested"
                    If lst.Count <> 2 Then VLA_Messages.RaiseMsg "vla-quasiquote-arity", "n", (lst.Count - 1)
                    Dim qqRes As Variant
                    AssignVar qqRes, Substitute(Nth(lst, 2), bindings, restName, restItems, True)
                    If IsObject(qqRes) Then Set Substitute = qqRes Else Substitute = qqRes
                    Exit Function
                Case "unquote"
                    If Not inQuasi Then VLA_Messages.RaiseMsg "vla-unquote-outside-quasiquote"
                    If lst.Count <> 2 Then VLA_Messages.RaiseMsg "vla-unquote-arity", "n", (lst.Count - 1)
                    Dim uqRes As Variant
                    ' VLA.5: a value is read here, so a loop is its results'
                    ' list, as the rest parameter is its whole list here.
                    AssignVar uqRes, SubstituteValue(Nth(lst, 2), bindings, restName, restItems)
                    If IsObject(uqRes) Then Set Substitute = uqRes Else Substitute = uqRes
                    Exit Function
                Case "unquote-splicing"
                    ' Valid unquote-splicing is intercepted by the
                    ' per-element loop below, as a list element, before
                    ' it would ever reach here as t itself - reaching
                    ' here means either no enclosing shield at all, or
                    ' a shielded-but-wrong position (a bare value, not
                    ' a list element).
                    If Not inQuasi Then VLA_Messages.RaiseMsg "vla-unquote-splicing-outside-quasiquote"
                    VLA_Messages.RaiseMsg "vla-unquote-splicing-not-element"
                Case "car", "cdr", "cddr", "cons", "list", "null?", "eq?", "equal?", "quote-if", "cond", _
                     "+expand", "=expand", "<>expand", ">expand", "<expand", ">=expand", "<=expand", _
                     "-expand", "*expand", "quotient-expand", "mod-expand", "&expand", "characters-expand", _
                     "address-expand", "row-expand", "column-expand", "column-letters-expand", "assoc-expand"
                    ' LISTOPS: unlike symbol, NOT recognized while
                    ' inQuasi is True - a shielded (car x) stays ordinary
                    ' shielded list data, same as everything else under
                    ' a quasiquote, reachable only via an explicit
                    ' (unquote (car x)). symbol's own exemption doesn't
                    ' generalize here: an unbound symbol argument is
                    ' always a safe passthrough, but car on arbitrary
                    ' shielded template data is not. When inQuasi is
                    ' True, execution falls through (no Exit Function)
                    ' to the ordinary per-element walk below.
                    If Not inQuasi Then
                        Dim lprim As Variant
                        AssignVar lprim, EvalListopsPrim(h, lst, bindings, restName, restItems)
                        If IsObject(lprim) Then Set Substitute = lprim Else Substitute = lprim
                        Exit Function
                    End If
                Case "each", "each-index"
                    ' VLA.5: a loop standing alone is a statement, its results
                    ' one begin (EachFormBegin), taken here at its head before
                    ' the element walk below, where the rest parameter in its
                    ' binding would be spliced. Shielded data while inQuasi, as
                    ' the primitives are. A helper, so this frame gains no local.
                    If Not inQuasi Then
                        Set Substitute = EachFormBegin(lst, bindings, restName, restItems)
                        Exit Function
                    End If
            End Select
        End If
    End If

    Dim outc As New Collection
    Dim e As Variant
    Dim spliced As Boolean
    For Each e In lst
        spliced = False
        If inQuasi And ListHeadIs(e, "unquote-splicing") Then
            ' QUASIQUOTE: the explicit escape - generalizes the bare-
            ' restName splice below beyond just the rest param, but
            ' only inside a shield, where that implicit splice is off.
            Dim splLst As Collection
            Set splLst = e
            If splLst.Count <> 2 Then VLA_Messages.RaiseMsg "vla-unquote-splicing-arity", "n", (splLst.Count - 1)
            Dim splVal As Variant
            ' VLA.5: a value is read here too, so a loop splices its results.
            AssignVar splVal, SubstituteValue(Nth(splLst, 2), bindings, restName, restItems)
            If Not IsList(splVal) Then VLA_Messages.RaiseMsg "vla-unquote-splicing-not-list"
            ' PNTH.0: indexed, not For Each - splVal may be a VlaSlice
            ' (the splice argument can legitimately resolve through a
            ' cdr/cddr, e.g. ,@(cdr lst)), and VlaSlice has no
            ' enumerator to walk with For Each.
            Dim splC As Object
            Set splC = splVal
            Dim si As Long
            For si = 1 To splC.Count
                outc.Add splC.Item(si)
            Next
            spliced = True
        ElseIf Not inQuasi Then
            If IsSym(e) Then
                If Len(restName) > 0 Then
                    If VLA_Identity.Fold(CStr(e)) = restName Then
                        ' Rest param as a list element: splice its forms in.
                        Dim ri As Variant
                        For Each ri In restItems
                            outc.Add ri
                        Next
                        spliced = True
                    End If
                End If
            ElseIf IsEachForm(e) Then
                ' VLA.5: a loop as a list element: its results spliced in, as
                ' the rest parameter's forms are (EachFormInto).
                EachFormInto outc, e, bindings, restName, restItems
                spliced = True
            End If
        End If
        If Not spliced Then outc.Add Substitute(e, bindings, restName, restItems, inQuasi)
    Next
    TagLine outc, 0   ' S3.1: template copies are EXPLICITLY untagged -
                      ' spliced user forms above keep their own tags,
                      ' which is the attribution doctrine unchanged
    Set Substitute = outc
End Function

' QUASIQUOTE: (symbol A [B C...]) fuses each argument's resolved text
' into one new symbol atom, e.g. (symbol "make-" d) -> the atom
' make-bold when d is bound to bold. Generalizes what G11's make-{d}
' and the English matcher's own FormSubstitute (VLA_English.bas)
' already do in miniature, in one place, for one caller - those two
' stay untouched; they serve a different caller (English pattern-
' matching time) at a different layer.
'
' Called from two sites: Substitute above (inside a defmacro template,
' with that macro's real bindings/restName/restItems) and ExpandMacros
' (standalone, anywhere in ordinary program text, with empty tables -
' symbol is the one primitive of the QUASIQUOTE bundle that needs no
' ambient environment, since an unbound symbol argument already passes
' through as itself via Substitute's own scalar branch above).
'
' At least one argument is required (lst.Count >= 2, head + args) -
' zero arguments has nothing to fuse and no value to return. A single
' argument is explicitly legal, not an error: it's a true no-op when
' the argument is already a symbol, but a real string-literal ->
' symbol coercion when it isn't - falls out of this same loop with no
' special-casing.
' PNTH.0: lst widened Collection -> Object - ExpandMacros' own
' trampoline can re-dispatch here on data that survived a prior cdr/
' cddr (VlaSlice), not just a freshly-parsed form. Body only ever reads
' .Count/Nth(lst,k), both VlaSlice-safe.
Private Function FuseSymbol(lst As Object, bindings As Collection, ByVal restName As String, restItems As Collection) As String
    If lst.Count < 2 Then VLA_Messages.RaiseMsg "vla-symbol-arity"
    Dim fused As String
    Dim k As Long
    For k = 2 To lst.Count
        Dim piece As Variant
        AssignVar piece, Substitute(Nth(lst, k), bindings, restName, restItems)
        If IsList(piece) Then VLA_Messages.RaiseMsg "vla-symbol-arg-is-list", "n", (k - 1)
        Dim ptext As String
        If Left$(CStr(piece), 1) = Chr$(34) Then
            ptext = StrLitContent(piece)
        Else
            ptext = CStr(piece)
        End If
        fused = fused & ptext
    Next
    FuseSymbol = fused
End Function

' =====================================================================
'  VLA.5: the loop at expand time, (each (name list) form ...) and
'  (each-index (index name list start) form ...). A form of the walks,
'  beside formula-expand, and no LISTOPS primitive: its forms are a
'  template substituted once for each element of the list, the name bound
'  to the element, and its results take the rest parameter's own rule -
'  spliced where the loop is a list's element (EachFormInto), the list of
'  them where a value is read (SubstituteValue), and one begin where it
'  stands alone as a statement (EachFormBegin). No gensym and no eval: the
'  names are the author's, and the loop reads nothing but the source.
' =====================================================================

' VLA.5: whether a form is a loop - a list headed by each or each-index.
' Asked of every list the walks meet, so a head of another length is passed
' over before it is folded: ListHeadIs twice would fold four texts a list.
Private Function IsEachForm(v As Variant) As Boolean
    If Not IsObject(v) Then Exit Function
    Dim lst As Object
    Set lst = v
    If lst.Count = 0 Then Exit Function
    If IsObject(lst.Item(1)) Then Exit Function
    Dim h As String
    h = CStr(lst.Item(1))
    If Len(h) <> 4 And Len(h) <> 10 Then Exit Function
    h = VLA_Identity.Fold(h)
    IsEachForm = (h = "each" Or h = "each-index")
End Function

' VLA.5: a loop standing alone, a statement - a macro's whole template, a
' branch of quote-if or cond, a form of the source: its results as one
' begin, tagged 0 as every list the macro system makes is (S3.1), so
' ExpandMacros' trampoline unwraps it.
Private Function EachFormBegin(eachForm As Variant, bindings As Collection, ByVal restName As String, restItems As Collection) As Collection
    Dim results As Collection
    Set results = EachFormResults(eachForm, bindings, restName, restItems)
    Dim outc As New Collection
    outc.Add "begin"
    Dim r As Variant
    For Each r In results
        outc.Add r
    Next
    TagLine outc, 0
    Set EachFormBegin = outc
End Function

' VLA.5: a loop as a list's element: its results added to outc in order,
' spliced as the rest parameter's forms are.
Private Sub EachFormInto(outc As Collection, eachForm As Variant, bindings As Collection, ByVal restName As String, restItems As Collection)
    Dim results As Collection
    Set results = EachFormResults(eachForm, bindings, restName, restItems)
    Dim r As Variant
    For Each r In results
        outc.Add r
    Next
End Sub

' VLA.5: a form where a value is read - a primitive's argument
' (ResolveListopsArg), unquote's, unquote-splicing's: a loop there is the
' list of its results, as the rest parameter is its whole list there; any
' other form as Substitute makes it.
Private Function SubstituteValue(t As Variant, bindings As Collection, ByVal restName As String, restItems As Collection) As Variant
    If IsEachForm(t) Then
        Set SubstituteValue = EachFormResults(t, bindings, restName, restItems)
        Exit Function
    End If
    Dim v As Variant
    AssignVar v, Substitute(t, bindings, restName, restItems)
    If IsObject(v) Then Set SubstituteValue = v Else SubstituteValue = v
End Function

' VLA.5: a loop's results, a new list tagged 0. The binding read whole, by
' its shape (EachBinding); each name a loop binds checked (EachCheckName);
' the list resolved as a primitive's argument is (EachList); each-index's
' start a whole number (EachStartValue). Then, for each element in order,
' every form of the loop substituted as a list's element is (a loop among
' them spliced, the rest parameter spliced), under a scope of the name
' bound to the element - and the index to the element's count, the start
' up by one - beside the template's own bindings its forms read: those,
' found once by a walk of the forms (EachReadBindings), since a Collection
' cannot list its keys to be copied.
Private Function EachFormResults(eachForm As Variant, bindings As Collection, ByVal restName As String, restItems As Collection) As Collection
    Dim lst As Object
    Set lst = eachForm
    Dim op As String
    op = VLA_Identity.Fold(CStr(lst.Item(1)))
    Dim binding As Object
    Set binding = EachBinding(lst, op)
    Dim indexed As Boolean
    indexed = (op = "each-index")
    Dim indexName As String, varName As String
    Dim indexFold As String, varFold As String
    Dim listAt As Long
    If indexed Then
        indexName = CStr(binding.Item(1))
        varName = CStr(binding.Item(2))
        listAt = 3
        EachCheckName op, indexName, bindings, restName
        indexFold = VLA_Identity.Fold(indexName)
    Else
        varName = CStr(binding.Item(1))
        listAt = 2
    End If
    EachCheckName op, varName, bindings, restName
    varFold = VLA_Identity.Fold(varName)
    If indexed Then
        If varFold = indexFold Then VLA_Messages.RaiseMsg "vla-each-name-bound", "op", op, "word", varName
    End If
    Dim walked As Object
    Set walked = EachList(op, binding.Item(listAt), bindings, restName, restItems)
    Dim startN As Double
    If indexed Then startN = EachStartValue(op, binding.Item(4), bindings, restName, restItems)

    Dim readKeys As New Collection, readVals As New Collection, readSeen As New Collection
    Dim fi As Long
    For fi = 3 To lst.Count
        EachReadBindings lst.Item(fi), bindings, varFold, indexFold, readKeys, readVals, readSeen
    Next

    Dim results As New Collection
    Dim scope As Collection
    Dim wi As Long, ki As Long
    For wi = 1 To walked.Count
        Set scope = New Collection          ' trap 2: a new scope each element
        scope.Add walked.Item(wi), varFold
        If indexed Then scope.Add Trim$(Str$(startN + wi - 1)), indexFold
        For ki = 1 To readKeys.Count
            scope.Add readVals.Item(ki), CStr(readKeys.Item(ki))
        Next
        For fi = 3 To lst.Count
            EachBodyFormInto results, lst.Item(fi), scope, restName, restItems
        Next
    Next
    TagLine results, 0
    Set EachFormResults = results
End Function

' VLA.5: the loop's binding, or the refusal that shows its shape: (name
' list) for each, (index name list start) for each-index, each name a
' symbol, and at least one form after it.
Private Function EachBinding(lst As Object, ByVal op As String) As Object
    Dim want As Long
    If op = "each-index" Then want = 4 Else want = 2
    Dim ok As Boolean
    Dim b As Object
    ok = (lst.Count >= 3)
    If ok Then ok = IsList(lst.Item(2))
    If ok Then
        Set b = lst.Item(2)
        ok = (b.Count = want)
    End If
    If ok Then ok = IsSym(b.Item(1))
    If ok And want = 4 Then ok = IsSym(b.Item(2))
    If Not ok Then
        If want = 4 Then
            VLA_Messages.RaiseMsg "vla-each-shape", "op", op, "binding", "(index name list start)"
        Else
            VLA_Messages.RaiseMsg "vla-each-shape", "op", op, "binding", "(name list)"
        End If
    End If
    Set EachBinding = b
End Function

' VLA.5: a name a loop binds, refused when it is an expand-time name (the
' table defmacro's names are refused by) or a name the template already
' binds: the macro's parameter, its rest, an enclosing loop's name.
Private Sub EachCheckName(ByVal op As String, ByVal word As String, bindings As Collection, ByVal restName As String)
    Dim wordFold As String
    wordFold = VLA_Identity.Fold(word)
    Dim desc As String
    desc = ExpandReservedDesc(wordFold)
    If Len(desc) > 0 Then VLA_Messages.RaiseMsg "vla-each-name-reserved", "op", op, "word", wordFold, "desc", desc
    Dim found As Boolean
    CollGet bindings, wordFold, found
    If Len(restName) > 0 Then
        If wordFold = restName Then found = True
    End If
    If found Then VLA_Messages.RaiseMsg "vla-each-name-bound", "op", op, "word", word
End Sub

' VLA.5: the list a loop walks, resolved as a primitive's argument is: a
' parameter, the rest parameter, a quoted list, a primitive's result or
' another loop's results. A list written as a macro's call is refused by
' name, since a primitive's argument is never expanded and the call would
' be walked as its symbols; an atom is refused as no list.
Private Function EachList(ByVal op As String, listForm As Variant, bindings As Collection, ByVal restName As String, restItems As Collection) As Object
    If IsList(listForm) Then
        Dim lf As Object
        Set lf = listForm
        If lf.Count > 0 Then
            If IsSym(lf.Item(1)) Then
                If Not GetMacro(CStr(lf.Item(1))) Is Nothing Then VLA_Messages.RaiseMsg "vla-each-list-a-call", "op", op, "head", CStr(lf.Item(1))
            End If
        End If
    End If
    Dim v As Variant
    AssignVar v, ResolveListopsArg(listForm, bindings, restName, restItems)
    If Not IsList(v) Then VLA_Messages.RaiseMsg "vla-expand-expected-list", "op", op, "value", CStr(v)
    Set EachList = v
End Function

' VLA.5: each-index's start, resolved as a primitive's argument and held
' to the arithmetic's gate: a number literal (IsNumericLiteralText, never
' VBA's IsNumeric) and a whole one, so the count is exact.
Private Function EachStartValue(ByVal op As String, startForm As Variant, bindings As Collection, ByVal restName As String, restItems As Collection) As Double
    Dim v As Variant
    AssignVar v, ResolveListopsArg(startForm, bindings, restName, restItems)
    If IsObject(v) Then VLA_Messages.RaiseMsg "vla-arith-expand-not-number-list", "op", op
    If Not IsNumericLiteralText(CStr(v)) Then VLA_Messages.RaiseMsg "vla-arith-expand-not-number-value", "op", op, "value", CStr(v)
    Dim n As Double
    n = Val(CStr(v))
    If n <> Int(n) Then VLA_Messages.RaiseMsg "vla-arith-expand-not-whole", "op", op, "value", CStr(v)
    EachStartValue = n
End Function

' VLA.5: the template's bindings a loop's forms read, each kept once with
' its value: every symbol of a form, a quoted form's included since
' Substitute substitutes inside a quote, the loop's own names left out
' (skipA, skipB), looked up in the bindings once (readSeen). A walk of the
' written form, as deep as its text, so the depth limit bounds it.
Private Sub EachReadBindings(f As Variant, bindings As Collection, ByVal skipA As String, ByVal skipB As String, readKeys As Collection, readVals As Collection, readSeen As Collection)
    If IsObject(f) Then
        Dim fl As Object
        Set fl = f
        Dim i As Long
        For i = 1 To fl.Count
            EachReadBindings fl.Item(i), bindings, skipA, skipB, readKeys, readVals, readSeen
        Next
        Exit Sub
    End If
    If Not IsSym(f) Then Exit Sub
    Dim k As String
    k = VLA_Identity.Fold(CStr(f))
    If k = skipA Or k = skipB Then Exit Sub
    Dim found As Boolean
    CollGet readSeen, k, found
    If found Then Exit Sub
    readSeen.Add k, k
    Dim v As Variant
    AssignVar v, CollGet(bindings, k, found)
    If Not found Then Exit Sub
    readKeys.Add k
    readVals.Add v
End Sub

' VLA.5: one of a loop's forms for one element, as a list's element is
' substituted: a loop spliced, the rest parameter spliced, any other form
' substituted under the element's scope.
Private Sub EachBodyFormInto(outc As Collection, f As Variant, scope As Collection, ByVal restName As String, restItems As Collection)
    If IsEachForm(f) Then
        EachFormInto outc, f, scope, restName, restItems
        Exit Sub
    End If
    If IsSym(f) And Len(restName) > 0 Then
        If VLA_Identity.Fold(CStr(f)) = restName Then
            Dim ri As Variant
            For Each ri In restItems
                outc.Add ri
            Next
            Exit Sub
        End If
    End If
    outc.Add Substitute(f, scope, restName, restItems)
End Sub

' =====================================================================
'  VLA.6: a name at expand time, (define name form). A form of the walk
'  that writes nothing where it stands: the walk passes it by as written
'  (DefinePassedBy), and once the form of the source it stands in has
'  expanded, ExpandSourceForm takes the defines from that form's
'  statements (DefinesTake), so that a word binds from the next form of
'  the source on, whatever order the trampoline expanded the form in. The
'  walk's atom case puts a define's form where its word stands as an
'  expression (DefineNamed), never at a list's head, never inside a quote,
'  a text or a primitive's argument. No gensym and no eval: the word is
'  the author's and the form the source's own, both gone at the pass.
' =====================================================================

' VLA.6: one form of the source to its fixpoint, and then, when the walk
' passed a define by, the form's defines taken from its statements and
' registered for the forms after it: the one place the three entry points
' expand a form of the source.
Private Function ExpandSourceForm(f As Variant) As Variant
    mDefineMet = False
    Set mFormulaTexts = Nothing
    Dim r As Variant
    AssignVar r, ExpandMacros(f, 0)
    If mDefineMet Then AssignVar r, DefinesTake(r)
    If IsObject(r) Then Set ExpandSourceForm = r Else ExpandSourceForm = r
End Function

' VLA.6: whether a form is a define, a list headed by define; a head of
' another length is passed over before it is folded, as IsEachForm's is.
Private Function IsDefineForm(v As Variant) As Boolean
    If Not IsObject(v) Then Exit Function
    Dim lst As Object
    Set lst = v
    If lst.Count = 0 Then Exit Function
    If IsObject(lst.Item(1)) Then Exit Function
    Dim h As String
    h = CStr(lst.Item(1))
    If Len(h) <> 6 Then Exit Function
    IsDefineForm = (VLA_Identity.Fold(h) = "define")
End Function

' VLA.6: a define the walk meets at a form's head, passed by as written for
' ExpandSourceForm to take. Inside a formula's expression it would be
' written as a function's call, so it is refused there.
Private Function DefinePassedBy(lst As Object) As Object
    If mFormulaExpandOpen Then VLA_Messages.RaiseMsg "vla-define-not-a-statement", "word", DefineWordOf(lst)
    mDefineMet = True
    Set DefinePassedBy = lst
End Function

' VLA.6: a standalone primitive's chosen branch is not expanded again (VLA.5,
' left as found), so a define it chooses is met at no head: noted here, as
' DefinePassedBy notes one, and refused inside a formula's expression.
Private Sub DefineNoteChosen(v As Variant)
    Dim held As Object
    Set held = DefineHeldIn(v)
    If held Is Nothing Then Exit Sub
    If mFormulaExpandOpen Then VLA_Messages.RaiseMsg "vla-define-not-a-statement", "word", DefineWordOf(held)
    mDefineMet = True
End Sub

' VLA.6: the first define a form holds, itself or at any depth, a quote's
' data passed over; Nothing for none.
Private Function DefineHeldIn(v As Variant) As Object
    If Not IsObject(v) Then Exit Function
    If ListHeadIs(v, "quote") Then Exit Function
    If IsDefineForm(v) Then
        Set DefineHeldIn = v
        Exit Function
    End If
    Dim lst As Object
    Set lst = v
    Dim inner As Object
    Dim i As Long
    For i = 1 To lst.Count
        Set inner = DefineHeldIn(lst.Item(i))
        If Not inner Is Nothing Then
            Set DefineHeldIn = inner
            Exit Function
        End If
    Next
End Function

' VLA.6: a define's word as a refusal names it: the word, or what stands
' where the word belongs, as written.
Private Function DefineWordOf(lst As Object) As String
    If lst.Count < 2 Then Exit Function
    DefineWordOf = WriteDatum(lst.Item(2))
End Function

' VLA.6: a form of the source, expanded, with its defines taken from it.
' Those among its statements - the form itself, or a begin's elements at
' any depth of begins - are taken in order, each leaving nothing where it
' stood (DefinesFromStatement); one anywhere else is refused
' (DefinesRefuseStray). Each is checked (DefineTakeOne); then a word of
' theirs used in this same form, where it is not yet bound, is refused,
' found in the result (DefinesRefuseUse) or in a text formula-expand wrote
' (DefinesRefuseFormulaUse), and so is a ring of words (DefinesRefuseRing);
' and they bind the forms of the source after this one.
Private Function DefinesTake(r As Variant) As Variant
    Dim found As New Collection
    Dim kept As Variant
    AssignVar kept, DefinesFromStatement(r, found)
    DefinesRefuseStray kept
    Dim taken As Object
    Set taken = CreateObject("Scripting.Dictionary")
    Dim d As Variant
    For Each d In found
        DefineTakeOne d, taken
    Next
    DefinesRefuseUse kept, taken
    DefinesRefuseFormulaUse taken
    Dim k As Variant
    For Each k In taken.Keys
        DefinesRefuseRing CStr(k), taken
    Next
    If mDefines Is Nothing Then Set mDefines = CreateObject("Scripting.Dictionary")
    For Each k In taken.Keys
        mDefines.Add k, taken.Item(k)
    Next
    If IsObject(kept) Then Set DefinesTake = kept Else DefinesTake = kept
End Function

' VLA.6: a statement with the defines among it taken out, in order, into
' found: a define itself leaves an empty begin; a begin is rebuilt without
' them, a begin among its elements the same, keeping its line; any other
' form is kept whole.
Private Function DefinesFromStatement(v As Variant, found As Collection) As Variant
    If IsDefineForm(v) Then
        found.Add v
        Dim leftNone As New Collection
        leftNone.Add "begin"
        TagLine leftNone, 0
        Set DefinesFromStatement = leftNone
        Exit Function
    End If
    If Not ListHeadIs(v, "begin") Then
        If IsObject(v) Then Set DefinesFromStatement = v Else DefinesFromStatement = v
        Exit Function
    End If
    Dim lst As Object
    Set lst = v
    Dim outc As New Collection
    outc.Add lst.Item(1)
    Dim i As Long
    For i = 2 To lst.Count
        If IsDefineForm(lst.Item(i)) Then
            found.Add lst.Item(i)
        Else
            outc.Add DefinesFromStatement(lst.Item(i), found)
        End If
    Next
    TagLine outc, FormLine(v)
    Set DefinesFromStatement = outc
End Function

' VLA.6: a define left anywhere in a form the statements' walk kept whole -
' a row's argument, a call's, a procedure's body - refused by name, since a
' define stands as a form of the source or a statement of a begin standing
' as one; a quote's data passed over.
Private Sub DefinesRefuseStray(v As Variant)
    If Not IsObject(v) Then Exit Sub
    If ListHeadIs(v, "quote") Then Exit Sub
    Dim lst As Object
    Set lst = v
    If IsDefineForm(v) Then VLA_Messages.RaiseMsg "vla-define-not-a-statement", "word", DefineWordOf(lst)
    Dim i As Long
    For i = 1 To lst.Count
        DefinesRefuseStray lst.Item(i)
    Next
End Sub

' VLA.6: one define taken from its form: its shape, (define name form), the
' name a symbol and the form holding no define (DefinesRefuseStray); its
' name held to the rules (DefineCheckName); then the word, folded, and its
' form as written into taken.
Private Sub DefineTakeOne(d As Variant, taken As Object)
    Dim lst As Object
    Set lst = d
    If lst.Count <> 3 Then VLA_Messages.RaiseMsg "vla-define-shape", "form", WriteDatum(d)
    If Not IsSym(lst.Item(2)) Then VLA_Messages.RaiseMsg "vla-define-shape", "form", WriteDatum(d)
    DefinesRefuseStray lst.Item(3)
    Dim word As String
    word = CStr(lst.Item(2))
    DefineCheckName word, taken
    taken.Add VLA_Identity.Fold(word), lst.Item(3)
End Sub

' VLA.6: a define's word, refused by name when the reference scanner reads
' it whole as a reference or a name on a sheet, or it has the R1C1 shape
' Excel refuses a name (DefineReadsAs); when it is a number, a truth value,
' a keyword, a form of the head table, an expand-time name or a form of the
' formula syntax's own (DefineReservedDesc); and when it is a macro's name
' or defined already, in an earlier form or in this one.
Private Sub DefineCheckName(ByVal word As String, taken As Object)
    Dim what As String
    what = DefineReadsAs(word)
    If Len(what) > 0 Then VLA_Messages.RaiseMsg "vla-define-name-reference", "word", word, "what", what
    what = DefineReservedDesc(word)
    If Len(what) > 0 Then VLA_Messages.RaiseMsg "vla-define-name-reserved", "word", word, "desc", what
    If Not GetMacro(word) Is Nothing Then VLA_Messages.RaiseMsg "vla-define-name-taken", "word", word, "what", "a macro's name"
    If DefineBound(VLA_Identity.Fold(word), taken) Then VLA_Messages.RaiseMsg "vla-define-name-taken", "word", word, "what", "defined already"
End Sub

' VLA.6: what the reference scanner reads a define's word as, "" for a word
' of the source's own: a cell, a range, a column, a row, a span of sheets,
' a name on a sheet or any other reference it reads whole; or a reference
' in R1C1 notation, R, C, RC or R1C1, which Excel refuses as a name.
Private Function DefineReadsAs(ByVal word As String) As String
    Dim what As String
    Dim fr As FormulaRef
    If FsWholeReference(word, fr) Then
        Select Case fr.Kind
            Case "cell"
                what = "a cell"
            Case "range"
                what = "a range"
            Case "column"
                what = "a column"
            Case "row"
                what = "a row"
            Case "3d"
                what = "a span of sheets"
            Case "name"
                If Len(fr.SheetName) > 0 Then what = "a name on a sheet"
            Case Else
                what = "a reference"
        End Select
    End If
    If Len(what) = 0 Then
        If VLA_Refers.RefersR1C1Shaped(word) Then what = "a reference in R1C1 notation"
    End If
    DefineReadsAs = what
End Function

' VLA.6: what a define's word already is in the language, "" for a word of
' the source's own: a number, a truth value, a keyword argument, a form of
' the head table or one of its aliases, or defmacro; an expand-time name
' (ExpandReservedDesc); or a form of the formula syntax's own.
Private Function DefineReservedDesc(ByVal word As String) As String
    Dim wordFold As String
    Dim desc As String
    wordFold = VLA_Identity.Fold(word)
    If FsIsNumeral(word) Or FsIsNegativeNumeral(word) Or FsVbaNumeral(word) Then
        desc = "a number"
    ElseIf wordFold = "true" Or wordFold = "false" Then
        desc = "a truth value"
    ElseIf IsKeywordArg(word) Then
        desc = "a keyword argument"
    ElseIf Not VLA_HeadTable.VlaHeadTableRow(wordFold) Is Nothing Then
        desc = "a form of the language"
    ElseIf wordFold = "defmacro" Then
        desc = "a form of the language"
    ElseIf Len(ExpandReservedDesc(wordFold)) > 0 Then
        desc = "the " & ExpandReservedDesc(wordFold)
    ElseIf wordFold = "error" Or wordFold = "ref" Or wordFold = "last" Then
        desc = "a form of the formula syntax's own"
    End If
    DefineReservedDesc = desc
End Function

' VLA.6: a word a define of this form binds, standing in the form's result
' where the walk would have put its form - the result itself, a list's
' element after an atom head, never a quote's - refused: in the form that
' defines it a word is not yet bound.
Private Sub DefinesRefuseUse(v As Variant, taken As Object)
    If Not IsObject(v) Then
        If IsSym(v) Then
            If taken.Exists(VLA_Identity.Fold(CStr(v))) Then VLA_Messages.RaiseMsg "vla-define-used-in-its-form", "word", CStr(v)
        End If
        Exit Sub
    End If
    If ListHeadIs(v, "quote") Then Exit Sub
    Dim lst As Object
    Set lst = v
    Dim i As Long
    For i = 1 To lst.Count
        If i > 1 Or IsObject(lst.Item(i)) Then DefinesRefuseUse lst.Item(i), taken
    Next
End Sub

' VLA.6: a word a define of this form binds, written by formula-expand into
' one of this form's formulas as one of Excel's names, since it was not yet
' bound there: each text formula-expand wrote in the form read again by
' RefersScan, and a name with no sheet that is such a word refused.
Private Sub DefinesRefuseFormulaUse(taken As Object)
    If mFormulaTexts Is Nothing Then Exit Sub
    Dim t As Variant
    Dim refs() As FormulaRef
    Dim n As Long, i As Long
    For Each t In mFormulaTexts
        n = VLA_Refers.RefersScan(CStr(t), refs)
        For i = 1 To n
            If refs(i).Kind = "name" And Len(refs(i).SheetName) = 0 Then
                If taken.Exists(VLA_Identity.Fold(refs(i).Part)) Then VLA_Messages.RaiseMsg "vla-define-used-in-its-form", "word", refs(i).Part
            End If
        Next
    Next
End Sub

' VLA.6: a ring of words, a define whose form comes back to its own word
' through the words its forms stand for, refused when its last define
' registers, the ring named. The words a define's form would have replaced
' (DefineWordsIn) are followed once each, breadth first, through every word
' a define binds, in an earlier form or this one; the ring is the way back.
Private Sub DefinesRefuseRing(ByVal startName As String, taken As Object)
    Dim cameFrom As Object
    Set cameFrom = CreateObject("Scripting.Dictionary")
    Dim queue As New Collection
    queue.Add startName
    Dim queueAt As Long
    queueAt = 1
    Dim atName As String
    Dim wordsIn As Collection
    Dim wordAt As Variant
    Do While queueAt <= queue.Count
        atName = CStr(queue.Item(queueAt))
        queueAt = queueAt + 1
        Set wordsIn = New Collection
        DefineWordsIn DefineValueOf(atName, taken), wordsIn
        For Each wordAt In wordsIn
            If CStr(wordAt) = startName Then VLA_Messages.RaiseMsg "vla-define-cycle", "word", startName, "ring", DefineRingText(startName, atName, cameFrom)
            If DefineBound(CStr(wordAt), taken) Then
                If Not cameFrom.Exists(CStr(wordAt)) Then
                    cameFrom.Add CStr(wordAt), atName
                    queue.Add CStr(wordAt)
                End If
            End If
        Next
    Loop
End Sub

' VLA.6: the words a define's form would have replaced where the walk meets
' them - the form itself when a symbol, a list's elements after an atom
' head, never a quote's - folded, in order, into wordsIn.
Private Sub DefineWordsIn(v As Variant, wordsIn As Collection)
    If Not IsObject(v) Then
        If IsSym(v) Then wordsIn.Add VLA_Identity.Fold(CStr(v))
        Exit Sub
    End If
    If ListHeadIs(v, "quote") Then Exit Sub
    Dim lst As Object
    Set lst = v
    Dim i As Long
    For i = 1 To lst.Count
        If i > 1 Or IsObject(lst.Item(i)) Then DefineWordsIn lst.Item(i), wordsIn
    Next
End Sub

' VLA.6: whether a folded word is bound by a define, of this form or of an
' earlier one.
Private Function DefineBound(ByVal wordFold As String, taken As Object) As Boolean
    If taken.Exists(wordFold) Then
        DefineBound = True
    ElseIf Not mDefines Is Nothing Then
        DefineBound = mDefines.Exists(wordFold)
    End If
End Function

' VLA.6: the form a folded word stands for, a define of this form first; the
' caller has seen that a define binds it (DefineBound).
Private Function DefineValueOf(ByVal wordFold As String, taken As Object) As Variant
    If taken.Exists(wordFold) Then
        If IsObject(taken.Item(wordFold)) Then Set DefineValueOf = taken.Item(wordFold) Else DefineValueOf = taken.Item(wordFold)
    Else
        If IsObject(mDefines.Item(wordFold)) Then Set DefineValueOf = mDefines.Item(wordFold) Else DefineValueOf = mDefines.Item(wordFold)
    End If
End Function

' VLA.6: a ring as its refusal names it, from its first word around to it
' again, startName -> ... -> lastName -> startName, by the word each was
' reached from.
Private Function DefineRingText(ByVal startName As String, ByVal lastName As String, cameFrom As Object) As String
    Dim ringText As String
    Dim nm As String
    ringText = lastName & " -> " & startName
    nm = lastName
    Do While nm <> startName
        nm = CStr(cameFrom.Item(nm))
        ringText = nm & " -> " & ringText
    Loop
    DefineRingText = ringText
End Function

' VLA.6: whether an atom is a word a define binds: a symbol, never a text,
' found in mDefines, which the walk has seen is not Nothing.
Private Function DefineNamed(v As Variant) As Boolean
    If IsObject(v) Then Exit Function
    Dim s As String
    s = CStr(v)
    If Len(s) = 0 Then Exit Function
    If Left$(s, 1) = Chr$(34) Then Exit Function
    DefineNamed = mDefines.Exists(VLA_Identity.Fold(s))
End Function

' VLA.6: a text formula-expand wrote, kept for the form ExpandSourceForm
' expands, so that a word a define of that form binds, which the formula
' wrote as one of Excel's names, is refused (DefinesRefuseFormulaUse).
Private Sub FormulaTextKept(ByVal syntaxText As String)
    If mFormulaTexts Is Nothing Then Set mFormulaTexts = New Collection
    mFormulaTexts.Add syntaxText
End Sub

' VLA.6: (sheet-expand reference) and (range-expand reference), forms of the
' walk beside formula-expand: the argument expanded to its fixpoint, so
' that a word a define binds is replaced first, then read as one reference
' on a sheet, and its sheet's name or its cell or range answered as a text
' (FsReferencePartText), the two texts ref takes; a string atom, as
' formula-expand's is. A helper of its own reads the reference, so that this
' frame, on the walk's way down, holds no record.
Private Function ExpandReferencePart(ByVal op As String, lst As Object, ByVal depth As Long, ByVal walkDepth As Long) As String
    If lst.Count <> 2 Then VLA_Messages.RaiseMsg "vla-expand-arity", "op", op, "want", 1, "n", (lst.Count - 1)
    Dim v As Variant
    AssignVar v, ExpandMacros(lst.Item(2), depth, walkDepth + 1)
    ExpandReferencePart = Chr$(34) & FsReferencePartText(v, op)
End Function

' LISTOPS.0: shared dispatcher for the twenty-eight expand-time list/
' arithmetic/conditional primitives (the original nine, LISTOPS-EXPAND's
' own seven, cond, VLA.3's ten and VLA.5's assoc-expand), called from both
' recognition sites (Substitute's per-template walk and ExpandMacros's
' standalone walk) exactly the way FuseSymbol already serves `symbol` from
' both. Caller has already confirmed headFold is one of the twenty-eight.
' PNTH.0: lst widened Collection -> Object, same reason as FuseSymbol
' just above - passed straight through to whichever Eval* this
' dispatches to, all of which accept Object on this same parameter now.
Private Function EvalListopsPrim(ByVal headFold As String, lst As Object, bindings As Collection, ByVal restName As String, restItems As Collection) As Variant
    Dim result As Variant
    Select Case headFold
        Case "car"
            AssignVar result, EvalCar(lst, bindings, restName, restItems)
        Case "cdr"
            Set result = EvalCdr(lst, bindings, restName, restItems)
        Case "cddr"
            Set result = EvalCddr(lst, bindings, restName, restItems)
        Case "cons"
            Set result = EvalCons(lst, bindings, restName, restItems)
        Case "list"
            Set result = EvalList(lst, bindings, restName, restItems)
        Case "null?"
            result = EvalNullQ(lst, bindings, restName, restItems)
        Case "eq?"
            result = EvalEq(lst, bindings, restName, restItems)
        Case "equal?"
            result = EvalEqual(lst, bindings, restName, restItems)
        Case "quote-if"
            AssignVar result, EvalQuoteIf(lst, bindings, restName, restItems)
        Case "cond"
            AssignVar result, EvalCond(lst, bindings, restName, restItems)
        Case "+expand", "=expand", "<>expand", ">expand", "<expand", ">=expand", "<=expand", _
             "-expand", "*expand", "quotient-expand", "mod-expand"
            result = EvalArithExpand(headFold, lst, bindings, restName, restItems)
        ' VLA.3: the expand-time object model's text and address primitives.
        Case "&expand"
            result = EvalConcatExpand(lst, bindings, restName, restItems)
        Case "characters-expand"
            Set result = EvalCharactersExpand(lst, bindings, restName, restItems)
        Case "address-expand"
            result = EvalAddressExpand(lst, bindings, restName, restItems)
        Case "row-expand", "column-expand"
            result = EvalAddressPartExpand(headFold, lst, bindings, restName, restItems)
        Case "column-letters-expand"
            result = EvalColumnLettersExpand(lst, bindings, restName, restItems)
        ' VLA.5: a piece found by its head.
        Case "assoc-expand"
            AssignVar result, EvalAssocExpand(lst, bindings, restName, restItems)
    End Select
    If IsObject(result) Then Set EvalListopsPrim = result Else EvalListopsPrim = result
End Function

' Resolves one LISTOPS primitive's own argument form - first through the
' ordinary Substitute walk (so a bound parameter name still resolves
' normally, and a nested car/cdr/cons/list call is already reduced by
' Substitute's own head-dispatch before this ever sees it), then unwraps
' exactly one top-level (quote X) wrapper if present: (quote ...) is how
' literal DATA is written in VLA source, and LISTOPS-PURITY's own wall
' requires these primitives to operate on that literal text, not the
' wrapper form - Substitute itself never does this unwrap (P.L5 needs
' quote preserved verbatim for ordinary macro templates reaching
' emission), so it is local to LISTOPS's own argument resolution only.
' Unwraps once, never recursively into nested elements - (quote (quote
' a)), "the list containing the symbols quote and a", unwraps to the
' 2-element list (quote a) as real data, not further.
Private Function ResolveListopsArg(argForm As Variant, bindings As Collection, ByVal restName As String, restItems As Collection) As Variant
    Dim v As Variant
    ' VLA.5: a loop where a primitive reads its argument is the list of its
    ' results, as the rest parameter is its whole list here - SubstituteValue's
    ' rule, written in place so that this frame on every primitive's path
    ' calls no further frame for it.
    If IsEachForm(argForm) Then
        Set v = EachFormResults(argForm, bindings, restName, restItems)
    Else
        AssignVar v, Substitute(argForm, bindings, restName, restItems)
    End If
    If ListHeadIs(v, "quote") Then
        ' PNTH.0: Object, not Collection - v (already-evaluated listops
        ' data, not a parsed form) could be a VlaSlice whose own first
        ' element happens to be the symbol "quote" as DATA, e.g. a cdr
        ' landing on (quote b c) inside (a quote b c).
        Dim qc As Object
        Set qc = v
        If qc.Count <> 2 Then VLA_Messages.RaiseMsg "vla-quote-arity"
        Dim inner As Variant
        AssignVar inner, Nth(qc, 2)
        If IsObject(inner) Then Set ResolveListopsArg = inner Else ResolveListopsArg = inner
        Exit Function
    End If
    If IsObject(v) Then Set ResolveListopsArg = v Else ResolveListopsArg = v
End Function

' (car lst) - lst must resolve to a non-empty list; returns its first
' element. Argument resolved through the same Substitute walk every
' LISTOPS primitive uses - no new evaluation model.
' PNTH.0: lst (the FORM being evaluated, e.g. the parsed "(car x)"
' itself) widened Collection -> Object on every Eval* below, same
' reasoning as FuseSymbol/EvalListopsPrim just above - ExpandMacros' own
' trampoline can re-dispatch a form check against data that survived a
' prior cdr/cddr, not only against freshly-parsed source. Every body
' below only ever reads lst.Count/lst.Item(1)/Nth(lst,k), all
' VlaSlice-safe; this is a widening for callers, not a behavior change.
Private Function EvalCar(lst As Object, bindings As Collection, ByVal restName As String, restItems As Collection) As Variant
    If lst.Count <> 2 Then VLA_Messages.RaiseMsg "vla-car-arity", "n", (lst.Count - 1)
    Dim v As Variant
    AssignVar v, ResolveListopsArg(Nth(lst, 2), bindings, restName, restItems)
    If Not IsList(v) Then VLA_Messages.RaiseMsg "vla-car-expected-list", "value", CStr(v)
    ' PNTH.0: Object, not Collection - v may now be a VlaSlice (an
    ' earlier cdr/cddr's own result), and .Count/.Item are all this
    ' needs from it.
    Dim c As Object
    Set c = v
    If c.Count = 0 Then VLA_Messages.RaiseMsg "vla-car-empty-list"
    Dim result As Variant
    AssignVar result, Nth(c, 1)
    If IsObject(result) Then Set EvalCar = result Else EvalCar = result
End Function

' PNTH.0: a read-only, O(1) VIEW (VlaSlice.cls) over c's items from
' position n+1 onward - was a full copy into a new Collection, the
' evidenced dominant cost in "expand" once P-DICT's own before/after
' ruled out GetMacro (BETA_REARVIEW.md's P-NTH entry has the full
' finding and the S3.1 safety argument for why a shared VIEW, not a
' copy, is still correct here). c may itself already be a VlaSlice (a
' prior cdr/cddr in the same chain) - flattened to point at the SAME
' root Backing with an adjusted Offset rather than wrapping a wrapper,
' so .Item/.Count stay O(1) no matter how many chained cdr's produced
' the input, not just the first one. TagLine ..., 0 unchanged -
' LISTOPS-PURITY's own binding note: cdr/cddr's result must never
' inherit a tag from the operand's own address, or the S3.1 stale-
' ObjPtr-tag bug reopens - still true here: the VlaSlice returned is a
' genuinely fresh object every call, its own distinct identity, never
' Backing's.
Private Function ListTail(c As Object, ByVal n As Long, ByVal callerName As String) As Object
    If c.Count < n Then VLA_Messages.RaiseMsg "vla-list-too-short", "caller", callerName, "n", n, "count", c.Count
    Dim sl As New VlaSlice
    If TypeName(c) = "VlaSlice" Then
        Dim srcSlice As VlaSlice
        Set srcSlice = c
        Set sl.Backing = srcSlice.Backing
        sl.Offset = srcSlice.Offset + n
    Else
        Dim srcColl As Collection
        Set srcColl = c
        Set sl.Backing = srcColl
        sl.Offset = n
    End If
    TagLine sl, 0
    Set ListTail = sl
End Function

' PNTH.0: return type widened Collection -> Object - ListTail now
' returns a VlaSlice (VlaSlice.cls), not a copied Collection.
Private Function EvalCdr(lst As Object, bindings As Collection, ByVal restName As String, restItems As Collection) As Object
    If lst.Count <> 2 Then VLA_Messages.RaiseMsg "vla-cdr-arity", "n", (lst.Count - 1)
    Dim v As Variant
    AssignVar v, ResolveListopsArg(Nth(lst, 2), bindings, restName, restItems)
    If Not IsList(v) Then VLA_Messages.RaiseMsg "vla-cdr-expected-list", "value", CStr(v)
    Dim c As Object
    Set c = v
    Set EvalCdr = ListTail(c, 1, "cdr")
End Function

Private Function EvalCddr(lst As Object, bindings As Collection, ByVal restName As String, restItems As Collection) As Object
    If lst.Count <> 2 Then VLA_Messages.RaiseMsg "vla-cddr-arity", "n", (lst.Count - 1)
    Dim v As Variant
    AssignVar v, ResolveListopsArg(Nth(lst, 2), bindings, restName, restItems)
    If Not IsList(v) Then VLA_Messages.RaiseMsg "vla-cddr-expected-list", "value", CStr(v)
    Dim c As Object
    Set c = v
    Set EvalCddr = ListTail(c, 2, "cddr")
End Function

' (cons item lst) - lst must resolve to a list; returns a NEW Collection
' of item followed by lst's own items. TagLine ..., 0, same reason as
' ListTail above. PNTH.0: still a full copy, deliberately NOT switched
' to structural sharing (VlaSlice's own header comment has the full
' reasoning - cons was never the evidenced hot path, and a flat view
' cannot represent "one prepended element ahead of a shared tail"
' without a real cons-cell chain, disproportionate for an unmeasured
' cost). Only widened to accept tailV as a VlaSlice without crashing:
' lst's own second argument may now legitimately be an earlier cdr's
' own result, e.g. (cons x (cdr lst)), a completely ordinary shape.
Private Function EvalCons(lst As Object, bindings As Collection, ByVal restName As String, restItems As Collection) As Collection
    If lst.Count <> 3 Then VLA_Messages.RaiseMsg "vla-cons-arity", "n", (lst.Count - 1)
    Dim itemV As Variant
    AssignVar itemV, ResolveListopsArg(Nth(lst, 2), bindings, restName, restItems)
    Dim tailV As Variant
    AssignVar tailV, ResolveListopsArg(Nth(lst, 3), bindings, restName, restItems)
    If Not IsList(tailV) Then VLA_Messages.RaiseMsg "vla-cons-second-not-list", "value", CStr(tailV)
    Dim tailC As Object
    Set tailC = tailV
    Dim outc As New Collection
    outc.Add itemV
    Dim ti As Long
    For ti = 1 To tailC.Count
        outc.Add tailC.Item(ti)
    Next
    TagLine outc, 0
    Set EvalCons = outc
End Function

' (list a b ...) - 0+ args, each resolved through ResolveListopsArg;
' returns a NEW Collection of the resolved values in order. 0 args ->
' empty list, for symmetry - (quote ()) stays the normal way to write a
' literal empty list. TagLine ..., 0, same reason as ListTail above.
Private Function EvalList(lst As Object, bindings As Collection, ByVal restName As String, restItems As Collection) As Collection
    Dim outc As New Collection
    Dim k As Long
    For k = 2 To lst.Count
        Dim piece As Variant
        AssignVar piece, ResolveListopsArg(Nth(lst, k), bindings, restName, restItems)
        outc.Add piece
    Next
    TagLine outc, 0
    Set EvalList = outc
End Function

Private Function BoolSym(ByVal b As Boolean) As String
    If b Then BoolSym = "true" Else BoolSym = "false"
End Function

' Compares two ATOMS (symbols or string literals) the way eq?/equal?
' both need to at the leaf level: symbols compared via
' VLA_Identity.Fold (SD-8's own identifier-invariance discipline),
' string literals by their exact stored content (never folded - a
' string is data, not an identifier), and a symbol is never equal to a
' string literal even when their text matches, preserving the type
' distinction IsSym/string-literal already enforces everywhere else in
' this module. Caller has already confirmed neither a nor b is a list.
Private Function AtomEqual(a As Variant, b As Variant) As Boolean
    Dim aIsSym As Boolean, bIsSym As Boolean
    aIsSym = IsSym(a)
    bIsSym = IsSym(b)
    If aIsSym <> bIsSym Then Exit Function
    If aIsSym Then
        AtomEqual = (VLA_Identity.Fold(CStr(a)) = VLA_Identity.Fold(CStr(b)))
    Else
        AtomEqual = (StrLitContent(a) = StrLitContent(b))
    End If
End Function

' (null? lst) - true iff lst resolves to a list of Count 0. A non-list
' atom is simply false, never an error - null? is a predicate over any
' well-formed value, not a list-only guard.
Private Function EvalNullQ(lst As Object, bindings As Collection, ByVal restName As String, restItems As Collection) As String
    If lst.Count <> 2 Then VLA_Messages.RaiseMsg "vla-nullq-arity", "n", (lst.Count - 1)
    Dim v As Variant
    AssignVar v, ResolveListopsArg(Nth(lst, 2), bindings, restName, restItems)
    If IsList(v) Then
        Dim c As Object
        Set c = v
        EvalNullQ = BoolSym(c.Count = 0)
    Else
        EvalNullQ = "false"
    End If
End Function

' (eq? a b) - both must resolve to ATOMS (see AtomEqual above); either
' being a list raises. Shallow/atom-only comparison - equal? is the
' structural (list-recursive) generalization below.
Private Function EvalEq(lst As Object, bindings As Collection, ByVal restName As String, restItems As Collection) As String
    If lst.Count <> 3 Then VLA_Messages.RaiseMsg "vla-eq-arity", "n", (lst.Count - 1)
    Dim a As Variant, b As Variant
    AssignVar a, ResolveListopsArg(Nth(lst, 2), bindings, restName, restItems)
    AssignVar b, ResolveListopsArg(Nth(lst, 3), bindings, restName, restItems)
    If IsList(a) Then VLA_Messages.RaiseMsg "vla-eq-expected-atom"
    If IsList(b) Then VLA_Messages.RaiseMsg "vla-eq-expected-atom"
    EvalEq = BoolSym(AtomEqual(a, b))
End Function

' (equal? a b) - any shape; deep structural comparison (DeepEqual).
Private Function EvalEqual(lst As Object, bindings As Collection, ByVal restName As String, restItems As Collection) As String
    If lst.Count <> 3 Then VLA_Messages.RaiseMsg "vla-equal-arity", "n", (lst.Count - 1)
    Dim a As Variant, b As Variant
    AssignVar a, ResolveListopsArg(Nth(lst, 2), bindings, restName, restItems)
    AssignVar b, ResolveListopsArg(Nth(lst, 3), bindings, restName, restItems)
    EvalEqual = BoolSym(DeepEqual(a, b))
End Function

Private Function DeepEqual(a As Variant, b As Variant) As Boolean
    If IsList(a) <> IsList(b) Then Exit Function
    If IsList(a) Then
        ' PNTH.0: Object, not Collection - equal? may compare a cdr/cddr
        ' result (a VlaSlice) against ordinary list data.
        Dim ac As Object, bc As Object
        Set ac = a
        Set bc = b
        If ac.Count <> bc.Count Then Exit Function
        Dim i As Long
        For i = 1 To ac.Count
            If Not DeepEqual(ac.Item(i), bc.Item(i)) Then Exit Function
        Next
        DeepEqual = True
    Else
        DeepEqual = AtomEqual(a, b)
    End If
End Function

' (quote-if test then-form else-form) - the one special form in the
' bundle: only the SELECTED branch's form is ever resolved via
' Substitute, the other is never touched. test is resolved first and
' must come back as exactly the symbol "true" or "false" - anything
' else raises (strict, not heuristic - P.L5's own "quote skip is exact"
' precedent, applied here). This is what lets a self-recursive macro
' actually terminate instead of eagerly evaluating both its base case
' and its recursive call at every step.
Private Function EvalQuoteIf(lst As Object, bindings As Collection, ByVal restName As String, restItems As Collection) As Variant
    If lst.Count <> 4 Then VLA_Messages.RaiseMsg "vla-quote-if-arity", "n", (lst.Count - 1)
    Dim t As Variant
    AssignVar t, ResolveListopsArg(Nth(lst, 2), bindings, restName, restItems)
    ' IsObject checked FIRST, same convention SymText/StrLitContent
    ' already use (VLA.bas ~L1531/~L1556) - CStr on an Object raises its
    ' own unrelated type-mismatch error, which would mask this message.
    If IsObject(t) Then VLA_Messages.RaiseMsg "vla-quote-if-test-is-list"
    If Not IsSym(t) Then VLA_Messages.RaiseMsg "vla-quote-if-test-invalid", "value", CStr(t)
    Dim tf As String
    tf = VLA_Identity.Fold(CStr(t))
    Dim branch As Variant
    If tf = "true" Then
        AssignVar branch, Nth(lst, 3)
    ElseIf tf = "false" Then
        AssignVar branch, Nth(lst, 4)
    Else
        VLA_Messages.RaiseMsg "vla-quote-if-test-invalid", "value", CStr(t)
    End If
    Dim res As Variant
    AssignVar res, Substitute(branch, bindings, restName, restItems)
    If IsObject(res) Then Set EvalQuoteIf = res Else EvalQuoteIf = res
End Function

' LISTOPS-EXPAND: strict, LOCALE-INVARIANT numeric-literal check -
' deliberately NOT VBA's own IsNumeric, which accepts locale-specific
' separators (Application.International/regional Windows settings) and
' would let identical source text fold to a different literal result on
' two machines, exactly what LISTOPS-PURITY's own wall forbids (see
' BETA_REARVIEW.md's own LISTOPS-EXPAND entry - checked against the
' reader, not assumed: every atom in this system, numbers included, is
' stored as a plain string with no separate reader-time number type, so
' there is no existing classification to reuse). Accepts exactly an
' optional leading "-", digits, and an optional "." - nothing else: no
' thousands separators, no exponent notation, no leading "+".
Private Function IsNumericLiteralText(ByVal s As String) As Boolean
    If Len(s) = 0 Then Exit Function
    Dim i As Long
    i = 1
    If Mid$(s, 1, 1) = "-" Then i = 2
    If i > Len(s) Then Exit Function
    Dim sawDigit As Boolean, sawDot As Boolean
    Dim c As String
    Do While i <= Len(s)
        c = Mid$(s, i, 1)
        If c >= "0" And c <= "9" Then
            sawDigit = True
        ElseIf c = "." And Not sawDot Then
            sawDot = True
        Else
            Exit Function
        End If
        i = i + 1
    Loop
    IsNumericLiteralText = sawDigit
End Function

' (+expand a b), (=expand a b), (<>expand a b), (>expand a b),
' (<expand a b), (>=expand a b), (<=expand a b) - LISTOPS-EXPAND: expand-
' time constant folding over LITERAL numeric operands only, the
' arithmetic/comparison counterpart to car/cdr/cons's own structural-
' only construction. Both operands resolved through the same
' ResolveListopsArg every LISTOPS primitive uses, then validated against
' IsNumericLiteralText above (never VBA's own IsNumeric) - a non-numeric
' operand RAISES, never coerces, the same strictness eq?/quote-if
' already hold (no "smoothing"). Val()/Str$() throughout, not CDbl/CStr -
' both documented by Microsoft as locale-invariant (always a period
' decimal separator), the only way arithmetic folding keeps "the same
' source text folds to the same result on every machine" true. No
' TagLine concern, unlike cons/list - these return a SCALAR (a number or
' a boolean symbol), never a new Collection.
' PNTH.0: lst widened Collection -> Object - dispatched from
' EvalListopsPrim with the same Object-typed lst every other Eval*
' below accepts now.
' VLA.3: (-expand a b), (*expand a b), (quotient-expand a b) and
' (mod-expand a b) join the seven under the same gate. The last two are
' Excel's QUOTIENT (truncated toward zero, Fix) and MOD (the divisor's
' sign, a - b * Int(a / b)), over whole numbers only - a fraction refused
' by vla-arith-expand-not-whole, since QUOTIENT(0.3, 0.1) is 2 in binary
' floating point - and a zero divisor refused by
' vla-arith-expand-divide-by-zero. Never VBA's \ or Mod, which round their
' operands to a Long first.
Private Function EvalArithExpand(ByVal op As String, lst As Object, bindings As Collection, ByVal restName As String, restItems As Collection) As String
    If lst.Count <> 3 Then VLA_Messages.RaiseMsg "vla-arith-expand-arity", "op", op, "n", (lst.Count - 1)
    Dim a As Variant, b As Variant
    AssignVar a, ResolveListopsArg(Nth(lst, 2), bindings, restName, restItems)
    AssignVar b, ResolveListopsArg(Nth(lst, 3), bindings, restName, restItems)
    If IsObject(a) Then VLA_Messages.RaiseMsg "vla-arith-expand-not-number-list", "op", op
    If IsObject(b) Then VLA_Messages.RaiseMsg "vla-arith-expand-not-number-list", "op", op
    If Not IsNumericLiteralText(CStr(a)) Then VLA_Messages.RaiseMsg "vla-arith-expand-not-number-value", "op", op, "value", CStr(a)
    If Not IsNumericLiteralText(CStr(b)) Then VLA_Messages.RaiseMsg "vla-arith-expand-not-number-value", "op", op, "value", CStr(b)
    Dim na As Double, nb As Double
    na = Val(CStr(a))
    nb = Val(CStr(b))
    Select Case op
        Case "+expand"
            EvalArithExpand = Trim$(Str$(na + nb))
        Case "=expand"
            EvalArithExpand = BoolSym(na = nb)
        Case "<>expand"
            EvalArithExpand = BoolSym(na <> nb)
        Case ">expand"
            EvalArithExpand = BoolSym(na > nb)
        Case "<expand"
            EvalArithExpand = BoolSym(na < nb)
        Case ">=expand"
            EvalArithExpand = BoolSym(na >= nb)
        Case "<=expand"
            EvalArithExpand = BoolSym(na <= nb)
        Case "-expand"
            EvalArithExpand = Trim$(Str$(na - nb))
        Case "*expand"
            EvalArithExpand = Trim$(Str$(na * nb))
        Case "quotient-expand", "mod-expand"
            If na <> Int(na) Then VLA_Messages.RaiseMsg "vla-arith-expand-not-whole", "op", op, "value", CStr(a)
            If nb <> Int(nb) Then VLA_Messages.RaiseMsg "vla-arith-expand-not-whole", "op", op, "value", CStr(b)
            If nb = 0 Then VLA_Messages.RaiseMsg "vla-arith-expand-divide-by-zero", "op", op
            If op = "quotient-expand" Then
                EvalArithExpand = Trim$(Str$(Fix(na / nb)))
            Else
                EvalArithExpand = Trim$(Str$(na - nb * Int(na / nb)))
            End If
    End Select
End Function

' VLA.3: an atom's text as the object model's primitives read it - a
' string literal's content, or a symbol's (a number's) own text.
Private Function AtomTextExpand(v As Variant) As String
    Dim s As String
    s = CStr(v)
    If Left$(s, 1) = Chr$(34) Then
        AtomTextExpand = Mid$(s, 2)
    Else
        AtomTextExpand = s
    End If
End Function

' VLA.3: (&expand a b ...) - each argument resolved as every LISTOPS
' argument is (ResolveListopsArg) and its text joined: a string literal's
' content, a symbol's or a number's text. The result is always a string
' literal, never a symbol, so a macro writes an address or a formula's
' text with no leading-quote trick; symbol (FuseSymbol) keeps that trick
' for identifiers. At least one argument, as symbol takes.
Private Function EvalConcatExpand(lst As Object, bindings As Collection, ByVal restName As String, restItems As Collection) As String
    If lst.Count < 2 Then VLA_Messages.RaiseMsg "vla-concat-expand-arity"
    Dim joined As String
    Dim k As Long
    Dim piece As Variant
    For k = 2 To lst.Count
        AssignVar piece, ResolveListopsArg(Nth(lst, k), bindings, restName, restItems)
        If IsObject(piece) Then VLA_Messages.RaiseMsg "vla-expand-arg-is-list", "op", "&expand", "n", (k - 1)
        joined = joined & AtomTextExpand(piece)
    Next
    EvalConcatExpand = Chr$(34) & joined
End Function

' VLA.3: (characters-expand text) - the text split into one-character
' string literals in order, a NEW Collection tagged 0 (EvalList's reason).
' A character is a Unicode scalar value: a UTF-16 surrogate pair, a high
' half (D800-DBFF) then a low (DC00-DFFF), is one character, as vla-lang's
' chars() counts it, so an emoji is one item in both implementations and
' half a pair is never written to a UTF-8 file. The empty text gives the
' empty list.
Private Function EvalCharactersExpand(lst As Object, bindings As Collection, ByVal restName As String, restItems As Collection) As Collection
    If lst.Count <> 2 Then VLA_Messages.RaiseMsg "vla-expand-arity", "op", "characters-expand", "want", 1, "n", (lst.Count - 1)
    Dim v As Variant
    AssignVar v, ResolveListopsArg(Nth(lst, 2), bindings, restName, restItems)
    If IsObject(v) Then VLA_Messages.RaiseMsg "vla-expand-arg-is-list", "op", "characters-expand", "n", 1
    Dim s As String
    s = AtomTextExpand(v)
    Dim outc As New Collection
    Dim i As Long, w As Long, w2 As Long
    Dim unit As String
    i = 1
    Do While i <= Len(s)
        unit = Mid$(s, i, 1)
        w = AscW(unit) And &HFFFF&
        If w >= &HD800& And w <= &HDBFF& And i < Len(s) Then
            w2 = AscW(Mid$(s, i + 1, 1)) And &HFFFF&
            If w2 >= &HDC00& And w2 <= &HDFFF& Then unit = Mid$(s, i, 2)
        End If
        outc.Add Chr$(34) & unit
        i = i + Len(unit)
    Loop
    TagLine outc, 0
    Set EvalCharactersExpand = outc
End Function

' VLA.3: (address-expand row column) - a cell's address, relative, as
' Excel's ADDRESS(row, column, 4) writes it: B3. Each argument a number
' literal, read as EvalArithExpand reads one (both resolved, then both
' checked as lists, then both as numbers), then a whole number within
' Excel's bounds: a row 1 to 1048576, a column 1 to 16384. The result a
' string literal, the column's letters VLA_Refers'.
Private Function EvalAddressExpand(lst As Object, bindings As Collection, ByVal restName As String, restItems As Collection) As String
    If lst.Count <> 3 Then VLA_Messages.RaiseMsg "vla-expand-arity", "op", "address-expand", "want", 2, "n", (lst.Count - 1)
    Dim a As Variant, b As Variant
    AssignVar a, ResolveListopsArg(Nth(lst, 2), bindings, restName, restItems)
    AssignVar b, ResolveListopsArg(Nth(lst, 3), bindings, restName, restItems)
    If IsObject(a) Then VLA_Messages.RaiseMsg "vla-arith-expand-not-number-list", "op", "address-expand"
    If IsObject(b) Then VLA_Messages.RaiseMsg "vla-arith-expand-not-number-list", "op", "address-expand"
    If Not IsNumericLiteralText(CStr(a)) Then VLA_Messages.RaiseMsg "vla-arith-expand-not-number-value", "op", "address-expand", "value", CStr(a)
    If Not IsNumericLiteralText(CStr(b)) Then VLA_Messages.RaiseMsg "vla-arith-expand-not-number-value", "op", "address-expand", "value", CStr(b)
    Dim r As Double, c As Double
    r = Val(CStr(a))
    c = Val(CStr(b))
    If r <> Int(r) Or r < 1 Or r > 1048576 Then VLA_Messages.RaiseMsg "vla-address-expand-row", "op", "address-expand", "value", CStr(a)
    If c <> Int(c) Or c < 1 Or c > 16384 Then VLA_Messages.RaiseMsg "vla-address-expand-column", "op", "address-expand", "value", CStr(b)
    EvalAddressExpand = Chr$(34) & VLA_Refers.RefersColumnLetters(CLng(c)) & Trim$(Str$(r))
End Function

' VLA.3: (row-expand address) and (column-expand address) - an address's
' row and column numbers, as Excel's ROW and COLUMN of the reference; and
' (column-expand letters), a column's letters alone, the reverse of
' column-letters-expand. The text is read by ParseAddressText's grammar.
Private Function EvalAddressPartExpand(ByVal op As String, lst As Object, bindings As Collection, ByVal restName As String, restItems As Collection) As String
    If lst.Count <> 2 Then VLA_Messages.RaiseMsg "vla-expand-arity", "op", op, "want", 1, "n", (lst.Count - 1)
    Dim v As Variant
    AssignVar v, ResolveListopsArg(Nth(lst, 2), bindings, restName, restItems)
    If IsObject(v) Then VLA_Messages.RaiseMsg "vla-expand-arg-is-list", "op", op, "n", 1
    Dim s As String
    s = AtomTextExpand(v)
    Dim rowNum As Long, colNum As Long
    If Not ParseAddressText(s, (op = "column-expand"), rowNum, colNum) Then
        VLA_Messages.RaiseMsg "vla-address-expand-not-address", "op", op, "value", s
    End If
    If op = "row-expand" Then
        EvalAddressPartExpand = Trim$(Str$(rowNum))
    Else
        EvalAddressPartExpand = Trim$(Str$(colNum))
    End If
End Function

' VLA.3: one grammar for an address, vla-lang's parse_address rule for
' rule: an optional $, one to three ASCII letters in either case, an
' optional $, then one to seven ASCII digits; the column at most XFD
' (VLA_Refers.RefersColumnNumber), the row 1 to 1048576. With
' lettersAlone, the letters with their optional $ and nothing after them
' are a column too, rowNum 0. Anything else - a sheet's name before it, a
' range, a space, a $ anywhere else - is not an address.
Private Function ParseAddressText(ByVal s As String, ByVal lettersAlone As Boolean, ByRef rowNum As Long, ByRef colNum As Long) As Boolean
    Dim n As Long, i As Long, j As Long, k As Long
    Dim c As String
    n = Len(s)
    i = 1
    If n >= 1 Then
        If Mid$(s, 1, 1) = "$" Then i = 2
    End If
    j = i
    Do While j <= n
        c = Mid$(s, j, 1)
        If (c >= "A" And c <= "Z") Or (c >= "a" And c <= "z") Then
            j = j + 1
        Else
            Exit Do
        End If
    Loop
    If j = i Or j - i > 3 Then Exit Function
    colNum = VLA_Refers.RefersColumnNumber(Mid$(s, i, j - i))
    If colNum = 0 Then Exit Function
    rowNum = 0
    If j > n Then
        ParseAddressText = lettersAlone
        Exit Function
    End If
    If Mid$(s, j, 1) = "$" Then j = j + 1
    k = j
    Do While k <= n
        c = Mid$(s, k, 1)
        If c >= "0" And c <= "9" Then
            k = k + 1
        Else
            Exit Do
        End If
    Loop
    If k <= n Or k = j Or k - j > 7 Then Exit Function
    rowNum = CLng(Mid$(s, j, k - j))
    If rowNum < 1 Or rowNum > 1048576 Then Exit Function
    ParseAddressText = True
End Function

' VLA.3: (column-letters-expand column) - a column's letters as a string
' literal, 1 = A ... 16384 = XFD, VLA_Refers' rule; the reverse of
' column-expand. The column a number literal and a whole number in bounds.
Private Function EvalColumnLettersExpand(lst As Object, bindings As Collection, ByVal restName As String, restItems As Collection) As String
    If lst.Count <> 2 Then VLA_Messages.RaiseMsg "vla-expand-arity", "op", "column-letters-expand", "want", 1, "n", (lst.Count - 1)
    Dim v As Variant
    AssignVar v, ResolveListopsArg(Nth(lst, 2), bindings, restName, restItems)
    If IsObject(v) Then VLA_Messages.RaiseMsg "vla-arith-expand-not-number-list", "op", "column-letters-expand"
    If Not IsNumericLiteralText(CStr(v)) Then VLA_Messages.RaiseMsg "vla-arith-expand-not-number-value", "op", "column-letters-expand", "value", CStr(v)
    Dim c As Double
    c = Val(CStr(v))
    If c <> Int(c) Or c < 1 Or c > 16384 Then VLA_Messages.RaiseMsg "vla-address-expand-column", "op", "column-letters-expand", "value", CStr(v)
    EvalColumnLettersExpand = Chr$(34) & VLA_Refers.RefersColumnLetters(CLng(c))
End Function

' VLA.5: (assoc-expand head entries) - the first entry of the list whose
' own head is head, compared as equal? compares (DeepEqual: an atom as eq?
' compares it, a symbol by its folded text and a string exactly, never one
' to the other, and a list by its structure), the entry itself; or the
' empty list, a NEW Collection tagged 0, when none is: Scheme's assoc. An
' entry is a list with a head, and one that is not is refused when the
' search reaches it; an entry past the one found is never looked at.
Private Function EvalAssocExpand(lst As Object, bindings As Collection, ByVal restName As String, restItems As Collection) As Variant
    If lst.Count <> 3 Then VLA_Messages.RaiseMsg "vla-expand-arity", "op", "assoc-expand", "want", 2, "n", (lst.Count - 1)
    Dim head As Variant
    AssignVar head, ResolveListopsArg(Nth(lst, 2), bindings, restName, restItems)
    Dim listed As Variant
    AssignVar listed, ResolveListopsArg(Nth(lst, 3), bindings, restName, restItems)
    If Not IsList(listed) Then VLA_Messages.RaiseMsg "vla-expand-expected-list", "op", "assoc-expand", "value", CStr(listed)
    Dim entries As Object
    Set entries = listed
    Dim entry As Object
    Dim i As Long
    For i = 1 To entries.Count
        If Not IsList(entries.Item(i)) Then VLA_Messages.RaiseMsg "vla-assoc-expand-entry", "value", CStr(entries.Item(i))
        Set entry = entries.Item(i)
        If entry.Count = 0 Then VLA_Messages.RaiseMsg "vla-assoc-expand-entry", "value", "()"
        If DeepEqual(head, entry.Item(1)) Then
            Set EvalAssocExpand = entry
            Exit Function
        End If
    Next
    Dim none As New Collection
    TagLine none, 0
    Set EvalAssocExpand = none
End Function

' (cond (test1 form1) (test2 form2) ... (else formN)) - the tenth
' engine primitive, generalizing quote-if's own strict two-branch shape
' to N clauses AND unifying the expand-time and runtime use cases under
' one name (BETA_REARVIEW.md's own COND entry - both halves adjudicated
' and built together, not split into two primitives, after the owner
' caught that a `cond` prelude macro and a `cond` engine primitive can
' never coexist under the same name: DefineMacro refuses shadowing
' outright, so whichever loaded second would silently never fire).
' Each clause tries EXPAND-TIME resolution first, exactly the way
' quote-if's own test does (ResolveListopsArg, the same eager fold
' car/cdr/null?/etc. already get wherever they appear) - a test that
' resolves to the literal symbol true/false is folded immediately, at
' compile time: true returns that clause's form as the WHOLE result,
' right there; false skips to the next clause; the untaken side is
' never touched, same purity property quote-if's own test already
' proves (see TestListops's own "quote-if never resolves the untaken
' branch" pin). UNLIKE quote-if, a test that does NOT resolve to
' true/false is not an error here - it is treated as a genuine RUNTIME
' expression (e.g. (> x 0), x a real variable), and cond defers: the
' raw, unresolved test (whatever ResolveListopsArg already reduced it
' to, which may itself be a partially-folded compound form) is spliced
' verbatim into a native (if test (then form) (else <remaining clauses,
' recursed>)), emitted and evaluated at RUNTIME exactly like a
' hand-written `if` - real Lisp `cond` semantics for the traditional
' use, no special casing needed for it. Every later clause still gets
' its OWN independent chance to fold or defer - deferring on clause 2
' does not force deferring on clause 3, so a later clause that DOES
' fold true gets hoisted directly into the enclosing else slot rather
' than wrapped in another redundant `if`.
' `else` is a SYNTACTIC keyword position (checked via IsSym/Fold
' directly, the same way EmitIf's own Select Case already recognizes
' then/elseif/else - NOT a LISTOPS predicate call; eq?/equal? compare
' DATA, else here is syntax, a different question) and must be the
' LAST clause or this raises - a misplaced else silently shadowing
' later clauses is exactly the class of mistake this project's own
' strictness-over-smoothing value catches loudly instead of letting
' ride. No match and no else expands to (begin), a deliberate no-op,
' not an error - mirrors walk-rows's own established base-case
' convention ((quote-if (null? lst) (begin) ...)) and real Lisp cond's
' own unspecified-when-nothing-matches behavior, and falls out of the
' recursion's own base case for free: an empty clause list ((cond)) or
' a clause list where every test folds false both simply reach that
' base case, no separate arity-minimum check needed at all.
' Shared limitation with quote-if, not a new one introduced here: a
' test that is ITSELF a bare bound template parameter (not a literal
' LISTOPS call written directly in the clause) can't be re-resolved
' through that indirection - the same "raw argument, no re-walk"
' property quote-if's own test already has (a bare-symbol reference
' returns its bound value directly, without recursively re-examining
' that value's own internal structure for further LISTOPS recognition)
' - only the whole test position needs to be a literal call for the
' fold to fire, exactly the same rule quote-if's own test already
' lives under. Untested here because it was never previously exercised
' for quote-if either, not because it's new.
' PNTH.0: lst widened Collection -> Object, same reason as
' EvalArithExpand just above.
Private Function EvalCond(lst As Object, bindings As Collection, ByVal restName As String, restItems As Collection) As Variant
    Dim r As Variant
    AssignVar r, EvalCondFrom(lst, 2, bindings, restName, restItems)
    If IsObject(r) Then Set EvalCond = r Else EvalCond = r
End Function

' Internal: cond's own recursion, one clause at a time from startIdx
' onward. Plain VBA recursion (not a chain of ExpandMacros macro
' applications, and not routed through Substitute's own list-rebuild
' either) - deliberately, so the fold-vs-defer decision at each clause
' is made exactly once, in one place, with the full picture (this
' clause's test, and the freedom to recurse into the rest either as
' more folding or as an else-chain) available at every step.
' PNTH.0: lst widened Collection -> Object - recurses on itself
' (startIdx+1) unchanged, and NthList (above) now accepts Object too.
Private Function EvalCondFrom(lst As Object, ByVal startIdx As Long, bindings As Collection, ByVal restName As String, restItems As Collection) As Variant
    If startIdx > lst.Count Then
        Dim emptyBegin As New Collection
        emptyBegin.Add "begin"
        TagLine emptyBegin, 0
        Set EvalCondFrom = emptyBegin
        Exit Function
    End If
    Dim cl As Collection
    Set cl = NthList(lst, startIdx)
    If cl.Count <> 2 Then VLA_Messages.RaiseMsg "vla-cond-clause-shape", "n", cl.Count
    Dim headRaw As Variant
    AssignVar headRaw, Nth(cl, 1)
    If Not IsObject(headRaw) Then
        If IsSym(headRaw) Then
            If VLA_Identity.Fold(CStr(headRaw)) = "else" Then
                If startIdx <> lst.Count Then VLA_Messages.RaiseMsg "vla-cond-else-not-last"
                Dim resE As Variant
                AssignVar resE, Substitute(Nth(cl, 2), bindings, restName, restItems)
                If IsObject(resE) Then Set EvalCondFrom = resE Else EvalCondFrom = resE
                Exit Function
            End If
        End If
    End If

    Dim t As Variant
    AssignVar t, ResolveListopsArg(Nth(cl, 1), bindings, restName, restItems)
    Dim foldable As Boolean, tf As String
    If Not IsObject(t) Then
        If IsSym(t) Then
            tf = VLA_Identity.Fold(CStr(t))
            foldable = (tf = "true" Or tf = "false")
        End If
    End If

    If foldable Then
        If tf = "true" Then
            Dim res As Variant
            AssignVar res, Substitute(Nth(cl, 2), bindings, restName, restItems)
            If IsObject(res) Then Set EvalCondFrom = res Else EvalCondFrom = res
        Else
            Dim skip As Variant
            AssignVar skip, EvalCondFrom(lst, startIdx + 1, bindings, restName, restItems)
            If IsObject(skip) Then Set EvalCondFrom = skip Else EvalCondFrom = skip
        End If
        Exit Function
    End If

    Dim rawForm As Variant
    AssignVar rawForm, Substitute(Nth(cl, 2), bindings, restName, restItems)
    Dim restForm As Variant
    AssignVar restForm, EvalCondFrom(lst, startIdx + 1, bindings, restName, restItems)

    Dim thenC As New Collection
    thenC.Add "then"
    thenC.Add rawForm
    TagLine thenC, 0
    Dim elseC As New Collection
    elseC.Add "else"
    elseC.Add restForm
    TagLine elseC, 0
    Dim ifC As New Collection
    ifC.Add "if"
    ifC.Add t
    ifC.Add thenC
    ifC.Add elseC
    TagLine ifC, 0
    Set EvalCondFrom = ifC
End Function

' =====================================================================
'  Emitter: top-level forms
' =====================================================================

' G12.1: VBA silently IGNORES a module-level declaration placed after
' the first procedure, and every later use of the name then dies at
' Option Explicit's untrappable "Variable not defined" modal -
' discovered live when the include splice put a library's procedures
' above the program's declarations (the owner's screenshot). The
' brief's own hazard list names this shape for hand-written modules;
' this guard names it for GENERATED ones, at transpile, in words,
' where Check can show them.
Private Sub TopoGuard(ByVal what As String)
    If mProcEmitted Then VLA_Messages.RaiseMsg "vla-module-level-decl-after-proc", "what", what
End Sub

Private Function EmitTop(f As Variant) As String
    If Not IsList(f) Then VLA_Messages.RaiseMsg "vla-top-level-not-list", "value", CStr(f)
    Dim lst As Collection
    Set lst = f
    Dim h As String
    h = ResolveHeadAlias(VLA_Identity.Fold(HeadSym(lst)))   ' LX.4
    Dim i As Long, r As String
    Select Case h
        Case "sub"
            mProcEmitted = True                    ' G12.1
            EmitTop = EmitProc(lst, False)
        Case "function"
            mProcEmitted = True                    ' G12.1
            EmitTop = EmitProc(lst, True)
        Case "dim"
            TopoGuard "dim"
            EmitTop = EmitDimCore(lst) & vbCrLf
        Case "const"
            TopoGuard "const"
            EmitTop = EmitConstCore(lst) & vbCrLf
        Case "raw"
            EmitTop = StrLitContent(Nth(lst, 2)) & vbCrLf
        Case "type"
            EmitTop = EmitTypeDef(lst, "")
        Case "enum"
            EmitTop = EmitEnumDef(lst, "")
        Case "public", "private"
            ' C2: visibility as a wrapper form - (private (sub f () ...)),
            ' (private (dim m Long)), (public (const k 5)), and likewise
            ' for type/enum. One form, six wrappables; the default
            ' everywhere else stays Public, exactly as before.
            EmitTop = EmitVisibility(lst, IIf(h = "public", "Public", "Private"))
        Case "begin"
            Dim sbB As String, sbBU As Long
            For i = 2 To lst.Count
                SbAdd sbB, sbBU, EmitTop(Nth(lst, i))
            Next
            EmitTop = SbText(sbB, sbBU)
        Case "deflambda"
            ' L14: Names.Add RUNS - registration is imperative.
            VLA_Messages.RaiseMsg "vla-deflambda-runs"
        Case "include"
            ' P.L7: reachable only when an include shares its line
            ' with other text - the whole-line rule, taught.
            VLA_Messages.RaiseMsg "vla-include-must-stand-alone"
        Case Else
            VLA_Messages.RaiseMsg "vla-unknown-top-level-form", "head", h
    End Select
End Function

Private Function EmitProc(lst As Collection, ByVal isFunc As Boolean, Optional ByVal vis As String = "Public") As String
    Dim rawName As String
    rawName = SymText(Nth(lst, 2))
    ' IN.7: a colon in a procedure name is never something a user (or
    ' any core-grammar path but this one) can write - CheckName's own
    ' refusal list has no ":" in it because English never hands one
    ' to a name; the only source is VLA_English.bas's own handler
    ' declarations ("on:sheet-change", "on:click:<slug>"), built
    ' exactly like a "To ...:" action's (sub ...) EXCEPT for this
    ' character, by design (BuildSub's own comment: "the same
    ' declaration shape a To ...: action already gets"). "on:click:"
    ' names are the one colon shape compiled parity now covers -
    ' EmitStmt's own "make-button" case wires an OnAction straight to
    ' one, by the same 'on:click:' & slug convention, no registry
    ' involved. Every OTHER colon shape (today, only
    ' "on:sheet-change") still has no compiled path - it would ride
    ' straight into "Sub on:sheet_change()" once SymName's own ":"
    ' fold ran on it - not a VLA-level refusal at all, a raw VBA
    ' syntax error, surfaced only when Application.Run tries to
    ' compile the injected module (S4's own "untrappable modal"
    ' territory, EmitCallStmt's own precedent for catching this class
    ' of thing HERE instead). Refused by words, at the one point both
    ' EmitTop's "sub"/"function" cases and EmitVisibility's wrapped
    ' forms already funnel through.
    If InStr(rawName, ":") > 0 Then
        If Left$(VLA_Identity.Fold(rawName), 9) <> "on:click:" Then
            VLA_Messages.RaiseMsg "vla-interpreter-only-handler", "name", rawName
        End If
    End If
    Dim name As String
    name = SymName(rawName)
    Dim params As Collection
    Set params = NthList(lst, 3)

    Dim bodyIdx As Long
    bodyIdx = 4
    Dim retType As String
    retType = "Variant"
    If isFunc Then
        If lst.Count >= 4 Then
            If Not IsObject(lst.Item(4)) Then   ' atom after params...
                ' P.L6: ...that is NOT a string literal = return type.
                ' (Before this pass a string here was silently mangled
                ' into a garbage type name; now it falls through to
                ' the docstring rule below - worded paths only.)
                If Left$(CStr(lst.Item(4)), 1) <> Chr$(34) Then
                    retType = SymName(CStr(lst.Item(4)))
                    bodyIdx = 5
                End If
            End If
        End If
    End If

    ' P.L6 (P-SUBDOC): an optional docstring between signature (and
    ' return type, for functions) and body - L17's rule verbatim: a
    ' string literal here is documentation ONLY when at least one
    ' more form follows; a lone string IS the body, and a bare
    ' string in statement position keeps meeting the emitter's
    ' existing worded refusal, so the rule adds no new silent path.
    ' The doc never reaches EmitBody - invariance by construction,
    ' L17's own clause. Recorded for EVERY procedure, documented or
    ' not (name as written, hyphens intact), so apropos can list a
    ' user's whole program; kind is "sub" or "function".
    Dim pdoc As String
    pdoc = ""
    If lst.Count >= bodyIdx + 1 Then
        Dim dCand As Variant
        AssignVar dCand, Nth(lst, bodyIdx)
        If Not IsList(dCand) Then
            If Left$(CStr(dCand), 1) = Chr$(34) Then
                pdoc = StrLitContent(dCand)
                bodyIdx = bodyIdx + 1
            End If
        End If
    End If
    RecordSubDoc SymText(Nth(lst, 2)), pdoc, IIf(isFunc, "function", "sub")

    ' L18 (TCO for VLA functions): a (return (self args...)) is a
    ' tail call BY CONSTRUCTION - return exits - so every self-
    ' return can rebind the parameters and jump to the top instead
    ' of growing the stack: recursion in the language, a loop in
    ' the metal, O(1) stack where VBA gives no TCO of its own.
    ' Armed ONLY when (a) this is a function, (b) EVERY parameter
    ' is the (byval name type) shape - rebinding a ByRef parameter
    ' would mutate the ORIGINAL caller's variable, a semantics
    ' change, so ByRef functions recurse normally (stated, not
    ' smoothed; correctness over optimization) - and (c) the body
    ' actually contains a self-return (otherwise not a byte of
    ' emission changes: the goldens' own guarantee). The rebind
    ' names (vla_tco_N) are emitter-reserved like vla_step and
    ' vla_fail - emitter privilege, not template law; all argument
    ' expressions evaluate into temps BEFORE any parameter is
    ' reassigned, so cross-referencing calls ((self b (- a 1)))
    ' stay exact. Non-tail self-calls ((return (+ 1 (self ...))))
    ' are not self-returns and keep true recursion - also exact.
    mTcoName = ""
    Set mTcoParams = Nothing
    If isFunc Then
        Dim allByVal As Boolean
        allByVal = True
        Dim pv As Variant
        Dim pvl As Collection
        ' L18.1: HeadSym and Nth take a TYPED Collection ByRef, and a
        ' Variant cannot pass into a ByRef typed parameter - the
        ' compile gate said so ("ByRef argument type mismatch").
        ' Set into a typed local first: the walkers' own Set lst = d
        ' pattern, applied here too (TcoScanForm and EmitReturn
        ' already had it; this block did not).
        For Each pv In params
            If Not IsObject(pv) Then
                allByVal = False
            Else
                Set pvl = pv
                If VLA_Identity.Fold(HeadSym(pvl)) <> "byval" Then allByVal = False
            End If
        Next
        If allByVal And TcoScan(lst, bodyIdx, name) Then
            mTcoName = VLA_Identity.Fold(name)
            Set mTcoParams = New Collection
            For Each pv In params
                Set pvl = pv
                mTcoParams.Add SymName(SymText(Nth(pvl, 2)))
            Next
        End If
    End If

    mFuncName = name
    mInFunction = isFunc

    Dim r As String
    r = vis & " " & IIf(isFunc, "Function", "Sub") & " " & name & "(" & EmitParams(params) & ")"
    If isFunc Then r = r & " As " & retType
    r = r & vbCrLf
    If Len(mTcoName) > 0 Then
        ' L18: the loop head. Temps only when there is more than one
        ' parameter (a single parameter rebinds directly).
        Dim ti As Long
        If mTcoParams.Count > 1 Then
            For ti = 1 To mTcoParams.Count
                r = r & "    Dim vla_tco_" & ti & " As Variant" & vbCrLf
            Next
        End If
        r = r & "vla_tco:" & vbCrLf
    End If
    r = r & EmitBody(lst, bodyIdx, 1)
    r = r & "End " & IIf(isFunc, "Function", "Sub") & vbCrLf
    mTcoName = ""
    Set mTcoParams = Nothing
    EmitProc = r
End Function

' L18: does the body contain a (return (self ...)) anywhere? A
' recursive walk over statement structure; only the return's DIRECT
' operand counts - a self-call nested deeper is not a tail call and
' keeps true recursion.
Private Function TcoScan(lst As Collection, ByVal fromIdx As Long, ByVal selfName As String) As Boolean
    Dim i As Long
    For i = fromIdx To lst.Count
        If TcoScanForm(Nth(lst, i), selfName) Then
            TcoScan = True
            Exit Function
        End If
    Next
End Function

Private Function TcoScanForm(v As Variant, ByVal selfName As String) As Boolean
    If Not IsList(v) Then Exit Function
    Dim lst As Collection
    Set lst = v
    If lst.Count = 0 Then Exit Function
    If Not IsSym(lst.Item(1)) Then Exit Function
    Dim h As String
    h = VLA_Identity.Fold(CStr(lst.Item(1)))
    If h = "quote" Then Exit Function          ' data, never code
    If h = "return" And lst.Count >= 2 Then
        If IsList(Nth(lst, 2)) Then
            Dim inner As Collection
            Set inner = Nth(lst, 2)
            If inner.Count > 0 Then
                If IsSym(inner.Item(1)) Then
                    If VLA_Identity.Fold(SymName(CStr(inner.Item(1)))) = VLA_Identity.Fold(selfName) Then
                        TcoScanForm = True
                        Exit Function
                    End If
                End If
            End If
        End If
    End If
    Dim e As Variant
    For Each e In lst
        If TcoScanForm(e, selfName) Then
            TcoScanForm = True
            Exit Function
        End If
    Next
End Function

' Param specs: name | (name type) | (byval name type) | (byref name type)
'            | (optional name type [default]) | (paramarray name)
Private Function EmitParams(params As Collection) As String
    Dim parts As String
    Dim p As Variant
    Dim one As String
    Dim pl As Collection
    Dim ph As String
    For Each p In params
        If Not IsObject(p) Then
            one = SymName(CStr(p))                    ' bare name = ByRef Variant
        Else
            Set pl = p
            ph = VLA_Identity.Fold(HeadSym(pl))
            Select Case ph
                Case "byval"
                    one = "ByVal " & SymName(SymText(Nth(pl, 2))) & " As " & SymName(SymText(Nth(pl, 3)))
                Case "byref"
                    one = "ByRef " & SymName(SymText(Nth(pl, 2))) & " As " & SymName(SymText(Nth(pl, 3)))
                Case "optional"
                    one = "Optional " & SymName(SymText(Nth(pl, 2))) & " As " & SymName(SymText(Nth(pl, 3)))
                    If pl.Count >= 4 Then one = one & " = " & EmitExpr(Nth(pl, 4))
                Case "paramarray"
                    one = "ParamArray " & SymName(SymText(Nth(pl, 2))) & "() As Variant"
                Case Else                              ' (name type)
                    one = SymName(SymText(Nth(pl, 1))) & " As " & SymName(SymText(Nth(pl, 2)))
            End Select
        End If
        If Len(parts) > 0 Then parts = parts & ", "
        parts = parts & one
    Next
    EmitParams = parts
End Function

' =====================================================================
'  Emitter: statements
' =====================================================================

Private Function EmitBody(lst As Collection, ByVal fromIdx As Long, ByVal ind As Long) As String
    ' V8: builder - the old r = r & EmitStmt(...) copied the whole
    ' body once per statement, quadratic in procedure length. A
    ' 5,000-sentence main lives or dies here.
    Dim i As Long
    Dim sb As String, sbU As Long
    For i = fromIdx To lst.Count
        SbAdd sb, sbU, EmitStmt(Nth(lst, i), ind)
    Next
    EmitBody = SbText(sb, sbU)
End Function

Private Function EmitStmt(s As Variant, ByVal ind As Long) As String
    Dim pad As String
    pad = String$(ind * 4, " ")
    If Not IsList(s) Then VLA_Messages.RaiseMsg "vla-bare-atom-statement", "value", CStr(s)
    Dim lst As Collection
    Set lst = s
    ' C2: the source map. Remember the mapped line (for the emitter's
    ' "near vla line N" on failure) and tag the emitted statement's
    ' first line below - except (begin ...), which only delegates and
    ' whose inner statements tag themselves.
    Dim srcLn As Long
    srcLn = FormLine(s)
    If srcLn > 0 Then mEmitLine = srcLn
    Dim h As String
    h = ResolveHeadAlias(VLA_Identity.Fold(HeadSym(lst)))   ' LX.4
    Dim r As String
    Dim k As String
    Dim i As Long
    Dim parts As String
    Dim piece As String

    Select Case h
        Case "dim"
            r = pad & EmitDimCore(lst) & vbCrLf
        Case "const"
            r = pad & EmitConstCore(lst) & vbCrLf
        Case "set!"
            ' SEC.15: exactly (set! (. obj formula) v) writes through the
            ' runtime's VlaSetFormula, the one formula sink of both
            ' backends (VLA_Interpreter's Case "formula" calls the same
            ' helper): it refuses a formula that reaches outside the
            ' workbook by name before writing, and writes through Formula2
            ' as this arm did directly until SEC.15 - Range.Formula
            ' auto-inserts "@" (implicit intersection) on anything that
            ' could spill, which would silently break every query engine's
            ' own "returns a spilled array" promise (SQL.1's own SD-4
            ' freeze) the moment Frazaro itself wrote the formula, not
            ' just a person's stray Ctrl+Shift+Enter. Scoped to exactly
            ' the "formula" member on a bare (. obj formula), not a generic
            ' rename: a read of the same form, or any other member, is
            ' untouched, matching the interpreter's own equally narrow arm
            ' (AS.8 parity). The helper's name is one the generated code
            ' calls, which check_engine_call_names.ps1 reads here and sets
            ' aside with the runtime's other Vla helpers.
            Dim wroteFormula2 As Boolean
            If ListHeadIs(Nth(lst, 2), ".") Then
                Dim setDotForm As Collection
                Set setDotForm = NthList(lst, 2)
                If setDotForm.Count = 3 Then
                    If VLA_Identity.Fold(SymText(Nth(setDotForm, 3))) = "formula" Then
                        r = pad & "Call VlaSetFormula(" & EmitExpr(Nth(setDotForm, 2)) & ", " & EmitExpr(Nth(lst, 3)) & ")" & vbCrLf
                        wroteFormula2 = True
                    End If
                End If
            End If
            If Not wroteFormula2 Then
                r = pad & EmitExpr(Nth(lst, 2)) & " = " & EmitExpr(Nth(lst, 3)) & vbCrLf
            End If
        Case "obj-set!"
            r = pad & "Set " & EmitExpr(Nth(lst, 2)) & " = " & EmitExpr(Nth(lst, 3)) & vbCrLf
        Case "if"
            r = EmitIf(lst, ind)
        Case "for"
            r = EmitFor(lst, ind)
        Case "for-each"
            r = EmitForEach(lst, ind)
        Case "for-each-row"
            r = EmitForEachRow(lst, ind)
        Case "while"
            r = pad & "Do While " & EmitExpr(Nth(lst, 2)) & vbCrLf & _
                EmitBody(lst, 3, ind + 1) & pad & "Loop" & vbCrLf
        Case "do-until"
            r = pad & "Do Until " & EmitExpr(Nth(lst, 2)) & vbCrLf & _
                EmitBody(lst, 3, ind + 1) & pad & "Loop" & vbCrLf
        Case "select"
            r = EmitSelect(lst, ind)
        Case "with"
            r = pad & "With " & EmitExpr(Nth(lst, 2)) & vbCrLf & _
                EmitBody(lst, 3, ind + 1) & pad & "End With" & vbCrLf
        Case "return"
            r = EmitReturn(lst, pad)
        Case "exit-sub"
            r = pad & "Exit Sub" & vbCrLf
        Case "exit-function"
            r = pad & "Exit Function" & vbCrLf
        Case "exit-for"
            r = pad & "Exit For" & vbCrLf
        Case "exit-do"
            r = pad & "Exit Do" & vbCrLf
        Case "redim"
            r = pad & EmitRedim(lst) & vbCrLf
        Case "type", "enum"
            VLA_Messages.RaiseMsg "vla-type-enum-module-level-only", "head", h
        Case "on-error"
            k = VLA_Identity.Fold(SymText(Nth(lst, 2)))
            If k = "resume-next" Then
                r = pad & "On Error Resume Next" & vbCrLf
            ElseIf k = "goto" Then
                r = pad & "On Error GoTo " & SymName(SymText(Nth(lst, 3))) & vbCrLf
            Else
                VLA_Messages.RaiseMsg "vla-on-error-bad-shape"
            End If
        Case "goto"
            r = pad & "GoTo " & SymName(SymText(Nth(lst, 2))) & vbCrLf
        Case "label"
            r = SymName(SymText(Nth(lst, 2))) & ":" & vbCrLf
        Case "quote"
            ' P.L5: an expression-only form in statement position -
            ' refuse with words where VlaTry can catch and teach,
            ' never emit toward the compile modal (the L8 rider's
            ' pattern).
            VLA_Messages.RaiseMsg "vla-quote-in-statement-position"
        Case "deflambda"
            ' L14: registers an Excel LAMBDA under a workbook Name.
            r = pad & EmitDeflambda(lst) & vbCrLf
        Case "include"
            ' P.L7: an include that survived to the emitter was not
            ' alone on its line - the splicer only recognizes the
            ' whole-line spelling. Refuse with the rule.
            VLA_Messages.RaiseMsg "vla-include-must-stand-alone"
        Case "doc"
            r = EmitDoc(lst, pad)
        Case "resume"
            If lst.Count = 1 Then
                r = pad & "Resume" & vbCrLf
            ElseIf VLA_Identity.Fold(SymText(Nth(lst, 2))) = "next" Then
                r = pad & "Resume Next" & vbCrLf
            Else
                r = pad & "Resume " & SymName(SymText(Nth(lst, 2))) & vbCrLf
            End If
        Case "debug-print"
            parts = ""
            For i = 2 To lst.Count
                piece = EmitExpr(Nth(lst, i))
                If Len(parts) > 0 Then parts = parts & "; "
                parts = parts & piece
            Next
            r = pad & "Debug.Print" & IIf(Len(parts) > 0, " " & parts, "") & vbCrLf
        Case "raw"
            r = pad & StrLitContent(Nth(lst, 2)) & vbCrLf
        Case "begin"
            r = EmitBody(lst, 2, ind)
        Case "call"
            r = pad & EmitCallStmt(lst, 2) & vbCrLf
        Case "."
            ' Member call on a computed object in statement position:
            ' (. (range "B2") clearcontents) -> Call range("B2").clearcontents
            r = pad & "Call " & EmitDotText(lst) & vbCrLf
        Case "at-line"
            ' V4: the three-layer source map's join, a zero-runtime
            ' annotation form: (at-line N stmt ...) emits the wrapped
            ' statements exactly as written, with N - the line in the
            ' originating source, one layer up - riding every map tag
            ' inside as src:N. Dynamic extent by save/restore: begin
            ' bodies and If arms inherit it, an inner at-line
            ' overrides within its own extent, and SEVERAL sibling
            ' forms may share one wrapper (the English Try construct
            ' translates one sentence to a train of forms). On an
            ' emitter error inside the wrapper the restore is
            ' deliberately skipped (the error propagates first), so
            ' the failure message can name the source line it was in.
            If lst.Count < 3 Then VLA_Messages.RaiseMsg "vla-at-line-arity"
            Dim atTxt As String
            atTxt = SymText(Nth(lst, 2))
            If Len(atTxt) = 0 Then VLA_Messages.RaiseMsg "vla-at-line-missing-number"
            If Not atTxt Like String$(Len(atTxt), "#") Then
                VLA_Messages.RaiseMsg "vla-at-line-not-a-number", "value", atTxt
            End If
            Dim atPrev As Long
            atPrev = mAtLine
            mAtLine = CLng(atTxt)
            For i = 3 To lst.Count
                r = r & EmitStmt(Nth(lst, i), ind)
            Next
            mAtLine = atPrev
        Case "gen-row"
            ' LISTOPS-PROVENANCE: a zero-runtime annotation form, same
            ' shape as at-line above but a DIFFERENT kind of provenance -
            ' which row of a GENERATOR's own source table this macro's
            ' body came from, not which line of the calling program.
            ' RegisterVocabMacro (VLA_English.bas) is the only writer:
            ' when a defmacro was itself produced from an (at-row label
            ' ...) vocabulary directive, it wraps the macro's own
            ' template body in (gen-row label ...) before storing it, so
            ' EVERY future call to that macro - wherever, in whatever
            ' program - carries the tag. Deliberately carries NO
            ' ObjPtr/TagLine involvement: label is ordinary DATA inside
            ' the template, copied by Substitute like any other literal,
            ' so it can never suffer the S3.1 heap-address staleness bug
            ' the line-tag system already had to fix once (see mGenRow's
            ' own declaration comment). Dynamic extent by save/restore,
            ' identical to at-line's.
            If lst.Count < 3 Then VLA_Messages.RaiseMsg "vla-gen-row-arity"
            Dim rowLabel As String
            rowLabel = StrLitContent(Nth(lst, 2))
            Dim rowPrev As String
            rowPrev = mGenRow
            mGenRow = rowLabel
            For i = 3 To lst.Count
                r = r & EmitStmt(Nth(lst, i), ind)
            Next
            mGenRow = rowPrev
        Case "then", "else", "elseif", "case", "case-else"
            VLA_Messages.RaiseMsg "vla-clause-outside-parent", "head", h
        Case "make-button"
            ' IN.7 (button-click half), compiled parity: caption and
            ' place come straight off this call (EmitExpr - the place
            ' is "(range {r})", a real Range expression, exactly what
            ' the interpreter's own MakeVlaButton trusts to know its
            ' own sheet). The target handler is 'on:click:' &
            ' ClickHandlerSlug(caption) - EXACTLY VLA_English.bas's
            ' own clickProcName scheme, independently recomputed here
            ' from nothing but the caption already sitting in this
            ' call (ClickHandlerSlug's own comment: this module never
            ' calls up into VLA_English.bas to ask). SubExists checks
            ' whether THIS transpile actually declared that handler -
            ' no match is a legal, decorative button (no OnAction),
            ' not a refusal (VlaDispatchButtonClick's own
            ' "unregistered = no-op" precedent, SubExists's own
            ' comment). The actual Buttons.Add/.Name/.OnAction work
            ' lives in one shared helper Sub (VLA_MAKEBUTTON_SUB,
            ' appended once by VlaTranspile below) rather than inline
            ' here, so its own "delete any same-caption button first"
            ' On Error Resume Next/GoTo 0 pair - VBA has no way to
            ' probe Buttons(name) without one - can never disturb an
            ' outer error handler a stepped/traced sub already has
            ' active (On Error GoTo 0 does not RESTORE a prior
            ' handler, it clears one - a hazard confined to the
            ' helper's own procedure, never main's).
            Dim mbCaption As String
            mbCaption = StrLitContent(Nth(lst, 2))
            Dim mbRaw As String
            mbRaw = "on:click:" & ClickHandlerSlug(mbCaption)
            Dim mbProcLit As String
            If SubExists(mbRaw) Then
                mbProcLit = Chr$(34) & SymName(mbRaw) & Chr$(34)
            Else
                mbProcLit = Chr$(34) & Chr$(34)
            End If
            r = pad & "Call " & VLA_MAKEBUTTON_SUB & "(" & EmitExpr(Nth(lst, 2)) & ", " & _
                EmitExpr(Nth(lst, 3)) & ", " & mbProcLit & ")" & vbCrLf
        Case Else
            ' Default: a procedure call in statement position.
            r = pad & EmitCallStmt(lst, 1) & vbCrLf
    End Select
    ' LISTOPS-PROVENANCE: the gate widens from "srcLn > 0" alone to
    ' "srcLn > 0 Or mGenRow set" - a macro-body statement has srcLn=0
    ' by S3.1's own deliberate design (Substitute's TagLine outc, 0),
    ' so without this OR, a gen-row-wrapped statement would never reach
    ' MapTag at all and the row label would silently never appear.
    ' MapTag itself (see its own srcLine>0 branch) never prints a lying
    ' "vla:0" - it omits the line-number half of the tag entirely when
    ' srcLn<=0, printing only vla-row: in that case.
    If (srcLn > 0 Or Len(mGenRow) > 0) And h <> "begin" And h <> "raw" And h <> "at-line" And h <> "gen-row" And Len(r) > 0 Then r = MapTag(r, srcLn)
    ' (begin only delegates; raw is the user's verbatim VBA - a
    ' trailing comment there could break a line continuation; and
    ' at-line's/gen-row's emitted text IS their inner statements,
    ' already tagged - tagging the wrapper would inject a duplicate
    ' into their first line. V4.)
    EmitStmt = r
End Function

Private Function EmitCallStmt(lst As Collection, ByVal nameIdx As Long) As String
    ' L8 rider (owner-found at the L11 smokes): an operator can never
    ' head a STATEMENT. Without this guard, (+ 1 2) in statement
    ' position emitted "Call +(1, 2)" - VBA's untrappable compile
    ' modal during Application.Run. Refuse at TRANSPILE time instead,
    ' with words and directions, where VlaTry can catch and teach.
    Dim opHd As String
    opHd = VLA_Identity.Fold(SymText(Nth(lst, nameIdx)))
    Select Case opHd
        Case "+", "-", "*", "&", "/", "\", "=", "<>", "<", ">", "<=", ">=", _
             "and", "or", "not", "xor", "mod", "is", "like", "imp", "eqv"
            VLA_Messages.RaiseMsg "vla-operator-in-statement-position", "op", opHd
    End Select
    ' P.L4 guard: a keyword can never head a form (statement side).
    If IsKeywordArg(opHd) Then VLA_Messages.RaiseMsg "vla-keyword-misuse", "msg", KwMisuseMsg(opHd, "the head of a statement")
    RefuseInterpreterOnlyCall opHd
    Dim name As String
    name = SymName(SymText(Nth(lst, nameIdx)))
    Dim args As String
    args = EmitArgs(lst, nameIdx + 1)
    If Len(args) = 0 Then
        EmitCallStmt = "Call " & name
    Else
        EmitCallStmt = "Call " & name & "(" & args & ")"
    End If
End Function

Private Function EmitReturn(lst As Collection, ByVal pad As String) As String
    If lst.Count >= 2 Then
        If Not mInFunction Then VLA_Messages.RaiseMsg "vla-return-outside-function"
        ' L18: a self-return under an armed TCO rebinds and jumps
        ' instead of calling - see EmitProc's comment for the law.
        If Len(mTcoName) > 0 And IsList(Nth(lst, 2)) Then
            Dim inner As Collection
            Set inner = Nth(lst, 2)
            If inner.Count > 0 Then
                If IsSym(inner.Item(1)) Then
                    If VLA_Identity.Fold(SymName(CStr(inner.Item(1)))) = mTcoName Then
                        If inner.Count - 1 <> mTcoParams.Count Then VLA_Messages.RaiseMsg "vla-tco-arity", "name", CStr(inner.Item(1)), "given", (inner.Count - 1), "declared", mTcoParams.Count
                        Dim r As String
                        Dim i As Long
                        If mTcoParams.Count = 1 Then
                            r = pad & mTcoParams.Item(1) & " = " & EmitExpr(Nth(inner, 2)) & vbCrLf
                        Else
                            ' Evaluate every argument BEFORE any
                            ' parameter changes - exactness under
                            ' cross-reference.
                            For i = 1 To mTcoParams.Count
                                r = r & pad & "vla_tco_" & i & " = " & EmitExpr(Nth(inner, i + 1)) & vbCrLf
                            Next
                            For i = 1 To mTcoParams.Count
                                r = r & pad & mTcoParams.Item(i) & " = vla_tco_" & i & vbCrLf
                            Next
                        End If
                        EmitReturn = r & pad & "GoTo vla_tco" & vbCrLf
                        Exit Function
                    End If
                End If
            End If
        End If
        EmitReturn = pad & mFuncName & " = " & EmitExpr(Nth(lst, 2)) & vbCrLf & _
                     pad & "Exit Function" & vbCrLf
    Else
        EmitReturn = pad & IIf(mInFunction, "Exit Function", "Exit Sub") & vbCrLf
    End If
End Function

' (if test (then s...) [(elseif test s...)]* [(else s...)])
Private Function EmitIf(lst As Collection, ByVal ind As Long) As String
    Dim pad As String
    pad = String$(ind * 4, " ")
    Dim r As String
    r = pad & "If " & EmitExpr(Nth(lst, 2)) & " Then" & vbCrLf
    Dim i As Long
    Dim cl As Collection
    Dim ch As String
    For i = 3 To lst.Count
        Set cl = NthList(lst, i)
        ch = VLA_Identity.Fold(HeadSym(cl))
        Select Case ch
            Case "then"
                r = r & EmitBody(cl, 2, ind + 1)
            Case "elseif"
                r = r & pad & "ElseIf " & EmitExpr(Nth(cl, 2)) & " Then" & vbCrLf & EmitBody(cl, 3, ind + 1)
            Case "else"
                r = r & pad & "Else" & vbCrLf & EmitBody(cl, 2, ind + 1)
            Case Else
                VLA_Messages.RaiseMsg "vla-if-bad-clause", "head", ch
        End Select
    Next
    r = r & pad & "End If" & vbCrLf
    EmitIf = r
End Function

' (for (i start end [step]) body...)
Private Function EmitFor(lst As Collection, ByVal ind As Long) As String
    Dim pad As String
    pad = String$(ind * 4, " ")
    Dim hdr As Collection
    Set hdr = NthList(lst, 2)
    Dim v As String
    v = SymName(SymText(Nth(hdr, 1)))
    Dim r As String
    r = pad & "For " & v & " = " & EmitExpr(Nth(hdr, 2)) & " To " & EmitExpr(Nth(hdr, 3))
    If hdr.Count >= 4 Then r = r & " Step " & EmitExpr(Nth(hdr, 4))
    r = r & vbCrLf & EmitBody(lst, 3, ind + 1) & pad & "Next " & v & vbCrLf
    EmitFor = r
End Function

' (for-each (x collection) body...)
Private Function EmitForEach(lst As Collection, ByVal ind As Long) As String
    Dim pad As String
    pad = String$(ind * 4, " ")
    Dim hdr As Collection
    Set hdr = NthList(lst, 2)
    Dim v As String
    v = SymName(SymText(Nth(hdr, 1)))
    Dim r As String
    r = pad & "For Each " & v & " In " & EmitExpr(Nth(hdr, 2)) & vbCrLf & _
        EmitBody(lst, 3, ind + 1) & pad & "Next " & v & vbCrLf
    EmitForEach = r
End Function

' (for-each-row (row rng) body...) - PF.4c: bulk-read rng once, walk
' its rows by index. `row` is not a real emitted variable at all -
' (row i)/(set! (row i) v) inside the body compile directly to
' arr(iVar, i) (EmitExpr's own dispatch, checked ahead of its normal
' Select Case - see mSlabRowVar's own declaration for the full
' reasoning and the real bug this redesign replaced). One bulk write-
' back after the loop.
' All-or-nothing commit falls out of the shape itself, not a special
' case: VlaSlabWrite is the only place anything reaches the live sheet,
' called exactly once, after the loop - a raised error anywhere in the
' body propagates straight past it (nothing committed), and a bare
' exit-for (needing NO special-casing of its own, unlike the first,
' broken draft of this function) only escapes the row For loop, still
' falling through to that same write-back call. Because (row i) writes
' straight into arr with no separate buffer to defer, whatever's been
' mutated up to and including the break's own row is already sitting
' in arr when exit-for fires - "break commits up to and including the
' break" falls out for free, not bolted on.
' Calls VlaSlabRead/VlaSlabWrite unqualified, not VLA_Runtime.-prefixed -
' this text may run inside a standalone export where the same functions
' live in Frazaro_EN_Runtime, not VLA_Runtime (VBA's own project-wide
' unqualified lookup finds whichever is actually present; a qualified
' reference would only ever resolve inside the add-in).
' Hidden bookkeeping variables get a deterministic per-occurrence
' suffix, not a fixed or random name - mSlabCounter's own header has
' the full no-gensym reasoning (a second for-each-row in the same Sub
' needs its OWN array/range/index names, or it would collide with the
' first's).
Private Function EmitForEachRow(lst As Collection, ByVal ind As Long) As String
    mSlabCounter = mSlabCounter + 1
    Dim n As String
    n = CStr(mSlabCounter)
    Dim pad As String
    pad = String$(ind * 4, " ")
    Dim hdr As Collection
    Set hdr = NthList(lst, 2)
    Dim v As String
    v = SymName(SymText(Nth(hdr, 1)))

    Dim rngVar As String
    Dim arrVar As String
    Dim iVar As String
    rngVar = "vlaSlabRange" & n
    arrVar = "vlaSlabArr" & n
    iVar = "vlaSlabI" & n

    Dim r As String
    r = pad & "Dim " & rngVar & " As Range" & vbCrLf
    r = r & pad & "Set " & rngVar & " = " & EmitExpr(Nth(hdr, 2)) & vbCrLf
    r = r & pad & "Dim " & arrVar & " As Variant" & vbCrLf
    r = r & pad & arrVar & " = VlaSlabRead(" & rngVar & ")" & vbCrLf
    r = r & pad & "Dim " & iVar & " As Long" & vbCrLf
    r = r & pad & "For " & iVar & " = LBound(" & arrVar & ", 1) To UBound(" & arrVar & ", 1)" & vbCrLf

    ' Dynamic extent, mAtLine/mGenRow's own precedent (both declared
    ' above, this file): save the OUTER binding (Nothing/empty when
    ' there isn't one) before overwriting it with this occurrence's own,
    ' restore it after the body is emitted - so a for-each-row nested
    ' inside another's body resolves (row i) against ITS OWN arr/iVar
    ' while active, and correctly reverts to the outer one once its own
    ' body emission finishes. A plain for/for-each/while/do-until
    ' nested in between never touches these, so (row i) keeps resolving
    ' correctly straight through any such nesting.
    Dim savedRowVar As String, savedArrVar As String, savedIVar As String
    savedRowVar = mSlabRowVar
    savedArrVar = mSlabArrVar
    savedIVar = mSlabIVar
    mSlabRowVar = VLA_Identity.Fold(v)
    mSlabArrVar = arrVar
    mSlabIVar = iVar

    r = r & EmitBody(lst, 3, ind + 1)

    mSlabRowVar = savedRowVar
    mSlabArrVar = savedArrVar
    mSlabIVar = savedIVar

    r = r & pad & "Next " & iVar & vbCrLf
    r = r & pad & "VlaSlabWrite " & arrVar & ", " & rngVar & vbCrLf
    EmitForEachRow = r
End Function

' (select expr (case (v1 v2 ...) body...) ... (case-else body...))
Private Function EmitSelect(lst As Collection, ByVal ind As Long) As String
    Dim pad As String, pad2 As String
    pad = String$(ind * 4, " ")
    pad2 = String$((ind + 1) * 4, " ")
    Dim r As String
    r = pad & "Select Case " & EmitExpr(Nth(lst, 2)) & vbCrLf
    Dim i As Long, j As Long
    Dim cl As Collection
    Dim ch As String
    Dim vals As Collection
    Dim vparts As String
    For i = 3 To lst.Count
        Set cl = NthList(lst, i)
        ch = VLA_Identity.Fold(HeadSym(cl))
        If ch = "case" Then
            Set vals = NthList(cl, 2)
            vparts = ""
            For j = 1 To vals.Count
                If Len(vparts) > 0 Then vparts = vparts & ", "
                vparts = vparts & EmitExpr(Nth(vals, j))
            Next
            r = r & pad2 & "Case " & vparts & vbCrLf & EmitBody(cl, 3, ind + 2)
        ElseIf ch = "case-else" Then
            r = r & pad2 & "Case Else" & vbCrLf & EmitBody(cl, 2, ind + 2)
        Else
            VLA_Messages.RaiseMsg "vla-select-bad-clause", "head", ch
        End If
    Next
    r = r & pad & "End Select" & vbCrLf
    EmitSelect = r
End Function

' (dim name [type])  ->  Dim name As type   (Variant if omitted)
Private Function EmitDimCore(lst As Collection) As String
    EmitDimCore = "Dim " & DimSpecOf(lst, 2)
End Function

' C2: the "name As Type" half of a declaration, shared by Dim, the
' visibility wrapper (Private m As Long), and Type members. nameIdx
' points at the name: 2 for (dim name ...), 1 for a (name ...) member.
' The type slot is either a type name or an (array ...) spec:
'   (array)                ->  name() As Variant     (dynamic)
'   (array Long)           ->  name() As Long        (dynamic)
'   (array Long 10)        ->  name(10) As Long
'   (array Long (to 1 5))  ->  name(1 To 5) As Long
'   (array Long 3 4)       ->  name(3, 4) As Long    (multi-dim)
' Bounds are expressions; the explicit (to lo hi) marker keeps a pair
' unmistakable from a computed bound.
Private Function DimSpecOf(lst As Collection, ByVal nameIdx As Long) As String
    Dim nm As String
    nm = SymName(SymText(Nth(lst, nameIdx)))
    If lst.Count < nameIdx + 1 Then
        DimSpecOf = nm & " As Variant"
        Exit Function
    End If
    Dim t As Variant
    AssignVar t, Nth(lst, nameIdx + 1)
    If Not IsList(t) Then
        DimSpecOf = nm & " As " & SymName(CStr(t))
        Exit Function
    End If
    Dim al As Collection
    Set al = t
    If VLA_Identity.Fold(HeadSym(al)) <> "array" Then
        VLA_Messages.RaiseMsg "vla-decl-bad-type", "name", nm, "head", HeadSym(al)
    End If
    Dim atype As String
    atype = "Variant"
    If al.Count >= 2 Then atype = SymName(SymText(Nth(al, 2)))
    DimSpecOf = nm & "(" & BoundList(al, 3) & ") As " & atype
End Function

' Bounds for array declarations and ReDim: each is an expression
' (upper bound) or (to lower upper); comma-joined for multi-dim.
Private Function BoundList(lst As Collection, ByVal fromIdx As Long) As String
    Dim i As Long
    Dim r As String
    Dim b As Variant
    Dim piece As String
    For i = fromIdx To lst.Count
        AssignVar b, Nth(lst, i)
        piece = ""
        If IsList(b) Then
            Dim bl As Collection
            Set bl = b
            If VLA_Identity.Fold(HeadSym(bl)) = "to" Then
                If bl.Count <> 3 Then VLA_Messages.RaiseMsg "vla-array-bound-arity"
                piece = EmitExpr(Nth(bl, 2)) & " To " & EmitExpr(Nth(bl, 3))
            Else
                piece = EmitExpr(b)
            End If
        Else
            piece = EmitExpr(b)
        End If
        If Len(r) > 0 Then r = r & ", "
        r = r & piece
    Next
    BoundList = r
End Function

' C2: (redim [preserve] name bound...) -> ReDim [Preserve ]name(bounds)
Private Function EmitRedim(lst As Collection) As String
    Dim i As Long
    i = 2
    Dim pres As String
    If Not IsObject(Nth(lst, 2)) Then
        If VLA_Identity.Fold(CStr(Nth(lst, 2))) = "preserve" Then
            pres = "Preserve "
            i = 3
        End If
    End If
    Dim nm As String
    nm = SymName(SymText(Nth(lst, i)))
    If lst.Count < i + 1 Then VLA_Messages.RaiseMsg "vla-redim-missing-bound", "name", nm
    EmitRedim = "ReDim " & pres & nm & "(" & BoundList(lst, i + 1) & ")"
End Function

' C2: (type Name member...) where member = (name Type) |
' (name (array ...)) | bare name (Variant). Module-level only, like VBA.
Private Function EmitTypeDef(lst As Collection, ByVal vis As String) As String
    Dim r As String
    r = vis & "Type " & SymName(SymText(Nth(lst, 2))) & vbCrLf
    Dim i As Long
    Dim m As Variant
    Dim tm As Collection
    For i = 3 To lst.Count
        AssignVar m, Nth(lst, i)
        If IsList(m) Then
            ' The Variant needs a typed hop before it can ride into
            ' DimSpecOf's ByRef Collection parameter - "ByRef argument
            ' type mismatch" otherwise, caught by the compile gate on
            ' C2's maiden compile. Hazard 11's family, one call-frame
            ' down again; EmitEnumDef below had the hop from birth.
            Set tm = m
            r = r & "    " & DimSpecOf(tm, 1) & vbCrLf
        Else
            r = r & "    " & SymName(CStr(m)) & " As Variant" & vbCrLf
        End If
    Next
    If lst.Count < 3 Then VLA_Messages.RaiseMsg "vla-type-no-members", "name", SymName(SymText(Nth(lst, 2)))
    EmitTypeDef = r & "End Type" & vbCrLf
End Function

' C2: (enum Name member...) where member = name | (name value).
Private Function EmitEnumDef(lst As Collection, ByVal vis As String) As String
    Dim r As String
    r = vis & "Enum " & SymName(SymText(Nth(lst, 2))) & vbCrLf
    Dim i As Long
    Dim m As Variant
    For i = 3 To lst.Count
        AssignVar m, Nth(lst, i)
        If IsList(m) Then
            Dim ml As Collection
            Set ml = m
            r = r & "    " & SymName(SymText(Nth(ml, 1))) & " = " & EmitExpr(Nth(ml, 2)) & vbCrLf
        Else
            r = r & "    " & SymName(CStr(m)) & vbCrLf
        End If
    Next
    If lst.Count < 3 Then VLA_Messages.RaiseMsg "vla-enum-no-members", "name", SymName(SymText(Nth(lst, 2)))
    EmitEnumDef = r & "End Enum" & vbCrLf
End Function

' C2: the visibility wrapper's dispatch.
Private Function EmitVisibility(lst As Collection, ByVal vis As String) As String
    Dim inner As Collection
    Set inner = NthList(lst, 2)
    Dim ih As String
    ih = VLA_Identity.Fold(HeadSym(inner))
    Select Case ih
        Case "sub"
            EmitVisibility = EmitProc(inner, False, vis)
        Case "function"
            EmitVisibility = EmitProc(inner, True, vis)
        Case "dim"
            ' Module-level visibility replaces the word Dim entirely:
            ' Private m As Long, not Private Dim m As Long.
            EmitVisibility = vis & " " & DimSpecOf(inner, 2) & vbCrLf
        Case "const"
            EmitVisibility = vis & " " & EmitConstCore(inner) & vbCrLf
        Case "type"
            EmitVisibility = EmitTypeDef(inner, vis & " ")
        Case "enum"
            EmitVisibility = EmitEnumDef(inner, vis & " ")
        Case Else
            VLA_Messages.RaiseMsg "vla-visibility-wraps-unknown", "vis", VLA_Identity.Fold(vis), "head", ih
    End Select
End Function

' (const name value) | (const name type value)
Private Function EmitConstCore(lst As Collection) As String
    If lst.Count = 3 Then
        EmitConstCore = "Const " & SymName(SymText(Nth(lst, 2))) & " = " & EmitExpr(Nth(lst, 3))
    Else
        EmitConstCore = "Const " & SymName(SymText(Nth(lst, 2))) & " As " & _
                        SymName(SymText(Nth(lst, 3))) & " = " & EmitExpr(Nth(lst, 4))
    End If
End Function

' =====================================================================
'  Emitter: expressions
' =====================================================================

Private Function EmitExpr(x As Variant) As String
    If Not IsList(x) Then
        Dim s As String
        s = CStr(x)
        If Left$(s, 1) = Chr$(34) Then
            EmitExpr = ToVbaString(s)
            Exit Function
        End If
        ' P.L4 guard: a keyword token is only legal inside an argument
        ' list, where EmitArgs consumes it BEFORE any EmitExpr call -
        ' so reaching here means it stands where a value belongs.
        If IsKeywordArg(s) Then VLA_Messages.RaiseMsg "vla-keyword-misuse", "msg", KwMisuseMsg(s, "a value on its own")
        Select Case VLA_Identity.Fold(s)
            Case "true": EmitExpr = "True"
            Case "false": EmitExpr = "False"
            Case "nothing": EmitExpr = "Nothing"
            Case "null": EmitExpr = "Null"
            Case "empty": EmitExpr = "Empty"
            Case Else: EmitExpr = SymName(s)
        End Select
        Exit Function
    End If

    Dim lst As Collection
    Set lst = x
    Dim h As String
    h = HeadSym(lst)
    Dim hl As String
    hl = ResolveHeadAlias(VLA_Identity.Fold(h))   ' LX.4

    ' PF.4c: (row i) inside an active for-each-row's own body - checked
    ' ahead of the ordinary dispatch below (mSlabRowVar's own
    ' declaration, above, has the full reasoning: this substitution IS
    ' the fix for a real bug the first draft of EmitForEachRow shipped
    ' with). h, not hl - a user's own chosen row-variable name is never
    ' subject to ResolveHeadAlias's language-level aliasing.
    If Len(mSlabRowVar) > 0 Then
        If VLA_Identity.Fold(h) = mSlabRowVar And lst.Count = 2 Then
            EmitExpr = mSlabArrVar & "(" & mSlabIVar & ", " & EmitExpr(Nth(lst, 2)) & ")"
            Exit Function
        End If
    End If

    Select Case hl
        Case "+", "*", "&", "/", "\", "=", "<>", "<", ">", "<=", ">=", _
             "and", "or", "xor", "mod", "is", "like", "imp", "eqv"
            EmitExpr = EmitChain(lst, OpDisplay(hl))
        Case "-"
            If lst.Count = 2 Then
                EmitExpr = "(-" & EmitExpr(Nth(lst, 2)) & ")"
            Else
                EmitExpr = EmitChain(lst, "-")
            End If
        Case "not"
            EmitExpr = "(Not " & EmitExpr(Nth(lst, 2)) & ")"
        Case "quote"
            ' P.L5: the data literal (see EmitQuote).
            EmitExpr = EmitQuote(lst)
        Case "include"
            ' P.L7: same rule as the statement side.
            VLA_Messages.RaiseMsg "vla-include-must-stand-alone"
        Case "new"
            EmitExpr = "New " & SymName(SymText(Nth(lst, 2)))
        Case "array"
            ' G6: list-valued slots ({name:cat-list}, VLA_English.bas)
            ' lower to this - a real VBA array literal, the runtime
            ' value a Tier-2 helper (pivot-rows, etc.) iterates over at
            ' Excel-automation time. Deliberately NOT a LISTOPS
            ' primitive: LISTOPS's own `list` operates on expand-time
            ' quote-literal data only and never becomes a runtime value
            ' the emitted program touches (LISTOPS-STDLIB's own "different
            ' family from the runtime accessors" line, the mirror image
            ' of it) - this is the opposite, ordinary Tier-1 runtime
            ' data, same family as `+`/string literals, not the
            ' METAMETAMACRO LINE. Given an explicit case (rather than
            ' relying on Case Else's generic function-call fallback,
            ' which would happen to also work here since VBA is
            ' case-insensitive) so it reads as a real, recognized
            ' primitive - AS.8's own operator-gap finding is exactly the
            ' cost of leaning on an implicit fallthrough for something
            ' that actually needs to be a language primitive.
            EmitExpr = "Array(" & EmitArgs(lst, 2) & ")"
        Case "interpolate"
            ' L-INTERPOLATE: mirrors VLA_Interpreter.bas's own EvalExpr
            ' Case "interpolate" - see this file's own LINTERPOLATE.0
            ' header note for the full design. Given an explicit case
            ' for the same reason "array" (immediately above) has one.
            EmitExpr = EmitInterpolateCall(lst)
        Case "."
            ' Member access on a computed object:
            ' (. (range "B2") font.bold)      -> range("B2").font.bold
            ' (. (range "A1") copy :destination x) -> range("A1").copy(Destination:=x)
            EmitExpr = EmitDotText(lst)
        Case Else
            ' Function call, array index, or default-property access -
            ' all the same syntax in VBA, so all the same syntax here.
            ' P.L4 guard: a keyword can never head a form.
            If IsKeywordArg(h) Then VLA_Messages.RaiseMsg "vla-keyword-misuse", "msg", KwMisuseMsg(h, "the head of an expression")
            RefuseInterpreterOnlyCall VLA_Identity.Fold(h)
            EmitExpr = SymName(h) & "(" & EmitArgs(lst, 2) & ")"
    End Select
End Function

' OPTIMIZE.3 slice 5: a helper only the interpreter can run. The command
' behind "Optimize cell C1." runs the OPTIMIZE engine, which lives in the
' add-in and is never injected beside a compiled program, so a compiled
' call to it could only fail later, at VBA's untrappable compile modal.
' Refused here, at TRANSPILE, in words, and with VLA_ERR_INTERPRETER_ONLY
' - IN.7's own number - so Check's transpile probe takes it as the
' expected refusal it is and the program still Interprets.
Private Sub RefuseInterpreterOnlyCall(ByVal foldedHead As String)
    If foldedHead = "vlaoptimizecell" Then VLA_Messages.RaiseMsg "vla-interpreter-only-command"
End Sub

' (. obj member args...) - shared by expression and statement forms.
' Zero args emit without parens (correct for property reads and most
' method calls); with args, keyword arguments work as everywhere else.
Private Function EmitDotText(lst As Collection) As String
    If lst.Count < 3 Then VLA_Messages.RaiseMsg "vla-dot-needs-object-member"
    Dim obj As String
    obj = EmitExpr(Nth(lst, 2))
    Dim mem As String
    ' P.L4 guard: the member slot names a member; a keyword token
    ' here means the object or member was left out of the form.
    Dim memTok As String
    memTok = SymText(Nth(lst, 3))
    If IsKeywordArg(memTok) Then VLA_Messages.RaiseMsg "vla-keyword-misuse", "msg", KwMisuseMsg(memTok, "a member name after '.'")
    mem = SymName(memTok)
    Dim args As String
    args = EmitArgs(lst, 4)
    If Len(args) > 0 Then
        EmitDotText = obj & "." & mem & "(" & args & ")"
    Else
        EmitDotText = obj & "." & mem
    End If
End Function

' =====================================================================
'  L14 (deflambda): two Lisps shake hands. (deflambda name (params)
'  ["doc"] body-expr) emits ONE auditable VBA statement that
'  registers an Excel LAMBDA under a workbook Name:
'      ThisWorkbook.Names.Add Name:="tax", RefersTo:="=LAMBDA(x, (x*0.2))"
'  The worksheet function itself contains ZERO VBA - sheets call it
'  as =tax(A1) with no macros in the loop. The body is a FORMULA
'  dialect of VLA expressions, and it consumes the bench: macro
'  expansion runs first (so (even? n) reaches the sheet as
'  MOD(n, 2)=0 - the P-layer in a worksheet formula), quote data
'  becomes an array constant ({1,2,3}), the docstring rides P.L6's
'  table (apropos: "(worksheet function)") AND the Name's own
'  Comment, and the stepper narrates the expansion like any form.
'  Formula-dialect facts, stated: chain operators render compact
'  and parenthesized ((a*b)); mod is Excel's MOD(a, b) (two
'  operands - more refuses); and/or/not are the AND/OR/NOT
'  functions; if is the THREE-ARG spelling (if test then-value
'  else-value) -> IF(t, a, b) - the statement-if's (then ...)
'  blocks refuse with that teaching; quote nests refuse (Excel
'  arrays do not nest); keyword tokens refuse (formulas have no
'  named arguments); statement heads refuse by name. Names mangle
'  by SymName (Excel names refuse hyphens exactly as VBA does);
'  the as-written name lives in the doc table. Operational notes:
'  ThisWorkbook is the generated module's own book (host under
'  injection, dev under a scratch); re-registering a Name REPLACES
'  its formula, so reruns are safe; Undo's snapshots do not cover
'  Names (a created worksheet function persists - stated limit);
'  and on Excels without LAMBDA the Name registers but evaluates
'  #NAME? on the sheet (the registration itself cannot tell -
'  stated limit).
' =====================================================================
Private Function EmitDeflambda(lst As Collection) As String
    If lst.Count < 4 Then VLA_Messages.RaiseMsg "vla-deflambda-arity"
    Dim nameTok As String
    nameTok = SymText(Nth(lst, 2))
    Dim params As Collection
    Set params = NthList(lst, 3)
    Dim bodyIdx As Long
    bodyIdx = 4
    Dim ldoc As String
    ldoc = ""
    If lst.Count >= 5 Then
        Dim dCand As Variant
        AssignVar dCand, Nth(lst, 4)
        If Not IsList(dCand) Then
            If Left$(CStr(dCand), 1) = Chr$(34) Then
                ldoc = StrLitContent(dCand)
                bodyIdx = 5
            End If
        End If
    End If
    If lst.Count <> bodyIdx Then VLA_Messages.RaiseMsg "vla-deflambda-body-not-one-formula", "name", nameTok, "n", (lst.Count - bodyIdx + 1)

    Dim plist As String
    Dim e As Variant
    For Each e In params
        If IsList(e) Then VLA_Messages.RaiseMsg "vla-deflambda-params-not-plain", "name", nameTok
        If Len(plist) > 0 Then plist = plist & ", "
        plist = plist & SymName(CStr(e))
    Next

    Dim formula As String
    formula = "=LAMBDA(" & plist & IIf(Len(plist) > 0, ", ", "") & EmitFormula(Nth(lst, bodyIdx)) & ")"

    RecordSubDoc nameTok, ldoc, "lambda"

    Dim addPart As String
    addPart = "ThisWorkbook.Names.Add"
    Dim argPart As String
    argPart = "Name:=""" & SymName(nameTok) & """, RefersTo:=""" & Replace(formula, """", """""") & """"
    If Len(ldoc) > 0 Then
        EmitDeflambda = addPart & "(" & argPart & ").Comment = " & ToVbaString(Chr$(34) & ldoc)
    Else
        EmitDeflambda = addPart & " " & argPart
    End If
End Function

' The formula dialect: VLA expression -> Excel formula text (real
' quote characters; the caller VBA-escapes once at the Add line).
Private Function EmitFormula(v As Variant) As String
    If Not IsList(v) Then
        Dim s As String
        s = CStr(v)
        If Left$(s, 1) = Chr$(34) Then
            EmitFormula = FormulaText(Mid$(s, 2))
            Exit Function
        End If
        If IsKeywordArg(s) Then VLA_Messages.RaiseMsg "vla-formula-no-named-args", "tok", s
        Select Case VLA_Identity.Fold(s)
            Case "true"
                EmitFormula = "TRUE"
            Case "false"
                EmitFormula = "FALSE"
            Case Else
                EmitFormula = SymName(s)
        End Select
        Exit Function
    End If
    Dim lst As Collection
    Set lst = v
    Dim h As String
    h = ResolveHeadAlias(VLA_Identity.Fold(HeadSym(lst)))   ' LX.4
    Dim r As String
    Dim i As Long
    Select Case h
        Case "+", "-", "*", "/", "&", "=", "<>", "<", ">", "<=", ">="
            ' F.7: a lone operand under - is a negation and keeps its
            ' sign, (-x), the shape EmitExpr's own - arm writes; the
            ' loop below writes the operator only between operands, so
            ' it used to drop the sign and write (x).
            If h = "-" And lst.Count = 2 Then
                EmitFormula = "(-" & EmitFormula(Nth(lst, 2)) & ")"
                Exit Function
            End If
            For i = 2 To lst.Count
                If Len(r) > 0 Then r = r & h
                r = r & EmitFormula(Nth(lst, i))
            Next
            EmitFormula = "(" & r & ")"
        Case "mod"
            If lst.Count <> 3 Then VLA_Messages.RaiseMsg "vla-formula-mod-arity"
            EmitFormula = "MOD(" & EmitFormula(Nth(lst, 2)) & ", " & EmitFormula(Nth(lst, 3)) & ")"
        Case "not"
            EmitFormula = "NOT(" & EmitFormula(Nth(lst, 2)) & ")"
        Case "and", "or"
            For i = 2 To lst.Count
                If Len(r) > 0 Then r = r & ", "
                r = r & EmitFormula(Nth(lst, i))
            Next
            EmitFormula = UCase$(h) & "(" & r & ")"
        Case "if"
            If lst.Count >= 3 Then
                If IsList(Nth(lst, 3)) Then
                    Dim thenHead As String
                    thenHead = ""
                    On Error Resume Next
                    thenHead = VLA_Identity.Fold(HeadSym(NthList(lst, 3)))
                    On Error GoTo 0
                    If thenHead = "then" Then VLA_Messages.RaiseMsg "vla-formula-if-then-block"
                End If
            End If
            If lst.Count < 3 Or lst.Count > 4 Then VLA_Messages.RaiseMsg "vla-formula-if-arity"
            r = "IF(" & EmitFormula(Nth(lst, 2)) & ", " & EmitFormula(Nth(lst, 3))
            If lst.Count = 4 Then r = r & ", " & EmitFormula(Nth(lst, 4))
            EmitFormula = r & ")"
        Case "quote"
            If lst.Count <> 2 Then VLA_Messages.RaiseMsg "vla-quote-arity"
            EmitFormula = FormulaQuote(Nth(lst, 2))
        Case "lambda"
            ' L14.1 rider (shipped with L18): INNER lambdas -
            ' REDUCE/SCAN/MAP take function arguments, so
            ' (lambda (params) body-expr) is formula-legal and
            ' emits Excel's own LAMBDA(p1, p2, body). Without this
            ' case the params list would have parsed as a call -
            ' the bricks library is its reason for existing.
            If lst.Count <> 3 Then VLA_Messages.RaiseMsg "vla-formula-lambda-arity"
            Dim lp As Collection
            Set lp = NthList(lst, 2)
            Dim pe As Variant
            For Each pe In lp
                If IsList(pe) Then VLA_Messages.RaiseMsg "vla-formula-lambda-params-not-plain"
                If Len(r) > 0 Then r = r & ", "
                r = r & SymName(CStr(pe))
            Next
            EmitFormula = "LAMBDA(" & r & IIf(Len(r) > 0, ", ", "") & EmitFormula(Nth(lst, 3)) & ")"
        Case "set!", "obj-set!", "dim", "begin", "while", "for", "for-each", "for-each-row", "debug-print", "on-error", "label", "goto", "deflambda", "include"
            VLA_Messages.RaiseMsg "vla-formula-is-statement", "head", h
        Case Else
            ' VLA.4 (F.7): a head the head table declares no formula's is
            ' refused by name, where this arm used to write it as a call,
            ' (new x) as new(x).
            If FormulaHeadRefused(h) Then VLA_Messages.RaiseMsg "vla-formula-head-refused", "head", HeadSym(lst)
            For i = 2 To lst.Count
                If Len(r) > 0 Then r = r & ", "
                r = r & EmitFormula(Nth(lst, i))
            Next
            EmitFormula = SymName(HeadSym(lst)) & "(" & r & ")"
    End Select
End Function

Private Function FormulaQuote(v As Variant) As String
    If IsList(v) Then
        Dim lst As Collection
        Set lst = v
        Dim r As String
        Dim e As Variant
        For Each e In lst
            If IsList(e) Then VLA_Messages.RaiseMsg "vla-formula-array-no-nest"
            If Len(r) > 0 Then r = r & ","
            r = r & FormulaQuote(e)
        Next
        FormulaQuote = "{" & r & "}"
        Exit Function
    End If
    ' P.L5's data classes, formula-flavored: strings stay strings,
    ' numbers stay numbers, every other atom is its text AS WRITTEN
    ' as a string - a bare symbol here is data, never a reference.
    Dim s As String
    s = CStr(v)
    If Left$(s, 1) = Chr$(34) Then
        FormulaQuote = FormulaText(Mid$(s, 2))
    ElseIf IsNumeric(s) Then
        FormulaQuote = s
    Else
        FormulaQuote = FormulaText(s)
    End If
End Function

' F.7: a text in a formula, spelled as Excel reads it - wrapped in quote
' marks, with each quote mark inside the text doubled. The one quoting
' rule for the dialect's three text arms (EmitFormula's string literal,
' FormulaQuote's string literal and its bare atom), where each used to
' wrap the text as it stood and a text holding a quote mark gave Excel a
' formula it could not read. EmitDeflambda then VBA-escapes the whole
' formula once at the Add line, which doubles these again, as it doubles
' every other quote mark in the formula.
Private Function FormulaText(ByVal content As String) As String
    FormulaText = Chr$(34) & Replace(content, Chr$(34), Chr$(34) & Chr$(34)) & Chr$(34)
End Function

' VLA.4 (F.7): whether the head table declares a head no formula's - a
' statement, a declaration, a clause, or one of VBA's own operators and
' forms. EmitFormula's Case Else and the formula syntax both ask it, so a
' form the table refuses is refused in a formula by name.
Private Function FormulaHeadRefused(ByVal foldedHead As String) As Boolean
    Dim htRow As Collection
    Set htRow = VLA_HeadTable.VlaHeadTableRow(foldedHead)
    If htRow Is Nothing Then Exit Function
    FormulaHeadRefused = (CStr(htRow.Item(VLA_HeadTable.HT_FORMULA)) <> "yes")
End Function

' VLA.4: the form a formula's head names, an alias taken to it by the head
' table itself. ResolveHeadAlias reads mHeadAliases, which only a compile
' sets, and formula-expand runs in the expander, which sets none: in a
' session no compile had run in, or once a popped context put it back to
' Nothing, (si ...) was written SI(...) where IF(...) is meant - the
' owner's pass found it, VlaWriteExpandGolden against the port's golden.
' The table's own lookup holds no state, as vla-lang's resolve_head_alias
' holds none, so the formula syntax writes the same text in every session.
Private Function FsCanonicalHead(ByVal foldedHead As String) As String
    FsCanonicalHead = foldedHead
    Dim htRow As Collection
    Set htRow = VLA_HeadTable.VlaHeadTableRow(foldedHead)
    If htRow Is Nothing Then Exit Function
    FsCanonicalHead = CStr(htRow.Item(VLA_HeadTable.HT_SYMBOL))
End Function

' =====================================================================
'  VLA.4: THE FORMULA SYNTAX'S WRITER - vla-lang's formula_syntax::write,
'  of which this is the reference (SD-18), held to it by the expand
'  golden through formula-expand (ExpandFormulaCall). An expression the
'  expander has finished with, written as a formula's canonical text in
'  Excel's invariant spelling, its leading = included: no space outside a
'  text, the fewest parentheses Excel's precedence allows, every function
'  name, truth value, error and column letter upper-case, a sheet quoted
'  only where RefersQuoteSheet says it must be, a number and a text as
'  written. Refused by name where an expression is no formula's.
'
'  Not EmitFormula, deflambda's dialect, which parenthesizes every
'  operation and renames a symbol through SymName (A1:B2 becomes the name
'  A1_B2); this one never renames: a symbol is a numeral, a truth value,
'  or a reference or a name RefersScan reads whole, and anything else is
'  refused. The precedence is Excel's, lowest first: the comparisons 1,
'  & 2, + and - 3, * and / 4, ^ 5, a sign 6, % 7, : 8 (KERNEL.8, slice
'  f), anything else 9; every
'  binary operator left-associative, a form of many operands its left
'  fold, as EmitChain's. A left operand is parenthesized when it binds
'  looser than its operator, a right one when looser or equally; a
'  sign's operand when it is any binary operation, since a sign binds
'  tighter than ^ ((- (^ x 2)) is -(x^2), (^ (- x) 2) is -x^2); %'s when
'  it is a sign or a binary operation.
' =====================================================================
Private Function FormulaSyntaxWrite(v As Variant) As String
    FormulaSyntaxWrite = "=" & FsText(v, 1)
End Function

Private Function FsText(v As Variant, ByVal depth As Long) As String
    If depth > VLA_FORM_DEPTH_LIMIT Then VLA_Messages.RaiseMsg "vla-form-too-deep", "limit", VLA_FORM_DEPTH_LIMIT, "loc", ""
    If IsList(v) Then
        FsText = FsList(v, depth)
        Exit Function
    End If
    Dim s As String
    s = CStr(v)
    If Left$(s, 1) = Chr$(34) Then
        s = Mid$(s, 2)
        If Len(s) > 255 Then VLA_Messages.RaiseMsg "vla-formula-text-too-long", "n", Len(s)
        FsText = FormulaText(s)
    Else
        FsText = FsAtom(s)
    End If
End Function

' A form's level of precedence, as the writer spells it.
Private Function FsLevel(v As Variant) As Long
    FsLevel = 9
    If Not IsList(v) Then
        If FsIsNegativeNumeral(CStr(v)) Then FsLevel = 6
        Exit Function
    End If
    Dim lst As Object
    Set lst = v
    If lst.Count = 0 Then Exit Function
    If Not IsSym(lst.Item(1)) Then Exit Function
    Dim h As String
    h = VLA_Identity.Fold(CStr(lst.Item(1)))
    If (h = "+" Or h = "-") And lst.Count = 2 Then
        FsLevel = 6
    ElseIf h = "%" And lst.Count = 2 Then
        FsLevel = 7
    ElseIf lst.Count >= 3 And FsBinaryLevel(h) > 0 Then
        FsLevel = FsBinaryLevel(h)
    End If
End Function

' A binary operator's level, 0 for any other head.
Private Function FsBinaryLevel(ByVal op As String) As Long
    Select Case op
        Case "=", "<>", "<", ">", "<=", ">="
            FsBinaryLevel = 1
        Case "&"
            FsBinaryLevel = 2
        Case "+", "-"
            FsBinaryLevel = 3
        Case "*", "/"
            FsBinaryLevel = 4
        Case "^"
            FsBinaryLevel = 5
        Case ":"
            FsBinaryLevel = 8
    End Select
End Function

' An operand, parenthesized when the level it sits at needs it.
Private Function FsOperand(v As Variant, ByVal atLevel As Long, ByVal isRight As Boolean, ByVal depth As Long) As String
    Dim t As String
    t = FsText(v, depth + 1)
    Dim lv As Long
    lv = FsLevel(v)
    If lv < atLevel Or (isRight And lv = atLevel) Then
        FsOperand = "(" & t & ")"
    Else
        FsOperand = t
    End If
End Function

' A symbol: a numeral, a truth value, a reference or a name, each in its
' canonical spelling; anything else refused.
Private Function FsAtom(ByVal s As String) As String
    If FsIsNumeral(s) Or FsIsNegativeNumeral(s) Then
        If Not FsNumeralFinite(s) Then VLA_Messages.RaiseMsg "vla-formula-numeral-unspelled", "tok", s
        FsAtom = s
        Exit Function
    End If
    Select Case VLA_Identity.Fold(s)
        Case "true"
            FsAtom = "TRUE"
            Exit Function
        Case "false"
            FsAtom = "FALSE"
            Exit Function
    End Select
    Dim fr As FormulaRef
    If FsWholeReference(s, fr) Then
        Select Case fr.Kind
            Case "cell", "range", "column", "row", "3d", "name"
                FsAtom = FsReferenceText(fr)
            Case Else
                VLA_Messages.RaiseMsg "vla-formula-unsupported", "construct", FsConstructOf(fr)
        End Select
        Exit Function
    End If
    If FsVbaNumeral(s) Then VLA_Messages.RaiseMsg "vla-formula-numeral-unspelled", "tok", s
    VLA_Messages.RaiseMsg "vla-formula-atom-unspelled", "tok", s
End Function

' The one record RefersScan reads over the whole of s, into fr; False
' when it reads none, more than one, or one that is not all of s.
Private Function FsWholeReference(ByVal s As String, ByRef fr As FormulaRef) As Boolean
    If Len(s) = 0 Then Exit Function
    If Left$(s, 1) = "=" Then Exit Function
    Dim found() As FormulaRef
    If VLA_Refers.RefersScan(s, found) <> 1 Then Exit Function
    If found(1).Pos <> 1 Then Exit Function
    If found(1).Span <> Len(s) Then Exit Function
    fr = found(1)
    FsWholeReference = True
End Function

' What a construct outside this version is called in a refusal.
Private Function FsConstructOf(ByRef fr As FormulaRef) As String
    Dim what As String
    Select Case fr.Kind
        Case "spill"
            what = "the spill reference"
        Case "structured"
            what = "the structured reference"
        Case "external"
            what = "the reference into another workbook"
        Case "broken"
            what = "the broken reference"
        Case Else
            what = "the reference"
    End Select
    FsConstructOf = what & " " & fr.Written
End Function

' A reference's canonical spelling: its sheet or span quoted only where it
' must be, then its part (FsPartText).
Private Function FsReferenceText(ByRef fr As FormulaRef) As String
    If Len(fr.SheetName) = 0 Then
        FsReferenceText = FsPartText(fr)
    Else
        FsReferenceText = FsSheetText(fr.SheetName) & "!" & FsPartText(fr)
    End If
End Function

' A reference's part, after its sheet, in its canonical spelling: its
' column letters upper-case, its $ marks as written; a name as written.
' VLA.6: apart from FsReferenceText, for range-expand.
Private Function FsPartText(ByRef fr As FormulaRef) As String
    Dim refPart As String
    If fr.Kind = "name" Or Len(fr.PartShape) = 0 Then
        refPart = fr.Part
    Else
        Select Case fr.PartShape
            Case "cell"
                refPart = FsCellText(fr.Col1, fr.ColAbs1, fr.Row1, fr.RowAbs1)
            Case "range"
                refPart = FsCellText(fr.Col1, fr.ColAbs1, fr.Row1, fr.RowAbs1) & ":" & FsCellText(fr.Col2, fr.ColAbs2, fr.Row2, fr.RowAbs2)
            Case "column"
                refPart = FsMark(fr.ColAbs1) & VLA_Refers.RefersColumnLetters(fr.Col1) & ":" & FsMark(fr.ColAbs2) & VLA_Refers.RefersColumnLetters(fr.Col2)
            Case "row"
                refPart = FsMark(fr.RowAbs1) & CStr(fr.Row1) & ":" & FsMark(fr.RowAbs2) & CStr(fr.Row2)
            Case Else
                refPart = fr.Part
        End Select
    End If
    FsPartText = refPart
End Function

Private Function FsCellText(ByVal colNum As Long, ByVal colAbs As Boolean, ByVal rowNum As Long, ByVal rowAbs As Boolean) As String
    FsCellText = FsMark(colAbs) & VLA_Refers.RefersColumnLetters(colNum) & FsMark(rowAbs) & CStr(rowNum)
End Function

Private Function FsMark(ByVal isAbs As Boolean) As String
    If isAbs Then FsMark = "$"
End Function

' A sheet, or a span of sheets, as a qualifier spells it: a span quoted
' whole when either end needs quotes, as Excel writes 'Jan 1:Dec'!A1.
Private Function FsSheetText(ByVal sheetSpelled As String) As String
    Dim colonAt As Long
    colonAt = InStr(1, sheetSpelled, ":", vbBinaryCompare)
    If colonAt = 0 Then
        FsSheetText = VLA_Refers.RefersQuoteSheet(sheetSpelled)
        Exit Function
    End If
    Dim sheetA As String, sheetB As String
    sheetA = Left$(sheetSpelled, colonAt - 1)
    sheetB = Mid$(sheetSpelled, colonAt + 1)
    If VLA_Refers.RefersQuoteSheet(sheetA) = sheetA And VLA_Refers.RefersQuoteSheet(sheetB) = sheetB Then
        FsSheetText = sheetA & ":" & sheetB
    Else
        FsSheetText = "'" & Replace(sheetA, "'", "''") & ":" & Replace(sheetB, "'", "''") & "'"
    End If
End Function

' A list: an operator, one of the syntax's own forms (error, ref, last,
' lambda), or a call; a head the head table refuses is refused by name.
Private Function FsList(v As Variant, ByVal depth As Long) As String
    Dim lst As Object
    Set lst = v
    If lst.Count = 0 Then VLA_Messages.RaiseMsg "vla-formula-head-not-a-function", "head", "()"
    If Not IsSym(lst.Item(1)) Then VLA_Messages.RaiseMsg "vla-formula-head-not-a-function", "head", WriteDatum(lst.Item(1))
    Dim headWritten As String, h As String
    headWritten = CStr(lst.Item(1))
    h = FsCanonicalHead(VLA_Identity.Fold(headWritten))
    Dim n As Long, i As Long, lv As Long
    Dim outText As String
    n = lst.Count - 1
    If (h = "+" Or h = "-") And n = 1 Then
        FsList = h & FsOperand(lst.Item(2), 6, False, depth)
        Exit Function
    End If
    If h = "%" Then
        If n <> 1 Then VLA_Messages.RaiseMsg "vla-formula-arity", "head", headWritten, "want", "one operand", "n", n
        FsList = FsOperand(lst.Item(2), 7, False, depth) & "%"
        Exit Function
    End If
    lv = FsBinaryLevel(h)
    If lv > 0 Then
        If n < 2 Then
            If h = "+" Or h = "-" Then
                VLA_Messages.RaiseMsg "vla-formula-arity", "head", headWritten, "want", "one operand or more", "n", n
            Else
                VLA_Messages.RaiseMsg "vla-formula-arity", "head", headWritten, "want", "two operands or more", "n", n
            End If
        End If
        outText = FsOperand(lst.Item(2), lv, False, depth)
        For i = 3 To lst.Count
            ' #N/A, the one error that ends in neither ! nor ?, is read with
            ' a / after it as one token by vla-lang's parser, so before a /
            ' it is parenthesized: x*(#N/A)/y.
            If h = "/" And Right$(outText, 4) = "#N/A" Then outText = Left$(outText, Len(outText) - 4) & "(#N/A)"
            outText = outText & h & FsOperand(lst.Item(i), lv, True, depth)
        Next
        FsList = outText
        Exit Function
    End If
    Select Case h
        Case "error"
            FsList = FsErrorLiteral(lst, headWritten)
            Exit Function
        Case "ref"
            FsList = FsRefForm(lst, headWritten)
            Exit Function
        Case "last"
            FsList = FsLastForm(lst, headWritten)
            Exit Function
        Case "lambda"
            FsList = FsLambda(lst, headWritten, depth)
            Exit Function
        Case "quote"
            FsList = FsQuote(lst, headWritten)
            Exit Function
    End Select
    If FormulaHeadRefused(h) Then VLA_Messages.RaiseMsg "vla-formula-head-refused", "head", headWritten
    Select Case h
        Case "if"
            If n < 2 Or n > 3 Then VLA_Messages.RaiseMsg "vla-formula-arity", "head", headWritten, "want", "two or three arguments", "n", n
        Case "not"
            If n <> 1 Then VLA_Messages.RaiseMsg "vla-formula-arity", "head", headWritten, "want", "one argument", "n", n
        Case "mod"
            If n <> 2 Then VLA_Messages.RaiseMsg "vla-formula-arity", "head", headWritten, "want", "two arguments", "n", n
    End Select
    Dim fnName As String
    If h = VLA_Identity.Fold(headWritten) Then fnName = headWritten Else fnName = h
    If Not FsIsFunctionName(fnName) Then VLA_Messages.RaiseMsg "vla-formula-head-not-a-function", "head", headWritten
    For i = 2 To lst.Count
        If i > 2 Then outText = outText & ","
        outText = outText & FsText(lst.Item(i), depth + 1)
    Next
    FsList = UCase$(fnName) & "(" & outText & ")"
End Function

' (error "#N/A"): one of the seven errors a formula holds.
Private Function FsErrorLiteral(lst As Object, ByVal headWritten As String) As String
    Dim a As Variant
    Dim canon As String
    If lst.Count = 2 Then
        AssignVar a, lst.Item(2)
        If FsIsStringAtom(a) Then
            canon = FsErrorCanonical(Mid$(CStr(a), 2))
            If Len(canon) = 0 Then VLA_Messages.RaiseMsg "vla-formula-unsupported", "construct", "the error " & Mid$(CStr(a), 2)
            FsErrorLiteral = canon
            Exit Function
        End If
    End If
    VLA_Messages.RaiseMsg "vla-formula-arity", "head", headWritten, "want", "one text, such as ""#N/A""", "n", (lst.Count - 1)
End Function

' Whether a form is an error literal, (error "#N/A"), a value an array
' constant holds (KERNEL.8, slice f).
Private Function FsIsErrorForm(v As Variant) As Boolean
    If Not IsList(v) Then Exit Function
    Dim lst As Object
    Set lst = v
    If lst.Count = 0 Then Exit Function
    If Not IsSym(lst.Item(1)) Then Exit Function
    FsIsErrorForm = (VLA_Identity.Fold(CStr(lst.Item(1))) = "error")
End Function

' Whether a form is a row of an array constant: a list that is no error
' literal.
Private Function FsIsRow(v As Variant) As Boolean
    If IsList(v) Then FsIsRow = Not FsIsErrorForm(v)
End Function

' (quote (1 2 3)), one row, or (quote ((1 2) (3 4))), rows: an array
' constant, {1,2,3} and {1,2;3,4} (KERNEL.8, slice f), each value a
' numeral, a text, a truth value or (error "#N/A"), every row as long as
' the first; anything else refused by name.
Private Function FsQuote(lst As Object, ByVal headWritten As String) As String
    Dim held As Variant
    If lst.Count = 2 Then AssignVar held, lst.Item(2)
    If lst.Count <> 2 Then VLA_Messages.RaiseMsg "vla-formula-arity", "head", headWritten, "want", "one list of values or of rows", "n", (lst.Count - 1)
    If Not IsList(held) Then VLA_Messages.RaiseMsg "vla-formula-arity", "head", headWritten, "want", "one list of values or of rows", "n", (lst.Count - 1)
    Dim heldItems As Object
    Set heldItems = held
    If heldItems.Count = 0 Then VLA_Messages.RaiseMsg "vla-formula-unsupported", "construct", "an empty array constant, (quote ())"
    Dim i As Long, j As Long, nRows As Long
    For i = 1 To heldItems.Count
        If FsIsRow(heldItems.Item(i)) Then nRows = nRows + 1
    Next
    If nRows > 0 And nRows < heldItems.Count Then VLA_Messages.RaiseMsg "vla-formula-unsupported", "construct", "an array constant mixing values and rows"
    ' Its rows: each a list of values, or the one row its values are.
    Dim arrRows As Collection
    Set arrRows = New Collection
    If nRows = 0 Then
        arrRows.Add heldItems
    Else
        For i = 1 To heldItems.Count
            arrRows.Add heldItems.Item(i)
        Next
    End If
    Dim rowWidth As Long
    rowWidth = arrRows.Item(1).Count
    Dim r As Variant
    For Each r In arrRows
        If rowWidth = 0 Or r.Count <> rowWidth Then VLA_Messages.RaiseMsg "vla-formula-unsupported", "construct", "an array constant whose rows differ in length"
    Next
    Dim outText As String, s As String, v As Variant, rowItems As Object, errList As Object
    outText = "{"
    For i = 1 To arrRows.Count
        If i > 1 Then outText = outText & ";"
        Set rowItems = arrRows.Item(i)
        For j = 1 To rowItems.Count
            If j > 1 Then outText = outText & ","
            AssignVar v, rowItems.Item(j)
            If FsIsErrorForm(v) Then
                Set errList = v
                outText = outText & FsErrorLiteral(errList, "error")
            ElseIf IsList(v) Then
                VLA_Messages.RaiseMsg "vla-formula-unsupported", "construct", WriteDatum(v) & " in an array constant"
            Else
                s = CStr(v)
                If Left$(s, 1) = Chr$(34) Then
                    s = Mid$(s, 2)
                    If Len(s) > 255 Then VLA_Messages.RaiseMsg "vla-formula-text-too-long", "n", Len(s)
                    outText = outText & FormulaText(s)
                ElseIf FsIsNumeral(s) Or FsIsNegativeNumeral(s) Or VLA_Identity.Fold(s) = "true" Or VLA_Identity.Fold(s) = "false" Then
                    outText = outText & FsAtom(s)
                Else
                    VLA_Messages.RaiseMsg "vla-formula-unsupported", "construct", WriteDatum(v) & " in an array constant"
                End If
            End If
        Next
    Next
    FsQuote = outText & "}"
End Function

' An error's text in its canonical spelling, "" for none of the seven:
' vla-lang's ErrorKind::parse and text.
Private Function FsErrorCanonical(ByVal errText As String) As String
    Select Case VLA_Identity.Fold(Trim$(errText))
        Case "#null!"
            FsErrorCanonical = "#NULL!"
        Case "#div/0!"
            FsErrorCanonical = "#DIV/0!"
        Case "#value!"
            FsErrorCanonical = "#VALUE!"
        Case "#ref!"
            FsErrorCanonical = "#REF!"
        Case "#name?"
            FsErrorCanonical = "#NAME?"
        Case "#num!"
            FsErrorCanonical = "#NUM!"
        Case "#n/a"
            FsErrorCanonical = "#N/A"
    End Select
End Function

' (ref "Sheet" "A1") or (ref "A1"): one reference or one name, built from
' a sheet and an address, the sheet quoted as it must be.
Private Function FsRefForm(lst As Object, ByVal headWritten As String) As String
    Dim a1 As Variant, a2 As Variant
    Dim built As String, isBuilt As Boolean
    If lst.Count = 2 Then
        AssignVar a1, lst.Item(2)
        If FsIsStringAtom(a1) Then
            built = Mid$(CStr(a1), 2)
            isBuilt = True
        End If
    ElseIf lst.Count = 3 Then
        AssignVar a1, lst.Item(2)
        AssignVar a2, lst.Item(3)
        If FsIsStringAtom(a1) And FsIsStringAtom(a2) Then
            built = FsSheetText(Mid$(CStr(a1), 2)) & "!" & Mid$(CStr(a2), 2)
            isBuilt = True
        End If
    End If
    If Not isBuilt Then VLA_Messages.RaiseMsg "vla-formula-arity", "head", headWritten, "want", "a sheet and an address as texts, or an address alone", "n", (lst.Count - 1)
    Dim fr As FormulaRef
    If FsWholeReference(built, fr) Then
        Select Case fr.Kind
            Case "cell", "range", "column", "row", "3d", "name"
                FsRefForm = FsReferenceText(fr)
                Exit Function
        End Select
    End If
    VLA_Messages.RaiseMsg "vla-formula-ref-invalid", "text", built
End Function

' VLA.6: (last X), X's twin: the one reference on a sheet X is
' (FsSheetReference), with .last after its sheet's name, quoted by the
' sheet rule, its marks as written - Input.last!B13 for Input!B13. A twin's
' twin is refused, since no sheet a grid may hold ends in .last.
Private Function FsLastForm(lst As Object, ByVal headWritten As String) As String
    If lst.Count <> 2 Then VLA_Messages.RaiseMsg "vla-formula-arity", "head", headWritten, "want", "one reference on a sheet", "n", (lst.Count - 1)
    Dim argForm As Variant
    AssignVar argForm, lst.Item(2)
    Dim fr As FormulaRef
    FsSheetReference argForm, "last", fr
    If FsIsTwinSheet(fr.SheetName) Then VLA_Messages.RaiseMsg "vla-formula-last-twin", "value", WriteDatum(argForm)
    fr.SheetName = fr.SheetName & ".last"
    FsLastForm = FsReferenceText(fr)
End Function

' VLA.6: whether a sheet's name is a twin's, ending in .last without case,
' as vla-lang's twin_base reads it.
Private Function FsIsTwinSheet(ByVal sheetSpelled As String) As Boolean
    If Len(sheetSpelled) <= 5 Then Exit Function
    FsIsTwinSheet = (VLA_Identity.Fold(Right$(sheetSpelled, 5)) = ".last")
End Function

' VLA.6: the text sheet-expand or range-expand answers, the sheet's name or
' the cell or range of the one reference on a sheet a form is.
Private Function FsReferencePartText(v As Variant, ByVal op As String) As String
    Dim fr As FormulaRef
    FsSheetReference v, op, fr
    If op = "sheet-expand" Then
        FsReferencePartText = fr.SheetName
    Else
        FsReferencePartText = FsPartText(fr)
    End If
End Function

' VLA.6: one reference on one sheet, as last, sheet-expand and range-expand
' read it, into fr: a symbol RefersScan reads whole as a cell, a range, a
' column or a row behind its sheet, or (ref sheet address), whose own
' refusals stand (FsRefForm). Refused by name: a reference with no sheet,
' since a formula does not know the sheet it will stand on when it is
' written; and a name, a span of sheets, a text or any other form.
Private Sub FsSheetReference(v As Variant, ByVal op As String, ByRef fr As FormulaRef)
    Dim spelled As String
    If IsList(v) Then
        If Not ListHeadIs(v, "ref") Then VLA_Messages.RaiseMsg "vla-reference-expected", "op", op, "value", WriteDatum(v), "what", "no reference"
        Dim refList As Object
        Set refList = v
        spelled = FsRefForm(refList, CStr(refList.Item(1)))
    Else
        spelled = CStr(v)
        If Left$(spelled, 1) = Chr$(34) Then VLA_Messages.RaiseMsg "vla-reference-expected", "op", op, "value", WriteDatum(v), "what", "a text"
    End If
    Dim what As String
    what = "no reference"
    If FsWholeReference(spelled, fr) Then
        Select Case fr.Kind
            Case "cell", "range", "column", "row"
                If Len(fr.SheetName) = 0 Then VLA_Messages.RaiseMsg "vla-reference-no-sheet", "op", op, "value", WriteDatum(v)
                Exit Sub
            Case "3d"
                what = "a span of sheets"
            Case "name"
                what = "a name"
        End Select
    End If
    VLA_Messages.RaiseMsg "vla-reference-expected", "op", op, "value", WriteDatum(v), "what", what
End Sub

' (lambda (x y) body): Excel's LAMBDA(x,y,body), each parameter a name.
Private Function FsLambda(lst As Object, ByVal headWritten As String, ByVal depth As Long) As String
    If lst.Count <> 3 Then VLA_Messages.RaiseMsg "vla-formula-arity", "head", headWritten, "want", "a list of parameters and one body", "n", (lst.Count - 1)
    Dim ps As Variant
    AssignVar ps, lst.Item(2)
    If Not IsList(ps) Then VLA_Messages.RaiseMsg "vla-formula-lambda-params", "tok", WriteDatum(ps)
    Dim psList As Object
    Set psList = ps
    Dim outText As String, i As Long
    Dim p As Variant
    Dim fr As FormulaRef
    outText = "LAMBDA("
    For i = 1 To psList.Count
        AssignVar p, psList.Item(i)
        If Not IsSym(p) Then VLA_Messages.RaiseMsg "vla-formula-lambda-params", "tok", WriteDatum(p)
        If Not FsWholeReference(CStr(p), fr) Then VLA_Messages.RaiseMsg "vla-formula-lambda-params", "tok", CStr(p)
        If fr.Kind <> "name" Then VLA_Messages.RaiseMsg "vla-formula-lambda-params", "tok", CStr(p)
        If Len(fr.SheetName) > 0 Then VLA_Messages.RaiseMsg "vla-formula-lambda-params", "tok", CStr(p)
        outText = outText & CStr(p) & ","
    Next
    FsLambda = outText & FsText(lst.Item(3), depth + 1) & ")"
End Function

Private Function FsIsStringAtom(v As Variant) As Boolean
    If IsList(v) Then Exit Function
    FsIsStringAtom = (Left$(CStr(v), 1) = Chr$(34))
End Function

' The formula's numeral, unsigned: digits with an optional fraction, or a
' fraction alone, then an optional exponent with digits.
Private Function FsIsNumeral(ByVal s As String) As Boolean
    Dim n As Long, i As Long, digitCount As Long, expStart As Long
    Dim c As String
    n = Len(s)
    i = 1
    Do While i <= n
        If Not FsIsDigit(Mid$(s, i, 1)) Then Exit Do
        i = i + 1
        digitCount = digitCount + 1
    Loop
    If i <= n Then
        If Mid$(s, i, 1) = "." Then
            i = i + 1
            Do While i <= n
                If Not FsIsDigit(Mid$(s, i, 1)) Then Exit Do
                i = i + 1
                digitCount = digitCount + 1
            Loop
        End If
    End If
    If digitCount = 0 Then Exit Function
    If i <= n Then
        c = Mid$(s, i, 1)
        If c = "e" Or c = "E" Then
            i = i + 1
            If i <= n Then
                c = Mid$(s, i, 1)
                If c = "+" Or c = "-" Then i = i + 1
            End If
            expStart = i
            Do While i <= n
                If Not FsIsDigit(Mid$(s, i, 1)) Then Exit Do
                i = i + 1
            Loop
            If i = expStart Then Exit Function
        End If
    End If
    FsIsNumeral = (i > n)
End Function

Private Function FsIsNegativeNumeral(ByVal s As String) As Boolean
    If Left$(s, 1) <> "-" Then Exit Function
    FsIsNegativeNumeral = FsIsNumeral(Mid$(s, 2))
End Function

' Whether a numeral's value fits a Double; Val reads it without a locale.
Private Function FsNumeralFinite(ByVal s As String) As Boolean
    Dim d As Double
    On Error GoTo tooLarge
    d = Val(s)
    FsNumeralFinite = True
    Exit Function
tooLarge:
    FsNumeralFinite = False
End Function

' A token VBA's IsNumeric reads as a number, read without a locale:
' vla-lang's intrinsics::is_numeric, mirrored, so that "a number a formula
' does not spell" (&HFF, 1d5, +5) is told from any other atom alike in
' both implementations, whatever Windows' regional settings say of a
' comma or a currency sign.
Private Function FsVbaNumeral(ByVal s As String) As Boolean
    Dim n As Long, i As Long, j As Long
    Dim c As String, radixCh As String
    Dim sawDigit As Boolean, sawDot As Boolean
    n = Len(s)
    If n = 0 Then Exit Function
    i = 1
    c = Left$(s, 1)
    If c = "+" Or c = "-" Then i = 2
    If i + 1 <= n Then
        If Mid$(s, i, 1) = "&" Then
            radixCh = UCase$(Mid$(s, i + 1, 1))
            If radixCh <> "H" And radixCh <> "O" Then Exit Function
            If i + 2 > n Then Exit Function
            For j = i + 2 To n
                c = UCase$(Mid$(s, j, 1))
                If radixCh = "H" Then
                    If Not (FsIsDigit(c) Or (c >= "A" And c <= "F")) Then Exit Function
                Else
                    If Not (c >= "0" And c <= "7") Then Exit Function
                End If
            Next
            FsVbaNumeral = True
            Exit Function
        End If
    End If
    Do While i <= n
        c = Mid$(s, i, 1)
        If FsIsDigit(c) Then
            sawDigit = True
        ElseIf c = "." And Not sawDot Then
            sawDot = True
        Else
            Exit Do
        End If
        i = i + 1
    Loop
    If Not sawDigit Then Exit Function
    If i > n Then
        FsVbaNumeral = True
        Exit Function
    End If
    c = Mid$(s, i, 1)
    If c <> "e" And c <> "E" And c <> "d" And c <> "D" Then Exit Function
    i = i + 1
    If i <= n Then
        c = Mid$(s, i, 1)
        If c = "+" Or c = "-" Then i = i + 1
    End If
    If i > n Then Exit Function
    For j = i To n
        If Not FsIsDigit(Mid$(s, j, 1)) Then Exit Function
    Next
    FsVbaNumeral = True
End Function

Private Function FsIsDigit(ByVal c As String) As Boolean
    If Len(c) <> 1 Then Exit Function
    FsIsDigit = (AscW(c) >= 48 And AscW(c) <= 57)
End Function

' A function's name as a formula spells one: an ASCII letter, then
' letters, digits, . and _.
Private Function FsIsFunctionName(ByVal s As String) As Boolean
    Dim i As Long, a As Long
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        a = AscW(Mid$(s, i, 1))
        If (a >= 65 And a <= 90) Or (a >= 97 And a <= 122) Then
            ' a letter, anywhere
        ElseIf i > 1 And ((a >= 48 And a <= 57) Or a = 46 Or a = 95) Then
            ' a digit, . or _ after the first
        Else
            Exit Function
        End If
    Next
    FsIsFunctionName = True
End Function

' P.L5 (P-QUOTE): (quote <datum>) -> one VBA literal expression.
' Data as literal instead of data as construction code: a list
' becomes Array(...), elements mapped by their READER class -
' string literals stay strings, numeric tokens stay numbers, and
' every other atom becomes its own text as a string, AS WRITTEN:
' no SymName mangling (data is not an identifier - (quote
' (hot-pink)) is Array("hot-pink")), no case folding, and keyword
' tokens are data here too. Quote is a FULL data context: the
' P.L4 guards never see its interior (nothing here rides
' EmitExpr/EmitArgs), and macro expansion never enters it (the
' walkers skip quote forms whole; DefineMacro refuses the name).
' Template parameters DO substitute inside a quote in a defmacro
' body - substitution runs when the enclosing macro expands,
' before the skip applies to the result - which is the template
' author's data-parameterization privilege, pinned. Honest limit:
' the literal emits on ONE line, and VBA caps a logical line
' around 1,023 characters - a very large table belongs in cells,
' not a quote. Reader shorthand ' is REFUSED for now (the
' delegated judgment, reasons on the ledger): one character of
' sugar against a new tokenizer surface; (quote ...) is the
' auditable spelling, and ' can arrive later without breaking it.
Private Function EmitQuote(lst As Collection) As String
    If lst.Count <> 2 Then VLA_Messages.RaiseMsg "vla-quote-arity"
    EmitQuote = QuoteDatum(Nth(lst, 2))
End Function

Private Function QuoteDatum(v As Variant) As String
    If IsList(v) Then
        Dim lst As Collection
        Set lst = v
        Dim r As String
        Dim e As Variant
        For Each e In lst
            If Len(r) > 0 Then r = r & ", "
            r = r & QuoteDatum(e)
        Next
        QuoteDatum = "Array(" & r & ")"
        Exit Function
    End If
    Dim s As String
    s = CStr(v)
    If Left$(s, 1) = Chr$(34) Then
        QuoteDatum = ToVbaString(s)
    ElseIf IsNumeric(s) Then
        ' The module's own numeric convention (SymName's test) - a
        ' numeric token passes through as itself.
        QuoteDatum = s
    Else
        ' Symbols cannot contain a double quote (the tokenizer
        ' breaks on it), so plain wrapping is exact.
        QuoteDatum = """" & s & """"
    End If
End Function

' P.L4 (P-KW): keyword arguments -> VBA named arguments. The honest
' scoping finding, recorded on the ledger: the general mechanism was
' ALREADY HERE - :kw value pairs become Named:=arguments below, and
' all three call paths ride it (dot-member calls via EmitDotText,
' expression calls via EmitExpr's generic path, statement calls via
' EmitCallStmt), so the vocabulary's :password / :destination /
' :paste shapes were never hardwired special cases - they are this
' mechanism's users, and nothing needed refactoring. What P.L4 adds
' is the GUARD FAMILY (the L8-rider pattern): a keyword token
' OUTSIDE an argument list used to emit broken VBA silently
' (SymName passes ':' through) and meet the untrappable compile
' modal downstream; it now refuses at transpile time with words and
' directions, at the three sites a stray keyword can land - a value
' on its own, the head of a form, a member name after '.'.
' Recorded boundary: NAME slots (dim / label / goto) keep their raw
' behavior - a keyword there is not guarded, same as any other
' illegal identifier in those slots.
Private Function KwMisuseMsg(ByVal tok As String, ByVal where As String) As String
    KwMisuseMsg = "'" & tok & "' is a keyword argument token - inside a call's argument list it pairs with the value after it ((f " & tok & " x) -> f " & SymName(Mid$(tok, 2)) & ":=x); it cannot be " & where
End Function

' Arguments, with :keyword value pairs becoming Named:=arguments.
Private Function EmitArgs(lst As Collection, ByVal fromIdx As Long) As String
    Dim i As Long
    Dim r As String
    Dim a As Variant
    Dim piece As String
    i = fromIdx
    Do While i <= lst.Count
        AssignVar a, Nth(lst, i)
        If IsKeywordArg(a) Then
            If i + 1 > lst.Count Then VLA_Messages.RaiseMsg "vla-keyword-arg-missing-value", "tok", CStr(a)
            piece = SymName(Mid$(CStr(a), 2)) & ":=" & EmitExpr(Nth(lst, i + 1))
            i = i + 2
        Else
            piece = EmitExpr(a)
            i = i + 1
        End If
        If Len(r) > 0 Then r = r & ", "
        r = r & piece
    Loop
    EmitArgs = r
End Function

' L-INTERPOLATE: EmitExpr's own "interpolate" case. tpl (index 2) is
' required to be a literal string NODE - checked structurally, exactly
' like "new"'s own type-name argument (EmitExpr's "new" case, above) -
' never handed to EmitExpr, since a computed expression has no
' compile-time text this function could scan for hole names.
Private Function EmitInterpolateCall(lst As Collection) As String
    Dim tmplNode As Variant
    AssignVar tmplNode, Nth(lst, 2)
    If IsList(tmplNode) Then
        VLA_Messages.RaiseMsg "vla-interpolate-template-must-be-literal", "got", "a nested form"
    End If
    Dim tmplTok As String
    tmplTok = CStr(tmplNode)
    If Left$(tmplTok, 1) <> Chr$(34) Then
        VLA_Messages.RaiseMsg "vla-interpolate-template-must-be-literal", "got", tmplTok
    End If
    Dim tpl As String
    tpl = Mid$(tmplTok, 2)
    Dim pieces As Collection
    Set pieces = EmitParseInterpolateTemplate(tpl)

    ' Every argument after tpl is a ":key value" pair (EmitArgs's own
    ' ":key value" -> "key:=value" convention, reused by shape only -
    ' this is a bespoke loop, not a call to EmitArgs, since interpolate
    ' owns its own refusals and builds a key->text map instead of
    ' "key:=value" source). Each value's EmitExpr text is computed
    ' exactly ONCE here, then reused for every occurrence of its hole
    ' below - EmitExpr is pure text generation, so repeating the call
    ' would be harmless, but computing it once keeps this loop's shape
    ' identical to VLA_Interpreter.bas's own EvalInterpolateCall, which
    ' has a real reason (side effects) to evaluate only once.
    Dim keyRaw() As String, keyFold() As String
    Dim valText() As String, usedArr() As Boolean
    Dim kCount As Long
    kCount = 0
    Dim i As Long
    i = 3
    Do While i <= lst.Count
        Dim tok As Variant
        AssignVar tok, Nth(lst, i)
        If Not IsKeywordArg(tok) Then
            VLA_Messages.RaiseMsg "vla-interpolate-expected-keyword-arg", "pos", (i - 2)
        End If
        If i + 1 > lst.Count Then
            VLA_Messages.RaiseMsg "vla-keyword-arg-missing-value", "tok", CStr(tok)
        End If
        Dim rawK As String
        rawK = Mid$(CStr(tok), 2)
        Dim foldK As String
        foldK = VLA_Identity.Fold(rawK)
        Dim vTxt As String
        vTxt = EmitExpr(Nth(lst, i + 1))
        Dim slot As Long
        slot = InterpolateKeyIndex(keyFold, kCount, foldK)
        If slot < 0 Then
            ReDim Preserve keyRaw(0 To kCount)
            ReDim Preserve keyFold(0 To kCount)
            ReDim Preserve valText(0 To kCount)
            ReDim Preserve usedArr(0 To kCount)
            keyRaw(kCount) = rawK
            keyFold(kCount) = foldK
            valText(kCount) = vTxt
            usedArr(kCount) = False
            kCount = kCount + 1
        Else
            keyRaw(slot) = rawK
            valText(slot) = vTxt
        End If
        i = i + 2
    Loop

    Dim r As String
    Dim p As Variant
    For Each p In pieces
        Dim ps As String
        ps = CStr(p)
        Dim piece As String
        If Left$(ps, 1) = Chr$(1) Then
            Dim hName As String
            hName = Mid$(ps, 2)
            Dim idx As Long
            idx = InterpolateKeyIndex(keyFold, kCount, VLA_Identity.Fold(hName))
            If idx < 0 Then
                VLA_Messages.RaiseMsg "vla-interpolate-unknown-key", "key", hName
            End If
            usedArr(idx) = True
            piece = valText(idx)
        Else
            piece = ToVbaString(Chr$(34) & ps)
        End If
        If Len(r) > 0 Then
            r = r & " & " & piece
        Else
            r = piece
        End If
    Next p
    If Len(r) = 0 Then r = Chr$(34) & Chr$(34)

    Dim j As Long
    For j = 0 To kCount - 1
        If Not usedArr(j) Then
            VLA_Messages.RaiseMsg "vla-interpolate-unused-argument", "key", keyRaw(j)
        End If
    Next j

    EmitInterpolateCall = r
End Function

' A simple linear scan (BETA_ROADMAP.md's own L-INTERPOLATE scoping
' text: "not a second grammar") over tpl, splitting it into
' literal-text pieces and named-hole pieces - VLA_Interpreter.bas's own
' EvalParseInterpolateTemplate, EXACTLY (P.L5/REPLEVAL.0's own
' precedent for a compile-time/run-time pair that must agree on a
' parse): a hole piece is marked with a leading Chr$(1) so the pieces
' Collection can hold both kinds as plain strings; "{{" escapes to a
' literal "{"; a bare or doubled "}" is always literal, since a "}" is
' only ever ambiguous while a hole is open, and this scanner already
' closes a hole the moment it sees one.
Private Function EmitParseInterpolateTemplate(ByVal tpl As String) As Collection
    Dim outc As New Collection
    Dim n As Long
    n = Len(tpl)
    Dim i As Long
    i = 1
    Dim lit As String
    lit = ""
    Do While i <= n
        Dim c As String
        c = Mid$(tpl, i, 1)
        If c = "{" Then
            If Mid$(tpl, i + 1, 1) = "{" Then
                lit = lit & "{"
                i = i + 2
            Else
                Dim closeAt As Long
                closeAt = InStr(i + 1, tpl, "}")
                If closeAt = 0 Then
                    VLA_Messages.RaiseMsg "vla-interpolate-unclosed-hole"
                End If
                Dim holeName As String
                holeName = Mid$(tpl, i + 1, closeAt - i - 1)
                If Len(holeName) = 0 Then
                    VLA_Messages.RaiseMsg "vla-interpolate-empty-hole-name"
                End If
                If Len(lit) > 0 Then
                    outc.Add lit
                    lit = ""
                End If
                outc.Add Chr$(1) & holeName
                i = closeAt + 1
            End If
        ElseIf c = "}" Then
            If Mid$(tpl, i + 1, 1) = "}" Then
                lit = lit & "}"
                i = i + 2
            Else
                lit = lit & "}"
                i = i + 1
            End If
        Else
            lit = lit & c
            i = i + 1
        End If
    Loop
    If Len(lit) > 0 Then outc.Add lit
    Set EmitParseInterpolateTemplate = outc
End Function

' Linear search - see VLA_Interpreter.bas's own InterpolateKeyIndex
' (its header note explains why this is not a keyed Collection/VlaDict
' lookup: an interpolate call's own keyword-argument count is always
' small).
Private Function InterpolateKeyIndex(arr() As String, ByVal n As Long, ByVal target As String) As Long
    Dim k As Long
    For k = 0 To n - 1
        If arr(k) = target Then
            InterpolateKeyIndex = k
            Exit Function
        End If
    Next k
    InterpolateKeyIndex = -1
End Function

Private Function EmitChain(lst As Collection, ByVal opText As String) As String
    If lst.Count < 3 Then VLA_Messages.RaiseMsg "vla-operator-needs-two-operands", "op", opText
    Dim i As Long
    Dim r As String
    r = "(" & EmitExpr(Nth(lst, 2))
    For i = 3 To lst.Count
        r = r & " " & opText & " " & EmitExpr(Nth(lst, i))
    Next
    EmitChain = r & ")"
End Function

Private Function OpDisplay(ByVal hl As String) As String
    Select Case hl
        Case "and": OpDisplay = "And"
        Case "or": OpDisplay = "Or"
        Case "xor": OpDisplay = "Xor"
        Case "mod": OpDisplay = "Mod"
        Case "is": OpDisplay = "Is"
        Case "like": OpDisplay = "Like"
        Case "imp": OpDisplay = "Imp"
        Case "eqv": OpDisplay = "Eqv"
        Case Else: OpDisplay = hl
    End Select
End Function

' LX.6: one non-ASCII code point -> its closest plain-ASCII letter(s).
' Case labels are Unicode code points (hex), never a literal accented
' character in this file's own text - the .bas import/export round
' trip is not guaranteed UTF-8 (LESSONS.md's string-literal hazard),
' so a hex label is the only way this table cannot be mis-transcribed
' by an encoding mismatch nobody would notice until it silently
' mismapped a letter. Covers Latin-1 Supplement and the common Latin
' Extended-A accented letters (Western/Central European scripts);
' anything else - a different script entirely - has no case here and
' TransliterateToAscii below drops it rather than guessing. "" = no
' mapping.
Private Function TransliterateChar(ByVal code As Long) As String
    Select Case code
        ' Multi-letter expansions, checked before the 1:1 table.
        Case &HC6, &HE6: TransliterateChar = "ae"      ' Æ æ
        Case &H152, &H153: TransliterateChar = "oe"    ' Œ œ
        Case &HDE, &HFE: TransliterateChar = "th"      ' Þ þ
        Case &HDF: TransliterateChar = "ss"            ' ß

        Case &HC0, &HC1, &HC2, &HC3, &HC4, &HC5, &H100, &H102, &H104: TransliterateChar = "A"
        Case &HE0, &HE1, &HE2, &HE3, &HE4, &HE5, &H101, &H103, &H105: TransliterateChar = "a"
        Case &HC7, &H106, &H108, &H10A, &H10C: TransliterateChar = "C"
        Case &HE7, &H107, &H109, &H10B, &H10D: TransliterateChar = "c"
        Case &H10E, &H110: TransliterateChar = "D"
        Case &H10F, &H111: TransliterateChar = "d"
        Case &HC8, &HC9, &HCA, &HCB, &H112, &H114, &H116, &H118, &H11A: TransliterateChar = "E"
        Case &HE8, &HE9, &HEA, &HEB, &H113, &H115, &H117, &H119, &H11B: TransliterateChar = "e"
        Case &H11C, &H11E, &H120, &H122: TransliterateChar = "G"
        Case &H11D, &H11F, &H121, &H123: TransliterateChar = "g"
        Case &H124, &H126: TransliterateChar = "H"
        Case &H125, &H127: TransliterateChar = "h"
        Case &HCC, &HCD, &HCE, &HCF, &H128, &H12A, &H12C, &H12E, &H130: TransliterateChar = "I"
        Case &HEC, &HED, &HEE, &HEF, &H129, &H12B, &H12D, &H12F, &H131: TransliterateChar = "i"
        Case &H134: TransliterateChar = "J"
        Case &H135: TransliterateChar = "j"
        Case &H136: TransliterateChar = "K"
        Case &H137: TransliterateChar = "k"
        Case &H139, &H13B, &H13D, &H13F, &H141: TransliterateChar = "L"
        Case &H13A, &H13C, &H13E, &H140, &H142: TransliterateChar = "l"
        Case &HD1, &H143, &H145, &H147: TransliterateChar = "N"
        Case &HF1, &H144, &H146, &H148: TransliterateChar = "n"
        Case &HD2, &HD3, &HD4, &HD5, &HD6, &HD8, &H14C, &H14E, &H150: TransliterateChar = "O"
        Case &HF2, &HF3, &HF4, &HF5, &HF6, &HF8, &H14D, &H14F, &H151: TransliterateChar = "o"
        Case &H154, &H156, &H158: TransliterateChar = "R"
        Case &H155, &H157, &H159: TransliterateChar = "r"
        Case &H15A, &H15C, &H15E, &H160: TransliterateChar = "S"
        Case &H15B, &H15D, &H15F, &H161: TransliterateChar = "s"
        Case &H162, &H164, &H166: TransliterateChar = "T"
        Case &H163, &H165, &H167: TransliterateChar = "t"
        Case &HD9, &HDA, &HDB, &HDC, &H168, &H16A, &H16C, &H16E, &H170, &H172: TransliterateChar = "U"
        Case &HF9, &HFA, &HFB, &HFC, &H169, &H16B, &H16D, &H16F, &H171, &H173: TransliterateChar = "u"
        Case &H174: TransliterateChar = "W"
        Case &H175: TransliterateChar = "w"
        Case &HDD, &H176, &H178: TransliterateChar = "Y"
        Case &HFD, &HFF, &H177: TransliterateChar = "y"
        Case &H179, &H17B, &H17D: TransliterateChar = "Z"
        Case &H17A, &H17C, &H17E: TransliterateChar = "z"
    End Select
End Function

' Only code points above ASCII (>127) are ever touched - every ASCII
' character (including '.', which SymName's caller relies on passing
' through untouched for member access like worksheet.cells) is
' appended exactly as before. This is deliberately narrower than "is
' this a legal identifier character" - it targets exactly the LX.6
' finding (non-ASCII passed straight into generated VBA source) and
' nothing else SymName already did.
Private Function TransliterateToAscii(ByVal s As String) As String
    Dim n As Long, i As Long, c As Long
    Dim buf As String
    n = Len(s)
    For i = 1 To n
        c = AscW(Mid$(s, i, 1))
        If c >= 0 And c <= 127 Then
            buf = buf & Chr$(c)
        Else
            buf = buf & TransliterateChar(c)   ' "" when no mapping exists - dropped, not guessed at
        End If
    Next i
    TransliterateToAscii = buf
End Function

' Identifier translation: numbers pass through; non-ASCII letters
' transliterate to their closest ASCII equivalent (LX.6 - ratified:
' degrade gracefully rather than pass an accented letter straight into
' generated VBA source, where it is not legal); hyphens become
' underscores (VBA identifiers can't contain '-'); dots are preserved
' so member access like worksheet.cells works unmodified. Refuses if
' transliteration leaves nothing - a name that was entirely a script
' this table has no mapping for (Cyrillic, CJK, ...) needs a name a
' human can still type, not a silently empty one.
Private Function SymName(ByVal s As String) As String
    If IsNumeric(s) Then
        SymName = s
        Exit Function
    End If
    Dim r As String
    ' IN.7 (button-click half): ':' also folds to '_' - the only
    ' source of a colon in a name is VLA_English.bas's own handler
    ' declarations ("on:sheet-change", "on:click:<slug>"; EmitProc's
    ' colon-name guard, above its own call site, is what decides
    ' WHICH of those may reach here at all).
    r = Replace(Replace(TransliterateToAscii(s), "-", "_"), ":", "_")
    If Len(r) = 0 Then
        VLA_Messages.RaiseMsg "vla-name-no-plain-letters", "value", s
    End If
    SymName = r
End Function

' IN.7 (button-click half): twin of the identical helper in
' VLA_English.bas ('ClickHandlerSlug') - each module stays self-
' contained (VLA_English.bas's own CollGet comment is this codebase's
' precedent for that duplication). EmitStmt's "make-button" case
' recomputes 'on:click:' & this SAME slug from nothing but the
' caption in its own call, to find the handler VLA_English.bas
' declared under that exact name - see the other twin's comment for
' why this can't just be SymName (dot-preserving) or a cross-module
' call (Layer 1 cannot call up into Layer 3, REBUILD.md).
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

' IN.7 (button-click half): the compiled module's own button-creation
' helper, VlaTranspile's own conditional append target
' (VLA_MAKEBUTTON_SUB) - a compiled twin of VLA_Interpreter.bas's
' MakeVlaButton, self-contained (ThisWorkbook here means the EXPORTED
' workbook this code physically runs in, correct for a program meant
' to survive Compile/export with the add-in closed - never the add-
' in's own ThisWorkbook the interpreter path wires to). Re-running the
' same caption on the same sheet replaces the old button outright
' (MakeVlaButton's own "current run wins" precedent) rather than
' erroring on Excel's own "duplicate shape name" refusal the SECOND
' time a compiled Sub runs. vlaProc empty means EmitStmt's SubExists
' check found no declared handler for this caption - .OnAction is
' simply never set, a decorative button, not a reference to a Sub
' that does not exist.
Private Function MakeButtonHelperText() As String
    Dim r As String
    r = "Private Sub " & VLA_MAKEBUTTON_SUB & "(ByVal vlaCaption As String, ByVal vlaPlace As Range, ByVal vlaProc As String)" & vbCrLf & _
        "    Dim vlaWs As Worksheet" & vbCrLf & _
        "    Set vlaWs = vlaPlace.Worksheet" & vbCrLf & _
        "    On Error Resume Next" & vbCrLf & _
        "    vlaWs.Buttons(vlaCaption).Delete" & vbCrLf & _
        "    On Error GoTo 0" & vbCrLf & _
        "    Dim vlaBtn As Button" & vbCrLf & _
        "    Set vlaBtn = vlaWs.Buttons.Add(vlaPlace.Left, vlaPlace.Top, 120, 24)" & vbCrLf & _
        "    vlaBtn.Caption = vlaCaption" & vbCrLf & _
        "    vlaBtn.Name = vlaCaption" & vbCrLf & _
        "    If Len(vlaProc) > 0 Then" & vbCrLf & _
        "        vlaBtn.OnAction = ""'"" & ThisWorkbook.Name & ""'!"" & vlaProc" & vbCrLf & _
        "    End If" & vbCrLf & _
        "End Sub" & vbCrLf
    MakeButtonHelperText = r
End Function

Private Function ToVbaString(ByVal tok As String) As String
    Dim c As String
    c = Mid$(tok, 2)
    ToVbaString = """" & Replace(c, """", """""") & """"
End Function
