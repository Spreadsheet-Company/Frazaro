Attribute VB_Name = "VLA_Prolog"
Option Explicit
Public Const VLA_PROLOG_VERSION As String = "PROLOG.9"
' (This constant read "PROLOG.7" until PROLOG.9. PROLOG.8 documented
' itself at each site and added no header block of its own, and the bump
' went with it. Nothing reads this constant - it is a marker for whoever
' opens the module - so the staleness cost nothing; it is corrected here
' rather than left to grow.)
'
' PROLOG.9: the six ISO type-test goals - written `var?`, `nonvar?`,
' `atom?`, `number?`, `atomic?`, `compound?` - and `between/3`. The
' roadmap lists them in one bullet. They are TWO DIFFERENT SHAPES and
' are built as two arms:
'
'   the six type tests - DETERMINISTIC. Exactly one outcome, bind
'       nothing, thread envN/envT into the continuation UNCHANGED. The
'       shape `not` and the six comparisons already have.
'   between/3         - a GENERATOR, and the FIRST one in this dispatch.
'       It binds X to each value in turn and BACKTRACKS, so it recurses
'       into SolveGoalList once per value the way the candidates loop
'       does, not once per goal the way every other arm does.
'
' Lumping them together would have given the type tests machinery they
' do not need and `between` a shape that cannot enumerate.
'
' WHAT `atom`/`number` HAD TO DECIDE, AND WHAT IT DID NOT.
' A TEXT cell becomes Chr$(34) & value (TableCellToTerm, below); a
' NUMERIC cell becomes a bare number. So `(number X)` where X came from
' a cell reading 42, and `(atom X)` where X is `"eng`, are questions
' about whether the quoted-string marker is part of a term's IDENTITY -
' which is exactly what PROLOG.10 exists to adjudicate and has NOT yet
' answered.
'
' THE QUESTION MARK, decided late and on purpose: `(atom bob)` in an
' S-expression language is visually identical to a compound DATA term,
' and this codebase's macro layer (VLA.bas) already reserves null?, eq?
' and equal? alongside car/cdr/cons/list - a question mark on exactly the
' primitives that ask a question. TypeTestKindFor (below) carries the
' full reasoning, and TypeTestIsoSpellingFor (below it) is why writing
' the bare ISO `(atom X)` gets a refusal naming the right spelling
' instead of the silent dead end an unknown predicate would be.
'
' PROLOG.9's LOCAL READING, which does NOT close PROLOG.10: the marker
' is identity-bearing, so a marked leaf is an ATOM, never a number.
' `(atom? "eng)` is True, `(number? "42)` is False, and both are
' `atomic?`.
' The reason is coherence with what already shipped, not a preference
' about strings: PROLOG.8 shipped `"42 \== 42` (pinned by two tests), so
' a `number("42)` answering True would classify a term into a class that
' holds no term it is identical to. Every other engine that could have
' answered was consulted rather than assumed - and they DISAGREE, which
' is the whole reason PROLOG.10 is open:
'
'   EvalArithTerm (below) STRIPS the marker before testing numeric-ness,
'       so `(is X "42")` computes 42 and `(> "42" 41)` succeeds today.
'   UnifyTwoWay/TermsIdentical compare marker-INCLUDED, so `"42` and 42
'       are different terms today.
'
' Arithmetic already says "same", identity already says "different", and
' PROLOG.9 sides with IDENTITY because classification is an identity
' question, not an arithmetic one. This reading sits on PROLOG.10's
' option A, which is the base its own recommendation (D on top of A) is
' built on, so option D changes nothing here. Option B WOULD reopen it,
' and the blast radius was measured rather than guessed: a
' transliteration of this classifier run over 25 term shapes x 6
' predicates before a single import showed option B moves EXACTLY TWO
' cases - a marked leaf whose stripped text parses as a number (`"42`,
' `"-3.5`). It never touches var/nonvar/atomic/compound, and both
' readings agree the term is `atomic`. Only the atom/number split moves.
'
' The marker is read in ONE place, LeafIsNumberTerm (below), for the
' reason UnificationBindsOutward (below) exists: a cross-cutting
' judgement written down once cannot be re-derived three ways by three
' callers and drift. `atom`, `number` and `atomic` all ask it.
'
' Classification runs on the RAW text, before LeafText strips anything -
' EvalArithTerm's own already-documented discipline. A capitalised
' quoted string like `"Hello` must never reach IsVarAtom with its marker
' gone, or a real string value is misreported as an unbound variable.
'
' `between`'s THREE decisions the roadmap does not settle:
'
'   MODE. X unbound -> generate. X bound to a whole number -> TEST, the
'       way real Prolog's between/3 is semi-deterministic in that mode
'       (`(between 1 10 5)` succeeds). Low and High must both resolve to
'       numbers; they go through EvalArithTerm, so they may be
'       EXPRESSIONS (consistent with the comparison family) and an
'       unbound or non-numeric bound inherits the two existing
'       {form}-templated arithmetic refusals, attributed to
'       "(between ...)". X bound to a non-number, or to a number that is
'       not whole, is refused BY NAME rather than failed: `(between 1 10
'       2.5)` answering a bare False would be indistinguishable from
'       `(between 1 10 25)` answering False, and the user could not tell
'       "out of range" from "not the kind of thing between talks about".
'       Refusing is also what keeps test mode and generate mode the SAME
'       relation - test mode succeeds exactly on values generate mode
'       would produce.
'   EMPTY RANGE. Low > High simply FAILS - zero solutions, no error. It
'       is an empty range, exactly like an empty candidate list, and
'       `(between 1 N X)` with N bound to 0 must yield no rows rather
'       than stopping the query.
'   BOUNDING. Measured, not assumed. PROLOG_MAX_STEPS is 120 and is a
'       TOTAL-RESOLUTION-WORK ceiling (SolveIsolated's own header says
'       so), charged once per candidate by the loop below. So every
'       generated value is charged one step, exactly as a candidate is,
'       and `between` gets NO private budget of its own - a second
'       ceiling would let a query spend 120 steps on everything else
'       plus an unrelated allowance on enumeration. That alone already
'       makes `(between 1 1000000 X)` terminate, but it would terminate
'       into prolog-step-ceiling, whose text blames "a rule that
'       recurses without ever reaching a base case" - a confidently
'       wrong answer for a user whose rules are all fine. So the range
'       is ALSO checked UP FRONT, before a single value is generated,
'       and refused by its own name against the same ceiling.
'
'       The honest cost, stated rather than hidden: `between`'s usable
'       range is therefore capped at 120, which is small, and that is a
'       limit of the ENGINE'S ceiling rather than of this item.
'       PROLOG_MAX_STEPS was tuned against a genuinely non-terminating
'       rule and its own declaration comment demands real profiling data
'       before it moves; raising it to make a feature look better is
'       exactly the trade that comment forbids, so it is NOT raised
'       here. Note the asymmetry this creates, which is correct and is
'       pinned by a test: `(between 1 1000000 5)` SUCCEEDS, because test
'       mode enumerates nothing and so has no range to refuse.
'
' CUT, which the roadmap does not mention at all. `between`'s loop is a
' loop between a cut's own origin and its firing site, so it must stop
' generating the moment cutActive comes back True - the candidates loop
' below documents why every such loop must. It never ABSORBS the signal
' (never clears it): absorption belongs only to the loop that selected
' the clause the `!` sits inside, identified by myStep = cutTargetBarrier,
' and `between` selects no clause and creates no barrier. So it is
' exactly the "any other loop leaves it set" case already described
' there. Without this, `!` silently fails to prune a generator.
'
' A representation edge, pinned rather than special-cased: this engine
' has no list type, so findall's own Bag is a Collection and an EMPTY
' bag is a zero-length one. `(compound EmptyBag)` is therefore True
' here, where ISO's `[]` is atomic. Classifying by representation is the
' honest reading when the representation is all there is; inventing a
' special case for zero-length would make `compound` mean something the
' rest of the module does not.
'
' PROLOG.7: the six comparison operators (<, >, =<, >=, =:=, =\=) as
' goals in their own right. PROLOG.5.1 shipped arithmetic strictly as the
' BINDING form `(is Var Expr)`, so there was no way to TEST two numbers
' against each other - the README's own staffing example had to key on an
' exact dept/cert match rather than a threshold, precisely because there
' was no `(> Salary 80000)` to write.
'
' A comparison is a GOAL, not an expression: it succeeds or fails, and it
' never binds. That single sentence decides the whole shape. It is new
' dispatch in SolveGoalList's own chain, sibling to not/findall/!, and NOT
' a new case inside any arithmetic function - there is no value to produce
' and nothing to unify a result against, so envN/envT are threaded into
' the continuation UNCHANGED, exactly as `not` already does and unlike
' `is`, which must clone (isN/isT) because it does bind.
'
' Each side is evaluated by EvalArithTerm - the same recursive walker
' `(is ...)`'s own right-hand side already uses, which dereferences
' through the env and refuses an unbound or non-numeric leaf by name.
' NOT VLA_Relation.ComputeArithmetic directly: that is the shared leaf
' substrate EvalArithTerm itself delegates its innermost binary op to,
' it never walks an env, and reaching for it here would have re-
' implemented the walk rather than reused it.
'
' The comparison itself reuses VLA_Relation.CompareValues with
' bothNumeric:=True - this project's own shared six-operator vocabulary,
' already carrying SQL's and DATALOG's comparisons. Only the SPELLING
' differs, so ComparisonOpFor (below) is a pure translation table:
' Prolog's =< is that vocabulary's <=, and =:= / =\= are its = / <>. Two
' Doubles enter it, so its own AsInvariantDouble is an identity step and
' no text round-trip or locale question arises.
'
' =:= and =\= are NUMERIC equality only. `=`/`\=` (unification) and
' `==`/`\==` (structural identity) are PROLOG.8's, deliberately untouched
' here and deliberately still unreserved - see IsReservedPredicateName.
'
' Two disciplines inherited rather than re-decided: the six names join
' IsReservedPredicateName (so they refuse as predicate or table-column
' names, the forward-reservation precedent PROLOG.5.1 set for the other
' four), and the refusals EvalArithTerm raises now carry {form} so a user
' who wrote (> Salary 80000) is never told about (is ...) - a form they
' did not write. tools/check_prolog_form_attribution.ps1 holds that
' mechanically; VLA_Messages.bas's own PROLOG.7 note has the reasoning.
'
' SolveComparison is factored OUT of the dispatch for the identical live-
' caught reason HarvestFindallBag already was: every local declared
' directly in SolveGoalList bloats its per-step recursive stack frame.
' The dispatch arm therefore declares none of its own.
'
' PROLOG.6: table-sourced and named-column facts. Reuses DATALOG.3's
' column-scoped table-argument resolution (VLA_Relation.TableArgResolve)
' and DATALOG.5's named-column-atom desugaring PATTERN directly, per
' this item's own roadmap wording ("both pieces of substrate it needs
' already exist and are already proven live") - confirmed by reading
' both, not assumed still true.
'
' A real, deliberate DIVERGENCE from DATALOG's own reuse, found by
' reading DATALOG.5's header closely before assuming its shape carries
' over unchanged: table rows here are read via `VLA_Relation.RangeToRows`
' (SQL.1's own bag-preserving reader), never `RelFromRange` (a
' deduping SET) - PROLOG.3's own already-decided stance, "facts are a
' bag, not deduped at storage" (matching SQL, not DATALOG's own
' fixpoint-necessitated dedup), applies identically to table-sourced
' facts as to hand-authored ones; reusing `RelFromRange` here would have
' silently dropped a genuinely duplicate source row.
'
' A second real divergence, found while wiring value conversion, not
' assumed: DATALOG keeps a table-sourced value's own real Excel Value2
' type all the way through to its own spilled output (never converting
' it to text at all) - PROLOG's own substitution environment can't do
' that (PROLOG.5.1's own header: "only ever holds a bare String or a
' Collection", the introduction of a third raw type being "exactly the
' kind of substrate mismatch" already found and avoided once for
' VLA_Relation's own tuple storage). Every table cell is therefore
' converted to a PROLOG-native leaf at fact-construction time
' (`TableCellToTerm`, below): a real Excel numeric type
' (`VLA_Relation.ValueIsNumericType`) renders via `Trim$(Str$(CDbl(v)))`
' - VBA's `Str$` is genuinely locale-invariant (always '.' as the
' decimal separator, unlike `CStr` on a Double, which respects the
' user's own regional settings) - `CStr`, this section's own SD-4
' locale-invariance freeze applied to a WRITE path for the first time in
' this codebase (every prior invariant helper, `Val`/`IsInvariantNumericString`,
' only ever READS text; nothing before this item needed to WRITE a
' locale-invariant number back out from a raw Excel type). Every other
' value (text, a live Excel Boolean, a blank cell) is marked with the
' SAME leading-`Chr$(34)` quoted-atom convention `VlaReadForms`'s own
' reader already uses for a quoted string literal in program text -
' load-bearing, not cosmetic: a table's own text column routinely holds
' capitalized values ("Bob", "Chicago") that `IsVarAtom` would otherwise
' misdetect as an unbound VARIABLE rather than ground data the moment
' its first letter is uppercase, exactly the class of bug this session's
' own PROLOG.5.x line has already found and fixed more than once
' elsewhere. Confirmed, not merely defended by symmetry: `TermHasVariable`
' already forces a HAND-AUTHORED fact to write any capitalized ground
' constant quoted too (`(fact (employee "Bob" eng))`, never a bare
' `Bob`), so this is the ONLY spelling a same-content hand-authored fact
' could ever have used anyway - a query written against a mix of
' hand-authored and table-sourced facts for the identical predicate
' behaves identically either way, not a two-tier system. A live Excel
' Boolean (`True`/`False`, also capitalized) gets this SAME treatment for
' free, since `ValueIsNumericType` excludes `vbBoolean` - no separate
' case needed.
'
' A real, pre-existing gap in `VLA_Unify.UnifyTwoWay` found while
' wiring this - confirmed by hand-tracing, not assumed harmless -  and
' fixed, narrowly: `SolveGoalList`'s own candidates loop used to unify
' the WHOLE query-goal term against the WHOLE freshened clause-head term,
' position 1 (the predicate name itself) included, via `UnifyTwoWay`'s
' own fully generic per-element recursion, which has no notion that
' position 1 of a PREDICATE-APPLICATION term is a functor, never a
' variable candidate - the SAME rule `FreshenTerm`/`CollectVars`/
' `TermHasVariable`/`ResolveTermDeep`/`SubstituteTemplate` all already
' enforce, `UnifyTwoWay` alone never did. Harmless whenever both sides
' spell the predicate identically (the "same text" fast path short-
' circuits before ever binding anything), but a capitalized predicate
' name - routine for a table-sourced one, e.g. a Table literally named
' "Employees" - reaching this loop with EVEN ONE mismatched-case
' occurrence (a query written `(Employees ...)` against a stored head
' `(employees ...)`, `clauseDict`'s own FOLDED-key lookup already having
' correctly matched them as the SAME predicate) would have `UnifyTwoWay`
' misdetect the capitalized side as an unbound VARIABLE and silently
' BIND it, real but confirmed-harmless-in-practice pollution (the stray
' binding is never one of `freeVarNames`, since `CollectVars` already
' skips position 1 correctly, so it never surfaces in output) - not
' fixed by patching `UnifyTwoWay` itself, which is deliberately generic
' and MUST stay that way (findall's own harvested Bag is a genuinely
' functor-less list where position 1 IS an ordinary, potentially-
' variable element - `ResolveTermDeep`'s own header has the full
' reasoning for why `UnifyTwoWay` can never safely special-case position
' 1 unconditionally). Fixed instead at the ONE call site that is
' STRUCTURALLY GUARANTEED to always be comparing two predicate-
' application terms of already-equal arity (clauseDict's own lookup, and
' `RecordArity`'s own cross-definition consistency check, both already
' guarantee this) - a new `UnifyArgsOnly` (below) unifies only positions
' 2..N, skipping the predicate-name position entirely as pure redundant
' (and, in the mismatched-case scenario, actively wrong) work.
' Provably non-regressing for the matched-case scenario (a same-text
' position-1 comparison always succeeded anyway) and strictly correct
' for the mismatched-case one (no stray binding at all, not just a
' harmless one).
'
' `prolog-tables-not-yet-supported` is retired outright at this item's
' own build time (this file's own `unify-glue-multiple-slots` precedent,
' VLA_Messages.bas's own header, `prolog-cut-not-yet-supported`'s own
' PROLOG.5.4 retirement the SAME pattern applied a second time) - tables
' are real now, no code path can raise it anymore.
'
' A genuinely NEW, deliberately accepted limitation, named rather than
' silently risked: a table-sourced predicate's own arity (its column
' count) is never cross-checked against a HAND-AUTHORED `(fact ...)`/
' `(rule ...)` clause sharing the SAME predicate name - `RecordArity`'s
' own cross-definition consistency check (PROLOG.3's own mechanism)
' stays scoped to `clausesText` alone, exactly as before this item. A
' genuine mismatch (a table with 3 columns, a hand-authored fact for the
' same name with 2 arguments) is not refused by name; it simply never
' unifies (`UnifyTwoWay`'s own `.Count` check), the two clause shapes
' silently never matching each other - not a crash, not silently WRONG,
' just less diagnostically helpful than a named refusal would be. Left
' this way deliberately: fixing it would need `predArity` threaded as a
' THIRD shared object (alongside the new `clauseDict`/`headerMap`
' threading this item already adds) for a genuinely narrow scenario
' (deliberately mixing a table and conflicting hand-authored arity for
' ONE predicate) - real profiling/demand, not speculative building,
' `DATALOG`'s own "profile first" doctrine applied to a refusal instead
' of a cache.
'
' `ParseProgram` (below) changes shape twice for this item, both
' confirmed necessary before writing either: `clauseDict` is now an
' INPUT the caller creates and may pre-populate (table-sourced facts,
' added before `ParseProgram` ever runs) rather than something
' `ParseProgram` allocates fresh itself - the identical "table rows
' first, then whatever the clauses text itself defines, in written
' order" sequencing PROLOG.4's own "one predicate, fact AND rule
' together" precedent already established, extended to a third clause
' source; and a new `headerMap` parameter (predName folded ->
' `VLA_Relation.RangeColumnNames`' own (folded, original) pair
' Collection) threads through to a new `DesugarBodyItem`, called on
' every RULE BODY item and QUERY CONJUNCT (never a fact/rule HEAD -
' DATALOG.5's own explicit restriction, repeated here for the identical
' reason: a rule head DEFINES a new predicate's own output positions,
' it never REFERENCES an existing table row, so keyed syntax has no
' meaning there).
'
' `DesugarBodyItem`'s own classification, worked out from first
' principles rather than ported blindly, because PROLOG's own term
' grammar creates a genuine ambiguity DATALOG never had: DATALOG forbids
' compound-term arguments entirely, so a 2-element-list argument is
' UNAMBIGUOUSLY either a keyed-pair attempt or a parse error - there is
' no third legitimate reading. PROLOG genuinely allows compound-term
' arguments as ordinary data (`(fact (pair a b))` is completely
' legitimate, unrelated to any table), so a 2-element-list argument to
' an ORDINARY (non-table-sourced) predicate must never even be
' CONSIDERED for keyed-atom desugaring - checked here by looking the
' predicate up in `headerMap` FIRST, before examining its own arguments
' at all: not found there -> the atom is returned completely untouched,
' whatever compound terms its own arguments may legitimately hold. Only
' once a predicate IS confirmed table-sourced do arguments get
' classified: every argument bare -> already positional, untouched;
' every argument a clean 2-element (bareHeader, value) pair -> a genuine
' keyed atom, resolved against `headerMap`; anything else (a bare/list
' mix, or a list that fails the strict pair shape) -> refused by name
' (`prolog-atom-mixed-keying`) - a single unified refusal, unlike
' DATALOG's own two-case split, because PROLOG has no separate,
' pre-existing "no compound-term nesting" ban to silently fall through
' to for the second case the way DATALOG does; a malformed pair shape
' against a KNOWN-table-sourced predicate is refused here directly.
'
' An unmentioned column's own anonymous variable (`AnonymousColumnVarName`
' -> a gensym-free, deterministic name derived only from the atom's own
' written occurrence - the body-item/query-conjunct loop index the SAME
' way DATALOG.5's own `bodyItemIndex` already is, plus the column's own
' 1-based header position, this project's own standing veto on
' gensym-style identifier invention, unchanged) needs NO special
' handling anywhere in the solver, confirmed rather than assumed: it
' starts with an uppercase letter, so `IsVarAtom` already treats it as
' an ordinary variable, and `FreshenTerm`'s own existing generic
' recursion already freshens it per clause invocation like any other
' rule-body variable, for free. The ONE place it DOES need a change:
' `CollectVars` (a QUERY conjunct's own keyed atom, unlike a rule body's,
' is never freshened and its own free variables ARE the query's own
' output columns) gains one more skip-shape check - a variable whose own
' name starts with the reserved `VlaAnon` stem is never collected as a
' free output column, the identical "must never leak" instinct already
' applied to `not`/`findall`'s own discarded sub-search bindings,
' PROLOG.5.2's own header, applied here to a variable that DOES bind in
' the shared, non-discarded env but was never something the user asked
' to see.
'
' `PROLOG(clauses, ParamArray tables())` is now a thin wrapper (`SQL.1`'s
' own worksheet-function-around-a-testable-core shape, `DATALOG`'s own
' `DatalogRun`/`DATALOG()` split, precedented and reused rather than
' invented fresh): it resolves each table argument's own name
' (`TableArgName`, below, `VLA_Datalog.TableArgName`'s own exact shape,
' turning `VLA_Relation.TableArgResolve`'s shared category codes back
' into PROLOG's own wording), reads its rows, converts and stores them
' as ordinary fact clauses, captures its header into `headerMap`, then
' delegates to a new `Public Function PrologRun(clausesText, clauseDict,
' headerMap)` - the real engine, directly callable with a HAND-BUILT
' `clauseDict`/`headerMap` and no live workbook at all, exactly
' `DatalogRun`'s own precedent, and exactly what makes
' `TestPrologKeyedAtoms` (below) a PURE test of the desugaring mechanism
' itself rather than a host-required one.
'
' PROLOG.5.4: cut (`!`). The only one of the four PROLOG.5.x sub-items
' bullet (cut, is/arithmetic, negation-as-failure, findall/aggregation) -
' split into PROLOG.5.1-5.4, built and owner-tested one at a time, the
' same rhythm PROLOG.1-4 themselves already used rather than one
' uninterrupted pass: 5.1 (is/arithmetic, this item) is purely additive,
' a new body-item KIND with no control-flow change at all; 5.2
' (negation-as-failure) needs a genuinely new bounded sub-call that runs
' the solver in isolation just to check existence, then inverts it; 5.3
' (findall) reuses 5.2's own sub-call machinery directly, harvesting
' solutions instead of just checking they exist; 5.4 (cut) is last and
' gets its own item on purpose - the only one of the four that performs
' surgery on SolveGoalList's own backtracking loop itself rather than
' adding a self-contained new body-item kind, the highest-risk piece and
' the one most likely to regress PROLOG.4's own already-working, owner-
' tested search if built in the same pass as the other three.
'
' PROLOG.5.4: cut (`!`). The only one of the four PROLOG.5.x sub-items
' that performs surgery on SolveGoalList's own `For Each clauseRec In
' candidates` loop itself (the ONLY real choice point this engine has -
' is/not/findall's own dispatch branches are all deterministic, exactly
' one outcome per attempt, nothing to prune) rather than adding a
' self-contained new body-item kind the way is/not/findall each did -
' this item's own roadmap bullet named this as the highest-risk piece,
' confirmed while building it: two new ByRef parameters thread a cut
' signal - `cutActive As Boolean`, `cutTargetBarrier As Long` - through
' EVERY SolveGoalList call, alongside stepsTaken, the smallest signal
' shape that works (no new Collection/record type).
'
' What one cut prunes, worked out concretely before writing any code:
' real Prolog cut commits to BOTH (a) the CURRENT predicate call's own
' remaining candidate clauses, and (b) every choice point opened SINCE
' entering the clause body `!` appears in - meaning EARLIER goals in the
' SAME clause body also stop being retriable, not just the clause
' selection itself. Concretely: every `For Each candidates` loop
' instance on the call stack between (inclusive) the loop that selected
' THIS clause and the point `!` fires must stop trying further
' candidates; the loop that selected this clause is the one that must
' ABSORB the signal (stop propagating it further up), since an even
' more outer clause's own choice points are OUTSIDE this cut's scope.
'
' The real design fork this item's own scoping had to resolve, not
' assumed: SolveGoalList's own `goals` parameter is a FLAT list, merging
' a clause's own freshened body items with `rest` (whatever followed the
' call in the CALLER's own list) into one undifferentiated Collection -
' PROLOG.4's own design, unchanged since. A single scalar "current
' barrier" parameter, updated once per clause splice, cannot correctly
' track this: once a NESTED predicate call's own body is spliced in
' front of `rest`, the SAME merged list now contains goals belonging to
' TWO different clause activations at once (the nested call's own body,
' and whatever originally followed the goal that invoked it), and a `!`
' appearing in the OUTER portion needs the OUTER barrier, not the nested
' one - hand-traced with a concrete two-level example (a clause calling
' a second predicate, then cutting, where that second predicate's OWN
' clause also happens to use cut) before concluding a single "current"
' scalar is insufficient in general.
'
' Solved without a parallel per-goal barrier Collection (which would
' have worked but adds a new threaded structure this item's own roadmap
' guidance asked to avoid if a primitive suffices): the barrier is
' encoded directly INTO the cut atom itself, at the exact moment
' FreshenTerm (unchanged in shape) already rewrites a clause's own body
' items with a per-invocation suffix (`"#" & stepsTaken`, PROLOG.4's own
' mechanism, already deterministic and already unique per clause try) -
' a bare "!" now takes the SAME freshening branch a variable atom
' already does (`s & suffix`), turning `!` into e.g. "!#37" wherever it
' is spliced from a clause body, permanently carrying its OWN clause
' activation's identity with it no matter how deep the merged list gets
' processed afterward. No new counter, no new parallel Collection - the
' existing per-clause suffix already uniquely identifies "which clause
' activation," reused rather than duplicated. A bare, un-suffixed "!" -
' only ever possible directly in a (query ...) conjunct, since
' queryConjuncts are never freshened (ParseProgram's own long-standing
' comment) - is barrier 0, meaning "cut back to the very start of the
' query"; no real choice-point loop's own barrier (always the stepsTaken
' value AFTER at least one increment, so always >= 1) can ever equal 0,
' so a top-level cut's signal simply never gets formally "absorbed" and
' instead propagates, unread, all the way back to PROLOG()'s own
' top-level call - harmless and correct by construction, not a special
' case: every actual ancestor `For Each` loop between the query's own
' start and the cut still sees cutActive=True after its own recursive
' call returns and still breaks (below), which is the only observable
' effect a barrier-0 cut is supposed to have; nothing reads cutActive
' after PROLOG()'s own top-level SolveGoalList call returns.
'
' The `For Each clauseRec In candidates` loop's own protocol, the one
' piece of real surgery: `myStep` (the stepsTaken value for THIS
' specific candidate try, captured once - the same value already used
' to build `suffix`) is compared to `cutTargetBarrier` only AFTER the
' recursive call for this candidate returns, and only when `cutActive`
' is True. Whenever `cutActive` is True at that point, this loop
' unconditionally stops trying further candidates (`Exit For`) - every
' loop between the origin and the firing site needs this, and since
' stepsTaken only ever increases, EVERY loop that can legitimately
' observe an active signal here already has `myStep >= cutTargetBarrier`
' by construction (proved by hand-tracing, not assumed - above). Only
' when `myStep = cutTargetBarrier` exactly - this IS the loop that
' selected the clause `!` fired inside - is the signal also CLEARED
' before returning, so it never escapes to affect this loop's own
' caller (an even more outer clause activation, genuinely outside this
' cut's scope); every other loop in between leaves it set, letting it
' keep propagating up to that exact origin.
'
' `is`/`not`/`findall`'s own dispatch branches need no new logic at all,
' confirmed rather than assumed: none of them is a choice point (each is
' deterministic, exactly one outcome per attempt, PROLOG.5.1/5.2/5.3's
' own headers already establish this), so each simply threads the two
' new ByRef parameters through its own single forward recursive call
' unchanged.
'
' Cut's own opacity across `not`/`findall`'s own isolated sub-search -
' the real-Prolog rule that a cut inside `\+ Goal` or `findall`'s own
' Goal argument is confined to THAT Goal's own proof and can never
' escape to prune the outer query's choice points - confirmed to fall
' out for free from the existing recursion shape, not requiring any
' change to SolveIsolated's own contract beyond adding its own local
' pair: SolveIsolated (PROLOG.5.2/5.3) already runs the isolated Goal
' through a FRESH, wholly separate SolveGoalList call tree (a cloned
' env, PROLOG.5.2's own header has the full reasoning) - it now also
' declares its own fresh, LOCAL `subCutActive`/`subCutTargetBarrier`
' pair rather than accepting or forwarding the caller's own, so any cut
' signal born while proving the isolated Goal is entirely self-contained
' within SolveIsolated's own stack frame and is simply discarded, unread,
' the moment SolveIsolated's own SolveGoalList call returns - the
' identical "unread by construction" shape a barrier-0 top-level cut
' already relies on above, not a new mechanism.
'
' A cut inside a CONJUNCTIVE QUERY (not a rule body) is meaningful, real
' Prolog's own behavior for a directly-typed query (as if the whole
' query were the body of one anonymous, always-selected clause) -
' confirmed, not assumed: `(query (p X) ! (q X))` commits to the first
' solution of `(p X)` exactly as `(rule (r X) (p X) ! (q X))` would
' commit to it inside a rule, via the barrier-0 mechanism above
' producing identical pruning behavior at every ancestor loop between
' the query's own start and the cut, even though the signal itself is
' never formally "absorbed" by an equality match.
'
' A genuinely NEW, deliberately accepted limitation, named rather than
' silently risked: FreshenTerm now treats ANY bare "!" leaf, anywhere in
' a clause, as a cut marker unconditionally - it cannot distinguish "!"
' used as a CONTROL construct (a body-item position) from "!" used as an
' ordinary DATA atom (e.g. a fact argument literally named "!"), since
' FreshenTerm's own tree walk carries no positional context beyond
' "position 1 is a functor, skip it." Using "!" as a plain data value is
' therefore unsupported from this item onward - a rare usage even in
' real Prolog, and IsReservedPredicateName already forbade "!" as a
' predicate NAME from PROLOG.5.1 onward; detecting "is this specific
' occurrence a control position or a data position" would need exactly
' the same context-tracking this item's own barrier design deliberately
' avoided adding elsewhere, so this is named here rather than guarded
' against in code.
'
' `prolog-cut-not-yet-supported` (VLA_Messages.bas) is retired outright,
' not left dead or repurposed - no code path can raise it anymore, and
' this project's own precedent (`unify-glue-multiple-slots`, replacing
' the former `english-render-glue-multiple-slots`, VLA_Messages.bas's
' own header there) is to remove a superseded id from the catalogue
' rather than keep an unreachable entry around.
'
' PROLOG_MAX_STEPS (120, unchanged from PROLOG.5.3): this item adds two
' ByRef parameters to SolveGoalList's own signature (cheap, the same
' class of cost `stepsTaken` itself already is) plus one Long local
' (`myStep`) in the candidates loop and one Long local (the parsed
' barrier) in the new cut-dispatch branch - deliberately far lighter
' than PROLOG.5.3's own seven-local incident that forced 200->120, per
' this item's own roadmap warning to prefer a separate function past one
' or two new locals directly inside SolveGoalList's own dispatch. NOT
' assumed safe on that basis alone, though: PROLOG.5.3's own regression
' was caught only by actually re-running the deliberately non-terminating
' rule test to real exhaustion after the fact, not by reasoning about
' local counts in advance - this needs the identical live re-check
' before being trusted, not skipped because the change looks small.
'
' PROLOG.5.3: `(findall Template Goal Bag)` - harvests EVERY solution of
' Goal (not just whether one exists) and binds Bag to a single list term
' - one element per solution, each element Template's own structure with
' every one of ITS OWN variables substituted for that solution's resolved
' value. Reuses PROLOG.5.2's own bounded sub-call microscope directly,
' per this item's own roadmap wording, not a fresh parallel mechanism:
' SolveNegation (PROLOG.5.2) is refactored into a thin wrapper over a new
' shared `SolveIsolated` - clone envN/envT, discard regardless of
' outcome, run SolveGoalList on Goal alone, share stepsTaken (ByRef) with
' the caller - and `not`'s own dispatch call site does not change AT
' ALL, only ever passing an empty freeVarNames where PROLOG.5.3 now
' passes Template's own collected variables instead - exactly the one
' difference this item's own roadmap wording named ahead of time.
'
' Template is plain DATA, never a goal, and must NEVER be goal-
' interpreted the way CollectVars (below) treats a query conjunct -
' confirmed by construction, not assumed: a Template that happens to
' itself contain a literal `(not X)` or `(findall ...)` sub-term (e.g.
' `(findall (not X) (foo X) Bag)`, building a list of negation-shaped
' terms as ordinary data - a legitimate real-Prolog idiom) must still
' have every one of ITS OWN variables collected honestly. Reusing
' CollectVars for this would silently under-collect such a Template
' (CollectVars' own `not`/`findall` skip-shape checks exist specifically
' for GOAL interpretation, the opposite job), so Template's own variables
' are collected by a separate, deliberately simpler `CollectTemplateVars`
' with none of that special-casing - the same generic recursion
' CollectVars itself had before PROLOG.5.2 added its first skip-shape
' check.
'
' A real, pre-existing gap found while building this, not merely
' assumed away: `SolveGoalList`'s own base case (below) only ever
' shallow-resolved a free variable's name via `EnvWalkInto`'s own
' top-level chain walk - correct for a variable bound directly to a
' ground value or chained to another variable, but WRONG the moment that
' walk lands on a COMPOUND value that itself still contains an
' unresolved nested variable (e.g. a rule head `(r X (box X))` where X
' is bound to a real value only by a LATER body goal - `box`'s own
' argument would render as the raw freshened name "X#42" instead of the
' real value). Hand-traced against the ALREADY-COMMITTED PROLOG.4 code
' before touching anything, confirmed reachable (not hypothetical) -
' unification itself was never affected, since `UnifyTwoWay`
' (`VLA_Unify.bas`) already re-derefs correctly at every nesting level
' (every recursive call starts with its own `EnvWalkInto`); only this
' OUTPUT-time resolution, which walks a value one level and stops, had
' the gap. `findall`'s own Template harvesting needs full, correct deep
' resolution to work at all (a harvested list element that still showed
' raw internal variable names would be useless), so the fix - a new
' `ResolveTermDeep` (below, a Sub with a ByRef `dest` out param,
' `EnvWalkInto`'s own object/non-object discipline, not a Function
' return - the SAME "bare Variant assignment invokes an object's own
' default member" trap `EnvWalkInto`'s own header already documents)
' - is made ONCE, in the one shared spot both an ordinary top-level
' query's own free-variable output AND `findall`'s own per-solution
' harvest already go through, rather than patched twice. Strictly
' non-regressing, checked by hand-tracing PROLOG.3's own existing
' ground-compound-term round-trip test against the new code before
' relying on that claim: a value with no nested unresolved variable
' resolves to a structurally identical result either way (the new code
' rebuilds a fresh Collection with the same contents rather than
' returning the original object reference, which `VlaWriteForm`'s own
' structural serialization can't tell apart). A regression pin for the
' bug itself lives in `TestPrologRules` (PROLOG.4's own suite, since the
' gap was in PROLOG.4-era code, not `findall`-specific), and
' `TestPrologFindall` separately pins a Template that needs it.
'
' PROLOG.5.2: `(not Goal)` - negation-as-failure. Needs a genuinely new
' microscope, not reducible to is/arithmetic's own reuse-of-UnifyTwoWay
' shortcut: SolveNegation (below) runs SolveGoalList on Goal ALONE, in
' total isolation - a CLONED envN/envT (VLA_Unify.UnifyEnvClone) that is
' discarded regardless of outcome - just to check whether at least one
' solution exists, then inverts success/failure. Built once here,
' PROLOG.5.3's own findall reuses this exact call shape directly (only
' its own freeVarNames - the findall template's variables, not an empty
' Collection - and what it does with the harvested solutions afterward
' will differ).
'
' DATALOG.1's stratified negation is the nearest sibling in spirit (both
' refuse a case that can't be soundly enumerated) but not in mechanism,
' confirmed by reading DATALOG's own negation code before assuming the
' shape carries over: DATALOG proves stratifiability at PARSE time and
' keeps its own termination guarantee; PROLOG's negation-as-failure has
' no such proof available (a non-terminating sub-goal negated is simply
' non-terminating), so it inherits PROLOG.4's own RUNTIME step-budget
' refusal instead - SolveNegation threads stepsTaken THROUGH (ByRef, the
' SAME running total the outer search already carries), never a fresh
' sub-budget of its own, so `(not (loop X))` over a genuinely non-
' terminating rule hits the identical prolog-step-ceiling refusal
' ordinary recursion already does, rather than silently dodging it via a
' separate allowance.
'
' A real design fork, decided by hand-tracing rather than left to
' CollectVars' own existing generic recursion (the same recursion `is`
' relies on unchanged - PROLOG.5.1's own header below explains why THAT
' case needs no special handling): a variable appearing ONLY inside a
' negated goal's own argument must NEVER become a free query-output
' column. `is`'s own target Var genuinely binds in the shared,
' non-discarded outer env, so collecting it is correct; `not`'s own Goal
' binds only inside a clone that's thrown away no matter what, so a
' variable that appears NOWHERE else in the query would resolve to
' nothing and render as its own unresolved atom name instead of a real
' value - confirmed by hand-tracing `(query (not (foo X)))` with X
' appearing nowhere else: without this fix it would spill a one-column
' "X" header over a literal "X" string, not a real value. CollectVars
' (below) now recognizes `(not Goal)` and skips descending into Goal
' entirely - a variable that ALSO appears in an ordinary, non-negated
' conjunct is still collected correctly from THAT occurrence; this only
' ever skips the negated occurrence, never removes an already-collected
' name.
'
' What must never leak back out, checked explicitly rather than assumed
' safe: a binding made while proving Goal, on EITHER success or failure
' of the negation - handled by the clone/discard shape above, the exact
' words this item's own roadmap bullet used. An ERROR raised while
' proving Goal (an unbound-variable `is`, a divide-by-zero, an
' occurs-check violation) is the opposite case and must NOT be caught
' here - negation-as-failure only ever catches ordinary proof FAILURE
' (zero solutions), never a VBA error, so SolveNegation has no On Error
' of its own at all and lets any such error propagate all the way up to
' PROLOG()'s own top-level handler, exactly as if the same goal had been
' proved un-negated - real Prolog's own \+ behaves identically.
'
' `not` was already reserved as a predicate NAME from PROLOG.5.1 onward
' (IsReservedPredicateName, unchanged); this item is what makes it mean
' something as a body-item USE, both in a rule body (freshened per
' invocation by FreshenTerm's own existing generic recursion - no change
' needed there at all, since position 1 of any compound term, "not"
' included, was already treated as a non-variable functor position) and
' in a query conjunct.
'
' PROLOG.5.1: `(is Var Expr)` - Expr a recursive arithmetic expression
' over +/-/*//, evaluated bottom-up via `VLA_Relation.ComputeArithmetic`
' (confirmed present and exactly the shape needed - a mechanical
' "given two already-numeric operands and an operator, compute" with no
' storage/tuple concern at all), then UNIFIED against Var via the SAME
' `UnifyTwoWay` every other goal already uses - not a bespoke "bind a
' brand-new variable" mechanism the way DATALOG's own `(let Z (op X Y))`
' needs (DATALOG's `let` binds one-way because DATALOG's own engine has
' no real unifier; PROLOG already does, so `is` costs nothing extra
' here: a fresh variable binds, an already-bound one is checked for
' equality, exactly real Prolog's own `is/2` semantics, for free).
'
' Genuinely NEW dependency, not a violation of this module's own
' standing "LAYER 1 (VLA_Unify) - NOT VLA_Relation" design note below:
' that note was about fact STORAGE specifically (VLA_Relation's own
' tuple/dedup machinery genuinely doesn't fit a Prolog fact, PROLOG.3's
' own header has the full reasoning) - `ComputeArithmetic`/
' `IsInvariantNumericString`/`InvariantVal` are pure, storage-free
' mechanical helpers, orthogonal to that concern entirely, and this
' section's own "consumes rather than grows a twin" precedent (already
' applied three times: SQL.1's evaluator into DATALOG.4, DATALOG.5's
' header-map into SQL, DATALOG.4's own arithmetic hoisted from SQL.2)
' applies a fourth time rather than writing PROLOG's own third
' arithmetic evaluator.
'
' Unlike DATALOG's own flat `(let Z (op X Y))` (exactly one binary op,
' no nesting - a deliberate restriction DATALOG needs because it forbids
' compound terms entirely), PROLOG's own arithmetic expressions may
' NEST arbitrarily - `(is Total (+ (* 2 3) 4))` is legal - because
' compound terms are already a first-class citizen in this engine (the
' whole reason PROLOG exists rather than just using DATALOG), so
' restricting `is` to one flat binary op would be an ARTIFICIAL
' limitation copied from a different engine's own different constraint,
' not a real one - confirmed by reading `ComputeArithmetic`'s own
' contract before assuming DATALOG's shape carries over unchanged.
'
' Reserved words, forward-declared now rather than one at a time:
' `is`/`not`/`findall`/`!` may never be used as a predicate NAME in a
' `(fact ...)`/`(rule ...)` definition, checked at parse time
' (`prolog-reserved-predicate-name`) even though only `is` is built this
' item - reserving all four now means PROLOG.5.2-5.4 can never silently
' break a knowledge base someone already wrote against 5.1 alone, SD-4's
' own "a shipped spelling keeps its meaning" applied to what predicate
' NAMES may exist, not just to program syntax. A bare `!` appearing as a
' body item or query conjunct WAS refused by name
' (`prolog-cut-not-yet-supported`) at this item's own build time, rather
' than silently mismatched against a predicate literally named "!" -
' PROLOG.3's own `prolog-tables-not-yet-supported` precedent, applied to
' a forward-declared keyword instead of a forward-declared argument.
' SUPERSEDED at PROLOG.5.4's own build time (this module's own header,
' top of file) - `!` is now a real cut, and the refusal id is retired.
'
' A computed arithmetic result is CStr()'d into the environment
' immediately, never left as a raw VBA Double - this engine's own
' substitution environment only ever holds a bare String (atom) or a
' Collection (compound term), `VLA_Unify.bas`'s own native shape; every
' downstream consumer (`IsVarAtom`, `UnifyTwoWay`'s own ground-atom
' `CStr` comparison, `FreshenTerm`, `RenderBoundValue`) already assumes
' exactly those two kinds, and introducing a raw numeric type as a third
' would be exactly the kind of substrate mismatch this module's own
' PROLOG.3 header already found and avoided once for `VLA_Relation`'s
' own tuple storage.
'
' The item's own termination guarantee - PROLOG has none by construction,
' unlike DATALOG - and, secondarily, a safety margin below VBA's own real
' native call-stack ceiling (SolveGoalList, below, recurses one frame per
' resolution step along the currently-explored proof path, unlike
' DATALOG's flat MAX_ROUNDS=10000 round loop). Deliberately far smaller
' than that value. Declared here, in the module's own declarations
' section, not down by SolveGoalList where it's used - VBA compile
' error, live-caught: a bare Const/Dim statement can never appear
' BETWEEN two already-defined procedures (only another Sub/Function/
' Property declaration, or a comment, may follow one's own End Sub/End
' Function) - every other engine's own module-level Const (VLA_Sql.bas's
' SQL_CTE_MAX_ROUNDS and kin) already lives in its own declarations
' section for exactly this reason, not by convention alone.
'
' Live-test evidence, not a hypothetical: the first value tried (1000)
' was confirmed too high on the owner's own machine - the deliberately
' non-terminating `TestPrologRules` case ran VBA out of native call-stack
' space (a raw "Out of stack space", caught by PROLOG()'s own On Error
' handler and returned as an ugly "#PROLOG! Out of stack space" string,
' never an Excel crash - VBA's stack-overflow error IS trappable here,
' worth knowing - but the WRONG refusal, not the worded step-ceiling one)
' before stepsTaken could ever reach 1000. 200 was then the more
' conservative value, and held through PROLOG.4/PROLOG.5.1/PROLOG.5.2.
'
' PROLOG.5.3: 200 stopped being safe, live-caught the SAME way - not by
' findall usage itself, but by findall's own FIRST DRAFT inlining its
' whole harvest (several new locals: template, templateVars,
' harvestedTuples, bag, tup, bagN, bagT) directly into SolveGoalList's
' own dispatch. VBA allocates stack space for every local a procedure
' declares the moment it's entered, regardless of which branch actually
' runs - so those seven extra locals bloated EVERY recursive SolveGoalList
' frame, not just findall's own, and the SAME already-passing "genuinely
' non-terminating rule" test (no findall in it at all) started
' overflowing the native stack before stepsTaken ever reached 200. Fixed
' structurally first - the harvest moved into its own function
' (HarvestFindallBag, below SolveGoalList), so those locals live in a
' separate frame paid once per findall call, not once per resolution
' step - and THEN the ceiling itself lowered to 120, real headroom below
' 200 rather than tuned to the exact new failure point, the identical
' caution the 1000->200 drop already used. The general lesson, worth
' carrying into PROLOG.5.4 (cut) before it happens again: any new local
' added directly inside SolveGoalList's own dispatch, however small,
' taxes every recursive frame - prefer a separate function the moment a
' dispatch needs more than one or two.  If 120 is STILL too high on some
' machine, lower it further; if a real knowledge base someday needs
' deeper recursion than 120 steps can reach, that is the moment to
' revisit this number against real profiling data, DATALOG's own
' "profile first, don't build the cache speculatively" doctrine applied
' to a ceiling instead of a cache.
Private Const PROLOG_MAX_STEPS As Long = 120
' PROLOG.4: `(rule head body...)` forms, unification-driven SLD
' resolution with backtracking, no cut - docs/BETA_ROADMAP2.md's own
' PROLOG.4 entry. One predicate may now be defined by facts AND rules
' together: ParseProgram (below) appends EVERY clause - a `(fact ...)`
' becomes a clause with an empty body, a `(rule ...)` a clause with one
' or more - to the SAME per-predicate Collection in the clauseDict,
' in written order, regardless of which keyword produced it. SolveGoalList
' (below) does not special-case "fact vs rule" at all as a result: it
' tries every clause for a predicate in turn, and a fact's own empty body
' just means zero new goals get spliced in before the search continues -
' the same code path a rule's freshened body takes when solving it
' finishes. This is the concrete answer PROLOG.3's own fact-only dict
' had none for.
'
' New machinery this item alone needed, confirmed against the real code
' rather than assumed: clause-variable freshening (FreshenTerm, below) -
' the same rule text used twice in one proof (directly, or through
' recursion) must never let two invocations' own variables collide.
' Found simpler than scoped while building it: no rename-map/dictionary
' is needed at all. Appending ONE shared suffix (the running step
' counter, stringified) to a variable's own ORIGINAL name deterministically
' reproduces the SAME fresh name at every occurrence of that variable
' within one clause invocation (two occurrences of X both become "X#37"),
' which is exactly what "rename apart" requires, with no lookup structure
' at all - see FreshenTerm's own header for the full reasoning.
'
' Reasoned explicitly, not merely asserted, on the one point this item's
' own roadmap bullet stated as settled and this build reopens: the
' roadmap text calls for "a real, explicit choice-point stack for
' backtracking, not native recursion... VBA has no continuations/
' generators to suspend a search and resume it later." That concern is
' real for an INTERACTIVE Prolog top-level (";" for the next solution,
' one at a time, on demand) - but SD-4's own frozen output contract
' never wants that: `PROLOG()` always returns the WHOLE solution set in
' one spilled array, eagerly. Plain recursion - the same shape
' PROLOG.3's own SolveConjuncts already used for facts-only backtracking,
' generalized here to also try rule clauses - collects every solution
' into one accumulator Collection as the search unwinds, with no need to
' ever SUSPEND a partial search and resume it later; a continuation/
' explicit-stack machine is machinery this item's own actual output
' contract never calls for. What genuinely changes once rules/recursion
' exist, and is real, is TERMINATION, not the backtracking mechanism:
' PROLOG.3's own recursion depth was bounded by the query's own fixed
' conjunct count (finite by construction, no rule ever re-derives a
' predicate from itself); once a rule's body can invoke the same or
' another rule recursively, recursion depth depends on the DATA/program's
' own recursive depth, which may be unbounded or genuinely non-
' terminating. Two guards follow from that, both new: PROLOG_MAX_STEPS
' (below) bounds total resolution work regardless of how deep or wide
' the search goes, refusing by name past the ceiling - SD-4's own frozen
' requirement, and the item's actual termination guarantee, since PROLOG
' has none by design; and because THIS recursion's own native VBA call
' depth grows with the proof's own depth (one SolveGoalList frame per
' resolution step along the currently-explored path, unlike DATALOG's
' flat round loop), PROLOG_MAX_STEPS is deliberately set far below
' DATALOG's own MAX_ROUNDS (10000) - a worded refusal must fire before
' VBA's own raw "Out of stack space" ever could. Confirmed live, not
' hypothetical: the first value tried (1000) genuinely did run VBA out
' of native call-stack space on the owner's own machine, on exactly the
' deliberately-non-terminating test built to exercise this - caught by
' PROLOG()'s own error handler (VBA's stack-overflow error is trappable
' here, so this degraded to an ugly "#PROLOG! Out of stack space" string,
' never an Excel crash) but the WRONG refusal, not the worded step-
' ceiling one. PROLOG_MAX_STEPS's own declaration comment (module
' declarations section, top of file) has the current value and the
' reasoning for it; if it is ever found too high again, lower it further
' - the explicit-stack machine this item's own top paragraph already
' argued against building is still not warranted by this alone.
'
' PROLOG.3: new module - the first real `=PROLOG(...)` worksheet
' function, docs/BETA_ROADMAP2.md's own PROLOG.3 entry. Ground facts
' only (no `(rule ...)` yet - that is PROLOG.4) and conjunctive queries
' - `(query (pred1 ...) (pred2 ...) ...)` - answered by enumeration over
' the stored facts, not resolution: no search/backtracking machinery
' exists yet, since nothing here can recurse (facts are always ground,
' there is no rule to re-derive a predicate from itself), so the whole
' search space is exactly conjuncts-deep and facts-per-conjunct-wide,
' finite by construction - no step-budget safety valve needed at this
' item, unlike PROLOG.4, which is where that guarantee actually stops
' holding.
'
' Program text is VLA's own S-expression syntax (`VLA.VlaReadForms`,
' `VLA_Datalog.bas`'s own precedent) - but, verified while scoping this
' rather than assumed, that reader hands back only generic nested-list
' forms, nothing DSL-specific, so a real interpretation layer is still
' needed on top of it. This is the exact boundary `DATALOG` was carved
' out to avoid crossing: `DATALOG`'s own `ParseAtom` (`VLA_Datalog.bas`)
' raises `datalog-compound-term` the moment any argument is itself a
' list; `TermPredName`/`TermHasVariable` below do the OPPOSITE on
' purpose - a compound-term argument (`(fact (likes tom (color red)))`)
' is an ordinary, accepted term, recursively, since compound-term
' nesting is exactly what makes Prolog's unification a superset of
' Datalog's flat positional matching.
'
' A real simplification found by reading `VLA_Unify.bas` closely before
' writing this, not assumed from `DATALOG`'s own shape: unlike
' `DATALOG`, this module needs NO "atom record" abstraction (isVar/text
' pairs, `VLA_Datalog.bas`'s own `ParseAtom` shape) at the TERM level at
' all. `VLA_Unify.UnifyTwoWay` already operates directly on raw parsed
' forms - a bare String (atom or variable) or a Collection (compound
' term/list) - exactly `VlaReadForms`'s own native shape, no
' intermediate record needed. `DATALOG` needed its own flat record
' specifically because it forbids compound terms and needs cheap
' positional access for its hash-join-based fixpoint; `PROLOG`, built
' on the already-generic unifier, does not carry that constraint. This
' module's own parsing layer is genuinely smaller than `DATALOG`'s own
' as a result - it validates shape, tracks arity, and checks facts are
' ground, but never re-represents a term at all.
'
' A real substrate check made before assuming reuse, not after:
' `VLA_Relation.bas`'s own tuple storage (`RelNew`/`RelTryAdd`) cannot
' hold a fact with a compound-term argument at all - `TupleKey`
' (`VLA_Relation.bas`) does `CStr(t(i))` unconditionally over every
' tuple slot, which raises the exact runtime error 450 `VLA_Unify.bas`'s
' own PROLOG.2 fix already lives-caught once, the moment any slot holds
' an object. Facts here are therefore stored in a plain
' `VLA_Runtime.VlaDictNew()` dict (predicate name, folded, matching
' `DATALOG`'s own `predName` folding - never the fact's own argument
' text, `VLA_Identity`'s SD-8 fold is for PREDICATE-shaped names only,
' `VLA_Unify.bas`'s own PROLOG.2 header has the grep-verified receipt)
' mapping to a plain Collection of ground terms per predicate - no
' `VLA_Relation` dependency at all, a genuine, deliberate non-reuse
' rather than a missed hoist.
'
' Conjuncts were solved by `SolveConjuncts` (PROLOG.3's own function,
' since SUPERSEDED - see this module's own PROLOG.4 header above; the
' name survives only in this historical paragraph): NOT `VLA_Relation.
' RelJoin`'s hash join, a real non-reuse checked before writing any
' code - a hash join keys on exact equality; unification can BIND a
' variable, an asymmetric operation a hash-equality key can't express.
' Instead, a plain recursive enumeration over each conjunct's own
' candidate facts in turn, threading ONE substitution environment
' across the whole attempt (`VLA_Unify.UnifyEnvClone` gives each
' candidate its own independent copy, so a failed or partial attempt
' can never leak a binding into a sibling one) - a real, working
' Cartesian-product-with-early-pruning MVP, exactly as predicted here,
' PROLOG.4's own SolveGoalList (below) superseded the enumeration
' strategy itself while carrying the underlying `UnifyTwoWay` primitive
' forward unchanged, per this paragraph's own original prediction.
'
' Facts are a bag, not deduped at storage time (`SQL`'s own "a table may
' legitimately contain duplicate rows" stance, not `DATALOG`'s
' necessarily-deduped fixpoint-derived relations) - a fact authored
' twice produces two identical proof paths and two identical rows in
' the output, matching real Prolog's own behavior (which does not
' auto-dedupe `findall`-style answers either), not an oversight.
'
' `=PROLOG(clauses, ParamArray tables())` - the signature
' `docs/BETA_ROADMAP2.md`'s own top paragraph froze 2026-08-31, ahead of
' this item existing, `SQL.1`'s own precedent (a real signature pinned
' before the item that first exposes it is built). `tables` was reserved
' for PROLOG.6's own live-table-sourced facts and refused by name here
' (`prolog-tables-not-yet-supported`) at this item's own build time,
' rather than silently ignored the moment anything was actually passed -
' `SQL.1`'s own "exactly one table for now" bluntness, the same shape.
' SUPERSEDED at PROLOG.6's own build time (this module's own header, top
' of file) - `tables` is real now, and the refusal id is retired.
'
' LAYER:     1 (VLA_Unify, VLA_Relation) - VLA_Relation joined PROLOG.5.1,
'            for its pure mechanical arithmetic primitive only, never for
'            fact/tuple storage; see this item's own header above for why
'            that is a genuinely new, deliberate dependency and not a
'            reversal of PROLOG.1-4's "NOT VLA_Relation" stance.
' MAY CALL:  VLA (VlaReadForms/VlaWriteForm), VLA_Identity (Fold -
'            predicate-NAME lookups only, never a fact's own argument
'            text, VLA_Unify.bas's own PROLOG.2 header has the
'            reasoning), VLA_Runtime (VlaDictNew/Set/Get/Has - predicate-
'            name-keyed fact storage and arity tracking only), VLA_Unify
'            (UnifyTwoWay/EnvWalkInto/UnifyEnvClone/IsVarAtom),
'            VLA_Relation (ComputeArithmetic/IsInvariantNumericString/
'            InvariantVal - PROLOG.5.1's own `is`, mechanical only;
'            PROLOG.6 adds TableArgResolve/RangeToRows/RangeColumnNames/
'            ValueIsNumericType - RangeToRows genuinely reads table ROWS
'            now, but never RelFromRange/a Relation object - PROLOG's own
'            facts are still never stored as a Relation, this module's
'            own PROLOG.3 stance unchanged, only the ROW-READING half of
'            VLA_Relation's own surface is now touched, not its
'            tuple/dedup storage half), VLA_Messages (RaiseMsg).
' SHIPS:     add-in only - the `=PROLOG(...)` worksheet function,
'            callable the moment the add-in loads, no VBProject trust.
' PAYS INTO: PROLOG.5.2-5.4 (negation-as-failure and findall reuse this
'            item's own choice-point-shaped recursion and
'            PROLOG_MAX_STEPS budget directly - a bounded sub-call is a
'            SolveGoalList invocation like any other; cut prunes
'            alternatives already on the call stack); PROLOG.6
'            (table-sourced facts join the same clauseDict, shipped).
' REASON:    the smallest slice that turns PROLOG.4's own real Prolog
'            engine into one that can compute, not just match and
'            recurse - before cut/negation/findall exist to prune or
'            extend the search space arithmetic alone doesn't touch.

' ---------------------------------------------------------------------
'  Parsing
' ---------------------------------------------------------------------

' A top-level form's own head symbol, folded - VLA_Datalog.bas's own
' TopHead, repeated (no direct dependency between the two modules, a
' four-line rule not worth coupling for).
Private Function TopHead(ByVal form As Variant) As String
    If Not IsObject(form) Then VLA_Messages.RaiseMsg "prolog-top-form-not-a-list"
    Dim lst As Collection
    Set lst = form
    If lst.Count < 1 Then VLA_Messages.RaiseMsg "prolog-top-form-empty"
    If IsObject(lst.Item(1)) Then VLA_Messages.RaiseMsg "prolog-top-form-bad-head"
    TopHead = VLA_Identity.Fold(CStr(lst.Item(1)))
End Function

' Cross-definition arity consistency - VLA_Datalog.bas's own
' RecordArity, repeated (same reasoning as TopHead above).
Private Sub RecordArity(ByVal predArity As Object, ByVal predName As String, ByVal arity As Long)
    If VLA_Runtime.VlaDictHas(predArity, predName) Then
        Dim prev As Long
        prev = VLA_Runtime.VlaDictGet(predArity, predName)
        If prev <> arity Then
            VLA_Messages.RaiseMsg "prolog-arity-mismatch", "predicate", predName, "a", prev, "b", arity
        End If
    Else
        VLA_Runtime.VlaDictSet predArity, predName, arity
    End If
End Sub

' A term's own predicate name (folded), validated as a real predicate
' application - (name arg1 arg2 ...), zero or more args (a nullary
' predicate like (fact (raining)) is legal here: unlike DATALOG, this
' module never sizes an array off arity, so DATALOG's own arity-0
' restriction - a RelNew ReDim hazard, VLA_Relation.bas's own header -
' simply does not apply). Registers/checks arity as a side effect.
Private Function TermPredName(ByVal term As Variant, ByVal ctx As String, predArity As Object) As String
    If Not IsObject(term) Then VLA_Messages.RaiseMsg "prolog-atom-not-a-list", "context", ctx
    Dim lst As Collection
    Set lst = term
    If lst.Count < 1 Then VLA_Messages.RaiseMsg "prolog-atom-empty", "context", ctx
    If IsObject(lst.Item(1)) Then VLA_Messages.RaiseMsg "prolog-predicate-name-not-symbol", "context", ctx
    Dim predName As String
    predName = VLA_Identity.Fold(CStr(lst.Item(1)))
    RecordArity predArity, predName, lst.Count - 1
    TermPredName = predName
End Function

' Does term contain a variable anywhere in its own ARGUMENT positions
' (recursively, through nested compound terms) - never checking
' position 1 at any nesting level, since that is always a functor/
' predicate name, never a data argument (real Prolog's own rule:
' compound(a, b)'s own head symbol is never itself a variable
' position). Facts must be fully ground at this item's own scope - a
' variable only ever belongs in a (query ...).
Private Function TermHasVariable(ByVal term As Variant, ByRef outVarName As String) As Boolean
    If Not IsObject(term) Then
        If VLA_Unify.IsVarAtom(CStr(term)) Then
            outVarName = CStr(term)
            TermHasVariable = True
        End If
        Exit Function
    End If
    Dim lst As Collection
    Set lst = term
    Dim i As Long
    For i = 2 To lst.Count
        If TermHasVariable(lst.Item(i), outVarName) Then
            TermHasVariable = True
            Exit Function
        End If
    Next i
End Function

' Collects every distinct variable name appearing in term's own
' argument positions (same position-1-is-never-a-variable rule as
' TermHasVariable), in first-occurrence order - a query's own free
' variable list, and this program's own output column order (SD-4's
' frozen contract: header row first, one column per free query
' variable).
Private Sub CollectVars(ByVal term As Variant, freeVarNames As Collection, ByVal isGoalPosition As Boolean)
    If Not IsObject(term) Then
        Dim s As String
        s = CStr(term)
        If VLA_Unify.IsVarAtom(s) Then
            ' PROLOG.6: a keyed atom's own auto-generated anonymous-
            ' column variable (DesugarBodyItem's own "VlaAnon..." naming,
            ' below - a query-level "don't care" slot for a table column
            ' the user's own keyed atom never mentioned) must never
            ' surface as a phantom output column - the identical "must
            ' never leak" instinct already applied to not/findall's own
            ' discarded sub-search bindings (PROLOG.5.2's own header),
            ' applied here to a variable that DOES bind in the shared,
            ' non-discarded env but was never something the user asked
            ' to see.
            If Left$(s, 7) = "VlaAnon" Then Exit Sub
            If Not VarAlreadyCollected(freeVarNames, s) Then freeVarNames.Add s
        End If
        Exit Sub
    End If
    Dim lst As Collection
    Set lst = term

    ' PROLOG.11: EVERY skip-shape below is a statement about a GOAL, and
    ' therefore applies only where a goal actually is - the top of a query
    ' conjunct. Below that, every position this procedure can reach is
    ' DATA, and data has no `not`, no `\==` and no `atom?`; it has a
    ' compound term whose functor happens to be spelled that way, whose
    ' variables bind like any other argument's and are real output
    ' columns. Skipping them there dropped a column the query genuinely
    ' produced.
    '
    ' The reason "below here is always data" holds, checked rather than
    ' assumed: the only nested GOAL positions in this language are `not`'s
    ' own argument and `findall`'s own Goal, and both are skipped outright
    ' rather than descended into - so no recursive call from here can ever
    ' land on a goal. That is why every one of them passes False, and why
    ' this needs a flag rather than a depth counter: it is not "how deep"
    ' but "is this a goal", and the answer below the top is always no.
    '
    ' Found while writing PROLOG.9, which widened the fault by adding six
    ' more names to skip on, and filed then rather than folded in because
    ' it changes `not`'s long-shipped behaviour in the nested case. The
    ' question mark PROLOG.9 later took made a data functor named `atom?`
    ' implausible, but `(likes X (not Y))` was always reachable and always
    ' wrong.
    If isGoalPosition Then
        ' PROLOG.5.2: `(not Goal)` - Goal's own variables are never collected
        ' as query output columns here. Negation-as-failure discards every
        ' binding made while proving Goal regardless of success or failure
        ' (SolveNegation, below), so a variable appearing ONLY inside a
        ' negated goal can never resolve to anything meaningful outside it -
        ' collecting it would surface its own unresolved atom name as a
        ' phantom output column instead of a real value (hand-traced in this
        ' module's own PROLOG.5.2 header). A variable that ALSO appears in an
        ' ordinary, non-negated conjunct elsewhere in the query is still
        ' collected correctly, from THAT other occurrence - this case only
        ' ever skips descending into `not`'s own argument, it never removes
        ' an already-collected name. IsObject(lst.Item(1)) checked before
        ' Fold(CStr(...)) - this module's own IsObject-first discipline,
        ' since lst.Item(1) may itself be a nested list for an arbitrary
        ' compound-term argument CollectVars recurses into (e.g. a fact's own
        ' `(color red)`-shaped argument), never assumed to be a plain symbol
        ' just because a `not`-headed term happens to be 2 long.
        ' PROLOG.9: the six type tests join this same 2-long skip, on exactly
        ' the reasoning `not` is skipped on. They bind NOTHING - every one of
        ' them is a test - so a variable appearing only inside one can never
        ' resolve to anything, and collecting it would spill its own raw atom
        ' name into the cells as though it were a value the query had found.
        '
        ' The hazard is real and not hypothetical, and it is sharper here
        ' than for `not`: `(query (var X))` SUCCEEDS - a free X is exactly
        ' what `var` is true of - so without this skip the one query most
        ' likely to be typed while learning the predicate would answer with a
        ' single column headed X containing the text "X". That is PROLOG.5.2's
        ' own phantom-column finding and PROLOG.8's own `(query (\== X Y))`
        ' repeat, reached a third time by a third route.
        '
        ' A variable that ALSO appears in an ordinary conjunct is still
        ' collected from THAT occurrence - `(query (emp N S) (number S))`
        ' reports both columns, because this skip only ever declines to
        ' DESCEND into the type test, it never removes an already-collected
        ' name.
        '
        ' KNOWN LIMIT, inherited rather than introduced: this skip is keyed on
        ' shape alone, so a compound term used as DATA whose functor happens
        ' to be one of these names - `(query (likes X (atom? Y)))` against a
        ' fact storing an `(atom? foo)` argument - has its Y dropped from the
        ' output columns though it genuinely binds. `not` and the three
        ' non-binding term-matching operators already behave this way.
        '
        ' The trailing question mark all but closes it for THESE six. Written
        ' bare, `atom` and `number` would have been plausible data functors
        ' and PROLOG.9 would have widened this limit measurably; `atom?` is
        ' not a name anyone reaches for when inventing data, so the six added
        ' here are now among the least likely of the reserved set to collide
        ' rather than the most. An unplanned second dividend of the spelling
        ' decision, recorded because the first version of this comment
        ' claimed the opposite and would otherwise have stayed wrong.
        '
        ' Fixing it properly still means tracking goal-versus-data position
        ' here the way CollectTemplateVars (below) already does for findall's
        ' Template, which is a change to this function's contract and to
        ' `not`'s long-shipped behaviour - its own item, not a fold-in.
        If lst.Count = 2 Then
            If Not IsObject(lst.Item(1)) Then
                Dim headTwo As String
                headTwo = VLA_Identity.Fold(CStr(lst.Item(1)))
                If headTwo = "not" Then Exit Sub
                If TypeTestKindFor(headTwo) <> "" Then Exit Sub
            End If
        End If
        ' PROLOG.8: `(\= A B)`, `(== A B)` and `(\== A B)` - the THREE of the
        ' four term-matching goals that bind nothing - are skipped on exactly
        ' the reasoning `not` (above) is skipped on, and `(= A B)`, the one
        ' that DOES bind, is deliberately not.
        '
        ' The hazard is real and not hypothetical: `(query (\== X Y))`
        ' SUCCEEDS - two distinct free variables are trivially not identical -
        ' so without this skip X and Y would be collected as output columns
        ' and then resolve to nothing, rendering their own raw atom names "X"
        ' and "Y" into the spilled cells as though those were values the
        ' query had found. That is PROLOG.5.2's own phantom-column finding,
        ' reached by a different route.
        '
        ' `=` is the deliberate exception because it genuinely binds in the
        ' shared, non-discarded environment - the same property that earns
        ' `is`'s own target variable and findall's own Bag their columns.
        ' `\=` is NOT an exception despite running the same binding primitive:
        ' it succeeds only when that unification fails, and runs it against a
        ' throwaway clone regardless, so it has nothing to contribute either.
        ' The fork is UnificationBindsOutward (below), asked rather than
        ' re-derived, so this skip and the dispatch's own clone-or-thread
        ' decision can never disagree about which of the four keep what they
        ' bind. Descending is the default: only a head word that table knows
        ' is ever skipped, and a variable that ALSO appears in an ordinary
        ' conjunct is still collected from THAT occurrence, exactly as with
        ' `not`.
        If lst.Count = 3 Then
            If Not IsObject(lst.Item(1)) Then
                Dim unifyKind As String
                unifyKind = UnificationOpFor(VLA_Identity.Fold(CStr(lst.Item(1))))
                If unifyKind <> "" Then
                    If Not UnificationBindsOutward(unifyKind) Then Exit Sub
                End If
            End If
        End If
        ' PROLOG.5.3: `(findall Template Goal Bag)` - the same reasoning as
        ' `not` above, applied to TWO argument positions instead of one:
        ' Template's own variables only ever matter INSIDE the isolated
        ' harvest sub-search (SolveIsolated, below - a fresh, separate
        ' collection, CollectTemplateVars, feeds THAT), and Goal's are
        ' discarded exactly like `not`'s own Goal. Only Bag - the one
        ' argument genuinely bound in the shared, non-discarded outer env,
        ' exactly like `is`'s own target Var - is descended into here.
        '
        ' PROLOG.9: `(between Low High X)` is ALSO 4 long and deliberately
        ' gets NO arm here - it wants the default descend below, and stating
        ' why is worth more than an arm that would do nothing. `between` is
        ' the one goal PROLOG.9 adds that genuinely BINDS in the shared,
        ' non-discarded environment, so X is a real output column exactly the
        ' way `is`'s own target variable and findall's own Bag are, and the
        ' default already collects it. Low and High are contributed only when
        ' they are variables, which is harmless in both directions: bound
        ' elsewhere, they were already collected from that occurrence and
        ' VarAlreadyCollected dedupes; never bound at all, EvalArithTerm
        ' refuses the goal by name before any row is produced, so no phantom
        ' column can survive to be rendered.
        If lst.Count = 4 Then
            If Not IsObject(lst.Item(1)) Then
                If VLA_Identity.Fold(CStr(lst.Item(1))) = "findall" Then
                    CollectVars lst.Item(4), freeVarNames, False
                    Exit Sub
                End If
            End If
        End If
    End If
    ' False on every recursive call, without exception - see the
    ' PROLOG.11 note above for why no position reachable from here is
    ' ever a goal.
    Dim i As Long
    For i = 2 To lst.Count
        CollectVars lst.Item(i), freeVarNames, False
    Next i
End Sub

' PROLOG.5.3: `findall`'s own Template is plain DATA, never a goal - it
' must never be goal-interpreted the way CollectVars (above) treats a
' query conjunct, so this is a separate, deliberately simpler collector
' with none of CollectVars' own `(not Goal)`/`(findall ...)` skip-shape
' special-casing: a Template that happens to itself CONTAIN a literal
' `(not X)` or `(findall ...)` sub-term (e.g. `(findall (not X) (foo X)
' Bag)`, building a list of negation-shaped terms as ordinary data - a
' legitimate real-Prolog idiom) must still have every one of its own
' variables collected honestly - the same position-1-is-never-a-variable
' rule CollectVars itself still follows, otherwise unchanged.
Private Sub CollectTemplateVars(ByVal term As Variant, freeVarNames As Collection)
    If Not IsObject(term) Then
        If VLA_Unify.IsVarAtom(CStr(term)) Then
            If Not VarAlreadyCollected(freeVarNames, CStr(term)) Then freeVarNames.Add CStr(term)
        End If
        Exit Sub
    End If
    Dim lst As Collection
    Set lst = term
    Dim i As Long
    For i = 2 To lst.Count
        CollectTemplateVars lst.Item(i), freeVarNames
    Next i
End Sub

Private Function VarAlreadyCollected(freeVarNames As Collection, ByVal varName As String) As Boolean
    Dim v As Variant
    For Each v In freeVarNames
        If CStr(v) = varName Then
            VarAlreadyCollected = True
            Exit Function
        End If
    Next v
End Function

' PROLOG.5.3: reconstructs one instantiated Template per harvested
' solution - Template's own structure, walked the identical way
' FreshenTerm (below) walks a clause (skip position 1, recurse into
' 2..N), but substituting each variable LEAF for its own resolved value
' (looked up by name in templateVars/solutionTuple, the SAME parallel
' pair SolveGoalList's own base case already produces for any
' freeVarNames list, ResolveTermDeep-fixed - see this module's own
' PROLOG.5.3 header) rather than appending a fresh suffix. IsObject
' branched explicitly before returning a looked-up value - the same
' "bare Variant assignment invokes an object's own default member" trap
' EnvWalkInto's own header documents, since solutionTuple.Item(k) may
' hold either a ground leaf or a resolved compound term. A variable NOT
' found in templateVars can't actually occur here - templateVars was
' collected FROM this exact template (CollectTemplateVars, above) - but
' the lookup fails safely (leaves the atom as its own literal text)
' rather than assuming that invariant blindly.
Private Function SubstituteTemplate(ByVal template As Variant, templateVars As Collection, solutionTuple As Collection) As Variant
    If Not IsObject(template) Then
        Dim s As String
        s = CStr(template)
        If VLA_Unify.IsVarAtom(s) Then
            Dim k As Long
            For k = 1 To templateVars.Count
                If CStr(templateVars.Item(k)) = s Then
                    If IsObject(solutionTuple.Item(k)) Then
                        Set SubstituteTemplate = solutionTuple.Item(k)
                    Else
                        SubstituteTemplate = solutionTuple.Item(k)
                    End If
                    Exit Function
                End If
            Next k
        End If
        SubstituteTemplate = template
        Exit Function
    End If
    Dim lst As Collection, outLst As New Collection
    Set lst = template
    outLst.Add lst.Item(1)
    Dim i As Long
    For i = 2 To lst.Count
        outLst.Add SubstituteTemplate(lst.Item(i), templateVars, solutionTuple)
    Next i
    Set SubstituteTemplate = outLst
End Function

Private Function GetOrCreateClauseList(ByVal clauseDict As Object, ByVal predName As String) As Collection
    If VLA_Runtime.VlaDictHas(clauseDict, predName) Then
        Set GetOrCreateClauseList = VLA_Runtime.VlaDictGet(clauseDict, predName)
    Else
        Dim lst As New Collection
        VLA_Runtime.VlaDictSet clauseDict, predName, lst
        Set GetOrCreateClauseList = lst
    End If
End Function

' A clause record: a 2-item Collection, Item(1) the head term, Item(2) a
' Collection of body terms (empty for a fact). One predicate's whole
' clause list - facts and rules together, in written order - lives in
' ONE Collection per predicate name in clauseDict (PROLOG.4: generalized
' from PROLOG.3's own fact-only Collection-of-ground-terms; a fact is now
' simply the degenerate zero-body case, not a separate representation).
Private Function MakeClause(ByVal headTerm As Variant, ByVal bodyItems As Collection) As Collection
    Dim rec As New Collection
    rec.Add headTerm
    rec.Add bodyItems
    Set MakeClause = rec
End Function

' `is`/`not`/`findall`/`!` and PROLOG.7's own six comparison operators may
' never be a predicate's own NAME - checked at every (fact ...)/(rule ...)
' DEFINITION site, never at a body item or query conjunct's own USE site,
' where each has its own special meaning instead (ValidateBodyItem,
' below). The four word-shaped names were reserved together at PROLOG.5.1
' even though only `is` was built then, so a knowledge base written
' against 5.1 could never be silently broken once the others shipped; the
' six operators are reserved here on exactly that precedent, one stage
' before SolveGoalList learns to dispatch them.
'
' Reserving them is also what turns DesugarBodyItem's own pass-through
' from an accident into a guarantee. A comparison goal survives desugaring
' today only because DesugarPredicateAtom bails when the head word is not
' a known table name - so without this, a predicate genuinely named ">"
' could capture a comparison goal instead of the dispatch arm getting it.
'
' Matched case-sensitively (this module sets no Option Compare Text),
' which is exact for the six - none of them contains a letter - and
' correct for the four word-shaped names at the (fact ...)/(rule ...)
' sites, where TermPredName has already folded predName before it arrives.
' The table-argument site (PROLOG, below) passes TableArgResolve's own raw
' name instead, but an Excel Table cannot be named "<" or "=:=", so the
' six are unaffected by that difference.
Private Function IsReservedPredicateName(ByVal predName As String) As Boolean
    Select Case predName
    Case "is", "not", "findall", "!", "between"
        ' PROLOG.9: `between` is a literal arm rather than a table of its
        ' own because it is ONE name - the shape is/not/findall already
        ' have. Its six type-test siblings arrive by table below, since
        ' they are six.
        IsReservedPredicateName = True
    Case Else
        ' PROLOG.7: the six comparison names are not repeated here. They
        ' live in exactly one place - ComparisonOpFor's own table, below -
        ' so reserving them and dispatching them can never disagree about
        ' which six they are. A second literal list would be free to drift
        ' from the first, and a name reserved but not dispatched (or the
        ' reverse) is precisely the silent failure this module's own
        ' forward-reservation discipline exists to prevent.
        '
        ' PROLOG.8: its own four names join on the identical terms, from
        ' their own table. Or, not OrElse-style short-circuiting - VBA has
        ' none - but both operands are pure lookups over a frozen Select
        ' Case, so evaluating the second when the first already answered
        ' True costs a jump and can have no effect of its own.
        '
        ' PROLOG.9: its own six type-test names join on the identical
        ' terms, from their own table. Three delegated tables now, and
        ' the same non-short-circuit note applies unchanged - all three
        ' are pure lookups over frozen Select Cases, so evaluating a
        ' later one after an earlier already answered True costs a jump
        ' and can have no effect of its own.
        '
        ' tools/check_prolog_reserved_names.ps1 reads the delegated table
        ' names straight out of this expression, then holds all three
        ' places that spell the reserved set - here, SolveGoalList's own
        ' dispatch, and prolog-reserved-predicate-name's own text - to
        ' naming the same predicates. PROLOG.9 taught it to walk the
        ' dispatch BACK to this function as well (its own rule D): before
        ' that it only ever walked outward from here, so an arm added to
        ' SolveGoalList without a matching reservation right here passed
        ' clean - verified by mutation, and precisely the mistake two new
        ' families at once invites.
        IsReservedPredicateName = (ComparisonOpFor(predName) <> "" Or UnificationOpFor(predName) <> "" Or TypeTestKindFor(predName) <> "" Or TypeTestIsoSpellingFor(predName) <> "")
    End Select
End Function

' PROLOG.7: a comparison goal's own predicate name -> the operator spelling
' VLA_Relation.CompareValues already understands, or "" if predName is not
' one of the six at all. A pure translation table, and the SINGLE place the
' six names are written down: IsReservedPredicateName (above),
' ValidateBodyItem, DesugarBodyItem and SolveGoalList's own dispatch all
' ask this function rather than repeating the list.
'
' Only the spelling differs between the two vocabularies. Prolog writes
' `=<` where this project's shared comparison substrate writes `<=` (real
' Prolog's own spelling, chosen so `=<` cannot be misread as an arrow),
' and Prolog's `=:=`/`=\=` are that substrate's ordinary `=`/`<>` - the
' NUMERIC equality pair, never `=`/`\=`, which are unification and belong
' to PROLOG.8. `<` and `>` coincide in both and are still routed through
' here rather than passed straight along, so no caller has to know which
' of the six happen to need translating.
Private Function ComparisonOpFor(ByVal predName As String) As String
    Select Case predName
    Case "<":   ComparisonOpFor = "<"
    Case ">":   ComparisonOpFor = ">"
    Case "=<":  ComparisonOpFor = "<="
    Case ">=":  ComparisonOpFor = ">="
    Case "=:=": ComparisonOpFor = "="
    Case "=\=": ComparisonOpFor = "<>"
    End Select
End Function

' PROLOG.8: a unification goal's own predicate name -> the KIND of test it
' performs, or "" if predName is not one of the four at all. The exact
' single-source shape ComparisonOpFor (above) established, and for the
' identical reason: IsReservedPredicateName, ValidateBodyItem,
' DesugarBodyItem, CollectVars and SolveGoalList's own dispatch all ask
' this function rather than repeating the list, so reserving a name and
' dispatching it can never disagree about which four they are.
'
' A KIND rather than an operator spelling, unlike its comparison sibling.
' These four do not reduce to one shared operation with a parameter the
' way the six comparisons reduce to CompareValues plus a spelling: they
' fork on two independent axes, and the caller needs both.
'
'   BINDING     - "unify"/"notunify" run UnifyTwoWay, which BINDS free
'                 variables; "identical"/"notidentical" run TermsIdentical,
'                 which never binds anything at all.
'   POLARITY    - "notunify"/"notidentical" are the negated twins.
'
' The binding axis is the one that matters most outside this table.
' `(= X 1)` genuinely binds X in the shared, non-discarded environment -
' so, alone among the four, its variables ARE real query output columns
' (CollectVars, above, descends into it and skips the other three), and
' its dispatch must thread a FRESH env clone into the continuation the way
' `is` does rather than the caller's own the way `not` does. Getting that
' fork backwards leaks a failed attempt's bindings into a sibling goal, or
' loses a successful one's.
Private Function UnificationOpFor(ByVal predName As String) As String
    Select Case predName
    Case "=":   UnificationOpFor = "unify"
    Case "\=":  UnificationOpFor = "notunify"
    Case "==":  UnificationOpFor = "identical"
    Case "\==": UnificationOpFor = "notidentical"
    End Select
End Function

' PROLOG.8: True for the ONE of the four whose bindings survive into the
' rest of the query - the single predicate every caller that cares about
' that asks, so the judgement is written down once rather than re-derived
' from a kind string at each site.
'
' "Outward" is the load-bearing word, and the reason this is not simply
' `kind = "unify" Or kind = "notunify"`. TWO of the four run UnifyTwoWay,
' the binding primitive - but `\=` succeeds precisely when that
' unification FAILS, and it runs the attempt against a throwaway clone
' either way, so nothing it binds ever reaches the continuation. Only `=`
' both binds and keeps what it bound. Reading this predicate as "runs the
' binding primitive" instead of "keeps its bindings" would put `\=`'s
' variables back into the query's own output columns as the phantom
' names CollectVars (above) exists to keep out, and would tell
' SolveGoalList to thread a clone forward from a goal that must not
' change the environment at all.
Private Function UnificationBindsOutward(ByVal kind As String) As Boolean
    UnificationBindsOutward = (kind = "unify")
End Function

' PROLOG.9: a type-test goal's own predicate name -> the kind of question
' it asks, or "" if predName is not one of the six at all. The same
' single-source shape ComparisonOpFor and UnificationOpFor (above)
' established, and for the identical reason: IsReservedPredicateName,
' ValidateBodyItem, DesugarBodyItem, CollectVars and SolveGoalList's own
' dispatch all ask this function rather than repeating the list, so
' reserving a name and dispatching it can never disagree about which six
' they are.
'
' A TRANSLATION table, exactly like ComparisonOpFor - the written name
' carries a trailing "?" that the internal kind does not, so `atom?` maps
' to the kind `atom`. SolveTypeTest (below) therefore never sees a
' question mark at all and its Select Case is spelled in kinds, which is
' what keeps the naming decision confined to this one table.
'
' WHY THE QUESTION MARK, since real Prolog writes `atom(X)`. This engine
' is an S-EXPRESSION language, and in one of those `(atom bob)` is
' visually identical to a compound DATA term - `(color red)`, `(f a)`,
' `(name Alice)` are all ordinary data in exactly that shape, and this
' module's own tests classify `(f a)` as data two lines from where they
' classify with `atom?`. Real Prolog has no such ambiguity because its
' syntax separates a goal from a term by position; ours does not, so the
' name has to carry the distinction instead.
'
' It is also this codebase's OWN existing convention rather than an
' import: VLA.bas's macro layer reserves car, cdr, cons, list, null?,
' eq? and equal?, and the three that end in "?" are exactly the three
' that ask a yes/no question. The rule already in force here is that an
' alphabetic primitive asking a question ends in "?" and one that
' produces a value does not - which is also why `between` (below) keeps
' no question mark: it GENERATES, like cons and list, and only tests as
' a second mode.
'
' The six bare ISO spellings are reserved too, and dispatched to a
' refusal that names the "?" form - TypeTestIsoSpellingFor, below. A
' Prolog author's first instinct is `(atom X)`, and an unreserved
' `(atom X)` would be an unknown predicate, which in SolveGoalList is a
' SILENT dead end - zero rows and no explanation, the exact
' confidently-wrong-answer class this project holds to be worse than a
' crash. So the ISO spelling teaches instead of failing.
'
' tools/check_prolog_reserved_names.ps1 reads this table's own Case arms
' to count the set it must find advertised.
Private Function TypeTestKindFor(ByVal predName As String) As String
    Select Case predName
    Case "var?":      TypeTestKindFor = "var"
    Case "nonvar?":   TypeTestKindFor = "nonvar"
    Case "atom?":     TypeTestKindFor = "atom"
    Case "number?":   TypeTestKindFor = "number"
    Case "atomic?":   TypeTestKindFor = "atomic"
    Case "compound?": TypeTestKindFor = "compound"
    End Select
End Function

' PROLOG.9: a bare ISO type-test name -> the spelling this engine
' actually uses, or "" if predName is not one of the six. Reserved and
' dispatched exactly like a real goal, but its dispatch RAISES rather
' than solves: `(atom X)` is refused by name and told to write
' `(atom? X)`.
'
' This table exists so that being wrong about the spelling is LOUD.
' Without it the six bare names would be ordinary unknown predicates,
' and an unknown predicate is a silent dead end by design (see
' SolveGoalList's own clauseDict lookup) - so a Prolog author writing
' the spelling their own language taught them would get an empty result
' and nothing to read. That is the same move IN.15 and PROLOG.7 already
' made, and the same one PROLOG.10's recommended option D proposes.
'
' Reserved as well as dispatched, not merely dispatched: a name the
' solver acts on but the parser does not reserve is a predicate a user
' can DEFINE and then have silently shadowed - the defect
' tools/check_prolog_reserved_names.ps1's own rule D exists to catch, and
' the reason these six appear in IsReservedPredicateName's Case Else
' alongside the other three tables.
Private Function TypeTestIsoSpellingFor(ByVal predName As String) As String
    Select Case predName
    Case "var":      TypeTestIsoSpellingFor = "var?"
    Case "nonvar":   TypeTestIsoSpellingFor = "nonvar?"
    Case "atom":     TypeTestIsoSpellingFor = "atom?"
    Case "number":   TypeTestIsoSpellingFor = "number?"
    Case "atomic":   TypeTestIsoSpellingFor = "atomic?"
    Case "compound": TypeTestIsoSpellingFor = "compound?"
    End Select
End Function

' PROLOG.9: THE one place this module decides whether a ground leaf is a
' NUMBER or an ATOM - which is to say, the one place PROLOG.9 reads the
' quoted-string marker at all. `atom`, `number` and `atomic` all ask it,
' so the judgement is written down once rather than re-derived at three
' sites, exactly the reason UnificationBindsOutward (above) exists.
'
' A leaf carrying the marker is an ATOM, never a number, whatever its
' text says. `(number "42)` is False even though the 42 is right there.
' This module's own PROLOG.9 header has the full reasoning and the
' measured blast radius; the short version is that PROLOG.8 shipped
' `"42 \== 42`, so answering True here would classify a term into a class
' containing no term it is identical to. It is PROLOG.9's LOCAL reading
' of a question PROLOG.10 owns and has not answered, compatible with
' PROLOG.10's options A and D and reopened by its option B.
'
' raw, never LeafText(raw): the marker is the whole signal, so stripping
' it first would destroy exactly the thing being tested. EvalArithTerm
' (below) makes the opposite choice deliberately and documents it there -
' it strips, because arithmetic wants the value; this asks about
' identity, and identity is what the marker carries.
Private Function LeafIsNumberTerm(ByVal raw As String) As Boolean
    If Left$(raw, 1) = Chr$(34) Then Exit Function
    LeafIsNumberTerm = VLA_Relation.IsInvariantNumericString(raw)
End Function

' PROLOG.9: a Double -> the ground numeric leaf this engine represents it
' as. Str$, never CStr: Str$ is genuinely locale-invariant in VBA (always
' "." for the decimal point, plus a leading space for a non-negative
' number that Trim$ removes), where CStr follows the machine's locale and
' would produce "3,5" on a comma-decimal machine.
'
' This is TableCellToTerm's own expression (below), deliberately, and the
' duplication is the point rather than a missed hoist: a value `between`
' generates must be the SAME TEXT a numeric table cell of that value
' becomes, or `(between 1 3 X) (emp X)` would silently match nothing
' against a table whose id column holds 1, 2 and 3. Two spellings of a
' number are two different ground atoms to UnifyTwoWay, which compares
' atoms text-for-text.
Private Function NumberToTerm(ByVal v As Double) As String
    NumberToTerm = Trim$(Str$(v))
End Function

' An `is`-expression's own STATIC shape, checked recursively at parse
' time - every operator symbol is one of the frozen four (+/-/*//) and
' has exactly two operands, exactly `ComputeArithmetic`'s own contract.
' Deliberately does NOT check whether a leaf is numeric-looking - a leaf
' may be a variable, unbound now and bound to a real number only once
' solving actually reaches this goal, so numeric-ness is a RUNTIME-only
' question (EvalArithTerm, below) - checking it twice would risk the two
' checks drifting out of sync, not add real safety.
'
' PROLOG.7: formLabel is the form the CALLER is validating on behalf of -
' "(is ...)" from ValidateBodyItem's own `is` arm, "(> ...)" and its five
' siblings from the comparison arm. It exists only to be substituted into
' this procedure's own refusals: an operand expression is shape-identical
' whichever form encloses it, so nothing about the checking changes, but a
' user who wrote (> Salary 80000) must not be told about (is ...), a form
' they never wrote. Passed down every recursive call so a refusal from an
' arbitrarily nested operand still names the outermost form the user
' actually typed. Pinned by tools/check_prolog_form_attribution.ps1.
Private Sub ValidateArithExpr(ByVal term As Variant, ByVal formLabel As String)
    If Not IsObject(term) Then Exit Sub
    Dim lst As Collection
    Set lst = term
    ' lst.Count checked BEFORE any lst.Item(1) access - a genuinely
    ' empty () operand (real input a careless author could write, e.g.
    ' (is X (+ () 3))) would otherwise raise a raw "Subscript out of
    ' range" the instant Item(1) is touched.
    If lst.Count < 1 Then VLA_Messages.RaiseMsg "prolog-arith-wrong-arity", "op", "()", "form", formLabel
    ' IsObject-first, never CStr on a value that might be one - this
    ' project's own documented trap (feedback_vba_error_and_loop_gotchas
    ' gotcha 5), applied here since an operator position could itself be
    ' a nested form (a genuinely malformed expression, not a real term).
    If IsObject(lst.Item(1)) Then VLA_Messages.RaiseMsg "prolog-arith-unknown-operator", "op", "(a nested form)", "form", formLabel
    Dim op As String
    op = CStr(lst.Item(1))
    Select Case op
    Case "+", "-", "*", "/"
        ' recognized
    Case Else
        VLA_Messages.RaiseMsg "prolog-arith-unknown-operator", "op", op, "form", formLabel
    End Select
    If lst.Count <> 3 Then VLA_Messages.RaiseMsg "prolog-arith-wrong-arity", "op", op, "form", formLabel
    ValidateArithExpr lst.Item(2), formLabel
    ValidateArithExpr lst.Item(3), formLabel
End Sub

' Validates ONE body item / query conjunct - both are the same kind of
' goal list slot, so one function serves either ctx. A bare "!" is cut
' (PROLOG.5.4) - a legal, fully-formed body item/query conjunct needing
' no further shape validation at all (it never takes arguments;
' IsReservedPredicateName already forbids "!" from ever being DEFINED as
' a predicate, so no real ambiguity with a same-named predicate is
' possible either). An `(is Var Expr)` form validates Expr's own shape (ValidateArithExpr)
' and is otherwise left exactly as parsed - CollectVars's own existing
' generic recursion already collects Var/Expr's own variables correctly
' with no changes needed (position 1 - "is" - is never a variable
' position, exactly the same rule that already applies to any ordinary
' predicate's own name). A `(not Goal)` form (PROLOG.5.2) requires
' exactly one nested Goal and validates it by recursing into
' ValidateBodyItem itself - Goal may be an ordinary predicate
' application, an `(is ...)` form, or even another `(not ...)` (nested
' negation is not disallowed - it is logically sound, if unusual, and
' this module's own recursive dispatch already handles it for free with
' no extra code); a bare "!" inside `not`/`findall` is likewise legal
' (real Prolog's own opaque-cut rule - see this module's own PROLOG.5.4
' header for why SolveIsolated already contains it correctly), and a
' malformed Goal is caught by that SAME recursive call, reusing
' prolog-atom-not-a-list rather than inventing not-specific twins.
' CollectVars (below), not this function, is what stops a
' negated goal's own variables from becoming a free query-output column
' - see this module's own PROLOG.5.2 header for why that decision lives
' there instead of here. A `(findall Template Goal Bag)` form (PROLOG.5.3)
' requires exactly three arguments and validates only Goal by the SAME
' recursive call `not` already uses - Template and Bag are left exactly
' as parsed, unvalidated, the identical stance `is`'s own Var target
' already takes (any term is legal there; UnifyTwoWay sorts out whether
' it actually unifies at solve time). Anything else is an ordinary
' predicate application, TermPredName unchanged.
Private Sub ValidateBodyItem(ByVal item As Variant, ByVal ctx As String, predArity As Object)
    If Not IsObject(item) Then
        If CStr(item) = "!" Then Exit Sub   ' PROLOG.5.4: cut - legal, no further shape to check
        VLA_Messages.RaiseMsg "prolog-atom-not-a-list", "context", ctx
        Exit Sub
    End If
    Dim lst As Collection
    Set lst = item
    ' Nested Ifs, never one combined "count >= 1 And Not IsObject(...)"
    ' expression - VBA's And does not short-circuit, so lst.Item(1) would
    ' evaluate regardless of the count check and raise a raw, unworded
    ' error on a genuinely empty () body item (real input a careless
    ' author could write) - this module's own PROLOG.2-era header already
    ' documents this exact class of trap (UnifyTwoWay's own G0 fix).
    If lst.Count >= 1 Then
        If Not IsObject(lst.Item(1)) Then
            Dim headWord As String
            headWord = VLA_Identity.Fold(CStr(lst.Item(1)))
            If headWord = "is" Then
                If lst.Count <> 3 Then VLA_Messages.RaiseMsg "prolog-is-bad-shape"
                ValidateArithExpr lst.Item(3), "(is ...)"
                Exit Sub
            ElseIf headWord = "not" Then
                If lst.Count <> 2 Then VLA_Messages.RaiseMsg "prolog-not-bad-shape"
                ValidateBodyItem lst.Item(2), ctx, predArity
                Exit Sub
            ElseIf headWord = "findall" Then
                If lst.Count <> 4 Then VLA_Messages.RaiseMsg "prolog-findall-bad-shape"
                ValidateBodyItem lst.Item(3), ctx, predArity
                Exit Sub
            ElseIf ComparisonOpFor(headWord) <> "" Then
                ' PROLOG.7: exactly two operands, each validated as an
                ' arithmetic expression by the SAME ValidateArithExpr
                ' `is` already uses on its own right-hand side - an
                ' operand is shape-identical whichever form encloses it.
                ' A bare variable or number operand is left alone (that
                ' function returns immediately on a non-Collection),
                ' since numeric-ness is a RUNTIME question here for
                ' exactly the reason it is for `is`: an operand may be a
                ' variable that is still unbound at parse time.
                ' formLabel is this goal's own head word, so every
                ' refusal from either side names the comparison the user
                ' actually wrote rather than `(is ...)`.
                If lst.Count <> 3 Then VLA_Messages.RaiseMsg "prolog-comparison-bad-shape", "form", "(" & headWord & " ...)"
                ValidateArithExpr lst.Item(2), "(" & headWord & " ...)"
                ValidateArithExpr lst.Item(3), "(" & headWord & " ...)"
                Exit Sub
            ElseIf UnificationOpFor(headWord) <> "" Then
                ' PROLOG.8: exactly two operands, and - unlike the
                ' comparison arm directly above - NOTHING else checked
                ' about them. That difference is the whole distinction
                ' between the two families. A comparison's operands are
                ' arithmetic EXPRESSIONS, so ValidateArithExpr can hold
                ' them to the frozen +/-/*// operator set at parse time.
                ' These four take arbitrary TERMS: `(= X (f Y))` unifies
                ' against the compound term `(f Y)`, which is ordinary
                ' data, and running an arithmetic shape check over it
                ' would refuse a correct program for using an operator
                ' that was never meant to be arithmetic in the first
                ' place.
                '
                ' For the same reason the operands are NOT put through
                ' TermPredName either, so a compound operand's own head
                ' word records no arity: `(f Y)` here is a term being
                ' matched, not a call to a predicate named f, and letting
                ' it constrain the real f/N would be findall's own
                ' Template mistake (CollectTemplateVars, above, exists
                ' as a separate collector for exactly this reason).
                '
                ' An empty () operand needs no guard of its own - it
                ' reaches UnifyTwoWay/TermsIdentical as a zero-length
                ' Collection and simply fails to match anything that is
                ' not also one, no crash and no special case.
                If lst.Count <> 3 Then VLA_Messages.RaiseMsg "prolog-unification-bad-shape", "form", "(" & headWord & " ...)"
                Exit Sub
            ElseIf TypeTestKindFor(headWord) <> "" Then
                ' PROLOG.9: exactly ONE operand, and - like the
                ' term-matching arm directly above, and unlike the
                ' comparison arm above that - nothing else checked about
                ' it. A type test takes an arbitrary TERM and asks what
                ' shape it is; running ValidateArithExpr over it would
                ' refuse `(compound (f Y))` for using an operator that
                ' was never meant to be arithmetic, which is the whole
                ' point of the form.
                '
                ' Nor is the operand put through TermPredName: `(f Y)`
                ' inside `(compound (f Y))` is a term being CLASSIFIED,
                ' not a call to a predicate named f, and letting it
                ' constrain the real f/N would be findall's own Template
                ' mistake a second time.
                '
                ' One shared refusal for all six, so it cannot name a
                ' form of its own - it takes {form} from the head word
                ' the user actually wrote, exactly as the comparison and
                ' term-matching arms do. Pinned by
                ' tools/check_prolog_form_attribution.ps1, whose
                ' multi-form baseline this id was added to BEFORE the
                ' code existed, and which failed on it until it did.
                If lst.Count <> 2 Then VLA_Messages.RaiseMsg "prolog-type-test-bad-shape", "form", "(" & headWord & " ...)"
                Exit Sub
            ElseIf headWord = "between" Then
                ' PROLOG.9: `(between Low High X)` - four elements, so
                ' the same arity check every sibling arm makes, but with
                ' its two halves validated DIFFERENTLY, which is the
                ' whole shape of the form.
                '
                ' Low and High are arithmetic EXPRESSIONS - they go
                ' through the same ValidateArithExpr `(is ...)` and the
                ' six comparisons already use, so `(between 1 (+ N 1) X)`
                ' is legal and a malformed bound is refused at parse time
                ' naming "(between ...)" rather than a form the user
                ' never wrote.
                '
                ' X is NOT, and must not be. It is the term to bind or
                ' test, so ValidateArithExpr over it would refuse a
                ' perfectly ordinary `(between 1 10 X)`'s own bare
                ' variable's compound sibling - and, worse, would report
                ' an ARITHMETIC complaint about a position that holds no
                ' expression. Whether X is usable is a RUNTIME question
                ' (it may be unbound now and bound later, or unbound on
                ' purpose because generating is the point), decided in
                ' SolveBetween below - the identical reasoning `is`'s own
                ' target variable already gets.
                If lst.Count <> 4 Then VLA_Messages.RaiseMsg "prolog-between-bad-shape"
                ValidateArithExpr lst.Item(2), "(between ...)"
                ValidateArithExpr lst.Item(3), "(between ...)"
                Exit Sub
            End If
        End If
    End If
    TermPredName item, ctx, predArity
End Sub

' PROLOG.6: an unmentioned column's own gensym-free, deterministic
' anonymous variable name - derived only from the atom's own written
' occurrence (itemIndex, the SAME rule-body/query-conjunct loop index
' DATALOG.5's own AnonymousColumnVarName already uses this way) and the
' column's own 1-based header position, never a counter, this project's
' own standing veto on gensym-style identifier invention. anonPrefix
' distinguishes a rule-body occurrence ("VlaAnonB", freshened per
' invocation like any other rule-body variable, FreshenTerm's own
' existing generic recursion, no special handling needed) from a query
' conjunct occurrence ("VlaAnonQ", never freshened, but excluded from
' free-output-column collection by CollectVars' own new skip-shape check,
' above) - both share the "VlaAnon" stem that check matches against.
Private Function AnonymousColumnVarName(ByVal anonPrefix As String, ByVal itemIndex As Long, ByVal colIndex As Long) As String
    AnonymousColumnVarName = anonPrefix & itemIndex & "C" & colIndex
End Function

' Comma-joined ORIGINAL (unfolded) column names, for prolog-unknown-
' column's own teaching-refusal wording - VLA_Datalog.JoinOriginalNames'
' own exact shape, repeated (a four-line function, not worth coupling
' two modules for).
Private Function JoinOriginalNames(ByVal headerPairs As Collection) As String
    Dim r As String
    Dim j As Long
    For j = 1 To headerPairs.Count
        Dim hp As Collection
        Set hp = headerPairs.Item(j)
        If j > 1 Then r = r & ", "
        r = r & CStr(hp.Item(2))
    Next j
    JoinOriginalNames = r
End Function

' PROLOG.6: recognizes a rule-body item or query conjunct whose own
' arguments are ALL (header value) pairs against a KNOWN table-sourced
' predicate, and rewrites it into the exact positional form it stands
' for - VLA_Datalog.DesugarBodyAtomForm's own reuse-DATALOG.5-directly
' pattern, but PROLOG's own classification order is deliberately
' different (headerMap checked FIRST, before any argument shape is even
' examined) - see this module's own PROLOG.6 header for the real
' ambiguity this resolves that DATALOG never had to (PROLOG allows
' compound-term arguments as ordinary data everywhere; DATALOG forbids
' them entirely). Recurses into `not`/`findall`'s own nested Goal
' argument (never their own Template/Bag/Expr positions, which are
' either plain data or never goal-shaped); `is` is left entirely
' untouched (its own Expr is never a predicate reference). dest ends up
' completely UNCHANGED (a copy of item) whenever no desugaring applies -
' the overwhelmingly common case (an ordinary positional atom, or an
' atom whose predicate isn't table-sourced at all).
'
' A Sub with a ByRef dest out param, not a Function - live-caught while
' first writing this as a Function returning Variant: item/the rebuilt
' (not ...)/(findall ...) wrapper/the desugared atom are ALL genuinely
' objects (Collections) whenever desugaring is at all reachable (item is
' already confirmed an object by the point ANY of these branches runs),
' so `DesugarBodyItem = item`/`= DesugarPredicateAtom(...)` - a bare
' Function-return assignment - hit the identical "invoke the object's
' own default member instead of copying the reference" trap
' EnvWalkInto/ResolveTermDeep's own headers already document, the SAME
' class of bug, just on a function's own return value rather than a
' plain variable - a Function's return assignment is still an
' assignment STATEMENT, following the identical Let/Set disambiguation
' rule. Caught before this ever reached a live run, not live-caught this
' time - found by re-checking every new assignment against this
' project's own already-documented trap list before calling this done.
Private Sub DesugarBodyItem(ByRef dest As Variant, ByVal item As Variant, ByVal anonPrefix As String, ByVal itemIndex As Long, ByVal headerMap As Object)
    If Not IsObject(item) Then
        dest = item
        Exit Sub
    End If
    Dim lst As Collection
    Set lst = item
    If lst.Count >= 1 Then
        If Not IsObject(lst.Item(1)) Then
            Dim headWord As String
            headWord = VLA_Identity.Fold(CStr(lst.Item(1)))
            If headWord = "is" Then
                Set dest = item
                Exit Sub
            ElseIf ComparisonOpFor(headWord) <> "" Then
                ' PROLOG.7: passed through untouched, exactly as `is` is
                ' just above. A comparison's two operands are arithmetic
                ' expressions, never predicate applications, so there is
                ' no keyed atom in there to resolve against headerMap.
                ' Stated as its own arm rather than left to fall through
                ' to DesugarPredicateAtom: that function would also pass
                ' it along today, but only INCIDENTALLY - because it bails
                ' when the head word is not a known table name - and an
                ' incidental pass-through is not a guarantee.
                Set dest = item
                Exit Sub
            ElseIf UnificationOpFor(headWord) <> "" Then
                ' PROLOG.8: passed through untouched, stated as its own
                ' arm for exactly the reason the comparison arm above
                ' states - DesugarPredicateAtom would also pass it along
                ' today, but only INCIDENTALLY, and an incidental
                ' pass-through is not a guarantee.
                '
                ' Incidental is ALL it is, and the honest reason is worth
                ' recording, because the tempting justification is wrong.
                ' A term-matching operand IS data - `(= X (name Alice))`
                ' unifies X against the compound term `(name Alice)`, not
                ' against a table - but DesugarPredicateAtom could never
                ' have rewritten that operand anyway: it folds
                ' `lst.Item(1)` ONLY, the body item's own head word, which
                ' for these four is always =/\=/==/\==, and Exit Subs the
                ' moment that is not a table name. It never descends into
                ' arguments. Nor can one of the four ever BE a table
                ' name - PROLOG() runs IsReservedPredicateName over every
                ' table argument before it is loaded. So this arm changes
                ' no behaviour that any program can observe; it fixes the
                ' guarantee in place so a later change to
                ' DesugarPredicateAtom cannot quietly start reaching in.
                Set dest = item
                Exit Sub
            ElseIf TypeTestKindFor(headWord) <> "" Or TypeTestIsoSpellingFor(headWord) <> "" Or headWord = "between" Then
                ' PROLOG.9: both new families passed through untouched,
                ' one arm because the reason is one reason - and stated
                ' rather than left to fall through, exactly as PROLOG.7's
                ' and PROLOG.8's arms above are, because
                ' DesugarPredicateAtom would pass them along today only
                ' INCIDENTALLY (it folds lst.Item(1) alone and bails the
                ' moment that is not a table name) and an incidental
                ' pass-through is not a guarantee.
                '
                ' There is nothing for headerMap to resolve in either.
                ' A type test's single operand is a term being
                ' classified, and `between`'s three are two arithmetic
                ' expressions and a variable - no keyed atom in any of
                ' them. Nor can one of these seven names ever BE a table
                ' name: PROLOG() runs IsReservedPredicateName over every
                ' table argument before it is loaded, and all seven are
                ' reserved as of this item.
                Set dest = item
                Exit Sub
            ElseIf headWord = "not" Then
                If lst.Count = 2 Then
                    Dim innerNot As Variant
                    DesugarBodyItem innerNot, lst.Item(2), anonPrefix, itemIndex, headerMap
                    Dim outNot As New Collection
                    outNot.Add lst.Item(1)
                    outNot.Add innerNot
                    Set dest = outNot
                Else
                    Set dest = item   ' malformed shape - ValidateBodyItem's own job to refuse
                End If
                Exit Sub
            ElseIf headWord = "findall" Then
                If lst.Count = 4 Then
                    Dim innerFindall As Variant
                    DesugarBodyItem innerFindall, lst.Item(3), anonPrefix, itemIndex, headerMap
                    Dim outFindall As New Collection
                    outFindall.Add lst.Item(1)
                    outFindall.Add lst.Item(2)
                    outFindall.Add innerFindall
                    outFindall.Add lst.Item(4)
                    Set dest = outFindall
                Else
                    Set dest = item
                End If
                Exit Sub
            End If
        End If
    End If
    DesugarPredicateAtom dest, item, anonPrefix, itemIndex, headerMap
End Sub

' The actual keyed-atom resolution, factored out of DesugarBodyItem's
' own dispatch - see this module's own PROLOG.6 header for the full
' classification reasoning (headerMap checked first; allBare/allPairs
' computed in one pass; anything in between refused by name). A Sub with
' a ByRef dest out param for the identical reason DesugarBodyItem
' (above) is - term is always an object by the time this is ever called.
Private Sub DesugarPredicateAtom(ByRef dest As Variant, ByVal term As Variant, ByVal anonPrefix As String, ByVal itemIndex As Long, ByVal headerMap As Object)
    Dim lst As Collection
    Set lst = term
    Set dest = term
    If lst.Count < 1 Then Exit Sub
    If IsObject(lst.Item(1)) Then Exit Sub   ' malformed - TermPredName's own job to refuse
    Dim predFolded As String
    predFolded = VLA_Identity.Fold(CStr(lst.Item(1)))
    If Not VLA_Runtime.VlaDictHas(headerMap, predFolded) Then Exit Sub   ' not table-sourced - never desugared

    Dim allBare As Boolean, allPairs As Boolean
    allBare = True
    allPairs = True
    Dim i As Long
    For i = 2 To lst.Count
        If IsObject(lst.Item(i)) Then
            allBare = False
            Dim pairLst As Collection
            Set pairLst = lst.Item(i)
            If pairLst.Count <> 2 Then
                allPairs = False
            ElseIf IsObject(pairLst.Item(1)) Then
                allPairs = False
            End If
        Else
            allPairs = False
        End If
    Next i

    If allBare Then Exit Sub   ' plain positional reference to a table-sourced predicate - unchanged
    If Not allPairs Then
        VLA_Messages.RaiseMsg "prolog-atom-mixed-keying", "predicate", predFolded
        Exit Sub
    End If

    Dim headerPairs As Collection
    Set headerPairs = VLA_Runtime.VlaDictGet(headerMap, predFolded)
    Dim arity As Long
    arity = headerPairs.Count
    Dim slots() As Variant
    ReDim slots(1 To arity)
    Dim filled() As Boolean
    ReDim filled(1 To arity)

    For i = 2 To lst.Count
        Dim pLst As Collection
        Set pLst = lst.Item(i)
        Dim hText As String
        hText = CStr(pLst.Item(1))
        Dim hFolded As String
        hFolded = VLA_Identity.Fold(hText)
        Dim pos As Long
        pos = 0
        Dim j As Long
        For j = 1 To headerPairs.Count
            Dim hpPair As Collection
            Set hpPair = headerPairs.Item(j)
            If StrComp(CStr(hpPair.Item(1)), hFolded, vbBinaryCompare) = 0 Then
                pos = j
                Exit For
            End If
        Next j
        If pos = 0 Then
            VLA_Messages.RaiseMsg "prolog-unknown-column", "predicate", predFolded, "column", hText, "columns", JoinOriginalNames(headerPairs)
        End If
        If filled(pos) Then
            VLA_Messages.RaiseMsg "prolog-keyed-column-repeated", "predicate", predFolded, "column", hText
        End If
        filled(pos) = True
        ' IsObject-branched, never a bare assignment - this module's own
        ' already-documented trap (EnvWalkInto's own header): a keyed
        ' pair's own value position may itself be a compound term (e.g.
        ' (name (box X))), and a plain Variant-array-element assignment
        ' from an expression that MIGHT be an object invokes the
        ' object's own default member instead of copying the reference.
        If IsObject(pLst.Item(2)) Then
            Set slots(pos) = pLst.Item(2)
        Else
            slots(pos) = pLst.Item(2)
        End If
    Next i

    Dim outLst As New Collection
    outLst.Add lst.Item(1)
    For j = 1 To arity
        If filled(j) Then
            outLst.Add slots(j)
        Else
            outLst.Add AnonymousColumnVarName(anonPrefix, itemIndex, j)
        End If
    Next j
    Set dest = outLst
End Sub

' clauseDict: predName (folded) -> Collection of clause records
' (MakeClause, above) - facts and rules interleaved in written order,
' exactly as authored; see this module's own header for why nothing here
' special-cases "a predicate defined by both." PROLOG.6: clauseDict is
' now an INPUT/OUTPUT parameter, already created (and possibly
' pre-populated with table-sourced fact clauses) by the CALLER, never
' allocated fresh here - see this module's own PROLOG.6 header for why
' ("table rows first, then whatever the clauses text itself defines, in
' written order"). headerMap (PROLOG.6, new): predName (folded) ->
' VLA_Relation.RangeColumnNames' own (folded, original) pair Collection,
' threaded to DesugarBodyItem for every rule-body item and query
' conjunct - never a fact/rule HEAD, DATALOG.5's own explicit
' restriction, repeated here for the identical reason. queryConjuncts:
' the query's own goal list, in written order, never freshened (only
' STORED clauses get freshened per invocation - see FreshenTerm/
' SolveGoalList, below). freeVarNames: every distinct variable across
' every conjunct, first-occurrence order - this program's own output
' column order.
Private Sub ParseProgram(ByVal clausesText As String, ByVal clauseDict As Object, _
                          ByRef queryConjuncts As Collection, _
                          ByRef freeVarNames As Collection, _
                          ByVal headerMap As Object)
    Set queryConjuncts = New Collection
    Set freeVarNames = New Collection
    Dim predArity As Object
    Set predArity = VLA_Runtime.VlaDictNew()

    Dim forms As Collection
    Set forms = VLA.VlaReadForms(clausesText)
    Dim queryCount As Long
    Dim f As Variant
    For Each f In forms
        Dim head As String
        head = TopHead(f)
        Dim lst As Collection
        Set lst = f
        Select Case head
        Case "fact"
            If lst.Count <> 2 Then VLA_Messages.RaiseMsg "prolog-fact-bad-shape"
            Dim predName As String
            predName = TermPredName(lst.Item(2), "a fact", predArity)
            If IsReservedPredicateName(predName) Then VLA_Messages.RaiseMsg "prolog-reserved-predicate-name", "name", predName
            Dim varName As String
            varName = ""
            If TermHasVariable(lst.Item(2), varName) Then
                VLA_Messages.RaiseMsg "prolog-fact-has-variable", "predicate", predName, "var", varName
            End If
            GetOrCreateClauseList(clauseDict, predName).Add MakeClause(lst.Item(2), New Collection)
        Case "rule"
            If lst.Count < 3 Then VLA_Messages.RaiseMsg "prolog-rule-needs-body"
            Dim rulePredName As String
            rulePredName = TermPredName(lst.Item(2), "a rule head", predArity)
            If IsReservedPredicateName(rulePredName) Then VLA_Messages.RaiseMsg "prolog-reserved-predicate-name", "name", rulePredName
            ' Fresh `Set ... New Collection` EVERY time this Case runs,
            ' not `Dim ... As New` - the exact As-New-in-a-loop trap
            ' VLA_Datalog.bas's own ParseProgram already documents (a
            ' SECOND (rule ...) form would otherwise find bodyItems
            ' already non-Nothing from the first and silently append
            ' onto - and share - that rule's own Collection instead of
            ' starting fresh).
            Dim bodyItems As Collection
            Set bodyItems = New Collection
            Dim bi As Long
            For bi = 3 To lst.Count
                ' PROLOG.6: desugared BEFORE validation - a keyed atom's
                ' own arity (once resolved) must be what ValidateBodyItem/
                ' TermPredName actually see, never the raw key-pair count.
                Dim desugaredBody As Variant
                DesugarBodyItem desugaredBody, lst.Item(bi), "VlaAnonB", bi, headerMap
                ' NOT CollectVars here - a rule body's own variables are
                ' LOCAL to that clause (freshened per invocation,
                ' SolveGoalList's own job), never a free OUTPUT column
                ' of the overall query; only a QUERY's own conjuncts
                ' (below) ever populate freeVarNames, unchanged from
                ' PROLOG.3/PROLOG.4.
                ValidateBodyItem desugaredBody, "a rule body", predArity
                bodyItems.Add desugaredBody
            Next bi
            GetOrCreateClauseList(clauseDict, rulePredName).Add MakeClause(lst.Item(2), bodyItems)
        Case "query"
            If lst.Count < 2 Then VLA_Messages.RaiseMsg "prolog-query-bad-shape"
            queryCount = queryCount + 1
            If queryCount > 1 Then VLA_Messages.RaiseMsg "prolog-query-ambiguous", "count", queryCount
            Dim qi As Long
            For qi = 2 To lst.Count
                Dim desugaredQuery As Variant
                DesugarBodyItem desugaredQuery, lst.Item(qi), "VlaAnonQ", qi, headerMap
                ValidateBodyItem desugaredQuery, "a query", predArity
                CollectVars desugaredQuery, freeVarNames, True
                queryConjuncts.Add desugaredQuery
            Next qi
        Case Else
            VLA_Messages.RaiseMsg "prolog-unknown-top-form", "head", head
        End Select
    Next f
    If queryCount = 0 Then VLA_Messages.RaiseMsg "prolog-query-missing"
End Sub

' ---------------------------------------------------------------------
'  Solving
' ---------------------------------------------------------------------

' A goal term's own predicate name, folded - a pure lookup, unlike
' TermPredName (above), which ALSO records/checks arity as a parse-time
' side effect. Every goal SolveGoalList ever sees already passed through
' TermPredName once during ParseProgram (a query conjunct, or - freshened
' - a rule body item), and FreshenTerm (below) never changes a term's own
' arity, so re-checking it here would only repeat work already done, not
' add a real safety net.
Private Function GoalPredName(ByVal term As Variant) As String
    Dim lst As Collection
    Set lst = term
    GoalPredName = VLA_Identity.Fold(CStr(lst.Item(1)))
End Function

' PROLOG.5.4: the barrier embedded in a (possibly freshened) cut atom -
' "!" itself (only ever possible in a query conjunct, never freshened)
' is barrier 0; "!#37" (FreshenTerm's own rewrite, above) is barrier 37 -
' this module's own header has the full reasoning for why the barrier is
' carried on the atom itself rather than a separate structure.
Private Function CutBarrierOf(ByVal raw As String) As Long
    If Len(raw) > 1 Then CutBarrierOf = CLng(Mid$(raw, 3))   ' skip "!#"
End Function

' PROLOG.6: unifies two predicate-application terms' own ARGUMENTS
' (position 2..N) only, never position 1 (the predicate name itself) -
' this module's own PROLOG.6 header has the full incident (a real,
' pre-existing UnifyTwoWay gap: its own fully generic per-element
' recursion has no notion that position 1 of a predicate-application
' term is a functor, never a variable candidate, and would misdetect a
' capitalized predicate name - routine for a table-sourced one - as an
' unbound variable). Safe ONLY because the caller (SolveGoalList's own
' candidates loop) already guarantees goalTerm/headTerm are predicate-
' application terms of already-EQUAL arity (clauseDict's own folded-key
' lookup, plus RecordArity's own cross-definition consistency check) -
' not a general-purpose replacement for UnifyTwoWay, which must stay
' fully generic for findall's own functor-less Bag (ResolveTermDeep's
' own header has why).
Private Function UnifyArgsOnly(ByVal goalTerm As Variant, ByVal headTerm As Variant, envN As Collection, envT As Collection) As Boolean
    Dim gLst As Collection, hLst As Collection
    Set gLst = goalTerm
    Set hLst = headTerm
    Dim k As Long
    For k = 2 To gLst.Count
        If Not VLA_Unify.UnifyTwoWay(gLst.Item(k), hLst.Item(k), envN, envT) Then Exit Function
    Next k
    UnifyArgsOnly = True
End Function

' Renames every variable atom within term by appending suffix (e.g.
' "#37") to its own original name. The functor/predicate-name position -
' item 1 of any compound term, at every nesting depth - is copied through
' UNCHANGED, never treated as a variable candidate: the identical
' position-1-is-never-a-variable rule TermHasVariable/CollectVars (above)
' already established, for the same reason (a predicate's own name is
' never itself a data argument).
'
' No rename-map/dictionary is needed at all - found while building this,
' not assumed up front: appending the SAME suffix to a variable's own
' original name deterministically reproduces the IDENTICAL fresh name at
' every occurrence of that variable within one clause invocation (X and
' X, wherever they both recur in one clause, both become "X#37"), which
' is exactly what "rename apart" requires; a second, DIFFERENT variable
' in the same clause (Y) keeps its own distinct base name too ("Y#37"),
' so distinct variables never collide with each other either. suffix is
' shared by every variable freshened within ONE clause invocation
' (SolveGoalList passes the SAME suffix to both the head and every body
' item of one candidate clause) and never reused across a DIFFERENT
' invocation - SolveGoalList's own running step count is what makes each
' invocation's own suffix unique, so no separate counter is needed either.
Private Function FreshenTerm(ByVal term As Variant, ByVal suffix As String) As Variant
    If Not IsObject(term) Then
        Dim s As String
        s = CStr(term)
        ' PROLOG.5.4: a bare "!" (cut) takes the SAME freshening branch a
        ' variable atom already does - see this module's own PROLOG.5.4
        ' header for why this is the whole mechanism cut's own barrier
        ' needs: the per-invocation suffix already uniquely identifies
        ' "which clause activation," so "!#37" permanently carries that
        ' identity with it, no new counter or parallel structure required.
        If VLA_Unify.IsVarAtom(s) Or s = "!" Then
            FreshenTerm = s & suffix
        Else
            FreshenTerm = term
        End If
        Exit Function
    End If
    Dim lst As Collection, outLst As New Collection
    Set lst = term
    outLst.Add lst.Item(1)
    Dim i As Long
    For i = 2 To lst.Count
        outLst.Add FreshenTerm(lst.Item(i), suffix)
    Next i
    Set FreshenTerm = outLst
End Function

' Evaluates an already-freshened `is`-expression to a Double, fully
' dereferencing every variable through envN/envT as it recurses -
' RUNTIME is where numeric-ness is actually decided (ValidateArithExpr,
' above, only checked operator/arity shape at parse time, since a leaf
' may be a still-unbound variable then). Never silently guesses: an
' unbound variable, a non-numeric ground value, or a divide-by-zero all
' refuse by name (RaiseMsg) rather than producing a wrong number -
' exactly real Prolog's own `is/2`, which raises on all three too.
'
' PROLOG.7: formLabel carries the enclosing form's own spelling, exactly
' as ValidateArithExpr's (above) does and for the identical reason - this
' evaluator serves `(is ...)` and all six comparison goals, so it cannot
' know which one is running and must be told rather than guess. Threaded
' through both recursive calls so a refusal raised deep inside a nested
' operand still names the form the user actually wrote.
Private Function EvalArithTerm(ByVal term As Variant, envN As Collection, envT As Collection, ByVal formLabel As String) As Double
    Dim w As Variant
    VLA_Unify.EnvWalkInto w, term, envN, envT
    If Not IsObject(w) Then
        Dim raw As String
        raw = CStr(w)
        ' A quoted-string atom (leading literal " marker, LeafText's own
        ' convention below) is NEVER a variable, checked on the RAW text
        ' before LeafText strips the marker - a capitalized quoted
        ' string like "Hello would otherwise false-positive IsVarAtom
        ' once its own marker is gone, misreporting a real string value
        ' as an unbound variable.
        If Left$(raw, 1) <> Chr$(34) And VLA_Unify.IsVarAtom(raw) Then
            VLA_Messages.RaiseMsg "prolog-arith-unbound-variable", "var", raw, "form", formLabel
        End If
        Dim s As String
        s = LeafText(raw)
        If Not VLA_Relation.IsInvariantNumericString(s) Then
            VLA_Messages.RaiseMsg "prolog-arith-not-numeric", "value", s, "form", formLabel
        End If
        EvalArithTerm = VLA_Relation.InvariantVal(s)
        Exit Function
    End If
    ' A LITERAL compound term written directly in the (is ...) expression
    ' is always `(op left right)` here, ValidateArithExpr's own parse-time
    ' guarantee - but w may instead be a variable that DEREFERENCED to an
    ' arbitrary compound term from elsewhere in the program (e.g. Y bound
    ' to (color red) by an unrelated fact) - ValidateArithExpr only ever
    ' checked the EXPRESSION'S OWN written shape, never what a variable
    ' might resolve to, so that possibility was never validated at parse
    ' time at all. Re-checked here, defensively, before touching
    ' lst.Item(1)/(2)/(3) - an arbitrary bound compound term could have
    ' any arity or a non-symbol functor, and touching those blind risks a
    ' raw "Subscript out of range" or a CStr-on-object crash instead of a
    ' worded refusal.
    Dim lst As Collection
    Set lst = w
    If lst.Count <> 3 Then VLA_Messages.RaiseMsg "prolog-arith-not-numeric", "value", RenderBoundValue(w), "form", formLabel
    If IsObject(lst.Item(1)) Then VLA_Messages.RaiseMsg "prolog-arith-not-numeric", "value", RenderBoundValue(w), "form", formLabel
    Dim op As String
    op = CStr(lst.Item(1))
    Select Case op
    Case "+", "-", "*", "/"
        ' recognized
    Case Else
        VLA_Messages.RaiseMsg "prolog-arith-not-numeric", "value", RenderBoundValue(w), "form", formLabel
    End Select
    Dim l As Double, r As Double
    l = EvalArithTerm(lst.Item(2), envN, envT, formLabel)
    r = EvalArithTerm(lst.Item(3), envN, envT, formLabel)
    Dim ok As Boolean, reason As String
    Dim computed As Variant
    computed = VLA_Relation.ComputeArithmetic(op, l, r, True, ok, reason)
    If Not ok Then VLA_Messages.RaiseMsg "prolog-arith-divide-by-zero", "form", formLabel
    EvalArithTerm = CDbl(computed)
End Function

' Recursively proves goals(1) against every clause - fact OR rule,
' interleaved in written order, see clauseDict's own header above -
' stored for its own predicate, threading envN/envT across the whole
' goal list (so a variable bound while proving an earlier goal
' constrains a later one - the whole point of a CONJUNCTIVE query,
' PROLOG.3's own point, unchanged) while giving each CANDIDATE its own
' independent env copy (UnifyEnvClone) so a failed or partial attempt at
' one candidate never leaks into the next or into a sibling. A
' successful clause application does not recurse on goals(2..) directly
' the way PROLOG.3's own SolveConjuncts advanced by a fixed index -
' instead it builds a NEW goal list (the clause's own freshened body
' items, followed by every goal after the one just proved) and recurses
' on THAT list from its own start; a fact's empty body means the new
' list is just goals(2..) unchanged, so a fact and a one-clause rule that
' immediately succeeds are indistinguishable to the caller, by
' construction. Reaching an EMPTY goal list means every original query
' conjunct - and every rule body goal spliced in along the way - has
' succeeded under this one env: record one solution by reading each free
' variable's own final bound value back out (EnvWalkInto), identical to
' PROLOG.3's own base case.
' PROLOG.5.4: cutActive/cutTargetBarrier are the whole cut signal - see
' this module's own PROLOG.5.4 header (top of file) for the full design.
' Both False/0 at the start of every call by construction (only ever set
' True by a cut atom's own dispatch, below, immediately before it
' returns; only ever cleared by the exact `For Each candidates` loop
' iteration whose own barrier the signal targets) - so every dispatch
' branch below that is NOT the candidates loop simply threads them
' through its own single forward call unchanged, with no logic of its
' own to add.
Private Sub SolveGoalList(ByVal goals As Collection, clauseDict As Object, _
                           envN As Collection, envT As Collection, _
                           freeVarNames As Collection, solutions As Collection, _
                           ByRef stepsTaken As Long, _
                           ByRef cutActive As Boolean, ByRef cutTargetBarrier As Long)
    If goals.Count = 0 Then
        Dim sol As New Collection
        Dim vn As Variant
        For Each vn In freeVarNames
            Dim resolved As Variant
            ' PROLOG.5.3: ResolveTermDeep, not a bare EnvWalkInto - see
            ' this module's own PROLOG.5.3 header for the real,
            ' pre-existing gap found and fixed here (a compound value
            ' that itself still contains an unresolved nested variable
            ' used to render with that inner variable's own raw name
            ' still showing).
            ResolveTermDeep resolved, CStr(vn), envN, envT
            sol.Add resolved
        Next vn
        solutions.Add sol
        Exit Sub
    End If

    Dim rest As New Collection
    Dim gi As Long
    For gi = 2 To goals.Count
        rest.Add goals.Item(gi)
    Next gi

    ' PROLOG.5.4: cut (`!`) - the ONLY non-object goal ValidateBodyItem
    ' can ever have let through this far (every other body item/query
    ' conjunct is always a list; a bare non-"!" atom was already refused
    ' at parse time as prolog-atom-not-a-list), so this check must come
    ' BEFORE GoalPredName (which assumes a Collection) rather than being
    ' dispatched by predName like is/not/findall. Deterministic like
    ' those three (exactly one outcome), but structurally different: it
    ' explores `rest` FIRST, under the signal still False/0 (exactly as
    ' if no cut had happened - every solution `rest` can still produce
    ' gets recorded normally), and only sets cutActive/cutTargetBarrier
    ' AFTER that exploration fully returns, immediately before this call
    ' itself returns - this is what makes the signal mean "as you
    ' unwind past me, stop trying alternatives" rather than "stop trying
    ' alternatives from here forward," and is why every candidates-loop
    ' frame below sees a clean, correctly-ordered signal.
    If Not IsObject(goals.Item(1)) Then
        stepsTaken = stepsTaken + 1
        If stepsTaken > PROLOG_MAX_STEPS Then VLA_Messages.RaiseMsg "prolog-step-ceiling", "steps", PROLOG_MAX_STEPS
        Dim cutBarrier As Long
        cutBarrier = CutBarrierOf(CStr(goals.Item(1)))
        SolveGoalList rest, clauseDict, envN, envT, freeVarNames, solutions, stepsTaken, cutActive, cutTargetBarrier
        cutActive = True
        cutTargetBarrier = cutBarrier
        Exit Sub
    End If

    Dim predName As String
    predName = GoalPredName(goals.Item(1))

    ' `is` is dispatched here, never through clauseDict - clauseDict can
    ' never actually contain a "is" entry anyway (IsReservedPredicateName
    ' forbids ever DEFINING one), so this check is unambiguous by
    ' construction, not a race against a same-named user predicate.
    ' Deterministic (no candidate enumeration, no backtracking - exactly
    ' one outcome per attempt), but still counted against
    ' PROLOG_MAX_STEPS and still run against a FRESH env clone, the same
    ' two disciplines every ordinary candidate try below already follows.
    If predName = "is" Then
        stepsTaken = stepsTaken + 1
        If stepsTaken > PROLOG_MAX_STEPS Then VLA_Messages.RaiseMsg "prolog-step-ceiling", "steps", PROLOG_MAX_STEPS
        Dim isGoal As Collection
        Set isGoal = goals.Item(1)
        Dim computedVal As Double
        computedVal = EvalArithTerm(isGoal.Item(3), envN, envT, "(is ...)")
        Dim isN As Collection, isT As Collection
        VLA_Unify.UnifyEnvClone envN, envT, isN, isT
        If VLA_Unify.UnifyTwoWay(isGoal.Item(2), CStr(computedVal), isN, isT) Then
            SolveGoalList rest, clauseDict, isN, isT, freeVarNames, solutions, stepsTaken, cutActive, cutTargetBarrier
        End If
        Exit Sub
    End If

    ' PROLOG.5.2: `(not Goal)` - dispatched here, never through
    ' clauseDict, the identical unambiguous-by-construction reasoning
    ' `is` (above) already uses (IsReservedPredicateName forbids ever
    ' DEFINING a predicate named "not"). SolveNegation proves Goal in
    ' total isolation and answers only True/False; envN/envT themselves
    ' are passed UNCHANGED into the continuation on success - never isN/
    ' isT-style fresh bindings the way `is` produces, since a successful
    ' negation binds nothing of its own to thread forward at all (see
    ' this module's own PROLOG.5.2 header for what SolveNegation's own
    ' clone/discard is protecting against).
    If predName = "not" Then
        stepsTaken = stepsTaken + 1
        If stepsTaken > PROLOG_MAX_STEPS Then VLA_Messages.RaiseMsg "prolog-step-ceiling", "steps", PROLOG_MAX_STEPS
        Dim notGoal As Collection
        Set notGoal = goals.Item(1)
        If Not SolveNegation(notGoal.Item(2), clauseDict, envN, envT, stepsTaken) Then
            SolveGoalList rest, clauseDict, envN, envT, freeVarNames, solutions, stepsTaken, cutActive, cutTargetBarrier
        End If
        Exit Sub
    End If

    ' PROLOG.5.3: `(findall Template Goal Bag)` - dispatched here, the
    ' identical unambiguous-by-construction reasoning `is`/`not` (above)
    ' already use. All of the actual harvesting - collecting Template's
    ' own variables, running Goal in isolation, substituting each
    ' solution back into Template's structure - lives in HarvestFindallBag
    ' (below), a SEPARATE function, deliberately, not inlined here even
    ' though it runs exactly once per dispatch: every local this dispatch
    ' declares directly adds to SolveGoalList's OWN per-call stack frame
    ' size, and SolveGoalList recurses once per resolution step - live-
    ' caught the hard way (this item's own first version inlined the
    ' whole harvest here, and the SAME "genuinely non-terminating rule"
    ' test that PROLOG.4/PROLOG.5.2 already tuned PROLOG_MAX_STEPS
    ' against started overflowing VBA's native call stack BEFORE
    ' stepsTaken ever reached the ceiling, on a query with no findall in
    ' it at all - proof that a heavier SolveGoalList frame, not a
    ' findall-specific cost, was the actual damage). Bag is unified
    ' against the harvested list in a FRESH clone (bagN/bagT), the
    ' identical isN/isT shape `is` (above) already uses - findall always
    ' succeeds or fails exactly once, deterministically, never
    ' backtracks over multiple Bag values, real Prolog's own findall/3
    ' semantics.
    If predName = "findall" Then
        stepsTaken = stepsTaken + 1
        If stepsTaken > PROLOG_MAX_STEPS Then VLA_Messages.RaiseMsg "prolog-step-ceiling", "steps", PROLOG_MAX_STEPS
        Dim findallGoal As Collection
        Set findallGoal = goals.Item(1)
        Dim bagN As Collection, bagT As Collection
        VLA_Unify.UnifyEnvClone envN, envT, bagN, bagT
        If VLA_Unify.UnifyTwoWay(findallGoal.Item(4), HarvestFindallBag(findallGoal, clauseDict, envN, envT, stepsTaken), bagN, bagT) Then
            SolveGoalList rest, clauseDict, bagN, bagT, freeVarNames, solutions, stepsTaken, cutActive, cutTargetBarrier
        End If
        Exit Sub
    End If

    ' PROLOG.7: the six comparison goals - dispatched here, the identical
    ' unambiguous-by-construction reasoning is/not/findall (above) already
    ' use, since IsReservedPredicateName forbids ever DEFINING a predicate
    ' with one of these names. Deterministic like those three (exactly one
    ' outcome, no candidate enumeration, no backtracking), still counted
    ' against PROLOG_MAX_STEPS.
    '
    ' envN/envT are threaded into the continuation UNCHANGED - never a
    ' fresh clone the way `is` produces one. A comparison binds nothing:
    ' it is a test, so there are no new bindings to carry forward, exactly
    ' the shape `not` (above) already has. It must also sit ABOVE the
    ' clauseDict lookup below: an unknown predicate there is a silent dead
    ' end rather than an error, so a comparison reaching it would quietly
    ' yield no rows instead of comparing anything.
    '
    ' No locals declared in this arm at all - the whole evaluation lives
    ' in SolveComparison (below), factored out for the same live-caught
    ' stack-frame reason findall's own dispatch documents above.
    If ComparisonOpFor(predName) <> "" Then
        stepsTaken = stepsTaken + 1
        If stepsTaken > PROLOG_MAX_STEPS Then VLA_Messages.RaiseMsg "prolog-step-ceiling", "steps", PROLOG_MAX_STEPS
        If SolveComparison(goals.Item(1), predName, envN, envT) Then
            SolveGoalList rest, clauseDict, envN, envT, freeVarNames, solutions, stepsTaken, cutActive, cutTargetBarrier
        End If
        Exit Sub
    End If

    ' PROLOG.8: the four term-matching goals (=, \=, ==, \==) - dispatched
    ' here, the identical unambiguous-by-construction reasoning every arm
    ' above already uses, since IsReservedPredicateName forbids ever
    ' DEFINING a predicate with one of these names. Deterministic like
    ' those (exactly one outcome, no candidate enumeration, no
    ' backtracking - real Prolog's own =/2, \=/2, ==/2 and \==/2 are all
    ' semi-deterministic), still counted against PROLOG_MAX_STEPS.
    '
    ' It must sit ABOVE the clauseDict lookup below for the reason
    ' PROLOG.7's arm states: an unknown predicate there is a silent dead
    ' end rather than an error, so one of these reaching it would quietly
    ' yield no rows instead of matching anything. That is not a
    ' hypothetical - it is precisely the behaviour PROLOG.7 PINNED, with
    ' `(query (= 1 1))` asserted FALSE, to catch =:= leaking into `=`
    ' before this item existed. That pin is re-pointed at real
    ' unification here.
    '
    ' Unlike every arm above, the continuation's environment is NOT fixed
    ' by which arm ran: `=` threads a fresh clone forward the way `is`
    ' does, the other three thread the caller's own the way `not` does,
    ' and SolveUnification (below) decides which and hands it back through
    ' uniN/uniT. Those two are the only locals this arm declares - the
    ' evaluation itself lives in that function, for the live-caught
    ' stack-frame reason findall's own dispatch documents above.
    If UnificationOpFor(predName) <> "" Then
        stepsTaken = stepsTaken + 1
        If stepsTaken > PROLOG_MAX_STEPS Then VLA_Messages.RaiseMsg "prolog-step-ceiling", "steps", PROLOG_MAX_STEPS
        Dim uniN As Collection, uniT As Collection
        If SolveUnification(goals.Item(1), UnificationOpFor(predName), envN, envT, uniN, uniT) Then
            SolveGoalList rest, clauseDict, uniN, uniT, freeVarNames, solutions, stepsTaken, cutActive, cutTargetBarrier
        End If
        Exit Sub
    End If

    ' PROLOG.9: the six type-test goals - dispatched here, the identical
    ' unambiguous-by-construction reasoning every arm above already uses,
    ' since IsReservedPredicateName forbids ever DEFINING a predicate with
    ' one of these names. Deterministic like all of them (exactly one
    ' outcome, no candidate enumeration, no backtracking), still counted
    ' against PROLOG_MAX_STEPS.
    '
    ' envN/envT are threaded into the continuation UNCHANGED - never a
    ' fresh clone the way `is` and `=` produce one. A type test binds
    ' nothing whatsoever: it looks at a term and answers, so there is
    ' nothing to carry forward, exactly the shape `not` and the six
    ' comparisons have. That is also why CollectVars (above) skips them.
    '
    ' It must sit ABOVE the clauseDict lookup below for the reason
    ' PROLOG.7's arm states: an unknown predicate there is a SILENT dead
    ' end rather than an error, so a type test reaching it would quietly
    ' answer "no rows" instead of classifying anything - and since half of
    ' these six are naturally written as goals expected to FAIL, that
    ' failure would be indistinguishable from a correct answer. This arm
    ' existing is what makes the difference observable.
    '
    ' No locals declared in this arm at all - the whole evaluation lives
    ' in SolveTypeTest (below), factored out for the same live-caught
    ' stack-frame reason findall's own dispatch documents above.
    If TypeTestKindFor(predName) <> "" Then
        stepsTaken = stepsTaken + 1
        If stepsTaken > PROLOG_MAX_STEPS Then VLA_Messages.RaiseMsg "prolog-step-ceiling", "steps", PROLOG_MAX_STEPS
        If SolveTypeTest(goals.Item(1), TypeTestKindFor(predName), envN, envT) Then
            SolveGoalList rest, clauseDict, envN, envT, freeVarNames, solutions, stepsTaken, cutActive, cutTargetBarrier
        End If
        Exit Sub
    End If

    ' PROLOG.9: the six BARE ISO spellings - `(atom X)` where this engine
    ' writes `(atom? X)`. Dispatched here, above the clauseDict lookup,
    ' for a reason sharper than any arm above it: BELOW that lookup an
    ' unknown predicate is a SILENT dead end, and these six names are
    ' precisely the ones a Prolog author will type first. Left
    ' undispatched they would answer "no rows" and explain nothing, which
    ' is the confidently-wrong-answer class this project holds to be
    ' worse than a crash.
    '
    ' This arm never solves anything - it always raises. That is the
    ' whole point, and it is why the check that holds reserved names to
    ' being dispatched is satisfied honestly rather than by exemption:
    ' the name IS reached, and what it does when reached is teach the
    ' spelling. No step is charged, because no resolution work happens.
    '
    ' Deliberately NOT also refused in ValidateBodyItem at parse time,
    ' which would make this arm dead code. Solve-time is where every
    ' other reserved name is dispatched, and the user sees the refusal in
    ' the cell either way.
    If TypeTestIsoSpellingFor(predName) <> "" Then
        VLA_Messages.RaiseMsg "prolog-type-test-iso-spelling", _
            "form", "(" & predName & " ...)", "fixed", "(" & TypeTestIsoSpellingFor(predName) & " ...)"
    End If

    ' PROLOG.9: `(between Low High X)` - dispatched here on the same
    ' unambiguous-by-construction reasoning, and above the clauseDict
    ' lookup for the same silent-dead-end reason, but STRUCTURALLY UNLIKE
    ' every arm above it. All of those are deterministic: they decide
    ' once and call SolveGoalList at most once. `between` in its
    ' generating mode is a CHOICE POINT - it must bind X to each value in
    ' turn and let the continuation run for every one - so it recurses
    ' once PER VALUE, which is the candidates loop's shape, not `is`'s.
    '
    ' That is why it hands SolveBetween (below) the things no other
    ' dispatch arm passes on: `rest`, clauseDict, freeVarNames, solutions
    ' and the cut signal. The enumeration cannot be done here and the
    ' continuation resumed afterwards, the way `is` unifies and then
    ' calls on; the two are interleaved.
    '
    ' The cut signal is threaded ByRef into that loop and honoured there,
    ' not here - this module's own PROLOG.9 header explains why a
    ' generator's loop must stop on cutActive and must never absorb it.
    ' The step this arm charges is for the GOAL; SolveBetween charges one
    ' more per value generated, so enumeration is counted as the real
    ' resolution work it is rather than riding free on a single goal's
    ' step.
    If predName = "between" Then
        stepsTaken = stepsTaken + 1
        If stepsTaken > PROLOG_MAX_STEPS Then VLA_Messages.RaiseMsg "prolog-step-ceiling", "steps", PROLOG_MAX_STEPS
        SolveBetween goals.Item(1), rest, clauseDict, envN, envT, freeVarNames, solutions, stepsTaken, cutActive, cutTargetBarrier
        Exit Sub
    End If

    If Not VLA_Runtime.VlaDictHas(clauseDict, predName) Then Exit Sub   ' no candidates - dead end, not an error

    Dim candidates As Collection
    Set candidates = VLA_Runtime.VlaDictGet(clauseDict, predName)
    Dim clauseRec As Variant
    For Each clauseRec In candidates
        stepsTaken = stepsTaken + 1
        If stepsTaken > PROLOG_MAX_STEPS Then VLA_Messages.RaiseMsg "prolog-step-ceiling", "steps", PROLOG_MAX_STEPS

        ' PROLOG.5.4: myStep identifies THIS candidate try - the same
        ' value `suffix` is built from - and is this loop's own barrier
        ' identity for cut's purposes (this module's own PROLOG.5.4
        ' header has the full reasoning).
        Dim myStep As Long
        myStep = stepsTaken
        Dim suffix As String
        suffix = "#" & CStr(myStep)

        Dim tryN As Collection, tryT As Collection
        VLA_Unify.UnifyEnvClone envN, envT, tryN, tryT
        ' PROLOG.6: UnifyArgsOnly, not a bare UnifyTwoWay over the WHOLE
        ' term - this module's own PROLOG.6 header has the full incident
        ' (a real, pre-existing UnifyTwoWay gap, found while wiring
        ' commonly-capitalized table-sourced predicate names).
        If UnifyArgsOnly(goals.Item(1), FreshenTerm(clauseRec.Item(1), suffix), tryN, tryT) Then
            Dim bodyItems As Collection
            Set bodyItems = clauseRec.Item(2)
            ' Fresh `Set ... New Collection` on EVERY candidate try, not
            ' `Dim ... As New` - live-caught, runtime error 13 masquerading
            ' as a step-ceiling refusal: `As New` only auto-instantiates
            ' once (the first time it's Nothing), so every candidate AFTER
            ' the first one in this SAME loop reused - and kept appending
            ' onto - the FIRST candidate's own already-consumed newGoals,
            ' corrupting every backtrack with duplicated/stale goals from
            ' every earlier sibling attempt. The exact As-New-in-a-loop
            ' trap this module's own ParseProgram (above) already guards
            ' against for `bodyItems` - missed here on the first pass.
            Dim newGoals As Collection
            Set newGoals = New Collection
            Dim bi As Variant
            For Each bi In bodyItems
                newGoals.Add FreshenTerm(bi, suffix)
            Next bi
            For gi = 1 To rest.Count
                newGoals.Add rest.Item(gi)
            Next gi
            SolveGoalList newGoals, clauseDict, tryN, tryT, freeVarNames, solutions, stepsTaken, cutActive, cutTargetBarrier
            ' PROLOG.5.4: the one real piece of surgery on this loop -
            ' whenever a cut signal is active after this candidate's own
            ' subtree returns, stop trying further candidates uncondit-
            ' ionally (every loop between a cut's own origin and its
            ' firing site must do this, proved by hand-tracing in this
            ' module's own PROLOG.5.4 header); absorb it (clear it) only
            ' when THIS loop is the one that actually selected the clause
            ' `!` fired inside (myStep = cutTargetBarrier exactly) - any
            ' other loop leaves it set, letting it keep propagating to
            ' its own true origin further up the call stack.
            If cutActive Then
                If myStep = cutTargetBarrier Then cutActive = False
                Exit For
            End If
        End If
    Next clauseRec
End Sub

' PROLOG.5.2/5.3: the bounded sub-call microscope both negation-as-
' failure and findall need - proves goal in TOTAL isolation (a CLONED
' envN/envT, VLA_Unify.UnifyEnvClone) and returns EVERY solution found,
' each one a Collection of resolved values (one per name in
' subFreeVarNames, PROLOG.4's own base-case shape, unchanged) - then
' discards that clone regardless of outcome, so a binding made while
' proving goal never leaks back into the caller's own env, whether goal
' succeeds OR fails, exactly real Prolog's own \+/findall semantics.
' Shares stepsTaken (ByRef) with the caller rather than starting a fresh
' budget - PROLOG_MAX_STEPS is a total-resolution-work ceiling for the
' WHOLE query, not a per-sub-call allowance, so a goal that doesn't
' terminate inside EITHER `not` or `findall` still hits the same worded
' prolog-step-ceiling refusal ordinary recursion already would, never a
' silent separate budget that could let it dodge the ceiling. No On
' Error of its own, deliberately: neither negation-as-failure nor
' findall catch a raised error, only ordinary proof failure (zero
' solutions) - an unbound variable inside an `is`, a divide-by-zero, or
' an occurs-check violation reached while proving goal must propagate
' all the way up to PROLOG()'s own top-level handler unchanged, exactly
' as if goal had been proved un-negated/un-harvested. Extracted from
' PROLOG.5.2's own original SolveNegation, unchanged in shape - only
' subFreeVarNames is new, callable with an empty Collection (`not`'s own
' case, existence only) or a real one (`findall`'s own Template
' variables, PROLOG.5.3) - `not`'s own dispatch call site did not
' change at all, per this item's own roadmap wording ("reuses PROLOG.
' 5.2's own bounded sub-call microscope directly").
Private Function SolveIsolated(ByVal goal As Variant, clauseDict As Object, _
                                envN As Collection, envT As Collection, _
                                subFreeVarNames As Collection, ByRef stepsTaken As Long) As Collection
    Dim subN As Collection, subT As Collection
    VLA_Unify.UnifyEnvClone envN, envT, subN, subT
    Dim subGoals As New Collection
    subGoals.Add goal
    Dim subSolutions As New Collection
    ' PROLOG.5.4: a fresh, LOCAL cut-signal pair, never accepted from or
    ' forwarded to the caller - this is the entire mechanism that makes a
    ' cut inside `goal` opaque to the outer search (real Prolog's own
    ' rule for \+ and findall's own Goal argument), confirmed to need no
    ' other change at all; this module's own PROLOG.5.4 header has the
    ' full reasoning.
    Dim subCutActive As Boolean, subCutTargetBarrier As Long
    SolveGoalList subGoals, clauseDict, subN, subT, subFreeVarNames, subSolutions, stepsTaken, subCutActive, subCutTargetBarrier
    Set SolveIsolated = subSolutions
End Function

' PROLOG.5.2: inverts existence - True iff proving goal in isolation
' (SolveIsolated, above) found at least one solution. No freeVarNames of
' its own: existence is all negation-as-failure answers, so an empty
' Collection is passed as the sub-call's own subFreeVarNames;
' SolveGoalList's own base case still runs correctly with zero free
' variables (it simply adds one empty "solution" record per proof, which
' is all Count needs).
Private Function SolveNegation(ByVal goal As Variant, clauseDict As Object, _
                                envN As Collection, envT As Collection, _
                                ByRef stepsTaken As Long) As Boolean
    Dim noFreeVars As New Collection
    SolveNegation = (SolveIsolated(goal, clauseDict, envN, envT, noFreeVars, stepsTaken).Count > 0)
End Function

' PROLOG.7: answers one comparison goal - True iff it succeeds. Factored
' out of SolveGoalList's own dispatch for the reason that dispatch states:
' its four locals live in this frame, entered once per comparison, instead
' of bloating the frame SolveGoalList pays on EVERY resolution step.
'
' Takes no clauseDict and no stepsTaken: unlike not/findall there is no
' sub-proof to run here, so nothing can recurse and nothing further needs
' counting - the dispatch already counted this goal's own step. envN/envT
' are read-only in practice; EvalArithTerm only ever WALKS them (it
' dereferences variables and never binds), so no clone is needed to keep
' this test from leaking bindings - there are none to leak.
'
' Both sides are evaluated BEFORE either is compared, so a refusal from
' the right-hand side is raised even when the left already decided the
' answer. That is deliberate: `(> 5 UnboundVar)` is a broken program
' whichever way the numbers fall, and short-circuiting would make whether
' the user hears about it depend on the left operand's value.
Private Function SolveComparison(ByVal goalTerm As Variant, ByVal predName As String, _
                                  envN As Collection, envT As Collection) As Boolean
    Dim lst As Collection
    Set lst = goalTerm
    Dim formLabel As String
    formLabel = "(" & predName & " ...)"
    Dim lv As Double, rv As Double
    lv = EvalArithTerm(lst.Item(2), envN, envT, formLabel)
    rv = EvalArithTerm(lst.Item(3), envN, envT, formLabel)
    SolveComparison = VLA_Relation.CompareValues(ComparisonOpFor(predName), lv, rv, True)
End Function

' PROLOG.8: answers one term-matching goal - True iff it succeeds - and
' hands back the environment the continuation must run under. Factored out
' of SolveGoalList's own dispatch for the reason that dispatch states: its
' locals live in this frame, entered once per goal, instead of bloating
' the frame SolveGoalList pays on EVERY resolution step.
'
' Takes no clauseDict and no stepsTaken. This is where the entry's own
' proposed shortcut was rejected: `(\= A B)` IS definitionally `(not (= A
' B))`, but DESUGARING it to that - rewriting the term at parse time, the
' way the roadmap suggested - would hand ValidateBodyItem and every
' refusal beneath it a `(not ...)` / `(= ...)` the user never wrote, which
' is exactly the misattribution PROLOG.7 spent an item and a check
' removing. It would also route a test that needs one UnifyTwoWay through
' a whole isolated sub-proof (SolveIsolated: an env clone, a goal list, a
' solutions collection and a recursive SolveGoalList), charging a second
' step against PROLOG_MAX_STEPS for it. The equivalence is real; it is
' honoured by running the same unification and inverting the answer, right
' here, where the form the user wrote is still known.
'
' THE ENVIRONMENT FORK, which is the whole reason this returns an env
' rather than a bare Boolean. `=` BINDS: on success its bindings must
' reach the rest of the query, so outN/outT come back as the fresh clone
' it bound into. The other three bind nothing outward and hand back the
' caller's own environment untouched, the shape `not` already has. The
' clone is not optional for either unifying operator, and not merely
' hygiene: UnifyTwoWay's own contract is that it "returns False with
' envN/envT left however far the walk got", so a FAILED `(= X Y)` run
' directly against envN/envT would leave that partial walk's bindings
' behind for the next sibling goal to inherit.
'
' `\=` clones for that reason and then discards the clone whichever way
' the attempt went - it reports only whether unification was possible,
' never what it would have bound.
'
' OCCURS CHECK - decided here, not inherited by accident. `(\= X (f X))`
' RAISES prolog-occurs-check, identically to `(= X (f X))`, because
' UnifyTwoWay raises it and nothing here catches it. That is deliberate:
' the refusal is about the TERM being unrepresentable - an infinite term
' "can never render, spill to a worksheet, or be audited", PROLOG.1's own
' standing decision - and that is equally true of `(f X)` whichever
' operator encloses it. Answering True instead would require wrapping a
' RaiseMsg in On Error Resume Next, and would make `=` and `\=` disagree
' about a program neither can actually represent. `==`/`\==` never reach
' a bind at all, so they never occurs-check: `(== X (f X))` is an
' ordinary False and `(\== X (f X))` an ordinary True. That asymmetry is
' correct, not an oversight - it falls out of the binding fork above.
Private Function SolveUnification(ByVal goalTerm As Variant, ByVal kind As String, _
                                   envN As Collection, envT As Collection, _
                                   ByRef outN As Collection, ByRef outT As Collection) As Boolean
    Dim lst As Collection
    Set lst = goalTerm

    ' The default for three of the four, and for every failure path: the
    ' continuation runs under the caller's own environment, unchanged.
    Set outN = envN
    Set outT = envT

    ' PROLOG.10: each of the four consults RefuseIfQuotedVersusBare
    ' (below) at exactly one moment - when the underlying match FAILED,
    ' whichever answer that failure produces. That single rule covers all
    ' four uniformly: `=` and `==` answer False on a failed match, `\=`
    ' and `\==` answer True on one, so "the answer was decided by a
    ' mismatch" is the same condition in every case and is checked in the
    ' same place. Never on a SUCCESSFUL match, where there is nothing
    ' confusable to report.
    Select Case kind
    Case "identical", "notidentical"
        Dim same As Boolean
        same = VLA_Unify.TermsIdentical(lst.Item(2), lst.Item(3), envN, envT)
        If Not same Then RefuseIfQuotedVersusBare lst.Item(2), lst.Item(3), envN, envT
        If kind = "identical" Then
            SolveUnification = same
        Else
            SolveUnification = Not same
        End If
        Exit Function
    End Select

    Dim tryN As Collection, tryT As Collection
    VLA_Unify.UnifyEnvClone envN, envT, tryN, tryT
    Dim unified As Boolean
    unified = VLA_Unify.UnifyTwoWay(lst.Item(2), lst.Item(3), tryN, tryT)

    ' envN/envT, never tryN/tryT: the clone may carry half-finished
    ' bindings left behind by the walk that just failed (UnifyTwoWay's own
    ' documented "returns False with envN/envT left however far the walk
    ' got"), and this report must describe the terms the USER compared,
    ' not the wreckage of the attempt.
    If Not unified Then RefuseIfQuotedVersusBare lst.Item(2), lst.Item(3), envN, envT

    If kind = "notunify" Then
        SolveUnification = Not unified
        Exit Function          ' clone discarded: `\=` reports possibility, never bindings
    End If

    SolveUnification = unified
    If unified Then
        Set outN = tryN
        Set outT = tryT
    End If
End Function

' PROLOG.10: the adjudication, built. `(\= D eng)` succeeding on EVERY row
' of a table whose Dept column reads eng is a confidently wrong answer -
' the marker makes `"eng` and `eng` different terms, correctly, but the
' user who wrote it meant the cell. Option D on top of option A: the
' semantics do not change (they ARE different terms, and `(== "eng" eng)`
' is still not true), only the SILENCE does.
'
' SCOPED TO THE FOUR EXPLICIT COMPARISON GOALS, and not to clause
' matching, which the entry's own draft would also have covered. Four
' independent reasons, all pointing the same way:
'
'   CORRECTNESS CLASS. `(\= D eng)` returning every row is WRONG. A plain
'       `(query (emp N eng))` returning zero rows is RIGHT - there is no
'       term `eng` in that program - merely unhelpful. This project's
'       doctrine is about confidently wrong answers, and it reaches the
'       first case and not the second.
'   MONOTONICITY. Raising during clause matching would let a
'       NON-MATCHING clause decide a query's fate: adding
'       `(fact (color "red"))` beside `(fact (color red))` would turn a
'       working `(query (color red))` into an error. Adding a fact must
'       never remove a solution. This engine is already non-monotonic
'       exactly where the user WRITES a non-monotonic operator - not,
'       \=, \==, ! - and has never been so implicitly.
'   PRECEDENT. PROLOG.9 already ruled on this distinction three times:
'       `(between 1 10 "5")` refuses (an explicit goal handed a value it
'       cannot use), `(number? "42")` answers False (a question whose job
'       is to answer), and clause matching stays silent. This is that
'       rule applied unchanged.
'   REVERSIBILITY. This is a strict subset of the wider reading, so it
'       can be widened later - silence into a message is the compatible
'       direction. The wider one could only be narrowed by retracting an
'       error class.
'
' The plain-query half stays deliberately open, recorded in the roadmap
' rather than half-built.
Private Sub RefuseIfQuotedVersusBare(ByVal a As Variant, ByVal b As Variant, _
                                      envN As Collection, envT As Collection)
    Dim sameText As String
    sameText = ""
    ' freeVarUnifies:=False - an explicit comparison is being asked about
    ' the terms AS THEY STAND. `(== X 1)` with X free fails for a reason
    ' that has nothing to do with markers, and must not be reported as
    ' though it did.
    If QuotedVersusBareClass(a, b, envN, envT, sameText, False) <> 1 Then Exit Sub
    RaiseQuotedVersusBare "prolog-quoted-versus-bare", sameText
End Sub

' PROLOG.12: PROLOG.10 closed the half where a comparison ANSWERED
' wrongly and left open the half where a plain query merely found
' nothing. This closes it, and does so without any of the four costs that
' ruled out doing it inside the solver.
'
' THE ARGUMENT THAT CHANGED. PROLOG.10 declined this on the ground that
' zero rows is a CORRECT answer, and that remains true - but "correct"
' was doing less work in that sentence than it looked. A query that found
' nothing AND contains a term differing from a stored one by only the
' quoting is not a user who wanted zero rows; it is a user who wrote the
' one mistake this engine's own conventions make easiest to write. The
' decisive objection to the unscoped version was MONOTONICITY, and it
' does not apply here: this can never turn a success into a failure,
' because it only ever runs when there were no solutions at all. Adding a
' fact that matches makes it go away.
'
' THE RESIDUAL COST, stated rather than buried: it can turn an empty
' result into an error. That is a real change and the price of closing
' the hole - but it is the benign direction (empty to explained), never
' the direction that breaks a working program.
'
' KNOWN LIMIT, by construction rather than oversight: this compares each
' QUERY conjunct against the stored clauses of its own predicate, so it
' sees a near-miss the user could have spotted by reading their own
' query. It does NOT see one that only appears after an earlier conjunct
' binds a variable, nor one inside a rule body - both would need the
' near-miss recorded during solving, which is exactly the threading this
' design exists to avoid. Pinned by tests that assert those cases stay
' silent, so the limit is recorded as behaviour rather than as a comment
' that could drift.
Private Sub DiagnoseQuotedVersusBare(ByVal queryConjuncts As Collection, clauseDict As Object)
    ' An EMPTY environment on purpose: solving is over and its bindings
    ' are gone, so every variable here reads as free - which is precisely
    ' the reading freeVarUnifies wants, since a variable position is one
    ' that would have unified rather than one that went wrong.
    Dim emptyN As New Collection, emptyT As New Collection
    Dim conj As Variant
    For Each conj In queryConjuncts
        ' IsObject first: a cut atom is the one non-object a query
        ' conjunct can be, and GoalPredName assumes a Collection.
        If IsObject(conj) Then
            Dim predName As String
            predName = GoalPredName(conj)
            ' A reserved goal - is/not/findall/a comparison/a type test -
            ' is never in clauseDict, so this loop skips all of them
            ' without needing to know their names.
            If VLA_Runtime.VlaDictHas(clauseDict, predName) Then
                Dim candidates As Collection
                Set candidates = VLA_Runtime.VlaDictGet(clauseDict, predName)
                Dim clauseRec As Variant
                For Each clauseRec In candidates
                    Dim sameText As String
                    sameText = ""
                    If QuotedVersusBareClass(conj, clauseRec.Item(1), emptyN, emptyT, sameText, True) = 1 Then
                        RaiseQuotedVersusBare "prolog-quoted-versus-bare-no-rows", sameText
                    End If
                Next clauseRec
            End If
        End If
    Next conj
End Sub

' PROLOG.10/PROLOG.12: the shared raise. `number` or `name` for the bare
' side, decided ONCE for both refusals so they can never disagree about
' what a number is. LeafIsNumberTerm is asked rather than
' IsInvariantNumericString directly, so this and PROLOG.9's own
' atom?/number? cannot drift apart either - and sameText is already
' marker-free, so it takes the plain numeric branch by construction.
Private Sub RaiseQuotedVersusBare(ByVal msgId As String, ByVal sameText As String)
    Dim kindWord As String
    If LeafIsNumberTerm(sameText) Then
        kindWord = "number"
    Else
        kindWord = "name"
    End If
    VLA_Messages.RaiseMsg msgId, "text", sameText, "kind", kindWord
End Sub

' PROLOG.10: how two terms differ, as one of three answers -
'   0  identical
'   1  differ ONLY by the quoted-string marker, at one or more leaves
'   2  differ in some other way as well
'
' The three-way answer is the whole point, and a two-way "is there a
' marker-only pair anywhere" test would have been wrong. `(= (f "eng")
' (g eng))` contains a marker-only pair at position 2 AND a genuine
' functor difference at position 1; the real reason those two do not
' unify is `f` versus `g`, so blaming the marker would be a confidently
' wrong DIAGNOSIS - the same defect class this refusal exists to remove,
' reintroduced one level up. Class 2 wins over class 1 for exactly that
' reason, and the walk stops as soon as it is reached.
'
' outText keeps the FIRST marker-only text found, so a term differing at
' several leaves names one of them rather than the last one looked at.
'
' CONTRACT, surfaced by the pre-import transliteration rather than by
' reading: outText is meaningful ONLY when this function returns 1. It is
' written SPECULATIVELY, the moment a marker-only leaf pair is seen, and
' a later position may then push the answer to 2 - `(f "eng a)` versus
' `(f eng b)` correctly returns 2 while leaving "eng" behind in outText.
' Harmless because the one caller tests the class before reading it, and
' recorded here so a second caller cannot quietly assume otherwise.
' PROLOG.12: freeVarUnifies picks which QUESTION is being asked, and the
' two callers genuinely ask different ones.
'
'   False - "how do these two terms, as they stand, differ?" An explicit
'           `=`/`\=`/`==`/`\==` compares what is written, so a free
'           variable facing a ground atom is a real difference.
'   True  - "would these two have unified, but for the quoting?" Clause
'           matching binds, so a variable position is one that would have
'           SUCCEEDED and must not count as a difference at all.
'
' The True reading is an approximation in the safe direction: it ignores
' that binding one variable twice can still fail - `(p X X)` against
' `(p a b)` reads as "would unify" here and does not. That can only ever
' SUPPRESS a report, never manufacture one, which is the right way round
' for a diagnostic that fires on an already-empty result.
Private Function QuotedVersusBareClass(ByVal a As Variant, ByVal b As Variant, _
                                        envN As Collection, envT As Collection, _
                                        ByRef outText As String, _
                                        ByVal freeVarUnifies As Boolean) As Long
    Dim aw As Variant, bw As Variant
    VLA_Unify.EnvWalkInto aw, a, envN, envT
    VLA_Unify.EnvWalkInto bw, b, envN, envT

    If IsObject(aw) <> IsObject(bw) Then
        ' A free variable unifies with a compound term too, so the
        ' mismatched-shape case needs the same exemption as the leaf case
        ' below. Nested Ifs, never a combined And - VBA does not
        ' short-circuit, and CStr on the Collection side would raise 450.
        If freeVarUnifies Then
            If Not IsObject(aw) Then
                If VLA_Unify.IsVarAtom(CStr(aw)) Then Exit Function      ' 0
            End If
            If Not IsObject(bw) Then
                If VLA_Unify.IsVarAtom(CStr(bw)) Then Exit Function      ' 0
            End If
        End If
        QuotedVersusBareClass = 2
        Exit Function
    End If

    If Not IsObject(aw) Then
        Dim x As String, y As String
        x = CStr(aw)
        y = CStr(bw)
        If x = y Then Exit Function          ' 0 - identical leaves
        If freeVarUnifies Then
            If VLA_Unify.IsVarAtom(x) Then Exit Function                 ' 0
            If VLA_Unify.IsVarAtom(y) Then Exit Function                 ' 0
        End If
        Dim t As String
        If LeavesDifferOnlyByMarker(x, y, t) Then
            If Len(outText) = 0 Then outText = t
            QuotedVersusBareClass = 1
        Else
            QuotedVersusBareClass = 2
        End If
        Exit Function
    End If

    Dim la As Collection, lb As Collection
    Set la = aw
    Set lb = bw
    If la.Count <> lb.Count Then
        QuotedVersusBareClass = 2
        Exit Function
    End If
    Dim worst As Long
    Dim i As Long
    For i = 1 To la.Count
        Dim c As Long
        c = QuotedVersusBareClass(la.Item(i), lb.Item(i), envN, envT, outText, freeVarUnifies)
        If c > worst Then worst = c
        If worst = 2 Then Exit For
    Next i
    QuotedVersusBareClass = worst
End Function

' PROLOG.10: True iff exactly one of the two leaves carries the
' quoted-string marker AND removing it makes them the same text - so
' `"eng` versus `eng` qualifies, `"eng` versus `"sales` does not (both
' marked), and `eng` versus `sales` does not (neither). Narrow by
' construction: two genuinely different values can never reach it.
Private Function LeavesDifferOnlyByMarker(ByVal x As String, ByVal y As String, _
                                           ByRef outText As String) As Boolean
    Dim xMarked As Boolean, yMarked As Boolean
    xMarked = (Left$(x, 1) = Chr$(34))
    yMarked = (Left$(y, 1) = Chr$(34))
    If xMarked = yMarked Then Exit Function
    If xMarked Then
        If Mid$(x, 2) <> y Then Exit Function
        outText = y
    Else
        If Mid$(y, 2) <> x Then Exit Function
        outText = x
    End If
    LeavesDifferOnlyByMarker = True
End Function

' PROLOG.9: answers one type-test goal - True iff the term is of the kind
' asked about - and binds nothing at all, so it takes envN/envT to
' dereference through and hands nothing back. Factored out of
' SolveGoalList's own dispatch for the live-caught stack-frame reason
' findall's dispatch documents.
'
' EnvWalkInto FIRST, always. Classification is a question about what a
' term IS, and an unresolved variable is not a term but a reference to
' one: without the walk, `(= X 1) (number X)` would look at the atom "X"
' and answer False, misclassifying a term already known to be 1. The
' walk chases a variable-to-variable chain all the way to whatever it is
' ultimately bound to, or back to itself if still free - so a var atom
' surviving the walk is GENUINELY unbound, which is precisely what `var`
' asks and the only reason this function can answer it by inspection.
'
' Then exactly three cases, in this order:
'
'   IsObject   - a Collection, this engine's only compound term. nonvar
'                and compound; never atomic, atom, number or var.
'   IsVarAtom  - a still-free variable. `var` alone. Checked on the RAW
'                text (see LeafIsNumberTerm above): a marked string like
'                `"Hello` must never be asked this question with its
'                marker stripped, or a real text value is reported as an
'                unbound variable - EvalArithTerm's own documented trap.
'   otherwise  - a ground leaf. nonvar and atomic always; then atom
'                versus number, the ONE question that needed deciding,
'                delegated whole to LeafIsNumberTerm so this function
'                holds no opinion about the marker of its own.
'
' The last two arms are each other's exact negation by construction
' rather than by a second lookup - atom is Not number for a ground leaf,
' which is what makes "every atomic term is an atom or a number, never
' both and never neither" true by construction here rather than by
' agreement between two tables that could drift.
' goalTerm is a Variant and is Set into a typed local, never declared As
' Collection in the signature - SolveComparison's and SolveUnification's
' own established shape (above). goals.Item(1) is a Variant, and handing
' one straight to a typed parameter is this project's own recorded VBA
' trap; the two sibling functions already route around it and this does
' not invent a third way.
Private Function SolveTypeTest(ByVal goalTerm As Variant, ByVal kind As String, _
                                envN As Collection, envT As Collection) As Boolean
    Dim lst As Collection
    Set lst = goalTerm
    Dim w As Variant
    VLA_Unify.EnvWalkInto w, lst.Item(2), envN, envT

    If IsObject(w) Then
        Select Case kind
        Case "nonvar", "compound": SolveTypeTest = True
        End Select
        Exit Function
    End If

    Dim raw As String
    raw = CStr(w)

    If VLA_Unify.IsVarAtom(raw) Then
        SolveTypeTest = (kind = "var")
        Exit Function
    End If

    Select Case kind
    Case "nonvar", "atomic": SolveTypeTest = True
    Case "number":           SolveTypeTest = LeafIsNumberTerm(raw)
    Case "atom":             SolveTypeTest = Not LeafIsNumberTerm(raw)
    End Select
End Function

' PROLOG.9: `(between Low High X)`. The one goal in this module that
' ENUMERATES, so unlike every sibling Solve* function it does not return
' a Boolean for the dispatch to act on - it drives the continuation
' itself, once per value, and therefore needs everything the candidates
' loop needs. Factored out of SolveGoalList for the same stack-frame
' reason as its siblings, which matters more here than anywhere: this
' function declares the most locals of any of them, and every one would
' otherwise be paid on EVERY recursive resolution step rather than once
' per between goal.
'
' Low and High through EvalArithTerm - the same walker `(is ...)` and the
' six comparisons use - so they may be expressions, and an unbound or
' non-numeric bound raises the existing {form}-templated arithmetic
' refusals naming "(between ...)" rather than a form the user never
' wrote. Whole numbers are then required of both: between counts, and a
' fractional bound has no next value.
'
' THE RANGE IS CHECKED BEFORE A SINGLE VALUE IS GENERATED, and only in
' generating mode. See this module's own PROLOG.9 header for why the
' ceiling is PROLOG_MAX_STEPS itself rather than a private budget, and
' why the up-front refusal exists even though the per-value step charge
' below already guarantees termination: without it `(between 1 1000000
' X)` stops with prolog-step-ceiling blaming a runaway rule the user does
' not have.
' betweenGoal is a Variant Set into a typed local for the reason
' SolveTypeTest (above) states. `rest` stays typed: it is SolveGoalList's
' own real Collection local, passed to SolveGoalList itself the same way
' already.
Private Sub SolveBetween(ByVal betweenGoal As Variant, ByVal rest As Collection, _
                          clauseDict As Object, envN As Collection, envT As Collection, _
                          freeVarNames As Collection, solutions As Collection, _
                          ByRef stepsTaken As Long, _
                          ByRef cutActive As Boolean, ByRef cutTargetBarrier As Long)
    Dim lst As Collection
    Set lst = betweenGoal
    Dim lowV As Double, highV As Double
    lowV = EvalArithTerm(lst.Item(2), envN, envT, "(between ...)")
    highV = EvalArithTerm(lst.Item(3), envN, envT, "(between ...)")
    If lowV <> Int(lowV) Then VLA_Messages.RaiseMsg "prolog-between-not-whole-number", "value", NumberToTerm(lowV)
    If highV <> Int(highV) Then VLA_Messages.RaiseMsg "prolog-between-not-whole-number", "value", NumberToTerm(highV)

    ' The third argument decides the MODE, so it is resolved before
    ' anything else is done with the range. EnvWalkInto, not a bare read:
    ' `(is N 5) (between 1 10 N)` must test 5, not look at the atom "N"
    ' and mistake a bound variable for a free one.
    Dim w As Variant
    VLA_Unify.EnvWalkInto w, lst.Item(4), envN, envT

    ' ---- TEST MODE: X already bound. Deterministic, enumerates nothing,
    ' and therefore has no range to refuse - `(between 1 1000000 5)`
    ' succeeds where the generating form of the same range would be
    ' refused. That asymmetry is deliberate and is pinned by a test.
    If IsObject(w) Then VLA_Messages.RaiseMsg "prolog-between-not-a-number", "value", RenderBoundValue(w)
    Dim rawX As String
    rawX = CStr(w)
    If Not VLA_Unify.IsVarAtom(rawX) Then
        ' A marked leaf is an ATOM, never a number - LeafIsNumberTerm
        ' (above) is asked rather than IsInvariantNumericString directly,
        ' so `between` cannot quietly adopt a different reading of the
        ' quoted-string marker from the one `number`/`atom` use. A text
        ' cell reading 5 is not the number 5 here, and says so by name.
        If Not LeafIsNumberTerm(rawX) Then VLA_Messages.RaiseMsg "prolog-between-not-a-number", "value", LeafText(rawX)
        Dim xv As Double
        xv = VLA_Relation.InvariantVal(rawX)
        If xv <> Int(xv) Then VLA_Messages.RaiseMsg "prolog-between-not-whole-number", "value", LeafText(rawX)
        If xv >= lowV And xv <= highV Then
            SolveGoalList rest, clauseDict, envN, envT, freeVarNames, solutions, stepsTaken, cutActive, cutTargetBarrier
        End If
        Exit Sub
    End If

    ' ---- GENERATING MODE: X free. An empty range (Low > High) simply
    ' produces nothing - zero solutions, never an error, exactly as a
    ' predicate with no matching candidate does. `(between 1 N X)` with N
    ' bound to 0 must yield no rows rather than stopping the query.
    If highV < lowV Then Exit Sub
    ' PROLOG_MAX_STEPS - 1, not PROLOG_MAX_STEPS: the dispatch above
    ' already charged ONE step for this goal before calling here, so a
    ' range of exactly the ceiling cannot fit and would fall through this
    ' check only to die at the step ceiling a value later - producing the
    ' misleading "a rule that recurses without ever reaching a base case"
    ' text that this refusal exists to replace. Off by one here would
    ' therefore not be a rounding detail; it would silently restore the
    ' exact defect. The largest range that fits is PROLOG_MAX_STEPS - 1.
    If (highV - lowV + 1) > (PROLOG_MAX_STEPS - 1) Then
        VLA_Messages.RaiseMsg "prolog-between-range-too-wide", _
            "low", NumberToTerm(lowV), "high", NumberToTerm(highV), _
            "count", NumberToTerm(highV - lowV + 1), "max", PROLOG_MAX_STEPS
    End If

    Dim v As Double
    For v = lowV To highV
        ' Charged per value, exactly as the candidates loop charges per
        ' candidate: a generated value IS a unit of resolution work, and
        ' PROLOG_MAX_STEPS is a total-work ceiling for the whole query
        ' (SolveIsolated's own header). The up-front range check above
        ' cannot replace this - it bounds THIS goal's enumeration, while
        ' this bounds the whole query, including several between goals
        ' nested inside each other, each individually under the range
        ' ceiling and together far over the work ceiling.
        stepsTaken = stepsTaken + 1
        If stepsTaken > PROLOG_MAX_STEPS Then VLA_Messages.RaiseMsg "prolog-step-ceiling", "steps", PROLOG_MAX_STEPS

        ' A FRESH clone per value, never one threaded through the loop -
        ' the identical discipline the candidates loop follows for the
        ' identical reason. X binds to this value only; the next value
        ' must find X free again, and a continuation that binds other
        ' variables while exploring THIS value must not leak them into
        ' the next. UnifyTwoWay is used rather than a direct env write so
        ' the binding goes through the one primitive that owns the
        ' occurs check and the env representation.
        Dim betN As Collection, betT As Collection
        VLA_Unify.UnifyEnvClone envN, envT, betN, betT
        If VLA_Unify.UnifyTwoWay(rawX, NumberToTerm(v), betN, betT) Then
            SolveGoalList rest, clauseDict, betN, betT, freeVarNames, solutions, stepsTaken, cutActive, cutTargetBarrier
        End If

        ' The cut signal, honoured exactly as every loop between a cut's
        ' own origin and its firing site must (the candidates loop below
        ' has the hand-traced reasoning). Stop generating unconditionally
        ' when it is active, and NEVER absorb it: absorption belongs only
        ' to the loop that selected the clause the `!` fired inside,
        ' identified by myStep = cutTargetBarrier, and this loop selects
        ' no clause and creates no barrier of its own. So it is the "any
        ' other loop leaves it set" case, letting the signal keep
        ' propagating to its true origin further up the stack. Without
        ' this, `!` after a between goal would fail to prune the
        ' generator and the cut would silently do nothing.
        If cutActive Then Exit For
    Next v
End Sub

' PROLOG.5.3: findall's own harvest, factored OUT of SolveGoalList's own
' dispatch deliberately - see that dispatch's own comment for the live-
' caught reason (every local declared directly in SolveGoalList bloats
' its own per-step recursive stack frame; this function's own locals
' live in a separate frame instead, paid only once per findall
' invocation, not once per resolution step). Collects Template's own
' variables (CollectTemplateVars, above - never goal-aware CollectVars,
' see this module's own PROLOG.5.3 header for why), runs Goal in
' isolation via SolveIsolated (the SAME sub-call microscope SolveNegation
' itself wraps), and substitutes each returned solution tuple back into
' Template's own structure (SubstituteTemplate, above) to produce one
' harvested list element.
Private Function HarvestFindallBag(ByVal findallGoal As Collection, clauseDict As Object, _
                                    envN As Collection, envT As Collection, _
                                    ByRef stepsTaken As Long) As Collection
    ' IsObject-branched, never a bare "template = findallGoal.Item(2)" -
    ' this module's own already-documented trap (EnvWalkInto's own
    ' header): a plain Variant assignment from an expression that MIGHT
    ' be an object (a compound Template, e.g. (pair Name Age)) invokes
    ' the object's own default member instead of copying the reference,
    ' rather than the bare-atom Template case (a plain variable like X)
    ' this would otherwise pass silently for.
    Dim template As Variant
    If IsObject(findallGoal.Item(2)) Then
        Set template = findallGoal.Item(2)
    Else
        template = findallGoal.Item(2)
    End If
    Dim templateVars As New Collection
    CollectTemplateVars template, templateVars
    Dim harvestedTuples As Collection
    Set harvestedTuples = SolveIsolated(findallGoal.Item(3), clauseDict, envN, envT, templateVars, stepsTaken)
    Dim bag As New Collection
    ' Collection, not Variant - SubstituteTemplate's own solutionTuple
    ' parameter is Collection (implicit ByRef, this file's own
    ' established convention for Collection params), and VBA's ByRef
    ' requires an exact type match, unlike ByVal - a Variant loop
    ' variable here compiled ("ByRef argument type mismatch"), live-
    ' caught. Every element harvestedTuples ever holds really is a
    ' Collection - SolveGoalList's own base case always builds `Dim sol
    ' As New Collection` regardless of how many free variables it holds.
    Dim tup As Collection
    For Each tup In harvestedTuples
        bag.Add SubstituteTemplate(template, templateVars, tup)
    Next tup
    Set HarvestFindallBag = bag
End Function

' PROLOG.5.3: fully resolves term through envN/envT, recursively - not
' just the top-level chain EnvWalkInto (VLA_Unify.bas) already follows,
' but every variable position NESTED INSIDE a compound value too. A
' real, pre-existing gap in SolveGoalList's own base case (above) found
' while building findall's own Template harvesting, which genuinely
' needs this - see this module's own PROLOG.5.3 header for the full
' incident (a rule head like `(r X (box X))`, where X is bound to a real
' value only by a LATER body goal, used to render with that inner
' variable's own raw freshened name still showing). dest is a ByRef out
' param, not a function return - EnvWalkInto's own precedent
' (VLA_Unify.bas), the identical reason: w may be either an object or
' not, genuinely ambiguous at the point of assignment (unlike
' FreshenTerm's own two branches, each individually type-known), so a
' bare "dest = w" would risk invoking an object's own default member
' instead of copying the reference on whichever branch actually runs.
' Unification itself was never affected by this gap - UnifyTwoWay
' already re-derefs correctly at every nesting level, since every
' recursive call starts with its own EnvWalkInto; only this OUTPUT-time
' resolution, which used to walk a value one level and stop, needed it.
'
' Deliberately does NOT follow FreshenTerm/CollectVars/SubstituteTemplate
' 's own "position 1 is a functor, never touched" convention - reasoned
' through explicitly, not copied by reflex, because that convention's
' own justification (position 1 of a compound TERM, as PARSED from
' program text, is ALWAYS a predicate/functor symbol, enforced at parse
' time by TermPredName/ValidateBodyItem) does not hold for every value
' this function may be asked to walk: findall's own harvested Bag
' (below) is a genuinely NEW kind of structure this engine did not have
' before PROLOG.5.3 - a plain, functor-less LIST where EVERY position,
' including the first, is an ordinary element, not a keyword. Recursing
' into position 1 unconditionally costs nothing extra for an ordinary
' functor-headed term (EnvWalkInto on a lowercase, non-variable atom
' like "color" or "pair" is a same-string no-op) and removes the need to
' reason about whether a Bag element could ever legitimately need
' resolving in that specific position - simpler and strictly safer than
' carving out a position-1 exception that would only be correct by
' relying on Bag always arriving here already fully resolved.
Private Sub ResolveTermDeep(ByRef dest As Variant, ByVal term As Variant, envN As Collection, envT As Collection)
    Dim w As Variant
    VLA_Unify.EnvWalkInto w, term, envN, envT
    If Not IsObject(w) Then
        dest = w
        Exit Sub
    End If
    Dim lst As Collection, outLst As New Collection
    Set lst = w
    Dim i As Long
    For i = 1 To lst.Count
        Dim childResolved As Variant
        ResolveTermDeep childResolved, lst.Item(i), envN, envT
        outLst.Add childResolved
    Next i
    Set dest = outLst
End Sub

' ---------------------------------------------------------------------
'  Output
' ---------------------------------------------------------------------

' A ground leaf's own display text - VLA's own reader marks a quoted-
' string token with a leading literal " character (VLA_Datalog.bas's
' own AtomText, the identical convention, repeated here since rendering
' a leaf to a cell has nothing to do with unification and doesn't
' belong in VLA_Unify.bas).
Private Function LeafText(ByVal raw As Variant) As String
    Dim s As String
    s = CStr(raw)
    If Left$(s, 1) = Chr$(34) Then
        LeafText = Mid$(s, 2)
    Else
        LeafText = s
    End If
End Function

Private Function RenderBoundValue(ByVal v As Variant) As String
    If IsObject(v) Then
        RenderBoundValue = VLA.VlaWriteForm(v)
    Else
        RenderBoundValue = LeafText(v)
    End If
End Function

' Header row first, one column per free query variable, one row per
' solution - SQL/DATALOG's own frozen SD-4 contract. A query with NO
' free variables (every conjunct fully ground, e.g. a pure yes/no
' check) has no meaningful column shape at all, so it collapses to a
' single boolean scalar instead - TRUE iff at least one solution was
' found - DATALOG's own "a zero-row result is a safe blank scalar, not
' a crash" instinct applied to the analogous degenerate shape here.
Private Function BuildSpilledArray(freeVarNames As Collection, solutions As Collection) As Variant
    Dim nCols As Long
    nCols = freeVarNames.Count
    If nCols = 0 Then
        BuildSpilledArray = (solutions.Count > 0)
        Exit Function
    End If
    Dim nRows As Long
    nRows = solutions.Count
    Dim arr() As Variant
    ReDim arr(1 To nRows + 1, 1 To nCols)
    Dim c As Long
    For c = 1 To nCols
        arr(1, c) = CStr(freeVarNames.Item(c))
    Next c
    Dim r As Long
    For r = 1 To nRows
        Dim sol As Collection
        Set sol = solutions.Item(r)
        For c = 1 To nCols
            arr(r + 1, c) = RenderBoundValue(sol.Item(c))
        Next c
    Next r
    BuildSpilledArray = arr
End Function

' ---------------------------------------------------------------------
'  The worksheet function
' ---------------------------------------------------------------------

' PROLOG.6: VLA_Datalog.TableArgName's own exact shape - VLA_Relation.
' TableArgResolve cannot itself call VLA_Messages (its own LAYER 0
' contract), so this thin wrapper turns its shared category code back
' into PROLOG's own wording.
Private Function TableArgName(ByVal v As Variant) As String
    Dim ok As Boolean, reason As String
    TableArgName = VLA_Relation.TableArgResolve(v, ok, reason)
    If ok Then Exit Function
    Select Case reason
    Case "not-a-range"
        VLA_Messages.RaiseMsg "prolog-table-not-a-range"
    Case "noncontiguous"
        VLA_Messages.RaiseMsg "prolog-table-noncontiguous-columns"
    Case "needs-a-name"
        VLA_Messages.RaiseMsg "prolog-table-needs-a-name"
    End Select
End Function

' PROLOG.6: a table cell's own real Excel Value2 -> a PROLOG-native leaf
' (a bare String, this engine's own only ground-leaf representation) -
' this module's own PROLOG.6 header has the full reasoning for both
' branches (locale-invariant numeric rendering via Str$, and the
' quoted-atom marker every non-numeric value gets so a capitalized text
' cell, or a live Excel Boolean, is never misdetected as an unbound
' variable).
Private Function TableCellToTerm(ByVal v As Variant) As String
    If VLA_Relation.ValueIsNumericType(v) Then
        TableCellToTerm = Trim$(Str$(CDbl(v)))
    Else
        TableCellToTerm = Chr$(34) & CStr(v)
    End If
End Function

' PROLOG.6: one table row (VLA_Relation.RangeToRows' own 1-based Variant
' array) -> one ground fact term, predName at position 1 exactly like a
' hand-authored (fact (predName ...)) - IsObject-branched only for
' predName (a plain String, never actually an object here, but Add's
' own Variant parameter needs no branching at all - only a bare
' assignment does, this module's own already-documented trap).
Private Function TableRowToFact(ByVal predName As String, ByRef row As Variant) As Collection
    Dim term As New Collection
    term.Add predName
    Dim c As Long
    For c = LBound(row) To UBound(row)
        term.Add TableCellToTerm(row(c))
    Next c
    Set TableRowToFact = term
End Function

' PROLOG.6: the real engine, directly callable with a HAND-BUILT
' clauseDict/headerMap and no live workbook at all - VLA_Datalog.
' DatalogRun's own precedent, `DATALOG()`'s own worksheet-function-
' around-a-testable-core split reused rather than invented fresh. This
' is what makes TestPrologKeyedAtoms a PURE test of the desugaring
' mechanism itself.
Public Function PrologRun(ByVal clausesText As String, ByVal clauseDict As Object, ByVal headerMap As Object) As Variant
    Dim queryConjuncts As Collection, freeVarNames As Collection
    ParseProgram clausesText, clauseDict, queryConjuncts, freeVarNames, headerMap

    Dim envN As Collection, envT As Collection
    Set envN = New Collection
    Set envT = New Collection
    Dim solutions As New Collection
    Dim stepsTaken As Long
    Dim cutActive As Boolean, cutTargetBarrier As Long
    SolveGoalList queryConjuncts, clauseDict, envN, envT, freeVarNames, solutions, stepsTaken, cutActive, cutTargetBarrier

    ' PROLOG.12: only when the whole query found NOTHING. This is the
    ' half PROLOG.10 deliberately left open, closed the one way that does
    ' not cost anything on a query that works: the scan runs POST HOC,
    ' after solving, so a successful query never pays for it and the
    ' solver itself is untouched - no state threaded through
    ' SolveGoalList, no growth in the recursive frame that PROLOG.5.3's
    ' own stack-overflow incident made expensive, and nothing at all on
    ' the backtracking hot path.
    '
    ' Bounded by construction: a query that completes with zero solutions
    ' has by definition stayed under PROLOG_MAX_STEPS, so there is no
    ' large-table case here - one that really did scan ten thousand rows
    ' raised the step ceiling long before reaching this line.
    If solutions.Count = 0 Then DiagnoseQuotedVersusBare queryConjuncts, clauseDict

    PrologRun = BuildSpilledArray(freeVarNames, solutions)
End Function

' The worksheet-facing entry point - =PROLOG(clauses, table1, table2,
' ...). PROLOG.6: each further argument is a live cell range (an Excel
' Table or a plain named range, TableArgName's own resolution, DATALOG's
' own exact convention) - every one of its own rows becomes an ordinary
' ground fact clause (TableRowToFact, above), added to clauseDict BEFORE
' clausesText is ever parsed ("table rows first, then whatever the
' clauses text itself defines, in written order" - this module's own
' PROLOG.6 header). Its own header row (VLA_Relation.RangeColumnNames)
' is captured into headerMap the identical way, feeding ParseProgram's
' own new keyed-atom desugaring.
Public Function PROLOG(ByVal clauses As String, ParamArray tables() As Variant) As Variant
    On Error GoTo fail
    Dim clauseDict As Object
    Set clauseDict = VLA_Runtime.VlaDictNew()
    Dim headerMap As Object
    Set headerMap = VLA_Runtime.VlaDictNew()

    Dim ti As Long
    For ti = LBound(tables) To UBound(tables)
        Dim nm As String
        nm = TableArgName(tables(ti))
        If IsReservedPredicateName(nm) Then VLA_Messages.RaiseMsg "prolog-reserved-predicate-name", "name", nm
        Dim rows As Collection
        Set rows = VLA_Relation.RangeToRows(tables(ti))
        Dim rw As Variant
        For Each rw In rows
            GetOrCreateClauseList(clauseDict, nm).Add MakeClause(TableRowToFact(nm, rw), New Collection)
        Next rw
        Dim colsOk As Boolean
        Dim cols As Collection
        Set cols = VLA_Relation.RangeColumnNames(tables(ti), colsOk)
        If colsOk Then VLA_Runtime.VlaDictSet headerMap, nm, cols
    Next ti

    PROLOG = PrologRun(clauses, clauseDict, headerMap)
    Exit Function
fail:
    PROLOG = "#PROLOG! " & Err.Description
End Function
