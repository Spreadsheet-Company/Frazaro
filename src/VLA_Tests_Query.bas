Attribute VB_Name = "VLA_Tests_Query"
Option Explicit
Public Const VLA_TESTS_QUERY_VERSION As String = "OPTIMIZE.1"
' DATALOG.15: TestSpillHeaders (new, pure - VLA_Relation's spill header
' check, RefersTo matching and text, the error-value reason, and every
' new refusal's words, with no live workbook) and TestSpillHostTable (new,
' host-required - a real dynamic-array spill, named through Name Manager,
' read by DATALOG, SQL and PROLOG through their real worksheet functions,
' with a part of a spill and a plain named range pinned unchanged).
' PROLOG.6: TestPrologKeyedAtoms (new, pure - VLA_Prolog.PrologRun called
' directly with a hand-built headerMap, no live workbook, DATALOG.5's own
' TestDatalogKeyedAtoms precedent) and TestPrologHostTable (new, host-
' required - VLA_Prolog.PROLOG called through the real worksheet
' function against a live Table). See each Sub's own header, immediately
' above it, for full coverage detail - a fully-keyed query resolving
' correctly, in derivation order; a non-table-sourced predicate's own
' compound-term arguments left untouched (the real ambiguity PROLOG's
' own compound-term-permitting grammar creates that DATALOG never had);
' every named keyed-atom refusal; a rule body composing full keying with
' real backtracking; two discriminating proofs that desugaring reaches
' inside both not's and findall's own nested Goal; a live Table's own
' rows becoming ordinary ground fact clauses; a numeric column rendering
' locale-invariantly (Str$, not CStr); a capitalized text column
' round-tripping as ground data, never misdetected as a variable; a
' capitalized Table name working end to end; and named-column syntax
' reached through the real =PROLOG(...) function against a live Table.
' PROLOG.5.4: TestPrologCut (new, pure - VLA_Prolog.PROLOG called
' directly). See its own header, immediately above its own Sub, for full
' coverage detail - cut committing to the first matching clause of the
' SAME predicate call it appears in; the flagship proof this item's own
' design required, hand-traced before being written as an assertion: cut
' also prunes EARLIER goals in the SAME clause body, not just the
' current predicate's own remaining clauses (a two-predicate rule body
' collapsing from four possible solutions to exactly one); the opposite
' boundary, equally load-bearing - cut does NOT escape past its own
' clause's own selection to prune a caller's later, unrelated choice
' points, nor does it affect goals to its OWN right in the same body,
' which keep their full normal backtracking; cut's own opacity across
' both `not`'s and `findall`'s isolated sub-search (real Prolog's own
' rule - confirmed to fall out for free from SolveIsolated's own fresh
' local cut-state, not requiring any escape-detection code), each proven
' by a query where a LEAKED signal would visibly prune the outer choice
' point and a correctly-opaque one would not; multiple cuts in one
' clause body composing without double-firing or corrupting each other's
' own barrier; and the shared PROLOG_MAX_STEPS budget re-proven against
' the identical genuinely-non-terminating rule every other PROLOG.5.x
' item already re-proves it against, now carrying the two new cut-signal
' parameters through every recursive frame.
' PROLOG.5.3: TestPrologFindall (new, pure - VLA_Prolog.PROLOG called
' directly). See its own header, immediately above its own Sub, for full
' coverage detail - a bare-variable Template harvesting every solution
' into a list, in derivation order; the zero-solutions case collapsing
' to an empty list, never an error; a COMPOUND Template (proving
' SubstituteTemplate's own reconstruction, not just a flat value list);
' findall composing with an OUTER already-bound variable across real
' backtracking (the classic "group by" idiom); an already-bound Bag
' argument genuinely CHECKED via unification, both matching and
' mismatching; the shared PROLOG_MAX_STEPS budget and the error-
' propagates-out cases (findall's own half of "what must never leak back
' out", the sibling of PROLOG.5.2's own); the real design fork this
' item's own build found - a Template that itself contains a literal
' `(not X)`-shaped sub-term, as plain DATA, still has X collected
' correctly, proving CollectTemplateVars' own deliberately-simpler,
' non-goal-aware recursion (not a reuse of CollectVars) was the right
' call; and every named (findall ...) shape refusal. A separate
' regression pin for a real, pre-existing PROLOG.4-era gap found while
' building this item (a free variable resolving to a compound term that
' itself still contains an unresolved nested variable) lives in
' TestPrologRules instead, since the bug was in PROLOG.4's own base
' case, not findall-specific.
' PROLOG.5.2: TestPrologNegation (new, pure - VLA_Prolog.PROLOG called
' directly). See its own header, immediately above its own Sub, for full
' coverage detail - negation succeeding/failing on a directly-ground
' goal both ways; negation as a computed FILTER over real backtracking
' with an already-bound variable (the correct, intended NAF usage
' pattern); negation reached through a RULE body with its own fresh
' existential variable, freshened per candidate (the orphan/parent
' shape); the real design fork this item's own scoping called for -
' hand-traced, not assumed - proving a variable appearing ONLY inside a
' negated goal never becomes a phantom output column, rather than
' silently spilling its own unresolved name as text; the shared
' PROLOG_MAX_STEPS budget, proven by running a genuinely non-terminating
' negated goal to real exhaustion rather than trusting it inherits the
' ceiling; an error (divide-by-zero) raised while proving a negated goal
' propagating all the way out rather than being swallowed as mere
' failure - the other half of "what must never leak back out"; not
' composing with is; nested (double) negation, both the TRUE and FALSE
' case, proving the recursive dispatch and CollectVars' own opacity rule
' both survive nesting; and every named (not ...) shape refusal,
' including one reused verbatim from ValidateBodyItem's own recursive
' call into a malformed inner Goal (atom-not-a-list) rather than a
' not-specific twin. (A bare ! as not's own goal was ALSO refused,
' cut-not-yet-supported, at this item's own build time - superseded at
' PROLOG.5.4, see TestPrologCut instead.)
' PROLOG.5.1: TestPrologArithmetic (new, pure - VLA_Prolog.PROLOG called
' directly). See its own header, immediately above its own Sub, for full
' coverage detail - a rule computing a fresh variable through a
' fact -> is chain; nested arithmetic evaluated bottom-up in a pure
' query (a shape DATALOG's own flat (let Z (op X Y)) structurally can't
' express); is checking an already-bound target via ordinary
' unification, both matching and mismatching (real Prolog's own is/2
' semantics, free from reusing UnifyTwoWay); two chained is calls in one
' rule body sharing the same threaded env; and every named refusal,
' including a non-numeric value reached through a RULE (not a hand-typed
' literal, proving the runtime re-check catches what parse-time
' shape-only validation structurally cannot) and all four reserved
' predicate names (is/not/findall/!) tried even though only is is built.
' PROLOG.4: TestPrologRules (new, pure - VLA_Prolog.PROLOG called
' directly, no live workbook needed). See its own header, immediately
' above its own Sub, for full coverage detail - a non-recursive rule (a
' free-variable and a boolean-scalar shape), a predicate defined by a
' fact AND a rule together (both firing, in written order), the classic
' ancestor/parent recursive transitive closure (asserted against every
' solution's own value AND order, the shape a clause-freshening bug
' would silently corrupt), a fully-ground recursive query that correctly
' terminates on FALSE, the malformed-(rule ...)-shape refusal, and the
' resolution-step ceiling exercised to REAL exhaustion via a genuinely
' non-terminating rule - deliberately NOT skipped the way SQL.7/DATALOG's
' own round-ceiling tests are, since a single PROLOG resolution step is
' cheap enough to actually run hundreds of, and doing so is the only way
' to learn empirically whether PROLOG_MAX_STEPS is already too high for a
' real machine's own native VBA call-stack depth (VLA_Prolog.bas's own
' module header has the full reasoning - confirmed live: 1000 genuinely
' overflowed, lowered to 200). Two more pins, owner-requested after
' owner-verified live handoff testing: a cyclic graph's own non-
' terminating reachability query (the ceiling generalizing beyond the
' single-clause loop shape to a real two-clause recursive predicate whose
' non-termination comes from the data, not the rule) and occurs-check
' reached through a real freshened rule head rather than only the
' hand-built primitive TestUnifyTwoWay (PROLOG.2) already covers.
' PROLOG.2: TestUnifyTwoWay (new, pure - hand-built form literals,
' VLA_Unify.UnifyTwoWay called directly). Threads ONE shared envN/envT
' pair across several successive calls in the same test, deliberately -
' proves variable-to-variable chaining (X unify Y, then Y bound, then X
' resolves transitively) and that both sides of one compound term
' carrying a variable in the same position get correctly linked, not
' just unified independently - PROLOG.4's own SLD resolution will
' thread one environment across a whole subgoal chain the same way, so
' this is proving that shape works, not just a test convenience. Also
' pins the occurs-check refusal (a variable can't bind to a term
' containing itself) and its own false-positive-avoidance case (a
' DIFFERENT variable's name appearing in the term must never trip it).
' PROLOG.1: TestUnify (new, pure - hand-built form literals, VLA_Unify.
' UnifyOneWay called directly) and TestGRenderUnify (new, pure - through
' the real EnglishRenderText entry point) join this file rather than a
' separate one, corrected mid-session from an initial wrong call to wire
' both into VlaSelfTest instead: VLA_Unify.bas is PROLOG.1, the whole
' point of this file's own "one file per roadmap SECTION, not one per
' engine module" header note above - QUERY AND LOGIC's tests stay
' together here regardless of which shipped feature (G-RENDER) a given
' piece of substrate also happens to serve. See each Sub's own header
' for what it pins; VLA_Unify.bas's own module header has the full
' PROLOG.1 story (the real repeated-variable-consistency bug this item
' found and fixed, verified against the loaded English corpus).
' SQL.3: TestSqlJoin (new, pure - hand-built table records, VLA_Sql.
' SqlRunJoin called directly, no live workbook needed at all - the
' identical seam TestSql's own SQL.1/SQL.2 suite already established)
' covers a basic two-table equi-join, qualified AND unqualified column
' resolution, SELECT * listing every joined table's own columns in join
' order, WHERE after a JOIN, a composite ON clause (one equi term plus
' one residual term in the SAME ON, correctly split and both honored), a
' THIRD table folded in by a second JOIN whose own ON references a
' column the FIRST join already added (left-to-right folding, proven not
' just asserted), DISTINCT after a JOIN, "INNER JOIN" and bare "JOIN"
' producing identical results, and every named refusal this item's own
' roadmap calls for: an unqualified column ambiguous across two joined
' tables, the same table named twice (no aliasing support), a JOIN
' naming a table that wasn't passed, LEFT JOIN refused by name rather
' than silently misread as INNER, and =SQL() called with zero table
' arguments at all. TestSqlHostTable gains a live two-table JOIN
' scenario, a real ListObject on each side - and retires its own OLD
' SQL.1-era "passing two table arguments is refused" check, an
' assumption SQL.3 deliberately overturns (multiple table arguments are
' now the whole point), replaced rather than left stale.
'
' DATALOG.5: TestDatalogKeyedAtoms (new, pure - a hand-built (Name,
' Salary, Dept) "staffing" array, fed through RelFromRange's own new
' hasHeader:=True flag plus a hand-built headerMap, so no live workbook
' is needed at all) proves the item's own motivating example (salary
' keyed, filtered > 80000), the one correctness property that would be
' silently wrong if an omitted column's own anonymous variable name
' were keyed on column position alone rather than atom occurrence too
' (two keyed atoms each omitting the same columns must derive the full
' cross product, never be silently joined on a shared anonymous name),
' a partial-keying + count combination, full-arity keying, and every
' parse-time refusal this item's own roadmap names by id - including
' the one highest-risk case named explicitly, a keyed pair whose own
' value position is itself a nested compound term, still refused
' through ParseAtom's own pre-existing check rather than silently
' accepted. TestDatalogHostTable gains a fourth scenario, necessarily
' host-required (a plain 2D array has no live header to fold at all): a
' mixed-case Table name AND mixed-case header (Staffing/Salary, the
' item's own words) queried through all-lowercase rule text, proving
' both fold invariantly through VLA_Identity.Fold.
'
' SQL.2: TestSql grows computed SELECT expressions, AS aliases, and
' DISTINCT - all pure, appended to the existing suite rather than a new
' Sub, since SQL.2 has no new host-specific mechanism of its own (unlike
' DATALOG.3's real column-scoping): a computed arithmetic column with
' its own AS alias (value-checked, not just row-counted, via the new
' ValueAt helper alongside CountRows/HeaderAt); AS renaming a bare
' column; an un-aliased computed column refused at parse time; DISTINCT
' on one column, on *, and combined with WHERE (proving WHERE filters
' BEFORE DISTINCT, not after); operator precedence (* before +) and its
' parenthesized override; unary minus; WHERE reusing the SAME
' arithmetic grammar with no parens needed; division by zero and a
' non-numeric operand refused; and the one documented SQL.2 limit (a
' comparison's own LEFT operand can never itself start with '(' at the
' very top of a condition) confirmed as a clean refusal, not a silent
' misparse.
'
' DATALOG.4: TestDatalogBuiltins - comparison filters and the (let ...)
' arithmetic binding form, entirely pure (arithmetic/comparison over
' Datalog facts needs no live workbook at all, unlike DATALOG.3's own
' column-scoping). Covers: a comparison over a TABLE-sourced (real
' Double, via RelFromRange fed a hand-built numeric 2D array) column;
' the SAME comparison over fact-block values (always plain VBA Strings)
' - deliberately chosen operands (9 vs 10 against a threshold of 8)
' that would give the WRONG answer under a naive lexical-text fallback,
' proving the numeric-looking-STRING policy actually fires; a text
' fallback comparison (<>); all four arithmetic operators; a RECURSIVE
' rule accumulating (let ...) across rounds (path-length over a 3-edge
' chain) - the one test that would have caught this item's own
' round-loop/ComputeStrata fix (BI_CMP/BI_LET are no longer valid
' semi-naive delta positions, unlike before DATALOG.4 when BI_POS was
' the only non-full-relation kind) had it been wrong, since a bug there
' would either mis-derive the longer paths or loop needlessly rather
' than fail loudly; a comparison composing with (not ...) in one rule
' body; and the full set of parse-time/eval-time refusals (unbound
' operand variables for both shapes, a reused let result variable,
' wrong operand counts, an unrecognized arithmetic operator, division
' by zero, a non-numeric operand, a malformed (let ...) shape, and a
' non-variable let result name).
'
' SQL.1: TestSql (pure - a hand-built columnNames/rows pair, calling
' VLA_Sql.SqlRun directly, VLA_Datalog.DatalogRun's own precedent) and
' TestSqlHostTable (host-required - a real ListObject, calling the
' actual =SQL(...) worksheet function). TestSql's own fact table
' DELIBERATELY includes a duplicate row (proving SQL's bag semantics
' survive - VLA_Relation.RangeToRows, unlike RelFromRange, never dedups)
' plus an AND/OR precedence case and a parenthesized override of it -
' the two things a hand-rolled tokenizer+parser is most likely to get
' subtly wrong.
'
' DATALOG.3: TestDatalogHostTable gains a third scenario - a contiguous
' 2-of-3-column slice of a live Table narrows the resulting relation's
' arity while the predicate name stays the table's own ListObject.Name,
' and a non-contiguous (Ctrl-selected, Union-built) column pick on the
' same table is refused rather than silently guessed. Necessarily host-
' required, like the other two scenarios here: a plain 2D array has no
' .ListObject to narrow at all, so this capability has no pure-test seam.
'
' DATALOG.1/.2: TestDatalogNegation (stratified `not`) and
' TestDatalogAggregation (`count`/`sum`) added, both pure. TestDSLs also
' gains TestDatalogHostTable - the first host-required sub in this
' module - a real ListObject, a real Name-Manager alias pointing at the
' same cells, and the actual =DATALOG(...) worksheet function called
' directly, live: proves a table reached through its alias still
' resolves only by its own ListObject.Name, and that a recursive rule
' sourced from a live Table's own DataBodyRange derives correctly - both
' real bugs this engine's MVP found that no pure test (RelFromRange fed
' a hand-built array) could ever have caught.

' =====================================================================
'  VLA_Tests_Query - the QUERY AND LOGIC section's own test suite,
'  split out before it ever shared a file with the language's own
'  tests, not after (VLA_Tests_Grammar.bas's own F.8 split happened at
'  4,945 lines and REBUILD.md's R4 budget; this module exists so that
'  crossing never has to happen a second time). SQL/PROLOG/OPTIMIZE's own
'  future test subs belong here too, next to TestDatalog, as each
'  engine ships - one file per roadmap SECTION, not one per engine
'  module.
'
'  Independent of VlaSelfTest, VLA_Tests_Host.bas's own shape rather
'  than VLA_Tests_Grammar.bas's: its own Private mPass/mFail/
'  mFailedNames, its own Private Report (an unqualified Report call
'  inside TestDatalog resolves HERE, module-locally, never reaching
'  VLA_Tests.Report - VBA's own name-resolution rule, not a special
'  case), and its own entry point, TestDSLs. VlaSelfTest never calls
'  into this module at all - the owner's own instruction: an
'  aspirational, appetite-boxed tranche should not lengthen or bog
'  down the language's own everyday test run, and should have exactly
'  one command a reader would guess to run it standalone.
'
'  TestDSLs now needs a live workbook (TestDatalogHostTable, below) -
'  every real bug this engine's own MVP found (a table's predicate name
'  resolving to its own ListObject.Name rather than a Name-Manager
'  alias pointing at the same cells; a recursive join failing when facts
'  came from a live Table) only ever showed up through the real
'  =DATALOG(...) path, never through RelFromRange fed a hand-built
'  array. VLA_Tests_Host.bas's own pattern (a real ListObject on a real
'  sheet, the actual worksheet function called directly, asserted on
'  live) - not its own dispatch list, on purpose: this tranche stays
'  isolated from VlaSelfTestHost too, TestDSLs its one command.
'
'  LAYER:     Harness (dev-only; never ships - see REBUILD.md SS3
'             Layer 5)
'  MAY CALL:  VLA_Relation, VLA_Datalog, VLA_Unify (UnifyOneWay,
'             TestUnify's own function under test), VLA_English
'             (EnglishResetGrammar/EnglishAddPhrase/EnglishRenderText,
'             TestGRenderUnify's own G-RENDER regression pin), VLA_Runtime
'             (VlaDictNew/Set/Get - building a table argument's
'             baseRelations dict directly, the same as any other caller of
'             VLA_Datalog.DatalogRun; VlaEnsureSheet, TestDatalogHostTable's
'             own scratch-sheet setup, VLA_Tests_Host.bas's own precedent).
'  SHIPS:     nowhere - VLA_DevRig.bas's reload list only, never
'             VLA_Build.bas's shipped manifest (VLA_Tests_Grammar.bas's
'             own precedent exactly).
'  PAYS INTO: the QUERY AND LOGIC section's own DATALOG item; keeps
'             VLA.bas/VLA_Interpreter.bas/VLA_Runtime.bas AND
'             VlaSelfTest's own everyday run untouched by a still-
'             aspirational, appetite-boxed tranche's own churn.
'  REASON:    the four engines are one section but not one concern
'             each - a shared substrate (VLA_Relation.bas) plus four
'             independent grammars, none of them demand-driven
'             (specimens: 1, the owner's own stated want, not corpus
'             pressure). Filing both their tests AND their verdict
'             separately from day one means the VLA engine's own test
'             files (VLA_Tests, VLA_Tests_Grammar, VLA_Tests_Host)
'             never have to notice this tranche exists, and a run of
'             VlaSelfTest never gets slower or noisier because of it.
' =====================================================================

Private mPass As Long
Private mFail As Long
Private mFailedNames As Collection

Private Sub Report(ByVal name As String, ByVal ok As Boolean, ByVal detail As String)
    If ok Then
        mPass = mPass + 1
        Debug.Print "  PASS  " & name
    Else
        mFail = mFail + 1
        Debug.Print "  FAIL  " & name & " - " & detail
        If Not mFailedNames Is Nothing Then mFailedNames.Add name & " - " & detail
    End If
End Sub

' THE command for QUERY AND LOGIC's whole test suite - run this, not
' VlaSelfTest, to exercise DATALOG (and, as they ship, SQL/PROLOG/
' OPTIMIZE). Needs a live workbook open (TestDatalogHostTable, below,
' builds and tears down a real ListObject); every other sub in this
' module stays pure.
Public Function TestDSLs() As Boolean
    mPass = 0
    mFail = 0
    Set mFailedNames = New Collection
    Debug.Print "===== VLA SELF-TEST (QUERY AND LOGIC / DSLs) ====="

    TestDatalog
    TestDatalogNegation
    TestDatalogAggregation
    TestDatalogBuiltins
    TestDatalogKeyedAtoms
    TestDatalogUnknownPredicate
    TestDatalogGroundQuery
    TestDatalogNegatedQuery
    TestDatalogTextTests
    TestDatalogHostTable
    TestSpillHeaders
    TestSpillHostTable
    TestOptimizeKey
    TestOptimizeForms
    TestOptimizeMemo
    TestOptimizeParity
    TestOptimizeHostTable
    TestUnify
    TestGRenderUnify
    TestUnifyTwoWay
    TestProlog
    TestPrologRules
    TestPrologArithmetic
    TestPrologArithmeticBreadth
    TestPrologComparison
    TestPrologUnification
    TestPrologTypeTests
    TestPrologTypeTestsRest
    TestPrologBetween
    TestPrologQuotedVersusBare
    TestPrologNegation
    TestPrologLists
    TestPrologFindall
    TestPrologCut
    TestPrologControl
    TestPrologText
    TestPrologTextParts
    TestPrologTextTests
    TestPrologImpure
    TestPrologRefusedArity
    TestPrologUnknownPredicate
    TestPrologBudgets
    TestPrologKeyedAtoms
    TestPrologHostTable
    TestPrologConditions
    TestTableArguments
    TestPrologClauses
    TestPrologQuestions
    TestPrologRangeLint
    TestPrologClosure
    TestPrologAnswerShapes
    TestPrologOwnShapes
    TestSql
    TestSqlJoin
    TestSqlSetOps
    TestSqlCte
    TestSqlHostTable
    TestTer8EmptyRules

    Debug.Print "===== DSL SELF-TEST: " & mPass & " passed, " & mFail & " failed ====="
    If mFail > 0 Then
        Debug.Print "===== FAILED TESTS (" & mFailedNames.Count & ") ====="
        Dim fn As Variant
        For Each fn In mFailedNames
            Debug.Print "  FAIL  " & fn
        Next fn
        Debug.Print "===== (paste the block above into a bug report) ====="
    End If
    TestDSLs = (mFail = 0)
End Function

' TER-8: DATALOG and PROLOG with rules or clauses text holding no form -
' empty, or only a comment, as a formula reading a blank cell gives - are
' refused by name. Until TER-8, VLA.bas's reader raised VBA's own
' "Subscript out of range" for such text, and each engine showed it raw.
Private Sub TestTer8EmptyRules()
    Dim result As Variant
    Dim r As String
    result = VLA_Datalog.DATALOG("")
    r = ResultDescribe(result)
    Report "ter-8: DATALOG over empty rules text says the text is empty", _
           ResultTextStartsWith(result, "#DATALOG!") And InStr(1, r, "rules text is empty", vbTextCompare) > 0, "got: " & r
    result = VLA_Datalog.DATALOG("; only a comment")
    r = ResultDescribe(result)
    Report "ter-8: ...and over rules text of only a comment", _
           ResultTextStartsWith(result, "#DATALOG!") And InStr(1, r, "rules text is empty", vbTextCompare) > 0, "got: " & r
    result = VLA_Prolog.PROLOG("")
    r = ResultDescribe(result)
    Report "ter-8: PROLOG over empty clauses text says the text is empty", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "clauses text is empty", vbTextCompare) > 0, "got: " & r
    result = VLA_Prolog.PROLOG("; only a comment")
    r = ResultDescribe(result)
    Report "ter-8: ...and over clauses text of only a comment", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "clauses text is empty", vbTextCompare) > 0, "got: " & r
End Sub

' DATALOG.0: pins VLA_Datalog.bas/VLA_Relation.bas's own MVP - facts,
' a join rule, recursive transitive closure (the item's own "org
' chart" killer case), a within-atom repeated variable, a table
' argument built from a plain 2D array (no live Range at all - the
' seam RelFromRange/DatalogRun exist to make testable), the spilled-
' array shape, and the parse-time/safety refusals. Entirely pure -
' nothing here touches a live workbook.
Private Sub TestDatalog()
    Dim result As Collection
    Dim rel As Collection

    Set result = VLA_Datalog.DatalogRun("(fact (parent tom bob)) (fact (parent bob liz)) (query parent)")
    Set rel = VLA_Runtime.VlaDictGet(result.Item(2), result.Item(1))
    Report "datalog: facts-only relation has 2 tuples", VLA_Relation.RelCount(rel) = 2, "got " & VLA_Relation.RelCount(rel)
    Report "datalog: a raw fact predicate (no defining rule) offers no header names", IsEmpty(result.Item(3)), "expected Empty, got " & TypeName(result.Item(3))
    Report "datalog: no (headless) directive leaves the flag False", CBool(result.Item(4)) = False, "expected False"

    Set result = VLA_Datalog.DatalogRun( _
        "(headless) (fact (parent tom bob)) (fact (parent bob liz)) (query parent)")
    Report "datalog: (headless) sets the flag DatalogRun reports back", CBool(result.Item(4)) = True, "expected True"
    Set rel = VLA_Runtime.VlaDictGet(result.Item(2), result.Item(1))
    Dim arrHeadless As Variant
    arrHeadless = VLA_Relation.RelToSpilledArray(rel, , CBool(result.Item(4)))
    Report "datalog: headless output has exactly N rows, no header row", _
           ResultRowCount(arrHeadless) = VLA_Relation.RelCount(rel) And (ResultCellIs(arrHeadless, 1, 1, "tom") Or ResultCellIs(arrHeadless, 1, 1, "bob")), _
           "shape mismatch"

    ' DATALOG.8: never_true was an UNDEFINED name standing for "a relation
    ' that matches nothing", and an undefined name now refuses. A DEFINED
    ' relation that matches nothing - (thing none), where thing holds only
    ' a - keeps this pin's own subject, a zero-row headless result.
    Set result = VLA_Datalog.DatalogRun("(headless) (fact (thing a)) (rule (nothing_here X) (thing X) (thing none)) (query nothing_here)")
    Set rel = VLA_Runtime.VlaDictGet(result.Item(2), result.Item(1))
    Dim arrEmpty As Variant
    arrEmpty = VLA_Relation.RelToSpilledArray(rel, , CBool(result.Item(4)))
    Report "datalog: a zero-row headless result is a safe blank scalar, not a crash", _
           VLA_Relation.RelCount(rel) = 0 And ResultTextIs(arrEmpty, ""), _
           "got " & TypeName(arrEmpty)

    Set result = VLA_Datalog.DatalogRun( _
        "(fact (parent tom bob)) (fact (parent bob liz))" & _
        " (rule (grandparent X Z) (parent X Y) (parent Y Z))" & _
        " (query grandparent)")
    Set rel = VLA_Runtime.VlaDictGet(result.Item(2), result.Item(1))
    Report "datalog: one-join rule derives exactly 1 tuple", VLA_Relation.RelCount(rel) = 1, "got " & VLA_Relation.RelCount(rel)
    Report "datalog: a rule-derived predicate offers its own head-variable names as headers", _
           Not IsEmpty(result.Item(3)) And CStr(result.Item(3)(1)) = "X" And CStr(result.Item(3)(2)) = "Z", _
           "expected (X, Z)"

    Set result = VLA_Datalog.DatalogRun( _
        "(fact (reports_to alice bob)) (fact (reports_to bob carol)) (fact (reports_to carol dave))" & _
        " (rule (indirect_report X Y) (reports_to X Y))" & _
        " (rule (indirect_report X Y) (reports_to X Z) (indirect_report Z Y))" & _
        " (query indirect_report)")
    Set rel = VLA_Runtime.VlaDictGet(result.Item(2), result.Item(1))
    Report "datalog: transitive closure over a 3-edge chain derives 6 tuples", VLA_Relation.RelCount(rel) = 6, "got " & VLA_Relation.RelCount(rel)

    Set result = VLA_Datalog.DatalogRun( _
        "(fact (edge a a)) (fact (edge a b))" & _
        " (rule (self_loop X) (edge X X))" & _
        " (query self_loop)")
    Set rel = VLA_Runtime.VlaDictGet(result.Item(2), result.Item(1))
    Report "datalog: (edge X X) keeps only the self-loop", VLA_Relation.RelCount(rel) = 1, "got " & VLA_Relation.RelCount(rel)

    Dim arr(1 To 2, 1 To 2) As Variant
    arr(1, 1) = "alice": arr(1, 2) = "bob"
    arr(2, 1) = "bob": arr(2, 2) = "carol"
    Dim baseRel As Collection
    Set baseRel = VLA_Relation.RelFromRange(arr)
    Dim bases As Object
    Set bases = VLA_Runtime.VlaDictNew()
    VLA_Runtime.VlaDictSet bases, "reports_to", baseRel
    Set result = VLA_Datalog.DatalogRun( _
        "(rule (indirect_report X Y) (reports_to X Y))" & _
        " (rule (indirect_report X Y) (reports_to X Z) (indirect_report Z Y))" & _
        " (query indirect_report)", bases)
    Set rel = VLA_Runtime.VlaDictGet(result.Item(2), result.Item(1))
    Report "datalog: table-sourced facts (no live Range) join and recurse correctly", VLA_Relation.RelCount(rel) = 3, "got " & VLA_Relation.RelCount(rel)

    Dim arr2 As Variant
    arr2 = VLA_Relation.RelToSpilledArray(rel, Array("A", "B"))
    Report "datalog: spilled array has a header row plus N data rows", _
           (ResultRowCount(arr2) = VLA_Relation.RelCount(rel) + 1) And ResultCellIs(arr2, 1, 1, "A") And ResultCellIs(arr2, 1, 2, "B"), _
           "shape mismatch"

    Dim raised As Boolean

    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Datalog.DatalogRun "(fact (p (q r))) (query p)"
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "datalog: a compound term (nested predicate) is refused, not silently accepted", raised, "no error raised"

    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Datalog.DatalogRun "(rule (foo X Y) (bar X)) (query foo)"
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "datalog: a head variable absent from the body is refused", raised, "no error raised"

    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Datalog.DatalogRun "(fact (p a b)) (fact (p a b c)) (query p)"
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "datalog: the same predicate used with two different arities is refused", raised, "no error raised"

    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Datalog.DatalogRun "(fact (p a))"
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "datalog: a program with no (query ...) is refused, not guessed", raised, "no error raised"

    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Datalog.DatalogRun "(headless extra) (fact (p a)) (query p)"
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "datalog: (headless extra) is refused, not silently accepted", raised, "no error raised"
End Sub

' DATALOG.1: stratified negation (`not`) - a basic anti-join, the
' variable-safety refusal (a negated atom's own variable never bound by
' an earlier positive atom), the stratifiability refusal (a predicate
' negatively depending on itself, directly and through a mutual cycle),
' and a real multi-stratum program (transitive reachability, then a
' negation over it) - the case stratification exists for: unreachable's
' own rule cannot run until reachable's ENTIRE recursive fixpoint (its
' own earlier stratum) is finished, not just "evaluated once."
Private Sub TestDatalogNegation()
    Dim result As Collection
    Dim rel As Collection

    Set result = VLA_Datalog.DatalogRun( _
        "(fact (person alice)) (fact (person bob)) (fact (person carol))" & _
        " (fact (banned bob))" & _
        " (rule (allowed X) (person X) (not (banned X)))" & _
        " (query allowed)")
    Set rel = VLA_Runtime.VlaDictGet(result.Item(2), result.Item(1))
    Report "datalog negation: (not (banned X)) excludes exactly the banned person", _
           VLA_Relation.RelCount(rel) = 2, "got " & VLA_Relation.RelCount(rel)

    Dim raised As Boolean

    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Datalog.DatalogRun "(fact (item a)) (rule (foo X) (not (bar X))) (query foo)"
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "datalog negation: a variable used ONLY under (not ...) is refused, not silently unbound", raised, "no error raised"

    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Datalog.DatalogRun "(fact (item a)) (rule (p X) (item X) (not (p X))) (query p)"
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "datalog negation: a predicate negating itself is refused as unstratifiable", raised, "no error raised"

    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Datalog.DatalogRun _
        "(fact (item a))" & _
        " (rule (p X) (item X) (not (q X)))" & _
        " (rule (q X) (item X) (not (p X)))" & _
        " (query p)"
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "datalog negation: a mutual negation cycle (p negates q, q negates p) is refused", raised, "no error raised"

    Set result = VLA_Datalog.DatalogRun( _
        "(fact (node a)) (fact (node b)) (fact (node c))" & _
        " (fact (edge a b)) (fact (edge b c))" & _
        " (rule (reachable X Y) (edge X Y))" & _
        " (rule (reachable X Z) (edge X Y) (reachable Y Z))" & _
        " (rule (unreachable X Y) (node X) (node Y) (not (reachable X Y)))" & _
        " (query unreachable)")
    Set rel = VLA_Runtime.VlaDictGet(result.Item(2), result.Item(1))
    Report "datalog negation: unreachable = all pairs minus reachable's own full recursive fixpoint (9 - 3 = 6)", _
           VLA_Relation.RelCount(rel) = 6, "got " & VLA_Relation.RelCount(rel)
End Sub

' DATALOG.2: grouped aggregation (`count`/`sum`) - a per-group count
' including a real ZERO group (carol has no sales at all), a per-group
' sum including a real zero-sum group, the parse-time refusals (sum with
' zero or two unbound "value" variables, a reused result variable, a
' non-variable result name, a malformed wrapper shape), and a real
' multi-stratum program (count over a recursively-derived predicate -
' reachcount cannot run until reachable's own full fixpoint, an earlier
' stratum, is finished). RelContainsTuple probes a specific (group,
' aggregate) pair directly rather than trusting row order, since a
' VlaDict's own enumeration order is never part of this engine's
' contract.
Private Sub TestDatalogAggregation()
    Dim result As Collection
    Dim rel As Collection
    Dim probe(1 To 2) As Variant

    Set result = VLA_Datalog.DatalogRun( _
        "(fact (person alice)) (fact (person bob)) (fact (person carol))" & _
        " (fact (sale alice widget)) (fact (sale alice gadget)) (fact (sale bob widget))" & _
        " (rule (salescount X N) (person X) (count N (sale X Y)))" & _
        " (query salescount)")
    Set rel = VLA_Runtime.VlaDictGet(result.Item(2), result.Item(1))
    Report "datalog count: one row per group, including a zero-sale person", VLA_Relation.RelCount(rel) = 3, "got " & VLA_Relation.RelCount(rel)
    probe(1) = "alice": probe(2) = 2
    Report "datalog count: alice's own 2 sales", VLA_Relation.RelContainsTuple(rel, probe), "(alice, 2) not found"
    probe(1) = "bob": probe(2) = 1
    Report "datalog count: bob's own 1 sale", VLA_Relation.RelContainsTuple(rel, probe), "(bob, 1) not found"
    probe(1) = "carol": probe(2) = 0
    Report "datalog count: carol's own real zero (no sales at all)", VLA_Relation.RelContainsTuple(rel, probe), "(carol, 0) not found"

    Set result = VLA_Datalog.DatalogRun( _
        "(fact (person alice)) (fact (person bob)) (fact (person carol))" & _
        " (fact (amount alice 10)) (fact (amount alice 15)) (fact (amount bob 7))" & _
        " (rule (total X S) (person X) (sum S (amount X V)))" & _
        " (query total)")
    Set rel = VLA_Runtime.VlaDictGet(result.Item(2), result.Item(1))
    Report "datalog sum: one row per group, including a zero-amount person", VLA_Relation.RelCount(rel) = 3, "got " & VLA_Relation.RelCount(rel)
    probe(1) = "alice": probe(2) = 25
    Report "datalog sum: alice's own 10+15", VLA_Relation.RelContainsTuple(rel, probe), "(alice, 25) not found"
    probe(1) = "bob": probe(2) = 7
    Report "datalog sum: bob's own single amount", VLA_Relation.RelContainsTuple(rel, probe), "(bob, 7) not found"
    probe(1) = "carol": probe(2) = 0
    Report "datalog sum: carol's own real zero (no amounts at all)", VLA_Relation.RelContainsTuple(rel, probe), "(carol, 0) not found"

    Dim raised As Boolean

    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Datalog.DatalogRun _
        "(fact (person alice)) (fact (amount alice ""10""))" & _
        " (rule (bad X S) (person X) (sum S (amount X ""10"")))" & _
        " (query bad)"
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "datalog sum: zero unbound value variables is refused", raised, "no error raised"

    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Datalog.DatalogRun _
        "(fact (person alice)) (fact (amount2 alice 1 2))" & _
        " (rule (bad X S) (person X) (sum S (amount2 X V W)))" & _
        " (query bad)"
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "datalog sum: two unbound value variables is refused", raised, "no error raised"

    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Datalog.DatalogRun _
        "(fact (person alice)) (fact (sale alice widget))" & _
        " (rule (bad X N) (person X) (count N (sale X Y)) (count N (sale X Z)))" & _
        " (query bad)"
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "datalog aggregate: a result variable reused from earlier in the body is refused", raised, "no error raised"

    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Datalog.DatalogRun _
        "(fact (person alice)) (fact (sale alice widget))" & _
        " (rule (bad X) (person X) (count n (sale X Y)))" & _
        " (query bad)"
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "datalog aggregate: a lowercase (non-variable) result name is refused", raised, "no error raised"

    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Datalog.DatalogRun _
        "(fact (person alice))" & _
        " (rule (bad X N) (person X) (count N))" & _
        " (query bad)"
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "datalog aggregate: a malformed (count ...) wrapper shape is refused", raised, "no error raised"

    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Datalog.DatalogRun _
        "(fact (item a))" & _
        " (rule (p X N) (item X) (count N (p X Y)))" & _
        " (query p)"
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "datalog aggregate: a predicate counting itself is refused as unstratifiable", raised, "no error raised"

    Set result = VLA_Datalog.DatalogRun( _
        "(fact (node a)) (fact (node b)) (fact (node c))" & _
        " (fact (edge a b)) (fact (edge b c))" & _
        " (rule (reachable X Y) (edge X Y))" & _
        " (rule (reachable X Z) (edge X Y) (reachable Y Z))" & _
        " (rule (reachcount X N) (node X) (count N (reachable X Y)))" & _
        " (query reachcount)")
    Set rel = VLA_Runtime.VlaDictGet(result.Item(2), result.Item(1))
    Report "datalog aggregate: reachcount needs reachable's own full recursive fixpoint first (3 rows)", _
           VLA_Relation.RelCount(rel) = 3, "got " & VLA_Relation.RelCount(rel)
    probe(1) = "a": probe(2) = 2
    Report "datalog aggregate: a reaches 2 nodes (b, c)", VLA_Relation.RelContainsTuple(rel, probe), "(a, 2) not found"
    probe(1) = "b": probe(2) = 1
    Report "datalog aggregate: b reaches 1 node (c)", VLA_Relation.RelContainsTuple(rel, probe), "(b, 1) not found"
    probe(1) = "c": probe(2) = 0
    Report "datalog aggregate: c reaches 0 nodes", VLA_Relation.RelContainsTuple(rel, probe), "(c, 0) not found"
End Sub

' DATALOG.4: comparison filters ((> X 50000) and the rest of the
' </<=/>/>=/=/<> set) and the (let Z (+ X Y)) arithmetic binding form -
' this module's own header note above has the full scenario list.
' Entirely pure, unlike TestDatalogHostTable below: neither shape needs
' a live workbook, so every case here goes straight through DatalogRun.
Private Sub TestDatalogBuiltins()
    Dim result As Collection
    Dim rel As Collection

    ' A comparison over a TABLE-sourced column - built here via
    ' RelFromRange fed a hand-built 2D array carrying REAL Double values
    ' (VBA numeric literals, never strings), the identical seam
    ' TestDatalog's own "table-sourced facts (no live Range)" case
    ' already established, so this needs no live workbook either.
    ' Proves VLA_Relation.ValueIsNumericType's own STRICT policy path
    ' (a real Excel-typed value) fires correctly through DATALOG's own
    ' lenient BuiltinOperandIsNumeric.
    Dim arrNum(1 To 3, 1 To 2) As Variant
    arrNum(1, 1) = "alice": arrNum(1, 2) = 90000
    arrNum(2, 1) = "bob": arrNum(2, 2) = 60000
    arrNum(3, 1) = "carol": arrNum(3, 2) = 95000
    Dim baseRelNum As Collection
    Set baseRelNum = VLA_Relation.RelFromRange(arrNum)
    Dim basesNum As Object
    Set basesNum = VLA_Runtime.VlaDictNew()
    VLA_Runtime.VlaDictSet basesNum, "staff", baseRelNum
    Set result = VLA_Datalog.DatalogRun( _
        "(rule (highearner X) (staff X S) (> S 80000))" & _
        " (query highearner)", basesNum)
    Set rel = VLA_Runtime.VlaDictGet(result.Item(2), result.Item(1))
    Report "datalog builtins: (> S 80000) over a table-sourced (real Double) column keeps alice+carol (2 rows)", _
           VLA_Relation.RelCount(rel) = 2, "got " & VLA_Relation.RelCount(rel)

    ' The SAME comparison shape, but every value now a fact-block
    ' constant (always a plain VBA String, never a real Excel type) -
    ' 9 and 10 against a threshold of 8, deliberately chosen because a
    ' naive lexical-text fallback would get "10 > 8" WRONG (StrComp
    ' compares "1" against "8" first and finds "10" < "8"), so this
    ' specifically proves BuiltinOperandIsNumeric's own numeric-LOOKING-
    ' STRING policy (VLA_Relation.IsInvariantNumericString) actually
    ' fires rather than silently falling back to text.
    Set result = VLA_Datalog.DatalogRun( _
        "(fact (score alice 9)) (fact (score bob 10))" & _
        " (rule (high_scorer X) (score X N) (> N 8))" & _
        " (query high_scorer)")
    Set rel = VLA_Runtime.VlaDictGet(result.Item(2), result.Item(1))
    Report "datalog builtins: (> N 8) compares numeric-looking fact STRINGS numerically, not lexically (9 and 10 both qualify - 2 rows)", _
           VLA_Relation.RelCount(rel) = 2, "got " & VLA_Relation.RelCount(rel)

    ' Text fallback: neither operand looks numeric, so this compares as
    ' case-sensitive text - <> excludes exactly the one equal value.
    Set result = VLA_Datalog.DatalogRun( _
        "(fact (tag apple)) (fact (tag banana)) (fact (tag cherry))" & _
        " (rule (not_banana X) (tag X) (<> X banana))" & _
        " (query not_banana)")
    Set rel = VLA_Runtime.VlaDictGet(result.Item(2), result.Item(1))
    Report "datalog builtins: (<> X banana) falls back to text compare, excludes exactly banana (2 rows)", _
           VLA_Relation.RelCount(rel) = 2, "got " & VLA_Relation.RelCount(rel)

    ' A comparison and (not ...) composing in the SAME rule body - alice
    ' clears the age filter and isn't banned; bob fails the age filter
    ' regardless of his own separate ban.
    Set result = VLA_Datalog.DatalogRun( _
        "(fact (person alice 25)) (fact (person bob 15)) (fact (banned bob))" & _
        " (rule (adult_allowed X) (person X Age) (> Age 18) (not (banned X)))" & _
        " (query adult_allowed)")
    Set rel = VLA_Runtime.VlaDictGet(result.Item(2), result.Item(1))
    Report "datalog builtins: a comparison and (not ...) compose in one rule body (alice only)", _
           VLA_Relation.RelCount(rel) = 1, "got " & VLA_Relation.RelCount(rel)

    ' The four arithmetic operators, one rule apiece.
    Dim probe1(1 To 1) As Variant

    Set result = VLA_Datalog.DatalogRun( _
        "(fact (pair 10 3)) (rule (added S) (pair A B) (let S (+ A B))) (query added)")
    Set rel = VLA_Runtime.VlaDictGet(result.Item(2), result.Item(1))
    probe1(1) = 13
    Report "datalog let: (+ 10 3) = 13", VLA_Relation.RelContainsTuple(rel, probe1), "13 not found"

    Set result = VLA_Datalog.DatalogRun( _
        "(fact (pair 10 3)) (rule (subbed S) (pair A B) (let S (- A B))) (query subbed)")
    Set rel = VLA_Runtime.VlaDictGet(result.Item(2), result.Item(1))
    probe1(1) = 7
    Report "datalog let: (- 10 3) = 7", VLA_Relation.RelContainsTuple(rel, probe1), "7 not found"

    Set result = VLA_Datalog.DatalogRun( _
        "(fact (pair 10 3)) (rule (multiplied S) (pair A B) (let S (* A B))) (query multiplied)")
    Set rel = VLA_Runtime.VlaDictGet(result.Item(2), result.Item(1))
    probe1(1) = 30
    Report "datalog let: (* 10 3) = 30", VLA_Relation.RelContainsTuple(rel, probe1), "30 not found"

    Set result = VLA_Datalog.DatalogRun( _
        "(fact (pair 10 4)) (rule (divided S) (pair A B) (let S (/ A B))) (query divided)")
    Set rel = VLA_Runtime.VlaDictGet(result.Item(2), result.Item(1))
    probe1(1) = 2.5
    Report "datalog let: (/ 10 4) = 2.5", VLA_Relation.RelContainsTuple(rel, probe1), "2.5 not found"

    ' ---- PROLOG.17: the widened operator set reaches DATALOG through the
    ' SAME VLA_Relation.ArithOpArity table and the same two compute
    ' functions PROLOG uses. These assert the VALUE, because a wrong
    ' operator here produces a plausible NUMBER in a cell rather than an
    ' error - the failure this whole item is shaped around.
    Set result = VLA_Datalog.DatalogRun( _
        "(fact (pair -7 3)) (rule (modded S) (pair A B) (let S (mod A B))) (query modded)")
    Set rel = VLA_Runtime.VlaDictGet(result.Item(2), result.Item(1))
    probe1(1) = 2
    Report "datalog let: (mod -7 3) = 2 - FLOORED here exactly as in PROLOG, not VBA's own truncating Mod (which gives -1)", _
           VLA_Relation.RelContainsTuple(rel, probe1), "2 not found"

    Set result = VLA_Datalog.DatalogRun( _
        "(fact (pair -7 3)) (rule (remmed S) (pair A B) (let S (rem A B))) (query remmed)")
    Set rel = VLA_Runtime.VlaDictGet(result.Item(2), result.Item(1))
    probe1(1) = -1
    Report "datalog let: (rem -7 3) = -1 - the twin, so DATALOG cannot have collapsed mod and rem into one", _
           VLA_Relation.RelContainsTuple(rel, probe1), "-1 not found"

    Set result = VLA_Datalog.DatalogRun( _
        "(fact (pair 7 2)) (rule (idiv S) (pair A B) (let S (// A B))) (query idiv)")
    Set rel = VLA_Runtime.VlaDictGet(result.Item(2), result.Item(1))
    probe1(1) = 3
    Report "datalog let: (// 7 2) = 3 - integer division, where (/ 7 2) is 3.5", _
           VLA_Relation.RelContainsTuple(rel, probe1), "3 not found"

    Set result = VLA_Datalog.DatalogRun( _
        "(fact (pair 3 7)) (rule (mx S) (pair A B) (let S (max A B))) (query mx)")
    Set rel = VLA_Runtime.VlaDictGet(result.Item(2), result.Item(1))
    probe1(1) = 7
    Report "datalog let: (max 3 7) = 7", VLA_Relation.RelContainsTuple(rel, probe1), "7 not found"

    ' ---- THE UNARY SHAPE, which DATALOG could not express at all before
    ' this item: its `let` arm shared a flat "exactly two operands" check
    ' with comparisons, so (let Z (abs X)) was unwritable rather than
    ' merely unimplemented.
    Set result = VLA_Datalog.DatalogRun( _
        "(fact (val -4)) (rule (absd S) (val A) (let S (abs A))) (query absd)")
    Set rel = VLA_Runtime.VlaDictGet(result.Item(2), result.Item(1))
    probe1(1) = 4
    Report "datalog let: (abs -4) = 4 - a UNARY let, which this engine's arity check made impossible before PROLOG.17", _
           VLA_Relation.RelContainsTuple(rel, probe1), "4 not found"

    Set result = VLA_Datalog.DatalogRun( _
        "(fact (val 2.5)) (rule (rnd S) (val A) (let S (round A))) (query rnd)")
    Set rel = VLA_Runtime.VlaDictGet(result.Item(2), result.Item(1))
    probe1(1) = 3
    Report "datalog let: (round 2.5) = 3 - half AWAY FROM ZERO, where VBA's own Round gives 2", _
           VLA_Relation.RelContainsTuple(rel, probe1), "3 not found"

    ' A RECURSIVE rule accumulating (let ...) across semi-naive rounds -
    ' path length over a 3-edge chain (a-b-c-d, weight 1 each). This is
    ' the one case that would have caught a wrong round-loop/ComputeStrata
    ' fix: BI_LET is no longer a valid semi-naive delta position (unlike
    ' BI_POS), and its own atom (predicate "+") must NOT be mistaken for
    ' a real relation by ComputeStrata's own dependency graph.
    Set result = VLA_Datalog.DatalogRun( _
        "(fact (edge a b 1)) (fact (edge b c 1)) (fact (edge c d 1))" & _
        " (rule (path X Y D) (edge X Y D))" & _
        " (rule (path X Z D) (edge X Y D1) (path Y Z D2) (let D (+ D1 D2)))" & _
        " (query path)")
    Set rel = VLA_Runtime.VlaDictGet(result.Item(2), result.Item(1))
    Report "datalog let: recursive path-length accumulation derives all 6 reachable pairs", _
           VLA_Relation.RelCount(rel) = 6, "got " & VLA_Relation.RelCount(rel)
    Dim probePath(1 To 3) As Variant
    probePath(1) = "a": probePath(2) = "d": probePath(3) = 3
    Report "datalog let: a-to-d path length accumulates to 3 across two recursive hops", _
           VLA_Relation.RelContainsTuple(rel, probePath), "(a, d, 3) not found"

    Dim raised As Boolean

    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Datalog.DatalogRun "(fact (item a)) (rule (foo X) (item X) (> Y 5)) (query foo)"
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "datalog builtins: a comparison referencing an unbound variable is refused", raised, "no error raised"

    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Datalog.DatalogRun "(fact (item a)) (rule (foo X Z) (item X) (let Z (+ Y 1))) (query foo)"
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "datalog builtins: a let operand referencing an unbound variable is refused", raised, "no error raised"

    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Datalog.DatalogRun "(fact (pair a 1)) (rule (foo X N) (pair X N) (let N (+ N 1))) (query foo)"
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "datalog builtins: a let result variable reused from earlier in the body is refused", raised, "no error raised"

    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Datalog.DatalogRun "(fact (item a)) (rule (foo X) (item X) (> X)) (query foo)"
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "datalog builtins: a comparison with only one operand is refused", raised, "no error raised"

    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Datalog.DatalogRun "(fact (item a)) (rule (foo X Z) (item X) (let Z (+ X))) (query foo)"
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "datalog builtins: a let arithmetic expression with only one operand is refused", raised, "no error raised"

    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Datalog.DatalogRun "(fact (pair a 1 2)) (rule (foo X Z) (pair X A B) (let Z (% A B))) (query foo)"
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "datalog builtins: an unrecognized arithmetic operator is refused by name", raised, "no error raised"

    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Datalog.DatalogRun "(fact (pair a 10 0)) (rule (foo X Z) (pair X A B) (let Z (/ A B))) (query foo)"
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "datalog builtins: division by zero is refused, not a raw runtime crash", raised, "no error raised"

    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Datalog.DatalogRun "(fact (pair a hello)) (rule (foo X Z) (pair X A) (let Z (+ A 1))) (query foo)"
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "datalog builtins: a non-numeric arithmetic operand is refused", raised, "no error raised"

    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Datalog.DatalogRun "(fact (item a)) (rule (foo X) (item X) (let)) (query foo)"
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "datalog builtins: a malformed (let ...) wrapper shape is refused", raised, "no error raised"

    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Datalog.DatalogRun "(fact (pair a 1 2)) (rule (foo X Z) (pair X A B) (let z (+ A B))) (query foo)"
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "datalog builtins: a lowercase (non-variable) let result name is refused", raised, "no error raised"
End Sub

' DATALOG.5: named-column atoms - a hand-built (Name, Salary, Dept)
' "staffing" table, fed through RelFromRange's own new hasHeader:=True
' flag (this item's own pure-test seam - no live workbook needed at
' all, unlike DATALOG.3's own column-scoping) plus a hand-built
' headerMap in the exact (folded, original) pair shape VLA_Relation.
' RangeColumnNames returns (SqlColPair, already built for SQL's own
' pure suite below, reused as-is). Proves the motivating example this
' item was found by name (BETA_ROADMAP2.md's own words - "Show every
' Staffing name whose salary is over 80000"), the one correctness
' property that would be silently wrong if the anonymous-column naming
' were keyed on column position ALONE rather than atom occurrence too
' (two keyed atoms, same table, each omitting salary/dept, must derive
' the full cross product - not be silently aliased together), a partial
' keying + count combination, full-arity keying, and every parse-time
' refusal this item's own roadmap names by id, including the one
' highest-risk case named explicitly: a keyed pair whose own VALUE
' position is itself a nested compound term must still be refused, not
' silently accepted through the new opening.
Private Sub TestDatalogKeyedAtoms()
    Dim result As Collection
    Dim rel As Collection

    Dim arr(1 To 4, 1 To 3) As Variant
    arr(1, 1) = "Name": arr(1, 2) = "Salary": arr(1, 3) = "Dept"
    arr(2, 1) = "alice": arr(2, 2) = 90000: arr(2, 3) = "eng"
    arr(3, 1) = "bob": arr(3, 2) = 70000: arr(3, 3) = "sales"
    arr(4, 1) = "carol": arr(4, 2) = 95000: arr(4, 3) = "eng"

    Dim baseRel As Collection
    Set baseRel = VLA_Relation.RelFromRange(arr, True)
    Report "datalog keyed: RelFromRange's own hasHeader:=True skips the header row (3 data rows, not 4)", _
           VLA_Relation.RelCount(baseRel) = 3, "got " & VLA_Relation.RelCount(baseRel)

    Dim bases As Object
    Set bases = VLA_Runtime.VlaDictNew()
    VLA_Runtime.VlaDictSet bases, "staffing", baseRel

    Dim staffingCols As New Collection
    staffingCols.Add SqlColPair("name", "Name")
    staffingCols.Add SqlColPair("salary", "Salary")
    staffingCols.Add SqlColPair("dept", "Dept")
    Dim headerMap As Object
    Set headerMap = VLA_Runtime.VlaDictNew()
    VLA_Runtime.VlaDictSet headerMap, "staffing", staffingCols

    Set result = VLA_Datalog.DatalogRun( _
        "(rule (rich X) (staffing (name X) (salary S)) (> S 80000))" & _
        " (query rich)", bases, headerMap)
    Set rel = VLA_Runtime.VlaDictGet(result.Item(2), result.Item(1))
    Report "datalog keyed: the item's own motivating example (salary keyed, > 80000) matches exactly alice and carol", _
           VLA_Relation.RelCount(rel) = 2, "got " & VLA_Relation.RelCount(rel)

    Set result = VLA_Datalog.DatalogRun( _
        "(rule (pair X Y) (staffing (name X)) (staffing (name Y)))" & _
        " (query pair)", bases, headerMap)
    Set rel = VLA_Runtime.VlaDictGet(result.Item(2), result.Item(1))
    Report "datalog keyed: two keyed atoms each omitting salary/dept derive the full 3x3 cross product (9) - never silently joined on a shared anonymous column name", _
           VLA_Relation.RelCount(rel) = 9, "got " & VLA_Relation.RelCount(rel)

    Set result = VLA_Datalog.DatalogRun( _
        "(fact (deptname eng)) (fact (deptname sales))" & _
        " (rule (deptcount D N) (deptname D) (count N (staffing (dept D))))" & _
        " (query deptcount)", bases, headerMap)
    Set rel = VLA_Runtime.VlaDictGet(result.Item(2), result.Item(1))
    Report "datalog keyed: count over a keyed atom naming only the group-by column derives one row per dept", _
           VLA_Relation.RelCount(rel) = 2, "got " & VLA_Relation.RelCount(rel)
    Dim probeD(1 To 2) As Variant
    probeD(1) = "eng": probeD(2) = 2
    Report "datalog keyed count: eng has 2 staffers (alice, carol)", VLA_Relation.RelContainsTuple(rel, probeD), "(eng, 2) not found"
    probeD(1) = "sales": probeD(2) = 1
    Report "datalog keyed count: sales has 1 staffer (bob)", VLA_Relation.RelContainsTuple(rel, probeD), "(sales, 1) not found"

    Set result = VLA_Datalog.DatalogRun( _
        "(rule (allkeyed N S D) (staffing (name N) (salary S) (dept D)))" & _
        " (query allkeyed)", bases, headerMap)
    Set rel = VLA_Runtime.VlaDictGet(result.Item(2), result.Item(1))
    Report "datalog keyed: keying every column explicitly still reports the table's own arity/rows (3)", _
           VLA_Relation.RelCount(rel) = 3, "got " & VLA_Relation.RelCount(rel)

    Dim raised As Boolean

    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Datalog.DatalogRun "(rule (bad X) (staffing X (salary S))) (query bad)", bases, headerMap
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "datalog keyed: mixing keyed and positional arguments in one atom is refused", raised, "no error raised"

    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Datalog.DatalogRun "(rule (bad X S) (staffing (name X) (bogus S))) (query bad)", bases, headerMap
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "datalog keyed: an unknown column name is refused", raised, "no error raised"

    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Datalog.DatalogRun "(rule (bad X Y) (staffing (name X) (name Y))) (query bad)", bases, headerMap
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "datalog keyed: keying the same column twice in one atom is refused", raised, "no error raised"

    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Datalog.DatalogRun "(fact (widget a b)) (rule (bad X) (widget (foo X))) (query bad)", bases, headerMap
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "datalog keyed: a keyed atom against a predicate with no header of its own (a fact block) is refused", raised, "no error raised"

    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Datalog.DatalogRun "(rule (bad X Y) (staffing (name X) (salary (foo Y)))) (query bad)", bases, headerMap
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "datalog keyed: a keyed pair whose own value is itself a nested compound term is still refused, not silently accepted", raised, "no error raised"

    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Datalog.DatalogRun "(rule (staffing (name X)) (staffing (name X))) (query staffing)", bases, headerMap
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "datalog keyed: keyed syntax in a rule HEAD is refused (heads stay positional by definition)", raised, "no error raised"
End Sub

' DATALOG.8: AN UNDEFINED PREDICATE REFUSES - statically, anywhere in the
' program, before any rule runs. Pure: no live workbook (a Table argument
' with no data rows is also pinned live, in TestDatalogHostTable).
'
' Every body case below used to answer SILENTLY, and two were confidently
' wrong rather than merely empty: `not` over a misspelled relation kept
' every row, so a banned person was listed as allowed, and count and sum
' over one answered 0. Each is asserted on refusal text NAMING the
' predicate, so none can pass by failing.
'
' The decision is pinned in its directions: a misspelling in a rule the
' query never uses refuses too (the whole program, the owner's call), it
' refuses before any rule runs, an operator's head is never read as a
' relation, and a keyed atom is refused as undefined only when nothing
' defines its name. The roadmap entry has the options and why these.
Private Sub TestDatalogUnknownPredicate()
    Dim result As Variant
    Dim r As String
    Dim d As String
    Dim q As String
    q = Chr$(34)

    ' ---- THE PLAIN CASE. It used to spill a header with nothing under it.
    result = VLA_Datalog.DATALOG("(fact (parent tom bob)) (rule (kid X) (parnet tom X)) (query kid)")
    r = ResultDescribe(result)
    Report "datalog.8: a misspelled relation in a rule body refuses by name, where it used to spill an empty column", _
           ResultTextStartsWith(result, "#DATALOG!") And InStr(1, r, "'parnet' is used in a rule", vbTextCompare) > 0, "got: " & r

    ' ---- THE CONFIDENTLY WRONG ALLOWED. tom IS banned; the misspelled
    ' negation read an empty relation and kept him.
    result = VLA_Datalog.DATALOG("(fact (person tom)) (fact (banned tom)) (rule (ok X) (person X) (not (bannd X))) (query ok)")
    r = ResultDescribe(result)
    Report "datalog.8: (not ...) over a misspelled relation refuses - it used to list tom, who is banned", _
           ResultTextStartsWith(result, "#DATALOG!") And InStr(1, r, "'bannd' is used in a rule", vbTextCompare) > 0, "got: " & r

    ' ---- THE ZEROS THAT COUNTED NOTHING. tom has a sale and an amount.
    result = VLA_Datalog.DATALOG("(fact (person tom)) (fact (sale tom widget)) (rule (sales X N) (person X) (count N (sael X Y))) (query sales)")
    r = ResultDescribe(result)
    Report "datalog.8: (count ...) over a misspelled relation refuses - it used to answer 0", _
           ResultTextStartsWith(result, "#DATALOG!") And InStr(1, r, "'sael' is used in a rule", vbTextCompare) > 0, "got: " & r
    result = VLA_Datalog.DATALOG("(fact (person tom)) (fact (amount tom 10)) (rule (total X S) (person X) (sum S (amont X V))) (query total)")
    r = ResultDescribe(result)
    Report "datalog.8: (sum ...) over a misspelled relation refuses - it used to answer 0", _
           ResultTextStartsWith(result, "#DATALOG!") And InStr(1, r, "'amont' is used in a rule", vbTextCompare) > 0, "got: " & r

    ' ---- THE WHOLE PROGRAM. Nothing asks about unused, and its
    ' misspelling refuses anyway - strict first, because relaxing this later
    ' only turns a refusal into an answer.
    result = VLA_Datalog.DATALOG("(fact (p a)) (rule (unused X) (sibling X)) (query p)")
    r = ResultDescribe(result)
    Report "datalog.8: a misspelling in a rule the query never uses refuses too - the whole program is checked", _
           ResultTextStartsWith(result, "#DATALOG!") And InStr(1, r, "'sibling' is used in a rule", vbTextCompare) > 0, "got: " & r

    ' ---- THE FIRST ONE WRITTEN is the one named: zz1 sits in a rule the
    ' query does not use, ahead of zz2 in the rule it does.
    result = VLA_Datalog.DATALOG("(fact (p a)) (rule (a1 X) (p X) (zz1 X)) (rule (a2 X) (p X) (zz2 X)) (query a2)")
    r = ResultDescribe(result)
    Report "datalog.8: of two misspellings, the first in the order written is named", _
           ResultTextStartsWith(result, "#DATALOG!") And InStr(1, r, "'zz1' is used in a rule", vbTextCompare) > 0 _
           And InStr(1, r, "zz2", vbTextCompare) = 0, "got: " & r

    ' ---- DEFINING THE NAME is what stops it, and the negation then answers
    ' correctly: tom is banned, so nobody is ok - a header and nothing under.
    result = VLA_Datalog.DATALOG("(fact (person tom)) (fact (banned tom)) (rule (ok X) (person X) (not (banned X))) (query ok)")
    Report "datalog.8: ...spelled right, the same rule answers - tom is banned, so the ok column is empty", _
           ResultRowCount(result) = 1 And ResultCellIs(result, 1, 1, "X"), "got: " & ResultDescribe(result)

    ' ---- AN OPERATOR IS NOT A RELATION. The heads of a let (+) and of a
    ' comparison (>) are symbols with no relation behind them, and are
    ' never checked as one.
    result = VLA_Datalog.DATALOG("(fact (pair 10 3)) (rule (big S) (pair A B) (let S (+ A B)) (> S 5)) (query big)")
    Report "datalog.8: a let and a comparison name no relation - the rule answers 13", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, "13"), "got: " & ResultDescribe(result)

    ' ---- THE QUERY POSITION keeps its own words.
    result = VLA_Datalog.DATALOG("(fact (p a)) (query pp)")
    r = ResultDescribe(result)
    Report "datalog.8: a misspelled query name still refuses with the query's own text", _
           ResultTextStartsWith(result, "#DATALOG!") And InStr(1, r, "(query pp) names a predicate with no facts", vbTextCompare) > 0, "got: " & r

    ' ---- OPTIMIZE.1's follow-up: BOTH refusals now say what names there
    ' WERE. Asked for after two steps of OPTIMIZE.1's own live pass were
    ' lost to a table whose name had gone to Excel's Name Box instead of
    ' its Table Name box: the refusal was correct and told the reader
    ' nothing they could act on. Same two programs as above, so the parity
    ' table's coverage is untouched.
    Report "datalog.8: the query refusal now lists the names that DO exist", _
           InStr(1, r, "The names this program does define, table arguments first, are: p.", vbTextCompare) > 0, _
           "got: " & r
    result = VLA_Datalog.DATALOG("(fact (parent tom bob)) (rule (kid X) (parnet tom X)) (query kid)")
    r = ResultDescribe(result)
    Report "datalog.8: and so does the rule-body refusal, which is the one a wrong table name hits", _
           InStr(1, r, "The names this program does define, table arguments first, are: parent, kid.", vbTextCompare) > 0, _
           "got: " & r
    ' A program with nothing defined at all says so, rather than trailing
    ' off after a colon.
    result = VLA_Datalog.DATALOG("(query pp)")
    r = ResultDescribe(result)
    Report "datalog.8: a program defining nothing at all says that, rather than an empty list", _
           InStr(1, r, "This program defines no names at all.", vbTextCompare) > 0, "got: " & r

    ' ---- BEFORE ANY RULE RUNS. bad divides by zero the moment it is
    ' evaluated; each misspelling is named instead, because nothing ran.
    result = VLA_Datalog.DATALOG("(fact (pair a 10 0)) (rule (bad X Z) (pair X A B) (let Z (/ A B))) (rule (ok X) (pair X A B) (typo X)) (query ok)")
    r = ResultDescribe(result)
    Report "datalog.8: a misspelled relation is named before a rule that divides by zero ever runs", _
           ResultTextStartsWith(result, "#DATALOG!") And InStr(1, r, "'typo' is used in a rule", vbTextCompare) > 0 _
           And InStr(1, r, "divide by zero", vbTextCompare) = 0, "got: " & r
    result = VLA_Datalog.DATALOG("(fact (pair a 10 0)) (rule (bad X Z) (pair X A B) (let Z (/ A B))) (query pp)")
    r = ResultDescribe(result)
    Report "datalog.8: ...and so is a misspelled query name - it used to be checked only after every rule had run", _
           ResultTextStartsWith(result, "#DATALOG!") And InStr(1, r, "(query pp) names a predicate", vbTextCompare) > 0 _
           And InStr(1, r, "divide by zero", vbTextCompare) = 0, "got: " & r

    ' ---- A TABLE WITH NO ROWS IS DEFINED. A header and nothing under it is
    ' the pure seam for a Table whose rows were all deleted; DATALOG()
    ' registers a relation per table argument, rows or none.
    Dim hdr(1 To 1, 1 To 2) As Variant
    hdr(1, 1) = "Name": hdr(1, 2) = "Shift"
    Dim basesEmpty As Object
    Set basesEmpty = VLA_Runtime.VlaDictNew()
    VLA_Runtime.VlaDictSet basesEmpty, "leave", VLA_Relation.RelFromRange(hdr, True)
    Dim leaveCols As New Collection
    leaveCols.Add SqlColPair("name", "Name")
    leaveCols.Add SqlColPair("shift", "Shift")
    Dim headerMapEmpty As Object
    Set headerMapEmpty = VLA_Runtime.VlaDictNew()
    VLA_Runtime.VlaDictSet headerMapEmpty, "leave", leaveCols
    Dim res As Collection
    Dim rel As Collection
    Set rel = VLA_Relation.RelNew(1)
    d = ""
    On Error Resume Next
    Err.Clear
    Set res = VLA_Datalog.DatalogRun("(fact (person tom)) (rule (ok X) (person X) (not (leave X X))) (query ok)", basesEmpty, headerMapEmpty)
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    If Len(d) = 0 Then Set rel = VLA_Runtime.VlaDictGet(res.Item(2), res.Item(1))
    Report "datalog.8: a table argument with no rows is defined - (not ...) over it keeps tom", _
           Len(d) = 0 And VLA_Relation.RelCount(rel) = 1, "got: '" & d & "', " & VLA_Relation.RelCount(rel) & " row(s)"
    ' ...and leaving the table out is exactly the misspelling.
    result = VLA_Datalog.DATALOG("(fact (person tom)) (rule (ok X) (person X) (not (leave X X))) (query ok)")
    r = ResultDescribe(result)
    Report "datalog.8: ...while the same rule with no leave table passed names leave", _
           ResultTextStartsWith(result, "#DATALOG!") And InStr(1, r, "'leave' is used in a rule", vbTextCompare) > 0, "got: " & r

    ' ---- A KEYED ATOM over a misspelled Table is a misspelling, not a
    ' relation missing its headers, which is what the refusal used to say.
    Dim staff(1 To 2, 1 To 2) As Variant
    staff(1, 1) = "Name": staff(1, 2) = "Level"
    staff(2, 1) = "Ann": staff(2, 2) = 3
    Dim basesStaff As Object
    Set basesStaff = VLA_Runtime.VlaDictNew()
    VLA_Runtime.VlaDictSet basesStaff, "staff", VLA_Relation.RelFromRange(staff, True)
    Dim staffCols As New Collection
    staffCols.Add SqlColPair("name", "Name")
    staffCols.Add SqlColPair("level", "Level")
    Dim headerMapStaff As Object
    Set headerMapStaff = VLA_Runtime.VlaDictNew()
    VLA_Runtime.VlaDictSet headerMapStaff, "staff", staffCols
    d = ""
    On Error Resume Next
    Err.Clear
    VLA_Datalog.DatalogRun "(rule (senior X) (staf (name X) (level L)) (> L 2)) (query senior)", basesStaff, headerMapStaff
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    Report "datalog.8: a keyed atom over a misspelled Table refuses as undefined, not as a relation without column headers", _
           InStr(1, d, "'staf' is used in a rule", vbTextCompare) > 0 And InStr(1, d, "column headers", vbTextCompare) = 0, "got: " & d

    ' ...while one over a relation that IS defined, by a fact block, keeps
    ' the header text - defined earlier in the text or later.
    d = ""
    On Error Resume Next
    Err.Clear
    VLA_Datalog.DatalogRun "(fact (widget a b)) (rule (bad X) (widget (foo X))) (query bad)"
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    Report "datalog.8: a keyed atom over a fact block still refuses as having no column headers", _
           InStr(1, d, "no column headers", vbTextCompare) > 0 And InStr(1, d, "is used in a rule", vbTextCompare) = 0, "got: " & d
    d = ""
    On Error Resume Next
    Err.Clear
    VLA_Datalog.DatalogRun "(rule (bad X) (widget (foo X))) (fact (widget a b)) (query bad)"
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    Report "datalog.8: ...and so does one whose fact block is written after the rule that keys it", _
           InStr(1, d, "no column headers", vbTextCompare) > 0 And InStr(1, d, "is used in a rule", vbTextCompare) = 0, "got: " & d

    ' ---- G-PROLOG's WHO QUESTION, the shape that found this item: the
    ' relation a writer names sits in a generated narrowing rule's BODY, so
    ' a misspelled can-drive used to spill an empty Who column.
    Dim rota(1 To 3, 1 To 2) As Variant
    rota(1, 1) = "Name": rota(1, 2) = "Shift"
    rota(2, 1) = "Bob": rota(2, 2) = "Night"
    rota(3, 1) = "Di": rota(3, 2) = "Day"
    Dim rotaCols As New Collection
    rotaCols.Add SqlColPair("name", "Name")
    rotaCols.Add SqlColPair("shift", "Shift")
    Dim headerMapRota As Object
    Set headerMapRota = VLA_Runtime.VlaDictNew()
    VLA_Runtime.VlaDictSet headerMapRota, "rota", rotaCols
    Dim rulesText As String
    rulesText = "(rule (can-cover Person Shift) (rota (name Person) (shift Shift)))"
    Dim basesRota As Object
    Set basesRota = VLA_Runtime.VlaDictNew()
    VLA_Runtime.VlaDictSet basesRota, "rota", VLA_Relation.RelFromRange(rota, True)
    d = ""
    On Error Resume Next
    Err.Clear
    VLA_Datalog.DatalogRun rulesText & " (rule (vla-ask-can-drive Who) (can-drive Who " & q & "Night" & q & ")) (query vla-ask-can-drive)", basesRota, headerMapRota
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    Report "datalog.8: G-PROLOG's who-question with a misspelled relation refuses by name, where it spilled an empty Who column", _
           InStr(1, d, "'can-drive' is used in a rule", vbTextCompare) > 0, "got: " & d
    ' ...and its twin, spelled right, answers Bob.
    Dim basesRota2 As Object
    Set basesRota2 = VLA_Runtime.VlaDictNew()
    VLA_Runtime.VlaDictSet basesRota2, "rota", VLA_Relation.RelFromRange(rota, True)
    Dim probeBob(1 To 1) As Variant
    probeBob(1) = "Bob"
    Set rel = VLA_Relation.RelNew(1)
    d = ""
    On Error Resume Next
    Err.Clear
    Set res = VLA_Datalog.DatalogRun(rulesText & " (rule (vla-ask-can-cover Who) (can-cover Who " & q & "Night" & q & ")) (query vla-ask-can-cover)", basesRota2, headerMapRota)
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    If Len(d) = 0 Then Set rel = VLA_Runtime.VlaDictGet(res.Item(2), res.Item(1))
    Report "datalog.8: ...and spelled right, the same question answers Bob alone", _
           Len(d) = 0 And VLA_Relation.RelCount(rel) = 1 And VLA_Relation.RelContainsTuple(rel, probeBob), _
           "got: '" & d & "', " & VLA_Relation.RelCount(rel) & " row(s)"
End Sub

' DATALOG.9: A QUERY WRITTEN AS ONE FACT ANSWERS TRUE OR FALSE. Pure.
'
' The shape G-PROLOG's whether questions write, (query (can-cover "Bob"
' "Night")), used to be refused here as not a plain name, so a question with
' no unknown went to PROLOG - which, over a closure whose data loops, never
' finishes. These pins hold the answer (a real Boolean, decided by the same
' match a rule body makes), the reason (a cycle answers both ways and
' stops), a query by name unchanged, and each shape still refused, by name,
' pointing at the spelling that works. Every combined check reads a value
' that exists even when the call refused (ResultBoolIs over Empty is False).
Private Sub TestDatalogGroundQuery()
    Dim result As Variant
    Dim r As String
    Dim d As String
    Dim ans As Variant
    Dim res As Collection
    Dim loopy As String
    loopy = "(fact (link ""A"" ""B"")) (fact (link ""B"" ""C"")) (fact (link ""C"" ""A"")) (fact (link ""C"" ""D"")) (rule (route X Y) (link X Y)) (rule (route X Y) (link X Z) (route Z Y))"

    ' ---- THE ANSWER, both ways, as a Boolean cell rather than text.
    result = VLA_Datalog.DATALOG("(fact (link ""A"" ""B"")) (query (link ""A"" ""B""))")
    Report "datalog.9: a query written as one fact answers TRUE when the fact holds - a Boolean, not text", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Datalog.DATALOG("(fact (link ""A"" ""B"")) (query (link ""B"" ""A""))")
    Report "datalog.9: ...and FALSE when it does not, where it used to be refused as not a plain name", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)

    ' ---- THE REASON. Closure over data with a cycle: PROLOG collects every
    ' proof first, so it never finished either question.
    result = VLA_Datalog.DATALOG(loopy & " (query (route ""A"" ""D""))")
    Report "datalog.9: a closure over a cycle answers TRUE and stops - A reaches D", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Datalog.DATALOG(loopy & " (query (route ""A"" ""E""))")
    Report "datalog.9: ...and FALSE and stops when nothing reaches, from inside the cycle", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)

    ' ---- THE SAME MATCH A RULE BODY MAKES: bare constants, and text
    ' compared case-sensitively.
    result = VLA_Datalog.DATALOG("(fact (parent tom bob)) (query (parent tom bob))")
    Report "datalog.9: bare constants are constants - (query (parent tom bob)) is TRUE", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Datalog.DATALOG("(fact (person ""Bob"")) (query (person ""bob""))")
    Report "datalog.9: text matches case-sensitively, as in a rule body - ""bob"" is not ""Bob""", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)

    ' ---- A NUMBER CELL, through a rule over a keyed Table: DatalogRun's
    ' fifth item carries the answer.
    Dim staff(1 To 3, 1 To 2) As Variant
    staff(1, 1) = "Name": staff(1, 2) = "Level"
    staff(2, 1) = "Ann": staff(2, 2) = 3
    staff(3, 1) = "Bob": staff(3, 2) = 1
    Dim staffCols As New Collection
    staffCols.Add SqlColPair("name", "Name")
    staffCols.Add SqlColPair("level", "Level")
    Dim headerMapStaff As Object
    Set headerMapStaff = VLA_Runtime.VlaDictNew()
    VLA_Runtime.VlaDictSet headerMapStaff, "staff", staffCols
    Dim basesAnn As Object
    Set basesAnn = VLA_Runtime.VlaDictNew()
    VLA_Runtime.VlaDictSet basesAnn, "staff", VLA_Relation.RelFromRange(staff, True)
    d = ""
    ans = Empty
    On Error Resume Next
    Err.Clear
    Set res = VLA_Datalog.DatalogRun("(rule (has-level P L) (staff (name P) (level L))) (query (has-level ""Ann"" 3))", basesAnn, headerMapStaff)
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    If Len(d) = 0 Then ans = res.Item(5)
    Report "datalog.9: a bare 3 matches a number cell through a rule over a keyed Table - Ann is at level 3", _
           ResultBoolIs(ans, True), "got: '" & d & "', " & ResultDescribe(ans)
    Dim basesBob As Object
    Set basesBob = VLA_Runtime.VlaDictNew()
    VLA_Runtime.VlaDictSet basesBob, "staff", VLA_Relation.RelFromRange(staff, True)
    d = ""
    ans = Empty
    On Error Resume Next
    Err.Clear
    Set res = VLA_Datalog.DatalogRun("(rule (has-level P L) (staff (name P) (level L))) (query (has-level ""Bob"" 3))", basesBob, headerMapStaff)
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    If Len(d) = 0 Then ans = res.Item(5)
    Report "datalog.9: ...and Bob, at level 1, is FALSE", _
           ResultBoolIs(ans, False), "got: '" & d & "', " & ResultDescribe(ans)

    ' ---- G-PROLOG's OWN WHETHER TAIL, the text slice 3 will route here.
    Dim rota(1 To 3, 1 To 2) As Variant
    rota(1, 1) = "Name": rota(1, 2) = "Shift"
    rota(2, 1) = "Bob": rota(2, 2) = "Night"
    rota(3, 1) = "Di": rota(3, 2) = "Day"
    Dim rotaCols As New Collection
    rotaCols.Add SqlColPair("name", "Name")
    rotaCols.Add SqlColPair("shift", "Shift")
    Dim headerMapRota As Object
    Set headerMapRota = VLA_Runtime.VlaDictNew()
    VLA_Runtime.VlaDictSet headerMapRota, "rota", rotaCols
    Dim basesRota As Object
    Set basesRota = VLA_Runtime.VlaDictNew()
    VLA_Runtime.VlaDictSet basesRota, "rota", VLA_Relation.RelFromRange(rota, True)
    d = ""
    ans = Empty
    On Error Resume Next
    Err.Clear
    Set res = VLA_Datalog.DatalogRun("(rule (can-cover Person Shift) (rota (name Person) (shift Shift))) (query (can-cover ""Bob"" ""Night""))", basesRota, headerMapRota)
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    If Len(d) = 0 Then ans = res.Item(5)
    Report "datalog.9: G-PROLOG's whether tail answers here - Bob can cover Night", _
           ResultBoolIs(ans, True), "got: '" & d & "', " & ResultDescribe(ans)

    ' ---- A QUERY BY NAME IS UNCHANGED: rows, and an Empty fifth item.
    result = VLA_Datalog.DATALOG("(fact (p a)) (fact (p b)) (query p)")
    Report "datalog.9: a query by name still spills its rows - a header and two", _
           ResultRowCount(result) = 3 And ResultCellIs(result, 2, 1, "a"), "got: " & ResultDescribe(result)
    d = ""
    ans = "unset"
    On Error Resume Next
    Err.Clear
    Set res = VLA_Datalog.DatalogRun("(fact (p a)) (query p)")
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    If Len(d) = 0 Then ans = res.Item(5)
    Report "datalog.9: ...and DatalogRun's fifth item is Empty for it", _
           Len(d) = 0 And IsEmpty(ans), "got: '" & d & "', " & TypeName(ans)

    ' ---- EACH SHAPE STILL REFUSED, by name.
    result = VLA_Datalog.DATALOG("(fact (link ""A"" ""B"")) (query (link ""A"" Who))")
    r = ResultDescribe(result)
    Report "datalog.9: a query atom holding a variable is refused, naming it and teaching the rule", _
           ResultTextStartsWith(result, "#DATALOG!") And InStr(1, r, "holds the variable 'Who'", vbTextCompare) > 0, "got: " & r
    Dim basesKeyed As Object
    Set basesKeyed = VLA_Runtime.VlaDictNew()
    VLA_Runtime.VlaDictSet basesKeyed, "staff", VLA_Relation.RelFromRange(staff, True)
    d = ""
    On Error Resume Next
    Err.Clear
    Set res = VLA_Datalog.DatalogRun("(query (staff (name ""Ann"") (level 3)))", basesKeyed, headerMapStaff)
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    Report "datalog.9: a keyed query atom is refused as keyed, not as nesting", _
           InStr(1, d, "keys its values by column name", vbTextCompare) > 0 And InStr(1, d, "nested", vbTextCompare) = 0, "got: " & d
    result = VLA_Datalog.DATALOG("(headless) (fact (p a)) (query (p a))")
    r = ResultDescribe(result)
    Report "datalog.9: (headless) before a query written as one fact is refused", _
           ResultTextStartsWith(result, "#DATALOG!") And InStr(1, r, "asks for rows without their header", vbTextCompare) > 0, "got: " & r
    result = VLA_Datalog.DATALOG("(fact (p a)) (query (p a)) (headless)")
    r = ResultDescribe(result)
    Report "datalog.9: ...and after it", _
           ResultTextStartsWith(result, "#DATALOG!") And InStr(1, r, "asks for rows without their header", vbTextCompare) > 0, "got: " & r
    result = VLA_Datalog.DATALOG("(fact (link ""A"" ""B"")) (query (link ""A""))")
    r = ResultDescribe(result)
    Report "datalog.9: a query atom with the wrong number of values is an arity mismatch", _
           ResultTextStartsWith(result, "#DATALOG!") And InStr(1, r, "used with 2 argument(s) in one place and 1", vbTextCompare) > 0, "got: " & r
    result = VLA_Datalog.DATALOG("(fact (p a)) (query (pp a))")
    r = ResultDescribe(result)
    Report "datalog.9: a query atom over a name nothing defines keeps the query's own words", _
           ResultTextStartsWith(result, "#DATALOG!") And InStr(1, r, "(query pp) names a predicate with no facts", vbTextCompare) > 0, "got: " & r
    result = VLA_Datalog.DATALOG("(fact (p a)) (rule (q X) (p X) (typo X)) (query (p a))")
    r = ResultDescribe(result)
    Report "datalog.9: the whole program is still checked - a misspelling in a rule the query never uses refuses", _
           ResultTextStartsWith(result, "#DATALOG!") And InStr(1, r, "'typo' is used in a rule", vbTextCompare) > 0, "got: " & r
    result = VLA_Datalog.DATALOG("(fact (p a)) (query (p a) (p b))")
    r = ResultDescribe(result)
    Report "datalog.9: two facts in one query stay refused - one thing per query", _
           ResultTextStartsWith(result, "#DATALOG!") And InStr(1, r, "takes exactly one thing", vbTextCompare) > 0, "got: " & r
    ' DATALOG.10 re-pointed this pin in place: one atom under (not ...) is a
    ' query shape of its own now, answered rather than refused as nesting.
    result = VLA_Datalog.DATALOG("(fact (p a)) (query (not (p a)))")
    Report "datalog.10: (query (not ...)) answers where it was refused as nesting - (not (p a)) beside the fact is FALSE", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)
    result = VLA_Datalog.DATALOG("(fact (p a)) (query (p))")
    r = ResultDescribe(result)
    Report "datalog.9: a query atom with no values is refused as having no arguments", _
           ResultTextStartsWith(result, "#DATALOG!") And InStr(1, r, "'p' has no arguments", vbTextCompare) > 0, "got: " & r
End Sub

' DATALOG.10: a query may be one atom under (not ...), and it answers TRUE
' when nothing matches (VLA_Datalog.bas's module header). Every expectation
' was derived by running its program through a transliteration of that
' module, and the same text through PROLOG's transliteration agreed wherever
' PROLOG finished.
Private Sub TestDatalogNegatedQuery()
    Dim result As Variant
    Dim r As String
    Dim d As String
    Dim ans As Variant
    Dim res As Collection
    Dim loopy As String
    Dim gaps As String
    loopy = "(fact (link ""A"" ""B"")) (fact (link ""B"" ""C"")) (fact (link ""C"" ""A"")) (fact (link ""C"" ""D"")) (rule (route X Y) (link X Y)) (rule (route X Y) (link X Z) (route Z Y))"
    gaps = "(fact (shift ""Day"")) (fact (shift ""Night"")) (fact (covered ""Day"")) (rule (gap S) (shift S) (not (covered S)))"

    ' ---- VALUES ONLY: DATALOG.9's ground atom, inverted, as a Boolean.
    result = VLA_Datalog.DATALOG("(fact (p a)) (query (not (p b)))")
    Report "datalog.10: (not ...) over a fact nothing states is TRUE - a Boolean, not text", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Datalog.DATALOG("(fact (p a)) (fact (p b)) (query (not (p b)))")
    Report "datalog.10: ...and FALSE over a fact the program states", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)

    ' ---- BLANKS: a blank is any value, and a blank repeated is the same value.
    result = VLA_Datalog.DATALOG("(fact (link ""A"" ""B"")) (query (not (link ""A"" X)))")
    Report "datalog.10: a blank is any value - something links from A, so nothing-from-A is FALSE", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)
    result = VLA_Datalog.DATALOG("(fact (link ""A"" ""B"")) (query (not (link ""B"" X)))")
    Report "datalog.10: ...and nothing links from B, so nothing-from-B is TRUE", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Datalog.DATALOG("(fact (link ""A"" ""B"")) (query (not (link X Y)))")
    Report "datalog.10: two blanks ask whether the relation holds any row - it does, so FALSE", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)
    result = VLA_Datalog.DATALOG("(fact (link ""A"" ""B"")) (query (not (link X X)))")
    Report "datalog.10: a repeated blank is one value - nothing links to itself, so TRUE", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Datalog.DATALOG("(fact (link ""A"" ""B"")) (fact (link ""C"" ""C"")) (query (not (link X X)))")
    Report "datalog.10: ...and FALSE once C links to itself", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)

    ' ---- OVER DATA THAT LOOPS, where PROLOG refuses by DEPTH either way.
    result = VLA_Datalog.DATALOG(loopy & " (query (not (route ""A"" X)))")
    Report "datalog.10: over a closure whose data loops - A reaches something, so FALSE, and it stops", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)
    result = VLA_Datalog.DATALOG(loopy & " (query (not (route X ""E"")))")
    Report "datalog.10: ...and nothing reaches E, so TRUE", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)

    ' ---- A RELATION HOLDING NO ROW, and EVERY as G-PROLOG slice 4 will ask it.
    result = VLA_Datalog.DATALOG("(fact (p a)) (rule (q X) (p X) (not (p X))) (query (not (q X)))")
    Report "datalog.10: a rule that derives nothing answers TRUE - no row can match", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Datalog.DATALOG(gaps & " (query (not (gap X)))")
    Report "datalog.10: every shift is covered is no shift is a gap - Night has no cover, so FALSE", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)
    result = VLA_Datalog.DATALOG(gaps & " (fact (covered ""Night"")) (query (not (gap X)))")
    Report "datalog.10: ...and TRUE once Night is covered", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)

    ' ---- A NUMBER CELL through a keyed Table: DatalogRun's fifth item.
    Dim staff(1 To 3, 1 To 2) As Variant
    staff(1, 1) = "Name": staff(1, 2) = "Level"
    staff(2, 1) = "Ann": staff(2, 2) = 3
    staff(3, 1) = "Bob": staff(3, 2) = 1
    Dim staffCols As New Collection
    staffCols.Add SqlColPair("name", "Name")
    staffCols.Add SqlColPair("level", "Level")
    Dim headerMapStaff As Object
    Set headerMapStaff = VLA_Runtime.VlaDictNew()
    VLA_Runtime.VlaDictSet headerMapStaff, "staff", staffCols
    Dim basesAnn As Object
    Set basesAnn = VLA_Runtime.VlaDictNew()
    VLA_Runtime.VlaDictSet basesAnn, "staff", VLA_Relation.RelFromRange(staff, True)
    d = ""
    ans = Empty
    On Error Resume Next
    Err.Clear
    Set res = VLA_Datalog.DatalogRun("(rule (has-level P L) (staff (name P) (level L))) (query (not (has-level ""Ann"" 3)))", basesAnn, headerMapStaff)
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    If Len(d) = 0 Then ans = res.Item(5)
    Report "datalog.10: a bare 3 matches a number cell under not - Ann is at level 3, so FALSE", _
           ResultBoolIs(ans, False), "got: '" & d & "', " & ResultDescribe(ans)
    Dim basesAt3 As Object
    Set basesAt3 = VLA_Runtime.VlaDictNew()
    VLA_Runtime.VlaDictSet basesAt3, "staff", VLA_Relation.RelFromRange(staff, True)
    d = ""
    ans = Empty
    On Error Resume Next
    Err.Clear
    Set res = VLA_Datalog.DatalogRun("(rule (has-level P L) (staff (name P) (level L))) (query (not (has-level Who 3)))", basesAt3, headerMapStaff)
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    If Len(d) = 0 Then ans = res.Item(5)
    Report "datalog.10: a blank beside a number - somebody is at level 3, so FALSE", _
           ResultBoolIs(ans, False), "got: '" & d & "', " & ResultDescribe(ans)
    Dim basesAt2 As Object
    Set basesAt2 = VLA_Runtime.VlaDictNew()
    VLA_Runtime.VlaDictSet basesAt2, "staff", VLA_Relation.RelFromRange(staff, True)
    d = ""
    ans = Empty
    On Error Resume Next
    Err.Clear
    Set res = VLA_Datalog.DatalogRun("(rule (has-level P L) (staff (name P) (level L))) (query (not (has-level Who 2)))", basesAt2, headerMapStaff)
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    If Len(d) = 0 Then ans = res.Item(5)
    Report "datalog.10: ...and nobody is at level 2, so TRUE", _
           ResultBoolIs(ans, True), "got: '" & d & "', " & ResultDescribe(ans)

    ' ---- A TABLE WITH NO ROWS: nothing in it can match.
    Dim hdr(1 To 1, 1 To 2) As Variant
    hdr(1, 1) = "Name": hdr(1, 2) = "Shift"
    Dim basesEmpty As Object
    Set basesEmpty = VLA_Runtime.VlaDictNew()
    VLA_Runtime.VlaDictSet basesEmpty, "leave", VLA_Relation.RelFromRange(hdr, True)
    d = ""
    ans = Empty
    On Error Resume Next
    Err.Clear
    Set res = VLA_Datalog.DatalogRun("(query (not (leave X Y)))", basesEmpty)
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    If Len(d) = 0 Then ans = res.Item(5)
    Report "datalog.10: a table argument with no rows answers TRUE", _
           ResultBoolIs(ans, True), "got: '" & d & "', " & ResultDescribe(ans)

    ' ---- STILL REFUSED, each in words that name the right thing.
    result = VLA_Datalog.DATALOG("(fact (p a)) (query (p X))")
    r = ResultDescribe(result)
    Report "datalog.10: a blank OUTSIDE not stays refused - PROLOG reads that text as a list of values", _
           ResultTextStartsWith(result, "#DATALOG!") And InStr(1, r, "holds the variable 'X'", vbTextCompare) > 0, "got: " & r
    result = VLA_Datalog.DATALOG("(fact (p a)) (query (not))")
    r = ResultDescribe(result)
    Report "datalog.10: (not) with nothing to negate says not takes one predicate form - not that a predicate 'not' has no arguments", _
           ResultTextStartsWith(result, "#DATALOG!") And InStr(1, r, "takes exactly one predicate form", vbTextCompare) > 0 And InStr(1, r, "has no arguments", vbTextCompare) = 0, "got: " & r
    result = VLA_Datalog.DATALOG("(fact (p a)) (fact (q b)) (query (not (p a) (q b)))")
    r = ResultDescribe(result)
    Report "datalog.10: ...and so does (not ...) with two forms", _
           ResultTextStartsWith(result, "#DATALOG!") And InStr(1, r, "takes exactly one predicate form", vbTextCompare) > 0, "got: " & r
    result = VLA_Datalog.DATALOG("(fact (p a)) (query (not p))")
    r = ResultDescribe(result)
    Report "datalog.10: a bare name under not asks for a predicate form - it is not read as a query named not", _
           ResultTextStartsWith(result, "#DATALOG!") And InStr(1, r, "expected a predicate form", vbTextCompare) > 0 And InStr(1, r, "(query not)", vbTextCompare) = 0, "got: " & r
    result = VLA_Datalog.DATALOG("(fact (p a)) (query (not (not (p a))))")
    r = ResultDescribe(result)
    Report "datalog.10: a double negation stays nesting - one not only", _
           ResultTextStartsWith(result, "#DATALOG!") And InStr(1, r, "nested inside another's argument, in a query", vbTextCompare) > 0, "got: " & r
    result = VLA_Datalog.DATALOG("(fact (p a)) (query (not (p)))")
    r = ResultDescribe(result)
    Report "datalog.10: an atom with no values under not is refused as having no arguments", _
           ResultTextStartsWith(result, "#DATALOG!") And InStr(1, r, "'p' has no arguments", vbTextCompare) > 0, "got: " & r
    Dim basesKeyed As Object
    Set basesKeyed = VLA_Runtime.VlaDictNew()
    VLA_Runtime.VlaDictSet basesKeyed, "staff", VLA_Relation.RelFromRange(staff, True)
    d = ""
    On Error Resume Next
    Err.Clear
    Set res = VLA_Datalog.DatalogRun("(query (not (staff (name ""Ann"") (level 3))))", basesKeyed, headerMapStaff)
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    Report "datalog.10: a keyed atom under not is refused as keyed, not as nesting", _
           InStr(1, d, "keys its values by column name", vbTextCompare) > 0 And InStr(1, d, "nested", vbTextCompare) = 0, "got: " & d
    result = VLA_Datalog.DATALOG("(headless) (fact (p a)) (query (not (p b)))")
    r = ResultDescribe(result)
    Report "datalog.10: (headless) beside a negated query is refused", _
           ResultTextStartsWith(result, "#DATALOG!") And InStr(1, r, "asks for rows without their header", vbTextCompare) > 0, "got: " & r
    result = VLA_Datalog.DATALOG("(fact (link ""A"" ""B"")) (query (not (link ""A"")))")
    r = ResultDescribe(result)
    Report "datalog.10: a negated atom with the wrong number of values is an arity mismatch", _
           ResultTextStartsWith(result, "#DATALOG!") And InStr(1, r, "used with 2 argument(s) in one place and 1", vbTextCompare) > 0, "got: " & r
    result = VLA_Datalog.DATALOG("(fact (p a)) (query (not (pp X)))")
    r = ResultDescribe(result)
    Report "datalog.10: a negated atom over a name nothing defines keeps the query's own words - never a silent TRUE", _
           ResultTextStartsWith(result, "#DATALOG!") And InStr(1, r, "(query pp) names a predicate with no facts", vbTextCompare) > 0, "got: " & r
    result = VLA_Datalog.DATALOG("(fact (p a)) (rule (q X) (p X) (typo X)) (query (not (p b)))")
    r = ResultDescribe(result)
    Report "datalog.10: the whole program is still checked - a misspelling in a rule the query never uses refuses", _
           ResultTextStartsWith(result, "#DATALOG!") And InStr(1, r, "'typo' is used in a rule", vbTextCompare) > 0, "got: " & r
    result = VLA_Datalog.DATALOG("(fact (p a)) (query (p a) (p b))")
    r = ResultDescribe(result)
    Report "datalog.10: the query-shape refusal teaches the negated spelling too", _
           ResultTextStartsWith(result, "#DATALOG!") And InStr(1, r, "to answer TRUE when nothing matches", vbTextCompare) > 0, "got: " & r
    result = VLA_Datalog.DATALOG("(fact (p a)) (query (""not"" (p a)))")
    r = ResultDescribe(result)
    Report "datalog.10: a quoted ""not"" is a name, never the wrapper - read as one fact, whose argument is nested", _
           ResultTextStartsWith(result, "#DATALOG!") And InStr(1, r, "nested inside another's argument, in a query", vbTextCompare) > 0, "got: " & r
End Sub

' DATALOG.11: TEXT TESTS, TEXTJOIN, AND PROLOG'S GOALS REFUSED BY NAME.
' (text-starts-with T P), (text-ends-with T P) and (text-contains T P) are
' filters read the way a comparison is; (textjoin R Sep (pred ...)) joins a
' column's values into one cell, count's and sum's shape. Pure: a number cell
' is a hand-built array through RelFromRange.
Private Sub TestDatalogTextTests()
    Dim result As Variant
    Dim r As String
    Dim d As String
    Dim res As Collection
    Dim accts As String
    Dim cover As String
    accts = "(fact (acct ""GL-4010"")) (fact (acct ""GL-4020"")) (fact (acct ""GL-5010""))"
    cover = "(fact (cc ""Ann"" ""Day"")) (fact (cc ""Bob"" ""Day"")) (fact (cc ""Ed"" ""Day"")) (fact (cc ""Bob"" ""Night"")) (fact (shift ""Day"")) (fact (shift ""Night"")) (fact (shift ""Weekend""))"

    ' ---- THE THREE TESTS, each a filter answering once per row.
    result = VLA_Datalog.DATALOG(accts & " (rule (revenue C) (acct C) (text-starts-with C ""GL-4"")) (query revenue)")
    Report "datalog.11: text-starts-with keeps the codes that start with GL-4", _
           ResultRowCount(result) = 3 And ResultCellIs(result, 2, 1, "GL-4010") And ResultCellIs(result, 3, 1, "GL-4020"), "got: " & ResultDescribe(result)
    result = VLA_Datalog.DATALOG(accts & " (rule (tens C) (acct C) (text-ends-with C ""10"")) (query tens)")
    Report "datalog.11: text-ends-with keeps the codes that end with 10", _
           ResultRowCount(result) = 3 And ResultCellIs(result, 2, 1, "GL-4010") And ResultCellIs(result, 3, 1, "GL-5010"), "got: " & ResultDescribe(result)
    result = VLA_Datalog.DATALOG("(fact (product ""Banana"")) (fact (product ""Cyan"")) (fact (product ""Apple"")) (rule (has-an N) (product N) (text-contains N ""an"")) (query has-an)")
    Report "datalog.11: text-contains answers once per row however often the part occurs - Banana once, then Cyan", _
           ResultRowCount(result) = 3 And ResultCellIs(result, 2, 1, "Banana") And ResultCellIs(result, 3, 1, "Cyan"), "got: " & ResultDescribe(result)
    result = VLA_Datalog.DATALOG(accts & " (rule (revenue C) (acct C) (text-starts-with C ""gl-4"")) (query revenue)")
    Report "datalog.11: case is exact - gl-4 starts no code, header only", _
           ResultRowCount(result) = 1, "got: " & ResultDescribe(result)
    result = VLA_Datalog.DATALOG(accts & " (rule (any-code C) (acct C) (text-contains C """")) (query any-code)")
    Report "datalog.11: every text contains the empty text", _
           ResultRowCount(result) = 4, "got: " & ResultDescribe(result)
    result = VLA_Datalog.DATALOG(accts & " (rule (other C) (acct C) (not (text-starts-with C ""GL-4""))) (query other)")
    Report "datalog.11: a text test under not answers inverted - GL-5010 alone", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, "GL-5010"), "got: " & ResultDescribe(result)
    result = VLA_Datalog.DATALOG("(fact (pair ""GL-4010"" ""GL-4"")) (fact (pair ""GL-5010"" ""GL-4"")) (rule (ok C) (pair C P) (text-starts-with C P)) (query ok)")
    Report "datalog.11: both operands may be variables an earlier atom bound", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, "GL-4010"), "got: " & ResultDescribe(result)

    ' ---- A NUMBER CELL is read as its canonical text, the same on every machine.
    Dim codes(1 To 4, 1 To 2) As Variant
    codes(1, 1) = "Code": codes(1, 2) = "Name"
    codes(2, 1) = 4010: codes(2, 2) = "Sales"
    codes(3, 1) = "007": codes(3, 2) = "Petty"
    codes(4, 1) = 0.5: codes(4, 2) = "Half"
    Dim codeCols As New Collection
    codeCols.Add SqlColPair("code", "Code")
    codeCols.Add SqlColPair("name", "Name")
    Dim codeHeaders As Object
    Set codeHeaders = VLA_Runtime.VlaDictNew()
    VLA_Runtime.VlaDictSet codeHeaders, "codes", codeCols
    Dim basesForty As Object
    Set basesForty = VLA_Runtime.VlaDictNew()
    VLA_Runtime.VlaDictSet basesForty, "codes", VLA_Relation.RelFromRange(codes, True)
    d = ""
    result = Empty
    On Error Resume Next
    Err.Clear
    Set res = VLA_Datalog.DatalogRun("(rule (forty N) (codes (code C) (name N)) (text-starts-with C 40)) (query forty)", basesForty, codeHeaders)
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    If Len(d) = 0 Then result = VLA_Relation.RelToSpilledArray(VLA_Runtime.VlaDictGet(res.Item(2), "forty"))
    Report "datalog.11: a number cell holding 4010 starts with 40 - Sales, and the text 007 does not", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, "Sales"), "got: '" & d & "', " & ResultDescribe(result)
    Dim basesHalf As Object
    Set basesHalf = VLA_Runtime.VlaDictNew()
    VLA_Runtime.VlaDictSet basesHalf, "codes", VLA_Relation.RelFromRange(codes, True)
    d = ""
    result = Empty
    On Error Resume Next
    Err.Clear
    Set res = VLA_Datalog.DatalogRun("(rule (half N) (codes (code C) (name N)) (text-starts-with C ""0.5"")) (query half)", basesHalf, codeHeaders)
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    If Len(d) = 0 Then result = VLA_Relation.RelToSpilledArray(VLA_Runtime.VlaDictGet(res.Item(2), "half"))
    Report "datalog.11: a number cell holding 0.5 reads 0.5, its leading zero kept", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, "Half"), "got: '" & d & "', " & ResultDescribe(result)
    Dim basesZeros As Object
    Set basesZeros = VLA_Runtime.VlaDictNew()
    VLA_Runtime.VlaDictSet basesZeros, "codes", VLA_Relation.RelFromRange(codes, True)
    d = ""
    result = Empty
    On Error Resume Next
    Err.Clear
    Set res = VLA_Datalog.DatalogRun("(rule (zeros N) (codes (code C) (name N)) (text-starts-with C ""00"")) (query zeros)", basesZeros, codeHeaders)
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    If Len(d) = 0 Then result = VLA_Relation.RelToSpilledArray(VLA_Runtime.VlaDictGet(res.Item(2), "zeros"))
    Report "datalog.11: a text cell keeps its own text - 007 starts with 00", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, "Petty"), "got: '" & d & "', " & ResultDescribe(result)

    ' ---- TEXTJOIN: a column's values in one cell per group, the empty group blank.
    result = VLA_Datalog.DATALOG(cover & " (rule (who-covers S W) (shift S) (textjoin W "", "" (cc P S))) (query who-covers)")
    Report "datalog.11: textjoin puts each shift's people in one cell, in the relation's own row order", _
           ResultRowCount(result) = 4 And ResultCellIs(result, 2, 2, "Ann, Bob, Ed") And ResultCellIs(result, 3, 2, "Bob"), "got: " & ResultDescribe(result)
    Report "datalog.11: ...and a shift nobody covers holds empty text, never a missing row", _
           ResultCellIs(result, 4, 1, "Weekend") And ResultCellIs(result, 4, 2, ""), "got: " & ResultDescribe(result)
    result = VLA_Datalog.DATALOG(cover & " (rule (who-covers S W) (shift S) (textjoin W ""; "" (cc P S))) (query who-covers)")
    Report "datalog.11: the separator is the one written - a semicolon", _
           ResultCellIs(result, 2, 2, "Ann; Bob; Ed"), "got: " & ResultDescribe(result)
    result = VLA_Datalog.DATALOG(cover & " (rule (who-covers S W) (shift S) (textjoin W """" (cc P S))) (query who-covers)")
    Report "datalog.11: ...and an empty separator joins with nothing between", _
           ResultCellIs(result, 2, 2, "AnnBobEd"), "got: " & ResultDescribe(result)
    result = VLA_Datalog.DATALOG(cover & " (fact (spare ""Bob"" ""Day"")) (rule (can P S) (cc P S)) (rule (can P S) (spare P S)) (rule (who-can S W) (shift S) (textjoin W "", "" (can P S))) (query who-can)")
    Report "datalog.11: someone who qualifies two ways is joined once - a relation holds each row once", _
           ResultCellIs(result, 2, 2, "Ann, Bob, Ed"), "got: " & ResultDescribe(result)
    Dim staff(1 To 5, 1 To 2) As Variant
    staff(1, 1) = "Name": staff(1, 2) = "Level"
    staff(2, 1) = "Ann": staff(2, 2) = 3
    staff(3, 1) = "Bob": staff(3, 2) = 3
    staff(4, 1) = "Bob": staff(4, 2) = 1
    staff(5, 1) = "Di": staff(5, 2) = 0.5
    Dim staffCols As New Collection
    staffCols.Add SqlColPair("name", "Name")
    staffCols.Add SqlColPair("level", "Level")
    Dim staffHeaders As Object
    Set staffHeaders = VLA_Runtime.VlaDictNew()
    VLA_Runtime.VlaDictSet staffHeaders, "staff", staffCols
    Dim basesLevels As Object
    Set basesLevels = VLA_Runtime.VlaDictNew()
    VLA_Runtime.VlaDictSet basesLevels, "staff", VLA_Relation.RelFromRange(staff, True)
    d = ""
    result = Empty
    On Error Resume Next
    Err.Clear
    Set res = VLA_Datalog.DatalogRun("(rule (levels N L) (staff (name N)) (textjoin L "", "" (staff (name N) (level V)))) (query levels)", basesLevels, staffHeaders)
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    If Len(d) = 0 Then result = VLA_Relation.RelToSpilledArray(VLA_Runtime.VlaDictGet(res.Item(2), "levels"))
    Report "datalog.11: number cells join as their canonical text - Bob 3, 1 and Di 0.5", _
           ResultRowCount(result) = 4 And ResultCellIs(result, 3, 2, "3, 1") And ResultCellIs(result, 4, 2, "0.5"), "got: '" & d & "', " & ResultDescribe(result)
    result = VLA_Datalog.DATALOG("(fact (p ""x"" """ & String$(16383, "a") & """)) (fact (p ""x"" """ & String$(16382, "b") & """)) (rule (j K W) (p K Z) (textjoin W "", "" (p K V))) (query j)")
    Report "datalog.11: BOUNDARY - a join of exactly 32,767 characters fits a cell", _
           ResultRowCount(result) = 2 And Len(ResultCellText(result, 2, 2)) = 32767, "got: " & ResultDescribe(result)
    result = VLA_Datalog.DATALOG("(fact (p ""x"" """ & String$(16383, "a") & """)) (fact (p ""x"" """ & String$(16383, "b") & """)) (rule (j K W) (p K Z) (textjoin W "", "" (p K V))) (query j)")
    r = ResultDescribe(result)
    Report "datalog.11: ...and one character more is refused by name, with the length", _
           ResultTextStartsWith(result, "#DATALOG!") And InStr(1, r, "would put 32768 characters into 'W'", vbTextCompare) > 0, "got: " & r

    ' ---- REFUSED BY NAME.
    result = VLA_Datalog.DATALOG(accts & " (rule (revenue C) (text-starts-with C ""GL-4"") (acct C)) (query revenue)")
    r = ResultDescribe(result)
    Report "datalog.11: a text test before anything binds its text is refused, naming the variable", _
           ResultTextStartsWith(result, "#DATALOG!") And InStr(1, r, "in 'text-starts-with', the variable 'C'", vbTextCompare) > 0, "got: " & r
    result = VLA_Datalog.DATALOG(accts & " (rule (revenue C) (acct C) (text-starts-with C)) (query revenue)")
    r = ResultDescribe(result)
    Report "datalog.11: a text test with one operand is refused, showing the two-operand form", _
           ResultTextStartsWith(result, "#DATALOG!") And InStr(1, r, "like (text-starts-with Code ""GL-4"") - found 1", vbTextCompare) > 0, "got: " & r
    result = VLA_Datalog.DATALOG(cover & " (rule (who-covers S W) (shift S) (textjoin W (cc P S))) (query who-covers)")
    r = ResultDescribe(result)
    Report "datalog.11: textjoin without its separator is refused, showing its shape", _
           ResultTextStartsWith(result, "#DATALOG!") And InStr(1, r, "a result variable, a separator and a predicate form", vbTextCompare) > 0, "got: " & r
    result = VLA_Datalog.DATALOG(cover & " (rule (who-covers S W) (shift S) (textjoin W Sep (cc P S))) (query who-covers)")
    r = ResultDescribe(result)
    Report "datalog.11: a separator that is a variable is refused", _
           ResultTextStartsWith(result, "#DATALOG!") And InStr(1, r, "is the separator, written out", vbTextCompare) > 0, "got: " & r
    result = VLA_Datalog.DATALOG(cover & " (rule (all W) (shift S) (textjoin W "", "" (cc P Q))) (query all)")
    r = ResultDescribe(result)
    Report "datalog.11: textjoin with two arguments left free is refused - which one is the value?", _
           ResultTextStartsWith(result, "#DATALOG!") And InStr(1, r, "(the value to join) - found 2", vbTextCompare) > 0, "got: " & r
    result = VLA_Datalog.DATALOG(cover & " (rule (who-covers S) (shift S) (textjoin S "", "" (cc P S))) (query who-covers)")
    r = ResultDescribe(result)
    Report "datalog.11: a textjoin result that is already bound is refused", _
           ResultTextStartsWith(result, "#DATALOG!") And InStr(1, r, "'S' is already bound earlier", vbTextCompare) > 0, "got: " & r
    result = VLA_Datalog.DATALOG(cover & " (rule (chain S W) (shift S) (textjoin W "", "" (chain X S))) (query chain)")
    r = ResultDescribe(result)
    Report "datalog.11: a relation joined into itself is refused as unstratifiable, naming textjoin", _
           ResultTextStartsWith(result, "#DATALOG!") And InStr(1, r, "(sum ...), or (textjoin ...)", vbTextCompare) > 0, "got: " & r

    ' ---- PROLOG'S GOALS: named as PROLOG's, pointing at DATALOG's own spelling.
    result = VLA_Datalog.DATALOG(accts & " (rule (revenue C) (acct C) (sub-atom C 0 L A ""GL-4"")) (query revenue)")
    r = ResultDescribe(result)
    Report "datalog.11: sub-atom is named as PROLOG's goal and points at text-starts-with - not a misspelling", _
           ResultTextStartsWith(result, "#DATALOG!") And InStr(1, r, "'sub-atom' is one of PROLOG's goals", vbTextCompare) > 0 And InStr(1, r, "(text-starts-with Text Start)", vbTextCompare) > 0 And InStr(1, r, "check the spelling", vbTextCompare) = 0, "got: " & r
    result = VLA_Datalog.DATALOG(accts & " (rule (kind C T) (acct C) (if (text-starts-with C ""GL-4"") (= T ""Revenue"") (= T ""Cost""))) (query kind)")
    r = ResultDescribe(result)
    Report "datalog.11: (if ...) is named as PROLOG's - not as nesting", _
           ResultTextStartsWith(result, "#DATALOG!") And InStr(1, r, "'if' is one of PROLOG's goals", vbTextCompare) > 0 And InStr(1, r, "one rule for each case", vbTextCompare) > 0 And InStr(1, r, "nested", vbTextCompare) = 0, "got: " & r
    result = VLA_Datalog.DATALOG(cover & " (rule (who-covers S L) (shift S) (findall P (cc P S) L)) (query who-covers)")
    r = ResultDescribe(result)
    Report "datalog.11: findall is named as PROLOG's and points at textjoin - not as mixed keying", _
           ResultTextStartsWith(result, "#DATALOG!") And InStr(1, r, "'findall' is one of PROLOG's goals", vbTextCompare) > 0 And InStr(1, r, "textjoin", vbTextCompare) > 0 And InStr(1, r, "keyed", vbTextCompare) = 0, "got: " & r
    result = VLA_Datalog.DATALOG("(fact (lvl ""Ann"" 3)) (rule (low P) (lvl P L) (=< L 2)) (query low)")
    r = ResultDescribe(result)
    Report "datalog.11: PROLOG's =< is named, pointing at DATALOG's <=", _
           ResultTextStartsWith(result, "#DATALOG!") And InStr(1, r, "'=<' is one of PROLOG's goals", vbTextCompare) > 0 And InStr(1, r, "< > <= >= = and <>", vbTextCompare) > 0, "got: " & r
    result = VLA_Datalog.DATALOG(accts & " (rule (first-code C) (acct C) !) (query first-code)")
    r = ResultDescribe(result)
    Report "datalog.11: a bare cut is named as PROLOG's", _
           ResultTextStartsWith(result, "#DATALOG!") And InStr(1, r, "'!' is one of PROLOG's goals", vbTextCompare) > 0, "got: " & r
    result = VLA_Datalog.DATALOG(accts & " (rule (odd C) (acct C) (not (sub-atom C 0 L A ""GL""))) (query odd)")
    r = ResultDescribe(result)
    Report "datalog.11: ...and so is a PROLOG goal under not", _
           ResultTextStartsWith(result, "#DATALOG!") And InStr(1, r, "'sub-atom' is one of PROLOG's goals", vbTextCompare) > 0, "got: " & r
    result = VLA_Datalog.DATALOG("(query (sub-atom ""abc"" 0 1 A ""a""))")
    r = ResultDescribe(result)
    Report "datalog.11: ...and one written as a query", _
           ResultTextStartsWith(result, "#DATALOG!") And InStr(1, r, "'sub-atom' is one of PROLOG's goals", vbTextCompare) > 0, "got: " & r
    result = VLA_Datalog.DATALOG("(fact (member ""Ann"" ""Ops"")) (rule (in-ops P) (member P ""Ops"")) (query in-ops)")
    Report "datalog.11: a program that DEFINES a relation called member keeps it - DATALOG never reserved the name", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, "Ann"), "got: " & ResultDescribe(result)

    ' ---- THE TEXT TESTS' OWN NAMES.
    result = VLA_Datalog.DATALOG("(fact (text-contains ""a"" ""b"")) (query text-contains)")
    r = ResultDescribe(result)
    Report "datalog.11: a fact named as a text test is refused", _
           ResultTextStartsWith(result, "#DATALOG!") And InStr(1, r, "'text-contains' is one of DATALOG's text tests", vbTextCompare) > 0, "got: " & r
    result = VLA_Datalog.DATALOG("(fact (pair ""a"" ""b"")) (rule (text-ends-with X Y) (pair X Y)) (query text-ends-with)")
    r = ResultDescribe(result)
    Report "datalog.11: ...and so is a rule head", _
           ResultTextStartsWith(result, "#DATALOG!") And InStr(1, r, "'text-ends-with' is one of DATALOG's text tests", vbTextCompare) > 0, "got: " & r
    Dim tbl(1 To 2, 1 To 2) As Variant
    tbl(1, 1) = "A": tbl(1, 2) = "B"
    tbl(2, 1) = "x": tbl(2, 2) = "y"
    Dim basesNamed As Object
    Set basesNamed = VLA_Runtime.VlaDictNew()
    VLA_Runtime.VlaDictSet basesNamed, "text-starts-with", VLA_Relation.RelFromRange(tbl, True)
    d = ""
    On Error Resume Next
    Err.Clear
    Set res = VLA_Datalog.DatalogRun("(fact (p ""a"")) (query p)", basesNamed)
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    Report "datalog.11: ...and so is a table argument", _
           InStr(1, d, "'text-starts-with' is one of DATALOG's text tests", vbTextCompare) > 0, "got: " & d
    result = VLA_Datalog.DATALOG(accts & " (query (text-starts-with ""GL-4010"" ""GL-4""))")
    r = ResultDescribe(result)
    Report "datalog.11: a text test asked on its own as a query is refused, teaching the rule spelling", _
           ResultTextStartsWith(result, "#DATALOG!") And InStr(1, r, "asks a text test on its own", vbTextCompare) > 0, "got: " & r
End Sub

' Host-required: every real bug this engine's MVP ever found (this
' module's own header note has the list) only ever showed up through a
' real Excel Table, never through RelFromRange fed a hand-built array -
' TestDatalog/TestDatalogNegation above are pure precisely because
' RelFromRange makes that possible, but that same seam is exactly why
' they could never have caught either bug. This sub builds a real
' ListObject on a real sheet and calls the actual =DATALOG(...) worksheet
' function (VLA_Datalog.DATALOG, called directly - same code Excel calls
' from a cell formula, no Application.Run indirection needed since this
' test lives in the same VBA project) against it, live.
'
' Three scenarios - the first two matching real bugs this MVP hit by
' name, the third proving DATALOG.3's own new capability:
'
' 1. A table's predicate name is ALWAYS its own ListObject.Name, never a
'    Name-Manager alias pointing at the same cells (VLA_Datalog.bas's
'    own TableArgName header, and this codebase's own house-style trap
'    list). Proven both directions on ONE live table: querying by the
'    table's real (folded) ListObject name through a Range obtained via
'    the ALIAS succeeds; querying by the alias's own name, through the
'    identical Range, is refused as an unknown predicate - the alias
'    changes nothing about which name DATALOG will accept.
'
' 2. A recursive rule (transitive closure) sourced from a live Table's
'    own DataBodyRange, not a hand-built array - RelFromRange's
'    ListObject-aware header-stripping path, exercised for real, feeding
'    RunFixpointForRules's own semi-naive recursion.
'
' 3. DATALOG.3's own column-scoped table arguments - necessarily host-
'    required too, since a plain 2D array has no .ListObject at all to
'    narrow: passing a CONTIGUOUS 2-of-3-column slice of a live Table
'    narrows the resulting relation's own arity to 2 (Salary excluded)
'    while the predicate name stays the table's own real ListObject.Name
'    regardless (proven implicitly - the query would fail under an
'    unknown-predicate refusal if TableArgName's own naming had shifted);
'    a non-contiguous (Ctrl-selected) column pick on the SAME table is
'    refused by name, not silently guessed.
Private Sub TestDatalogHostTable()
    Dim prior As Worksheet
    Set prior = ActiveSheet

    ' Clean slate - a prior failed run could have left either behind.
    On Error Resume Next
    ThisWorkbook.Names("VlaDatalogAliasTest").Delete
    On Error GoTo 0

    VlaEnsureSheet "VlaDatalogHostSheet"
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets("VlaDatalogHostSheet")
    ws.Activate
    Do While ws.ListObjects.Count > 0
        ws.ListObjects(1).Delete
    Loop
    ws.Cells.Clear

    ' Scenario 2's own table: a two-edge chain, positionally-faceted
    ' (arbitrary header text - DATALOG strips it and never reads it).
    ws.Range("A1:B1").Value = Array("P", "C")
    ws.Range("A2:B2").Value = Array("alice", "bob")
    ws.Range("A3:B3").Value = Array("bob", "carol")
    Dim loEdges As ListObject
    Set loEdges = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:B3"), , xlYes)
    loEdges.Name = "ReportsToHostTest1"

    Dim arrEdges As Variant
    arrEdges = VLA_Datalog.DATALOG( _
        "(headless)" & _
        " (rule (indirect X Y) (reportstohosttest1 X Y))" & _
        " (rule (indirect X Y) (reportstohosttest1 X Z) (indirect Z Y))" & _
        " (query indirect)", _
        loEdges.Range)
    ' Not IIf/And here - VBA's IIf and And both evaluate EVERY operand
    ' regardless of the condition (neither short-circuits), so
    ' UBound(arrEdges, 1) inside either would still run, and raise,
    ' on DATALOG's own zero-row "" fallback (a String, not an array).
    Dim okEdges As Boolean, detailEdges As String
    If IsArray(arrEdges) Then
        Dim nRowsEdges As Long
        nRowsEdges = UBound(arrEdges, 1) - LBound(arrEdges, 1) + 1
        okEdges = (nRowsEdges = 3)
        detailEdges = "got " & nRowsEdges & " rows"
    Else
        detailEdges = "got " & TypeName(arrEdges) & " = " & CStr(arrEdges)
    End If
    Report "datalog host: recursive rule over a LIVE Table's own facts derives transitive closure (3 rows)", _
           okEdges, detailEdges

    ' Scenario 1's own table: a one-column fact list, plus a Name-
    ' Manager alias pointing at the SAME cells under a DIFFERENT name.
    ws.Range("D1").Value = "Name"
    ws.Range("D2").Value = "alice"
    ws.Range("D3").Value = "bob"
    ws.Range("D4").Value = "carol"
    Dim loPeople As ListObject
    Set loPeople = ws.ListObjects.Add(xlSrcRange, ws.Range("D1:D4"), , xlYes)
    loPeople.Name = "PersonHostTest1"
    ThisWorkbook.Names.Add Name:="VlaDatalogAliasTest", RefersTo:="='" & ws.Name & "'!" & loPeople.Range.Address

    Dim aliasRange As Range
    Set aliasRange = ThisWorkbook.Names("VlaDatalogAliasTest").RefersToRange

    Dim arrByRealName As Variant
    arrByRealName = VLA_Datalog.DATALOG("(headless) (query personhosttest1)", aliasRange)
    Dim okRealName As Boolean, detailRealName As String
    If IsArray(arrByRealName) Then
        Dim nRowsRealName As Long
        nRowsRealName = UBound(arrByRealName, 1) - LBound(arrByRealName, 1) + 1
        okRealName = (nRowsRealName = 3)
        detailRealName = "got " & nRowsRealName & " rows"
    Else
        detailRealName = "got " & TypeName(arrByRealName) & " = " & CStr(arrByRealName)
    End If
    Report "datalog host: a table reached through a Name-Manager ALIAS is still queryable by its own real ListObject name", _
           okRealName, detailRealName

    Dim arrByAliasName As Variant
    arrByAliasName = VLA_Datalog.DATALOG("(headless) (query vladatalogaliastest)", aliasRange)
    Dim okAliasRefused As Boolean, detailAliasRefused As String
    If VarType(arrByAliasName) = vbString Then
        Dim sAliasName As String
        sAliasName = CStr(arrByAliasName)
        okAliasRefused = (Left$(sAliasName, 9) = "#DATALOG!")
        detailAliasRefused = "got """ & sAliasName & """"
    Else
        detailAliasRefused = "got a " & TypeName(arrByAliasName) & ", not a #DATALOG! error string"
    End If
    Report "datalog host: the SAME table is refused under its Name-Manager alias's own name, not silently matched", _
           okAliasRefused, detailAliasRefused

    ' Scenario 3's own table: three columns (Name, Dept, Salary) - wide
    ' enough that a 2-of-3 contiguous slice is a genuinely NARROWER
    ' selection than the whole table, not a degenerate single-column case.
    ws.Range("F1:H1").Value = Array("Name", "Dept", "Salary")
    ws.Range("F2:H2").Value = Array("alice", "eng", "100")
    ws.Range("F3:H3").Value = Array("bob", "sales", "200")
    Dim loWide As ListObject
    Set loWide = ws.ListObjects.Add(xlSrcRange, ws.Range("F1:H3"), , xlYes)
    loWide.Name = "WideHostTest1"

    Dim arrNarrow As Variant
    arrNarrow = VLA_Datalog.DATALOG("(headless) (query widehosttest1)", ws.Range("F1:G3"))
    Dim okNarrow As Boolean, detailNarrow As String
    If IsArray(arrNarrow) Then
        Dim nRowsNarrow As Long, nColsNarrow As Long
        nRowsNarrow = UBound(arrNarrow, 1) - LBound(arrNarrow, 1) + 1
        nColsNarrow = UBound(arrNarrow, 2) - LBound(arrNarrow, 2) + 1
        okNarrow = (nRowsNarrow = 2) And (nColsNarrow = 2) And _
                   (ResultCellIs(arrNarrow, LBound(arrNarrow, 1), LBound(arrNarrow, 2), "alice") Or _
                    ResultCellIs(arrNarrow, LBound(arrNarrow, 1), LBound(arrNarrow, 2), "bob"))
        detailNarrow = "got " & nRowsNarrow & " rows x " & nColsNarrow & " cols"
    Else
        detailNarrow = "got " & TypeName(arrNarrow) & " = " & CStr(arrNarrow)
    End If
    Report "datalog host: a contiguous 2-of-3-column slice of a live Table narrows arity to 2 (Salary excluded), same predicate name", _
           okNarrow, detailNarrow

    Dim arrNoncontig As Variant
    arrNoncontig = VLA_Datalog.DATALOG("(headless) (query widehosttest1)", _
        Union(ws.Range("F1:F3"), ws.Range("H1:H3")))
    Dim okNoncontigRefused As Boolean, detailNoncontigRefused As String
    If VarType(arrNoncontig) = vbString Then
        Dim sNoncontig As String
        sNoncontig = CStr(arrNoncontig)
        okNoncontigRefused = (Left$(sNoncontig, 9) = "#DATALOG!")
        detailNoncontigRefused = "got """ & sNoncontig & """"
    Else
        detailNoncontigRefused = "got a " & TypeName(arrNoncontig) & ", not a #DATALOG! error string"
    End If
    Report "datalog host: a non-contiguous (Ctrl-selected) column pick is refused, not silently guessed", _
           okNoncontigRefused, detailNoncontigRefused

    ' Scenario 4, DATALOG.5: a mixed-case Table name AND mixed-case
    ' header (Staffing/Salary, the item's own words) queried through
    ' all-lowercase keyed-atom rule text - proving both the predicate
    ' name and the column name fold invariantly through VLA_Identity.Fold
    ' (SD-8), the identical discipline SQL.1 already proved for column
    ' names, now exercised for DATALOG's own keyed-atom syntax.
    ' Necessarily host-required, like DATALOG.3's own column-scoping: a
    ' plain 2D array has no live header row of its own to fold against
    ' - TestDatalogKeyedAtoms (VLA_Datalog's own pure suite) exercises
    ' the desugaring mechanism itself instead, off a hand-folded
    ' headerMap the test builds by hand.
    ws.Range("J1:L1").Value = Array("Name", "Salary", "Dept")
    ws.Range("J2:L2").Value = Array("alice", 90000, "eng")
    ws.Range("J3:L3").Value = Array("bob", 70000, "sales")
    ws.Range("J4:L4").Value = Array("carol", 95000, "eng")
    Dim loStaffing As ListObject
    Set loStaffing = ws.ListObjects.Add(xlSrcRange, ws.Range("J1:L4"), , xlYes)
    loStaffing.Name = "StaffingHostTest1"

    Dim arrRich As Variant
    arrRich = VLA_Datalog.DATALOG( _
        "(headless)" & _
        " (rule (rich X) (staffinghosttest1 (name X) (salary S)) (> S 80000))" & _
        " (query rich)", _
        loStaffing.Range)
    Dim okRich As Boolean, detailRich As String
    If IsArray(arrRich) Then
        Dim nRowsRich As Long
        nRowsRich = UBound(arrRich, 1) - LBound(arrRich, 1) + 1
        okRich = (nRowsRich = 2)
        detailRich = "got " & nRowsRich & " rows"
    Else
        detailRich = "got " & TypeName(arrRich) & " = " & CStr(arrRich)
    End If
    Report "datalog host: a keyed atom against a mixed-case Table/header (Staffing/Salary) resolves through Fold, matching exactly the 2 staffers over 80000", _
           okRich, detailRich

    ' Scenario 5, DATALOG.8: a Table whose one data row is blank is still a
    ' DEFINED relation - DATALOG() registers one per table argument, rows or
    ' none - so the undefined-predicate refusal must never fire on it.
    ' PROLOG.22's own trap, pinned for DATALOG: a ListObject over a header
    ' and one blank row, which RelFromRange reads as zero tuples.
    ws.Range("N1").Value = "Name"
    Dim loEmpty As ListObject
    Set loEmpty = ws.ListObjects.Add(xlSrcRange, ws.Range("N1:N2"), , xlYes)
    loEmpty.Name = "EmptyDatalogHostTest1"

    Dim arrNotEmpty As Variant
    arrNotEmpty = VLA_Datalog.DATALOG( _
        "(headless) (rule (ok X) (personhosttest1 X) (not (emptydataloghosttest1 X))) (query ok)", _
        loPeople.Range, loEmpty.Range)
    Report "datalog.8 host: (not ...) over a Table with no data rows answers all 3 people, with no undefined-predicate refusal", _
           ResultRowCount(arrNotEmpty) = 3, "got: " & ResultDescribe(arrNotEmpty)

    Dim arrEmptyRule As Variant
    arrEmptyRule = VLA_Datalog.DATALOG("(rule (who X) (emptydataloghosttest1 X)) (query who)", loEmpty.Range)
    Report "datalog.8 host: a rule reading a Table with no data rows spills its header and nothing under it", _
           ResultRowCount(arrEmptyRule) = 1 And ResultCellIs(arrEmptyRule, 1, 1, "X"), "got: " & ResultDescribe(arrEmptyRule)

    On Error Resume Next
    ThisWorkbook.Names("VlaDatalogAliasTest").Delete
    On Error GoTo 0
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
    prior.Activate
End Sub

' DATALOG.15: a spill's header row as Value2 reads it, (1 To 1, 1 To n),
' built from the cells given - so a pure pin can hand SpillHeaderCheck a
' number, a blank or an error value without a live sheet.
Private Function SpillHdr(ParamArray cells() As Variant) As Variant
    Dim n As Long
    n = UBound(cells) - LBound(cells) + 1
    Dim h() As Variant
    ReDim h(1 To 1, 1 To n)
    Dim i As Long
    For i = 1 To n
        h(1, i) = cells(LBound(cells) + i - 1)
    Next i
    SpillHdr = h
End Function

' DATALOG.15: the words a table-argument reason becomes, captured before
' On Error GoTo 0 can reset them.
Private Function SpillRefusalText(ByVal reason As String, ByVal detail As String) As String
    Dim d As String
    On Error Resume Next
    Err.Clear
    VLA_Relation.RaiseTableArgRefusal reason, detail
    d = Err.Description
    On Error GoTo 0
    SpillRefusalText = d
End Function

' DATALOG.15: the host-free half of reading a spilled range as a table.
' SpillHeaderCheck is what keeps a spill whose first row is DATA (a
' SEQUENCE, a FILTER over a Table's body) from losing that row silently:
' a number, a blank or an error in the first row is refused by name, and
' an error value is checked without CStr, which would raise 13 on it.
' SpillRefersToMatches accepts each spelling of the spill reference Excel
' may hand back in a Name's RefersTo. And an argument that is an error
' value - what N2# gives when N2 is not spilling - is refused as one by
' all three engines, while a plain value keeps today's "not a range".
Private Sub TestSpillHeaders()
    Dim why As String, badCol As Long, otherCol As Long

    why = VLA_Relation.SpillHeaderCheck(SpillHdr("Name", "Shift"), badCol, otherCol)
    Report "datalog.15: a header row of two names passes", why = "" And badCol = 0, "got '" & why & "' at " & badCol
    why = VLA_Relation.SpillHeaderCheck(SpillHdr("Only"), badCol, otherCol)
    Report "datalog.15: a one-column header row passes", why = "", "got '" & why & "'"
    why = VLA_Relation.SpillHeaderCheck(SpillHdr("Name", Empty), badCol, otherCol)
    Report "datalog.15: a blank header is refused, at its column", why = "spill-header-blank" And badCol = 2, "got '" & why & "' at " & badCol
    why = VLA_Relation.SpillHeaderCheck(SpillHdr("Name", "   "), badCol, otherCol)
    Report "datalog.15: a header of only spaces is blank", why = "spill-header-blank" And badCol = 2, "got '" & why & "' at " & badCol
    why = VLA_Relation.SpillHeaderCheck(SpillHdr(1, "Shift"), badCol, otherCol)
    Report "datalog.15: a number in the first row is refused (a SEQUENCE's first row is data)", why = "spill-header-not-text" And badCol = 1, "got '" & why & "' at " & badCol
    why = VLA_Relation.SpillHeaderCheck(SpillHdr("Name", True), badCol, otherCol)
    Report "datalog.15: TRUE in the first row is refused", why = "spill-header-not-text" And badCol = 2, "got '" & why & "' at " & badCol
    why = VLA_Relation.SpillHeaderCheck(SpillHdr("Name", CVErr(2042)), badCol, otherCol)
    Report "datalog.15: an error in the first row is refused, not raised on", why = "spill-header-not-text" And badCol = 2, "got '" & why & "' at " & badCol
    why = VLA_Relation.SpillHeaderCheck(SpillHdr("Name", "Shift", "NAME"), badCol, otherCol)
    Report "datalog.15: two headers alike after folding are refused, naming both columns", _
           why = "spill-header-duplicate" And badCol = 3 And otherCol = 1, "got '" & why & "' at " & badCol & "/" & otherCol

    Report "datalog.15: =Sheet1!$N$2# is the spill reference", _
           VLA_Relation.SpillRefersToMatches("=Sheet1!$N$2#", "Sheet1", "$N$2"), ""
    Report "datalog.15: a quoted sheet name matches", _
           VLA_Relation.SpillRefersToMatches("='Q1 Plan'!$N$2#", "Q1 Plan", "$N$2"), ""
    Report "datalog.15: the file format's _xlfn.ANCHORARRAY spelling matches", _
           VLA_Relation.SpillRefersToMatches("=_xlfn.ANCHORARRAY(Sheet1!$N$2)", "Sheet1", "$N$2"), ""
    Report "datalog.15: ANCHORARRAY without _xlfn matches", _
           VLA_Relation.SpillRefersToMatches("=ANCHORARRAY(Sheet1!$N$2)", "Sheet1", "$N$2"), ""
    Report "datalog.15: a reference without $ signs matches", _
           VLA_Relation.SpillRefersToMatches("=Sheet1!N2#", "Sheet1", "$N$2"), ""
    Report "datalog.15: the sheet and cell compare case-insensitively", _
           VLA_Relation.SpillRefersToMatches("=sheet1!$n$2#", "Sheet1", "$N$2"), ""
    Report "datalog.15: a sheet name holding a quote matches", _
           VLA_Relation.SpillRefersToMatches("='It''s'!$A$1#", "It's", "$A$1"), ""
    Report "datalog.15: a static range over the same cells is not the spill reference", _
           Not VLA_Relation.SpillRefersToMatches("=Sheet1!$N$2:$P$9", "Sheet1", "$N$2"), ""
    Report "datalog.15: the same cell on another sheet does not match", _
           Not VLA_Relation.SpillRefersToMatches("=Sheet2!$N$2#", "Sheet1", "$N$2"), ""
    Report "datalog.15: another anchor does not match", _
           Not VLA_Relation.SpillRefersToMatches("=Sheet1!$N$3#", "Sheet1", "$N$2"), ""
    Report "datalog.15: a reference with no sheet does not match", _
           Not VLA_Relation.SpillRefersToMatches("=$N$2#", "Sheet1", "$N$2"), ""

    Report "datalog.15: the RefersTo a refusal tells the user to type", _
           VLA_Relation.SpillRefersToText("Sheet1", "$N$2") = "='Sheet1'!$N$2#", "got " & VLA_Relation.SpillRefersToText("Sheet1", "$N$2")
    Report "datalog.15: ...with a quote in the sheet name doubled", _
           VLA_Relation.SpillRefersToText("It's", "$A$1") = "='It''s'!$A$1#", "got " & VLA_Relation.SpillRefersToText("It's", "$A$1")
    Report "datalog.15: ...and what it tells the user to type is what the match accepts", _
           VLA_Relation.SpillRefersToMatches(VLA_Relation.SpillRefersToText("Q1 Plan", "$N$2"), "Q1 Plan", "$N$2"), ""

    Dim ok As Boolean, reason As String, detail As String, nm As String
    nm = VLA_Relation.TableArgResolve(CVErr(2023), ok, reason, detail)
    Report "datalog.15: an error value is its own reason, not 'not a range'", _
           (Not ok) And reason = "error-value", "got ok=" & ok & " reason=" & reason
    nm = VLA_Relation.TableArgResolve(1, ok, reason)
    Report "datalog.15: a plain value is still 'not a range' (and detail stays optional)", _
           (Not ok) And reason = "not-a-range", "got ok=" & ok & " reason=" & reason

    Dim result As Variant
    result = VLA_Datalog.DATALOG("(query p)", CVErr(2023))
    Report "datalog.15: DATALOG over an error value says it is one", _
           ResultTextStartsWith(result, "#DATALOG!") And InStr(1, ResultDescribe(result), "is an error value", vbTextCompare) > 0, "got: " & ResultDescribe(result)
    result = VLA_Sql.SQL("SELECT * FROM t", CVErr(2023))
    Report "datalog.15: SQL over an error value says it is one", _
           ResultTextStartsWith(result, "#SQL!") And InStr(1, ResultDescribe(result), "is an error value", vbTextCompare) > 0, "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (p X))", CVErr(2023))
    Report "datalog.15: PROLOG over an error value says it is one", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, ResultDescribe(result), "is an error value", vbTextCompare) > 0, "got: " & ResultDescribe(result)

    Dim d As String
    d = SpillRefusalText("spill-needs-a-name", "='Sheet1'!$N$2#")
    Report "datalog.15: a spill with no name is told the RefersTo to type", _
           InStr(1, d, "Name Manager", vbTextCompare) > 0 And InStr(1, d, "='Sheet1'!$N$2#", vbBinaryCompare) > 0, "got: " & d
    d = SpillRefusalText("spill-two-names", "roster, schedule")
    Report "datalog.15: a spill with two names names them", InStr(1, d, "(roster, schedule)", vbBinaryCompare) > 0, "got: " & d
    d = SpillRefusalText("spill-header-blank", "O2")
    Report "datalog.15: a blank header names its cell", InStr(1, d, "cell O2 is blank", vbTextCompare) > 0, "got: " & d
    d = SpillRefusalText("spill-header-not-text", "N2")
    Report "datalog.15: a data-looking first row names its cell", InStr(1, d, "cell N2 holds a number", vbTextCompare) > 0, "got: " & d
    d = SpillRefusalText("spill-header-duplicate", "N2 and P2")
    Report "datalog.15: duplicate headers name both cells", InStr(1, d, "cells N2 and P2 hold the same name", vbTextCompare) > 0, "got: " & d
    d = SpillRefusalText("zzq-no-such-reason", "")
    Report "datalog.15: a reason nobody words is refused as a bug, never a silent blank name", _
           InStr(1, d, "zzq-no-such-reason", vbBinaryCompare) > 0 And InStr(1, d, "bug", vbTextCompare) > 0, "got: " & d
End Sub

' DATALOG.15, host-required: a real dynamic-array spill read as a table.
' Each spill is an array constant entered with Formula2 (late-bound, so
' this module still compiles on an Excel without dynamic arrays), named
' through the workbook's Names with the spill reference itself, and read
' by the three engines' real worksheet functions from VBA - the Range a
' formula like =DATALOG(..., Schedule) hands them is the same object.
' Pinned unchanged beside it: a part of a spill, and a plain named range,
' both still read with their first row as a fact.
Private Sub TestSpillHostTable()
    Dim prior As Worksheet
    Set prior = ActiveSheet
    Dim nmList As Variant, nmItem As Variant
    nmList = Array("VlaSpillSched", "VlaSpillSched2", "VlaSpillEmpty", "VlaSpillNums", "VlaSpillDup", "VlaSpillPart", "VlaSpillPlain")
    On Error Resume Next
    For Each nmItem In nmList
        ActiveWorkbook.Names(CStr(nmItem)).Delete
    Next nmItem
    On Error GoTo 0

    VlaEnsureSheet "VlaSpillHostSheet"
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets("VlaSpillHostSheet")
    ws.Activate
    Do While ws.ListObjects.Count > 0
        ws.ListObjects(1).Delete
    Loop
    ws.Cells.Clear

    Dim q As String
    q = Chr$(34)
    Dim cellSched As Object, cellEmpty As Object, cellNums As Object, cellDup As Object
    Set cellSched = ws.Range("A1")
    Set cellEmpty = ws.Range("D1")
    Set cellNums = ws.Range("G1")
    Set cellDup = ws.Range("J1")
    ' Formula2 raises 438 on an Excel without dynamic arrays; the Nothing
    ' check below then reports it rather than the run dying here.
    On Error Resume Next
    cellSched.Formula2 = "={" & q & "Name" & q & "," & q & "Shift" & q & ";" & q & "Ann" & q & "," & q & "Mon" & q & ";" & _
                         q & "Bob" & q & "," & q & "Tue" & q & ";" & q & "Ann" & q & "," & q & "Wed" & q & "}"
    cellEmpty.Formula2 = "={" & q & "Name" & q & "," & q & "Shift" & q & "}"
    cellNums.Formula2 = "={1,2;3,4}"
    cellDup.Formula2 = "={" & q & "Name" & q & "," & q & "name" & q & ";" & q & "a" & q & "," & q & "b" & q & "}"
    On Error GoTo 0
    ws.Range("P1:Q1").Value = Array("Name", "Shift")
    ws.Range("P2:Q2").Value = Array("Cy", "Thu")
    ws.Range("P3:Q3").Value = Array("Di", "Fri")
    ws.Calculate

    Dim spillSched As Object, spillEmpty As Object, spillNums As Object, spillDup As Object
    On Error Resume Next
    Set spillSched = cellSched.SpillingToRange
    Set spillEmpty = cellEmpty.SpillingToRange
    Set spillNums = cellNums.SpillingToRange
    Set spillDup = cellDup.SpillingToRange
    On Error GoTo 0
    ' Is Nothing tested alone, never joined to .Address by And: VBA's And
    ' evaluates both sides, and .Address on Nothing raises 91.
    Dim spilled As Boolean
    spilled = Not (spillSched Is Nothing Or spillEmpty Is Nothing Or spillNums Is Nothing Or spillDup Is Nothing)
    Report "datalog.15 host: the four fixture formulas spill (an Excel with dynamic arrays)", spilled, "a SpillingToRange was Nothing"
    If Not spilled Then GoTo cleanup
    Report "datalog.15 host: the fixture spills where expected (A1:B4, D1:E1, G1:H2, J1:K2)", _
           spillSched.Address = "$A$1:$B$4" And spillEmpty.Address = "$D$1:$E$1" And spillNums.Address = "$G$1:$H$2" And spillDup.Address = "$J$1:$K$2", _
           "got " & spillSched.Address & " " & spillEmpty.Address & " " & spillNums.Address & " " & spillDup.Address

    Dim anchor As Object
    Set anchor = VLA_Relation.SpillAnchorOf(spillSched)
    Dim anchorAddr As String
    If Not anchor Is Nothing Then anchorAddr = anchor.Address
    Report "datalog.15 host: a spill's whole extent is recognised, with A1 as its anchor", anchorAddr = "$A$1", "got '" & anchorAddr & "'"
    Report "datalog.15 host: a part of a spill is not a spill", VLA_Relation.SpillAnchorOf(ws.Range("A1:B2")) Is Nothing, ""
    Report "datalog.15 host: plain cells are not a spill", VLA_Relation.SpillAnchorOf(ws.Range("P1:Q3")) Is Nothing, ""

    Dim result As Variant
    result = VLA_Datalog.DATALOG("(rule (who P) (vlaspillsched (name P))) (query who)", spillSched)
    Report "datalog.15 host: an unnamed spill is refused, told the RefersTo to type", _
           ResultTextStartsWith(result, "#DATALOG!") And InStr(1, ResultDescribe(result), "'VlaSpillHostSheet'!$A$1#", vbTextCompare) > 0, _
           "got: " & ResultDescribe(result)

    Dim addErr As String
    On Error Resume Next
    Err.Clear
    ws.Parent.Names.Add Name:="VlaSpillSched", RefersTo:="='VlaSpillHostSheet'!$A$1#"
    ws.Parent.Names.Add Name:="VlaSpillEmpty", RefersTo:="='VlaSpillHostSheet'!$D$1#"
    ws.Parent.Names.Add Name:="VlaSpillNums", RefersTo:="='VlaSpillHostSheet'!$G$1#"
    ws.Parent.Names.Add Name:="VlaSpillDup", RefersTo:="='VlaSpillHostSheet'!$J$1#"
    ws.Parent.Names.Add Name:="VlaSpillPart", RefersTo:="='VlaSpillHostSheet'!$A$1:$B$2"
    ws.Parent.Names.Add Name:="VlaSpillPlain", RefersTo:="='VlaSpillHostSheet'!$P$1:$Q$3"
    addErr = Err.Description
    On Error GoTo 0
    Report "datalog.15 host: Name Manager accepts a name referring to a spill", addErr = "", "got: " & addErr

    result = VLA_Datalog.DATALOG("(headless) (query vlaspillsched)", spillSched)
    Report "datalog.15 host: a named spill's first row is its headers, not a fact (3 rows, not 4)", _
           ResultRowCount(result) = 3, "got: " & ResultDescribe(result)
    result = VLA_Datalog.DATALOG("(rule (who P) (vlaspillsched (name P))) (query who)", spillSched)
    Report "datalog.15 host: a keyed atom reads the spill's header by name (Ann, Bob)", _
           ResultRowCount(result) = 3 And ResultCellIs(result, 1, 1, "P"), "got: " & ResultDescribe(result)
    result = VLA_Datalog.DATALOG("(rule (who P) (vlaspillempty (name P))) (query who)", spillEmpty)
    Report "datalog.15 host: a spill of its header row alone is a defined relation with no rows", _
           ResultRowCount(result) = 1 And ResultCellIs(result, 1, 1, "P"), "got: " & ResultDescribe(result)
    result = VLA_Datalog.DATALOG("(headless) (query vlaspillnums)", spillNums)
    Report "datalog.15 host: a spill whose first row is numbers is refused, naming G1", _
           ResultTextStartsWith(result, "#DATALOG!") And InStr(1, ResultDescribe(result), "cell G1 holds a number", vbTextCompare) > 0, _
           "got: " & ResultDescribe(result)
    result = VLA_Datalog.DATALOG("(headless) (query vlaspilldup)", spillDup)
    Report "datalog.15 host: a spill with two headers alike is refused, naming J1 and K1", _
           ResultTextStartsWith(result, "#DATALOG!") And InStr(1, ResultDescribe(result), "cells J1 and K1", vbTextCompare) > 0, _
           "got: " & ResultDescribe(result)

    result = VLA_Sql.SQL("SELECT Name FROM vlaspillsched WHERE Shift = 'Tue'", spillSched)
    Report "datalog.15 host: SQL reads the same spill, by its header names (Bob)", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 1, 1, "Name") And ResultCellIs(result, 2, 1, "Bob"), _
           "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (vlaspillsched (name P)))", spillSched)
    Report "datalog.15 host: PROLOG reads the same spill, keyed by header (3 rows under the header)", _
           ResultRowCount(result) = 4, "got: " & ResultDescribe(result)

    result = VLA_Datalog.DATALOG("(headless) (query vlaspillpart)", ws.Range("A1:B2"))
    Report "datalog.15 host: UNCHANGED - a named part of a spill reads its first row as a fact (2 rows)", _
           ResultRowCount(result) = 2, "got: " & ResultDescribe(result)
    result = VLA_Datalog.DATALOG("(headless) (query vlaspillplain)", ws.Range("P1:Q3"))
    Report "datalog.15 host: UNCHANGED - a plain named range reads its first row as a fact (3 rows)", _
           ResultRowCount(result) = 3, "got: " & ResultDescribe(result)

    On Error Resume Next
    Err.Clear
    ws.Parent.Names.Add Name:="VlaSpillSched2", RefersTo:="='VlaSpillHostSheet'!$A$1#"
    addErr = Err.Description
    On Error GoTo 0
    result = VLA_Datalog.DATALOG("(headless) (query vlaspillsched)", spillSched)
    Report "datalog.15 host: a spill with two different names is refused, naming both", _
           addErr = "" And ResultTextStartsWith(result, "#DATALOG!") And InStr(1, ResultDescribe(result), "vlaspillsched, vlaspillsched2", vbTextCompare) > 0, _
           "add: '" & addErr & "' got: " & ResultDescribe(result)

    ws.Range("S1").Value = "x"
    Dim notSpilling As Variant
    notSpilling = ws.Evaluate("S1#")
    result = VLA_Datalog.DATALOG("(headless) (query p)", notSpilling)
    Report "datalog.15 host: a # reference to a cell that is not spilling is refused as an error value", _
           IsError(notSpilling) And ResultTextStartsWith(result, "#DATALOG!") And InStr(1, ResultDescribe(result), "is an error value", vbTextCompare) > 0, _
           "Evaluate gave " & TypeName(notSpilling) & "; got: " & ResultDescribe(result)

cleanup:
    On Error Resume Next
    For Each nmItem In nmList
        ActiveWorkbook.Names(CStr(nmItem)).Delete
    Next nmItem
    On Error GoTo 0
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
    prior.Activate
End Sub

' ---------------------------------------------------------------------
'  OPTIMIZE.1, step 1: the memo key's encoder and its framing.
'
'  Three things are pinned here, and the FIRST two have known answers
'  rather than merely being self-consistent:
'
'  1. WTF-8, against SHA-256. Each digest below is the digest of the
'     byte sequence the design names, verified against .NET's own
'     SHA-256 over those same bytes by tools/check_hash_twin.ps1. So a
'     shared bug in VLA_Digest.bas and in this pin cannot cancel out -
'     the same discipline SEC.11's FIPS vectors already establish.
'
'  2. The framing, byte for byte, as one known-answer key. Fifty bytes
'     whose every field can be read by hand (the comment beside it
'     decodes them), so the framing is a SPEC and not whatever the
'     implementation happens to do.
'
'  3. The collisions. One pin per pair the framing prevents, asserted
'     as two DIFFERENT digests. These need no baseline: they are
'     properties, and a framing that lost one would fail here even if
'     every known answer above still matched.
'
'  Why here and not in VLA_Tests.bas beside TestSec11Digest: the
'  encoder exists for OPTIMIZE's memo, OPTIMIZE.1 owns the proof, and
'  TestDSLs is this tranche's one command - this module's own header
'  note about not lengthening VlaSelfTest's everyday run.
' ---------------------------------------------------------------------

' A code unit by number, without relying on how VBA reads an
' unsuffixed &H literal above 32767: ChrW's own documented range is
' -32768 to 65535, and this says which half is meant.
Private Function KeyCharW(ByVal code As Long) As String
    If code > 32767 Then
        KeyCharW = ChrW(code - 65536)
    Else
        KeyCharW = ChrW(code)
    End If
End Function

Private Function KeyByteText(ByRef b() As Byte, ByVal n As Long) As String
    Dim s As String, i As Long
    For i = 0 To n - 1
        If i > 0 Then s = s & " "
        s = s & Right$("0" & Hex$(b(i)), 2)
    Next i
    KeyByteText = s
End Function

' The key of a run of cells, in order - the collision pins' own shape.
Private Function KeyOfCells(ParamArray cells() As Variant) As String
    Dim buf() As Byte
    Dim m As Long
    Dim i As Long
    VLA_Digest.VlaKeyBegin buf, m
    For i = LBound(cells) To UBound(cells)
        VLA_Digest.VlaKeyAddCell buf, m, cells(i)
    Next i
    KeyOfCells = VLA_Digest.VlaKeyHex(buf, m)
End Function

' The key of a run of text fields, in order.
Private Function KeyOfTexts(ParamArray texts() As Variant) As String
    Dim buf() As Byte
    Dim m As Long
    Dim i As Long
    VLA_Digest.VlaKeyBegin buf, m
    For i = LBound(texts) To UBound(texts)
        VLA_Digest.VlaKeyAddText buf, m, CStr(texts(i))
    Next i
    KeyOfTexts = VLA_Digest.VlaKeyHex(buf, m)
End Function

' The key of a run of counts, in order.
Private Function KeyOfCounts(ParamArray nums() As Variant) As String
    Dim buf() As Byte
    Dim m As Long
    Dim i As Long
    VLA_Digest.VlaKeyBegin buf, m
    For i = LBound(nums) To UBound(nums)
        VLA_Digest.VlaKeyAddCount buf, m, CLng(nums(i))
    Next i
    KeyOfCounts = VLA_Digest.VlaKeyHex(buf, m)
End Function

Private Sub TestOptimizeKey()
    Dim b() As Byte
    Dim n As Long
    Dim buf() As Byte
    Dim m As Long
    Dim hexA As String

    ' --- 1. WTF-8, byte for byte, with SHA-256 known answers ---------

    n = VLA_Digest.VlaWtf8Encode(KeyCharW(&HE9), b)
    Report "optimize key: e-acute encodes as c3 a9", _
           KeyByteText(b, n) = "C3 A9", "got " & KeyByteText(b, n)
    Report "optimize key: SHA-256 over e-acute's two bytes", _
           VLA_Digest.VlaSha256Hex(b, n) = "4A99557E4033C3539DE2EB65472017CAD5F9557F7A0625A09F1C3F6E2BA69C4C", _
           "got " & VLA_Digest.VlaSha256Hex(b, n)

    ' One astral character is ONE code point of four bytes, from a
    ' surrogate PAIR - the case a naive per-code-unit encoder gets
    ' wrong in the other direction from the lone-surrogate case below.
    n = VLA_Digest.VlaWtf8Encode(KeyCharW(&HD83D&) & KeyCharW(&HDE00&), b)
    Report "optimize key: one astral character is four bytes (f0 9f 98 80)", _
           KeyByteText(b, n) = "F0 9F 98 80", "got " & KeyByteText(b, n)
    Report "optimize key: SHA-256 over the astral character's four bytes", _
           VLA_Digest.VlaSha256Hex(b, n) = "F0443A342C5EF54783A111B51BA56C938E474C32324D90C3A60C9C8E3A37E2D9", _
           "got " & VLA_Digest.VlaSha256Hex(b, n)

    n = VLA_Digest.VlaWtf8Encode("Zo" & KeyCharW(&HEB), b)
    Report "optimize key: SHA-256 over Zo-diaeresis (5a 6f c3 ab)", _
           VLA_Digest.VlaSha256Hex(b, n) = "C6A12698582FC1104EA24107A2D7268145FF06EF859707729D01FD060897F067", _
           "got " & VLA_Digest.VlaSha256Hex(b, n)

    ' THE reason this encoder exists. VLA_Loader.VlaUtf8Encode writes
    ' each of the next two as U+FFFD, which is right for a file and a
    ' collision for a key.
    n = VLA_Digest.VlaWtf8Encode(KeyCharW(&HD800&), b)
    Report "optimize key: a lone HIGH surrogate keeps its own bytes (ed a0 80)", _
           KeyByteText(b, n) = "ED A0 80", "got " & KeyByteText(b, n)
    Report "optimize key: SHA-256 over a lone high surrogate", _
           VLA_Digest.VlaSha256Hex(b, n) = "91A681B998555FB475479817B126C94E57E52011FA1842C5D188795A4A05226B", _
           "got " & VLA_Digest.VlaSha256Hex(b, n)

    n = VLA_Digest.VlaWtf8Encode(KeyCharW(&HDC00&), b)
    Report "optimize key: a lone LOW surrogate keeps its own bytes (ed b0 80)", _
           KeyByteText(b, n) = "ED B0 80", "got " & KeyByteText(b, n)
    Report "optimize key: SHA-256 over a lone low surrogate", _
           VLA_Digest.VlaSha256Hex(b, n) = "B2D612A08BEC1F41120EBD961F62EF19678375B5788C70D3F8F4C02E345ED412", _
           "got " & VLA_Digest.VlaSha256Hex(b, n)

    n = VLA_Digest.VlaWtf8Encode(KeyCharW(&HFFFD&), b)
    Report "optimize key: U+FFFD itself is ef bf bd", _
           KeyByteText(b, n) = "EF BF BD", "got " & KeyByteText(b, n)
    Report "optimize key: SHA-256 over U+FFFD", _
           VLA_Digest.VlaSha256Hex(b, n) = "83D544CCC223C057D2BF80D3F2A32982C32C3C0DB8E2674820DA5064783FB097", _
           "got " & VLA_Digest.VlaSha256Hex(b, n)

    ' The collision itself, as two known answers rather than as an
    ' inequality: these are the two strings the memo must never confuse.
    n = VLA_Digest.VlaWtf8Encode("a" & KeyCharW(&HD800&), b)
    Report "optimize key: SHA-256 over a + a lone high surrogate", _
           VLA_Digest.VlaSha256Hex(b, n) = "25819B9B43D499092EB2BE7B6F27AE28439EEE434CEA4490191AB4CCB8F3409C", _
           "got " & VLA_Digest.VlaSha256Hex(b, n)
    n = VLA_Digest.VlaWtf8Encode("a" & KeyCharW(&HFFFD&), b)
    Report "optimize key: SHA-256 over a + U+FFFD", _
           VLA_Digest.VlaSha256Hex(b, n) = "51D277510BA4BF97B25F12D38513C1B620A2A33FC83B3BEEEB0DD971BF429E6D", _
           "got " & VLA_Digest.VlaSha256Hex(b, n)

    n = VLA_Digest.VlaWtf8Encode("", b)
    Report "optimize key: the empty string encodes to no bytes at all", n = 0, "got " & n

    ' Plain ASCII is unchanged, so every key an ASCII-only workbook
    ' makes is what a plain UTF-8 encoder would have made.
    n = VLA_Digest.VlaWtf8Encode("abc", b)
    Report "optimize key: ASCII is byte-identical to UTF-8", _
           KeyByteText(b, n) = "61 62 63", "got " & KeyByteText(b, n)

    ' --- 2. a Double's eight bytes ------------------------------------

    n = VLA_Digest.VlaDoubleBytes(1#, b)
    Report "optimize key: 1 is 00 00 00 00 00 00 F0 3F", _
           n = 8 And KeyByteText(b, n) = "00 00 00 00 00 00 F0 3F", "got " & KeyByteText(b, n)
    n = VLA_Digest.VlaDoubleBytes(0#, b)
    Report "optimize key: 0 is eight zero bytes", _
           KeyByteText(b, n) = "00 00 00 00 00 00 00 00", "got " & KeyByteText(b, n)
    n = VLA_Digest.VlaDoubleBytes(-1#, b)
    Report "optimize key: -1 differs from 1 only in its sign bit", _
           KeyByteText(b, n) = "00 00 00 00 00 00 F0 BF", "got " & KeyByteText(b, n)
    n = VLA_Digest.VlaDoubleBytes(0.5, b)
    Report "optimize key: 0.5 is 00 00 00 00 00 00 E0 3F", _
           KeyByteText(b, n) = "00 00 00 00 00 00 E0 3F", "got " & KeyByteText(b, n)
    n = VLA_Digest.VlaDoubleBytes(1000000#, b)
    Report "optimize key: a million is 00 00 00 00 80 84 2E 41", _
           KeyByteText(b, n) = "00 00 00 00 80 84 2E 41", "got " & KeyByteText(b, n)

    ' --- 3. the framing, byte for byte --------------------------------
    '
    ' text "a"    01 00000001 61
    ' count 2     05 00000004 00000002
    ' number 1    02 00000008 0000000000 00F03F
    ' text "1"    01 00000001 31
    ' blank       00 00000000
    ' text ""     01 00000000
    ' TRUE        03 00000001 01
    '             = fifty bytes in all
    VLA_Digest.VlaKeyBegin buf, m
    VLA_Digest.VlaKeyAddText buf, m, "a"
    VLA_Digest.VlaKeyAddCount buf, m, 2
    VLA_Digest.VlaKeyAddCell buf, m, 1#
    VLA_Digest.VlaKeyAddCell buf, m, "1"
    VLA_Digest.VlaKeyAddCell buf, m, Empty
    VLA_Digest.VlaKeyAddCell buf, m, ""
    VLA_Digest.VlaKeyAddCell buf, m, True
    Report "optimize key: the sample key frames to exactly fifty bytes", m = 50, "got " & m
    Report "optimize key: the sample key's first field is text, length 1, 'a'", _
           KeyByteText(buf, 6) = "01 00 00 00 01 61", "got " & KeyByteText(buf, 6)
    Report "optimize key: the sample key's digest", _
           VLA_Digest.VlaKeyHex(buf, m) = "E4BF92CB66C89E6B0566D0C1FB8885AE0FA5423262BBFD7E694CF106641EFA53", _
           "got " & VLA_Digest.VlaKeyHex(buf, m)

    ' An empty key is the empty message's own digest - the same value
    ' SEC.11's own 'empty' vector holds, reached through this framing.
    VLA_Digest.VlaKeyBegin buf, m
    Report "optimize key: a key with no fields hashes the empty message", _
           m = 0 And VLA_Digest.VlaKeyHex(buf, m) = "E3B0C44298FC1C149AFBF4C8996FB92427AE41E4649B934CA495991B7852B855", _
           "got " & VLA_Digest.VlaKeyHex(buf, m)

    ' VlaKeyBegin must really restart, or a second key built into a
    ' reused buffer would carry the first one's bytes.
    VLA_Digest.VlaKeyBegin buf, m
    VLA_Digest.VlaKeyAddText buf, m, "x"
    hexA = VLA_Digest.VlaKeyHex(buf, m)
    VLA_Digest.VlaKeyBegin buf, m
    VLA_Digest.VlaKeyAddText buf, m, "x"
    Report "optimize key: VlaKeyBegin restarts a reused buffer", _
           VLA_Digest.VlaKeyHex(buf, m) = hexA, "the second key carried the first's bytes"

    ' --- 4. every collision the framing prevents ----------------------

    Report "optimize key: the number 1 and the text ""1"" differ", _
           KeyOfCells(1#) <> KeyOfCells("1"), "the type tag was lost"
    Report "optimize key: a blank cell and empty text differ", _
           KeyOfCells(Empty) <> KeyOfCells(""), "the type tag was lost"
    Report "optimize key: (""ab"",""c"") and (""a"",""bc"") differ", _
           KeyOfCells("ab", "c") <> KeyOfCells("a", "bc"), "the length frame was lost"
    Report "optimize key: TRUE and the text ""TRUE"" differ", _
           KeyOfCells(True) <> KeyOfCells("TRUE"), "the type tag was lost"
    Report "optimize key: TRUE and the number 1 differ", _
           KeyOfCells(True) <> KeyOfCells(1#), "the type tag was lost"
    Report "optimize key: FALSE and TRUE differ", _
           KeyOfCells(False) <> KeyOfCells(True), "the payload byte was lost"
    Report "optimize key: an error value and the text ""#N/A"" differ", _
           KeyOfCells(CVErr(2042)) <> KeyOfCells("#N/A"), "the type tag was lost"
    Report "optimize key: an error value and a blank differ", _
           KeyOfCells(CVErr(2042)) <> KeyOfCells(Empty), "the type tag was lost"
    Report "optimize key: two rows of one against one row of two differ", _
           KeyOfCounts(2, 1) <> KeyOfCounts(1, 2), "the shape was lost"
    Report "optimize key: tables (AB, C) and (A, BC) differ", _
           KeyOfTexts("AB", "C") <> KeyOfTexts("A", "BC"), "the length frame was lost"
    Report "optimize key: a count and text of the same bytes differ", _
           KeyOfCounts(1) <> KeyOfTexts("x"), "the type tag was lost"

    ' A lone surrogate against U+FFFD, at the KEY level rather than the
    ' byte level - the collision step 1 exists to prevent, stated the
    ' way OPTIMIZE's memo would meet it.
    Report "optimize key: a cell of a + a lone surrogate and one of a + U+FFFD differ", _
           KeyOfCells("a" & KeyCharW(&HD800&)) <> KeyOfCells("a" & KeyCharW(&HFFFD&)), _
           "WTF-8 was lost: a memo would hand one workbook's answer to another's inputs"

    ' The same cells in the same order must of course agree, or the
    ' memo would never hit at all.
    Report "optimize key: the same cells in the same order give one key", _
           KeyOfCells("a", 1#, Empty, True) = KeyOfCells("a", 1#, Empty, True), "not deterministic"
    Report "optimize key: the same cells in a different order differ", _
           KeyOfCells("a", "b") <> KeyOfCells("b", "a"), "order was lost"

    ' A Long and a Double of the same value are ONE key: the engine
    ' cannot tell them apart either, so a second key would only cost a
    ' needless re-solve.
    Report "optimize key: a Long 1 and a Double 1 give one key", _
           KeyOfCells(CLng(1)) = KeyOfCells(1#), "a needless re-solve"
End Sub

' ---------------------------------------------------------------------
'  OPTIMIZE.1: the parity pin - "DATALOG wearing the name", measured.
'
'  THE CLAIM. A program with no choice, no constraint and no objective
'  IS a DATALOG program, so =OPTIMIZE(...) must answer exactly what
'  =DATALOG(...) answers for every such program. The owner approved
'  this proof on 2026-09-18: every DATALOG test program in this module,
'  run through OPTIMIZE and required to answer identically, so that the
'  claim is measured rather than described.
'
'  WHAT "IDENTICALLY" MEANS HERE, stated because a weaker reading would
'  be easy: the same queried relation with the same tuples IN THE SAME
'  ORDER (order matters - standing decision 2 breaks ties by the
'  Tables' own row order, so a reordering would be a real difference),
'  the same head-variable names, the same headless flag, the same
'  Boolean answer - or both engines refusing, with the same error
'  number AND the same words.
'
'  The relations are compared by hashing them with OPTIMIZE.1's own
'  step-1 key (VLA_Digest's injective framing), which is why one
'  comparison covers arity, count, order and every cell's type at once,
'  and why it cannot raise on a cell it did not expect.
'
'  NO TABLE ARGUMENTS ARE SUPPLIED, on purpose. A program whose tables
'  its original test passed in refuses in BOTH engines, identically,
'  because its predicates are then undefined - which is still parity,
'  and is what lets this pin cover every program without a fixture per
'  program. 128 of the 177 answer outright; the rest agree on their
'  refusal, to the word.
'
'  tools/check_optimize_parity.ps1 is what keeps "every" true: it reads
'  every DatalogRun/DATALOG call site in this module and fails if one
'  of their programs is missing from the table below. Without it the
'  table would quietly become "every program as of OPTIMIZE.1" the
'  first time somebody adds a DATALOG test - the same defect
'  check_devrig_mods_parity.ps1 exists for.
' ---------------------------------------------------------------------

' A relation's contents as one digest: its arity, its row count, then
' every cell of every tuple in order, through the injective framing of
' OPTIMIZE.1's step 1. Two relations agree exactly when these agree.
Private Function ParityRelationKey(ByVal rel As Collection) As String
    Dim buf() As Byte
    Dim n As Long
    Dim tup As Variant
    Dim arr() As Variant
    Dim i As Long
    VLA_Digest.VlaKeyBegin buf, n
    VLA_Digest.VlaKeyAddCount buf, n, VLA_Relation.RelArity(rel)
    VLA_Digest.VlaKeyAddCount buf, n, VLA_Relation.RelCount(rel)
    For Each tup In VLA_Relation.RelTuples(rel)
        arr = tup
        For i = LBound(arr) To UBound(arr)
            VLA_Digest.VlaKeyAddCell buf, n, arr(i)
        Next i
    Next tup
    ParityRelationKey = VLA_Digest.VlaKeyHex(buf, n)
End Function

' The head-variable names either engine reports, as one digest -
' Empty (no defining rule) and an array of names both, without ever
' touching the array on the branch that admitted it might not be one.
Private Function ParityNamesKey(ByVal v As Variant) As String
    Dim buf() As Byte
    Dim n As Long
    Dim i As Long
    VLA_Digest.VlaKeyBegin buf, n
    If IsEmpty(v) Then
        VLA_Digest.VlaKeyAddText buf, n, "(no head names)"
    ElseIf IsArray(v) Then
        For i = LBound(v) To UBound(v)
            VLA_Digest.VlaKeyAddCell buf, n, v(i)
        Next i
    Else
        VLA_Digest.VlaKeyAddCell buf, n, v
    End If
    ParityNamesKey = VLA_Digest.VlaKeyHex(buf, n)
End Function

' "" when the two engines agree completely; otherwise what differed, in
' words. Returns rather than raises, so one genuine difference reports
' itself and the remaining programs still run.
'
' Err.Number and Err.Description are read BEFORE On Error GoTo 0, which
' clears them - a VBA trap this project has met before.
Private Function ParityVerdict(ByVal program As String) As String
    Dim dErr As Long, oErr As Long
    Dim dDesc As String, oDesc As String
    Dim dRes As Collection, oRes As Collection
    Dim dRel As Collection, oRel As Collection

    On Error Resume Next
    Err.Clear
    Set dRes = VLA_Datalog.DatalogRun(program)
    dErr = Err.Number
    dDesc = Err.Description
    On Error GoTo 0

    On Error Resume Next
    Err.Clear
    Set oRes = VLA_Optimize.OptimizeRun(program)
    oErr = Err.Number
    oDesc = Err.Description
    On Error GoTo 0

    If dErr <> 0 Or oErr <> 0 Then
        If dErr = 0 Then
            ParityVerdict = "DATALOG answered; OPTIMIZE refused with " & oDesc
        ElseIf oErr = 0 Then
            ParityVerdict = "DATALOG refused with " & dDesc & "; OPTIMIZE answered"
        ElseIf dErr <> oErr Then
            ParityVerdict = "both refused, different error numbers: " & dErr & " against " & oErr
        ElseIf StrComp(dDesc, oDesc, vbBinaryCompare) <> 0 Then
            ParityVerdict = "both refused, different words: [" & dDesc & "] against [" & oDesc & "]"
        End If
        Exit Function
    End If

    If StrComp(CStr(dRes.Item(1)), CStr(oRes.Item(1)), vbBinaryCompare) <> 0 Then
        ParityVerdict = "different queried predicate: " & CStr(dRes.Item(1)) & " against " & CStr(oRes.Item(1))
        Exit Function
    End If
    If CBool(dRes.Item(4)) <> CBool(oRes.Item(4)) Then
        ParityVerdict = "different headless flag"
        Exit Function
    End If
    If IsEmpty(dRes.Item(5)) <> IsEmpty(oRes.Item(5)) Then
        ParityVerdict = "one engine answered a Boolean and the other a table"
        Exit Function
    End If
    If Not IsEmpty(dRes.Item(5)) Then
        If CBool(dRes.Item(5)) <> CBool(oRes.Item(5)) Then
            ParityVerdict = "different Boolean answer"
            Exit Function
        End If
    End If
    If StrComp(ParityNamesKey(dRes.Item(3)), ParityNamesKey(oRes.Item(3)), vbBinaryCompare) <> 0 Then
        ParityVerdict = "different head-variable names"
        Exit Function
    End If
    Set dRel = VLA_Runtime.VlaDictGet(dRes.Item(2), CStr(dRes.Item(1)))
    Set oRel = VLA_Runtime.VlaDictGet(oRes.Item(2), CStr(oRes.Item(1)))
    If StrComp(ParityRelationKey(dRel), ParityRelationKey(oRel), vbBinaryCompare) <> 0 Then
        ParityVerdict = "different rows: " & VLA_Relation.RelCount(dRel) & " against " & VLA_Relation.RelCount(oRel)
        Exit Function
    End If
    ' OPTIMIZE's two extra items: a zero-choice program is the one
    ' answer there is, so it is proven best, and it says so.
    If CLng(oRes.Item(6)) <> VLA_Optimize.VLA_OPTIMIZE_PROVEN_BEST Then
        ParityVerdict = "OPTIMIZE reported result state " & CStr(oRes.Item(6)) & ", not proven best"
        Exit Function
    End If
End Function

Private Sub TestOptimizeParity()
    Dim programs As Collection
    Set programs = DatalogParityPrograms()
    Dim p As Variant
    Dim ix As Long
    Dim label As String
    Dim verdict As String
    For Each p In programs
        ix = ix + 1
        label = CStr(p)
        If Len(label) > 64 Then label = Left$(label, 64) & "..."
        verdict = ParityVerdict(CStr(p))
        Report "optimize parity " & ix & ": " & label, Len(verdict) = 0, verdict
    Next p
    ' The count itself is a pin: tools/check_optimize_parity.ps1 found
    ' 177 distinct programs in this module at OPTIMIZE.1, and it fails
    ' if a later DATALOG test adds one the table does not carry.
    Report "optimize parity: the table carries every DATALOG program (177 at OPTIMIZE.1)", _
           programs.Count >= 177, "got " & programs.Count
End Sub

' Every distinct DATALOG program this module names, from every
' DatalogRun and DATALOG call site. Maintained with
' tools/check_optimize_parity.ps1, which reads those call sites and
' fails when one is missing here; -Emit prints this list to paste.
Private Function DatalogParityPrograms() As Collection
    Dim p As Collection
    Set p = New Collection
    ' Generated by tools\check_optimize_parity.ps1 -Emit, then reviewed.
    ' 178 distinct DATALOG programs, every one this module names.
    p.Add " (fact (covered ""Night"")) (query (not (gap X)))"
    p.Add " (fact (spare ""Bob"" ""Day"")) (rule (can P S) (cc P S)) (rule (can P S) (spare P S)) (rule (who-can S W) (shift S) (textjoin W "", "" (can P S))) (query who-can)"
    p.Add " (query (not (gap X)))"
    p.Add " (query (not (route ""A"" X)))"
    p.Add " (query (not (route X ""E"")))"
    p.Add " (query (route ""A"" ""D""))"
    p.Add " (query (route ""A"" ""E""))"
    p.Add " (query (text-starts-with ""GL-4010"" ""GL-4""))"
    p.Add " (rule (all W) (shift S) (textjoin W "", "" (cc P Q))) (query all)"
    p.Add " (rule (any-code C) (acct C) (text-contains C """")) (query any-code)"
    p.Add " (rule (chain S W) (shift S) (textjoin W "", "" (chain X S))) (query chain)"
    p.Add " (rule (first-code C) (acct C) !) (query first-code)"
    p.Add " (rule (kind C T) (acct C) (if (text-starts-with C ""GL-4"") (= T ""Revenue"") (= T ""Cost""))) (query kind)"
    p.Add " (rule (odd C) (acct C) (not (sub-atom C 0 L A ""GL""))) (query odd)"
    p.Add " (rule (other C) (acct C) (not (text-starts-with C ""GL-4""))) (query other)"
    p.Add " (rule (revenue C) (acct C) (sub-atom C 0 L A ""GL-4"")) (query revenue)"
    p.Add " (rule (revenue C) (acct C) (text-starts-with C ""gl-4"")) (query revenue)"
    p.Add " (rule (revenue C) (acct C) (text-starts-with C)) (query revenue)"
    p.Add " (rule (revenue C) (text-starts-with C ""GL-4"") (acct C)) (query revenue)"
    p.Add " (rule (tens C) (acct C) (text-ends-with C ""10"")) (query tens)"
    p.Add " (rule (vla-ask-can-cover Who) (can-cover Who Night)) (query vla-ask-can-cover)"
    p.Add " (rule (vla-ask-can-drive Who) (can-drive Who Night)) (query vla-ask-can-drive)"
    p.Add " (rule (who-covers S L) (shift S) (findall P (cc P S) L)) (query who-covers)"
    p.Add " (rule (who-covers S W) (shift S) (textjoin W """" (cc P S))) (query who-covers)"
    p.Add " (rule (who-covers S W) (shift S) (textjoin W "", "" (cc P S))) (query who-covers)"
    p.Add " (rule (who-covers S W) (shift S) (textjoin W ""; "" (cc P S))) (query who-covers)"
    p.Add " (rule (who-covers S W) (shift S) (textjoin W (cc P S))) (query who-covers)"
    p.Add " (rule (who-covers S W) (shift S) (textjoin W Sep (cc P S))) (query who-covers)"
    p.Add " (rule (who-covers S) (shift S) (textjoin S "", "" (cc P S))) (query who-covers)"
    p.Add "(fact (deptname eng)) (fact (deptname sales)) (rule (deptcount D N) (deptname D) (count N (staffing (dept D)))) (query deptcount)"
    p.Add "(fact (edge a a)) (fact (edge a b)) (rule (self_loop X) (edge X X)) (query self_loop)"
    p.Add "(fact (edge a b 1)) (fact (edge b c 1)) (fact (edge c d 1)) (rule (path X Y D) (edge X Y D)) (rule (path X Z D) (edge X Y D1) (path Y Z D2) (let D (+ D1 D2))) (query path)"
    p.Add "(fact (item a)) (rule (foo X Z) (item X) (let Z (+ X))) (query foo)"
    p.Add "(fact (item a)) (rule (foo X Z) (item X) (let Z (+ Y 1))) (query foo)"
    p.Add "(fact (item a)) (rule (foo X) (item X) (> X)) (query foo)"
    p.Add "(fact (item a)) (rule (foo X) (item X) (> Y 5)) (query foo)"
    p.Add "(fact (item a)) (rule (foo X) (item X) (let)) (query foo)"
    p.Add "(fact (item a)) (rule (foo X) (not (bar X))) (query foo)"
    p.Add "(fact (item a)) (rule (p X N) (item X) (count N (p X Y))) (query p)"
    p.Add "(fact (item a)) (rule (p X) (item X) (not (p X))) (query p)"
    p.Add "(fact (item a)) (rule (p X) (item X) (not (q X))) (rule (q X) (item X) (not (p X))) (query p)"
    p.Add "(fact (link ""A"" ""B"")) (fact (link ""C"" ""C"")) (query (not (link X X)))"
    p.Add "(fact (link ""A"" ""B"")) (query (link ""A"" ""B""))"
    p.Add "(fact (link ""A"" ""B"")) (query (link ""A"" Who))"
    p.Add "(fact (link ""A"" ""B"")) (query (link ""A""))"
    p.Add "(fact (link ""A"" ""B"")) (query (link ""B"" ""A""))"
    p.Add "(fact (link ""A"" ""B"")) (query (not (link ""A"" X)))"
    p.Add "(fact (link ""A"" ""B"")) (query (not (link ""A"")))"
    p.Add "(fact (link ""A"" ""B"")) (query (not (link ""B"" X)))"
    p.Add "(fact (link ""A"" ""B"")) (query (not (link X X)))"
    p.Add "(fact (link ""A"" ""B"")) (query (not (link X Y)))"
    p.Add "(fact (lvl ""Ann"" 3)) (rule (low P) (lvl P L) (=< L 2)) (query low)"
    p.Add "(fact (member ""Ann"" ""Ops"")) (rule (in-ops P) (member P ""Ops"")) (query in-ops)"
    p.Add "(fact (node a)) (fact (node b)) (fact (node c)) (fact (edge a b)) (fact (edge b c)) (rule (reachable X Y) (edge X Y)) (rule (reachable X Z) (edge X Y) (reachable Y Z)) (rule (reachcount X N) (node X) (count N (reachable X Y))) (query reachcount)"
    p.Add "(fact (node a)) (fact (node b)) (fact (node c)) (fact (edge a b)) (fact (edge b c)) (rule (reachable X Y) (edge X Y)) (rule (reachable X Z) (edge X Y) (reachable Y Z)) (rule (unreachable X Y) (node X) (node Y) (not (reachable X Y))) (query unreachable)"
    p.Add "(fact (p ""a"")) (query p)"
    p.Add "(fact (p ""x"" ""a"")) (fact (p ""x"" ""b"")) (rule (j K W) (p K Z) (textjoin W "", "" (p K V))) (query j)"
    p.Add "(fact (p (q r))) (query p)"
    p.Add "(fact (p a b)) (fact (p a b c)) (query p)"
    p.Add "(fact (p a))"
    p.Add "(fact (p a)) (fact (p b)) (query (not (p b)))"
    p.Add "(fact (p a)) (fact (p b)) (query p)"
    p.Add "(fact (p a)) (fact (q b)) (query (not (p a) (q b)))"
    p.Add "(fact (p a)) (query (""not"" (p a)))"
    p.Add "(fact (p a)) (query (not (not (p a))))"
    p.Add "(fact (p a)) (query (not (p a)))"
    p.Add "(fact (p a)) (query (not (p b)))"
    p.Add "(fact (p a)) (query (not (p)))"
    p.Add "(fact (p a)) (query (not (pp X)))"
    p.Add "(fact (p a)) (query (not p))"
    p.Add "(fact (p a)) (query (not))"
    p.Add "(fact (p a)) (query (p a) (p b))"
    p.Add "(fact (p a)) (query (p a)) (headless)"
    p.Add "(fact (p a)) (query (p X))"
    p.Add "(fact (p a)) (query (p))"
    p.Add "(fact (p a)) (query (pp a))"
    p.Add "(fact (p a)) (query p)"
    p.Add "(fact (p a)) (query pp)"
    p.Add "(fact (p a)) (rule (a1 X) (p X) (zz1 X)) (rule (a2 X) (p X) (zz2 X)) (query a2)"
    p.Add "(fact (p a)) (rule (q X) (p X) (not (p X))) (query (not (q X)))"
    p.Add "(fact (p a)) (rule (q X) (p X) (typo X)) (query (not (p b)))"
    p.Add "(fact (p a)) (rule (q X) (p X) (typo X)) (query (p a))"
    p.Add "(fact (p a)) (rule (unused X) (sibling X)) (query p)"
    p.Add "(fact (pair ""a"" ""b"")) (rule (text-ends-with X Y) (pair X Y)) (query text-ends-with)"
    p.Add "(fact (pair ""GL-4010"" ""GL-4"")) (fact (pair ""GL-5010"" ""GL-4"")) (rule (ok C) (pair C P) (text-starts-with C P)) (query ok)"
    p.Add "(fact (pair 10 3)) (rule (added S) (pair A B) (let S (+ A B))) (query added)"
    p.Add "(fact (pair 10 3)) (rule (big S) (pair A B) (let S (+ A B)) (> S 5)) (query big)"
    p.Add "(fact (pair 10 3)) (rule (multiplied S) (pair A B) (let S (* A B))) (query multiplied)"
    p.Add "(fact (pair 10 3)) (rule (subbed S) (pair A B) (let S (- A B))) (query subbed)"
    p.Add "(fact (pair 10 4)) (rule (divided S) (pair A B) (let S (/ A B))) (query divided)"
    p.Add "(fact (pair 3 7)) (rule (mx S) (pair A B) (let S (max A B))) (query mx)"
    p.Add "(fact (pair 7 2)) (rule (idiv S) (pair A B) (let S (// A B))) (query idiv)"
    p.Add "(fact (pair -7 3)) (rule (modded S) (pair A B) (let S (mod A B))) (query modded)"
    p.Add "(fact (pair -7 3)) (rule (remmed S) (pair A B) (let S (rem A B))) (query remmed)"
    p.Add "(fact (pair a 1 2)) (rule (foo X Z) (pair X A B) (let Z (% A B))) (query foo)"
    p.Add "(fact (pair a 1 2)) (rule (foo X Z) (pair X A B) (let z (+ A B))) (query foo)"
    p.Add "(fact (pair a 1)) (rule (foo X N) (pair X N) (let N (+ N 1))) (query foo)"
    p.Add "(fact (pair a 10 0)) (rule (bad X Z) (pair X A B) (let Z (/ A B))) (query pp)"
    p.Add "(fact (pair a 10 0)) (rule (bad X Z) (pair X A B) (let Z (/ A B))) (rule (ok X) (pair X A B) (typo X)) (query ok)"
    p.Add "(fact (pair a 10 0)) (rule (foo X Z) (pair X A B) (let Z (/ A B))) (query foo)"
    p.Add "(fact (pair a hello)) (rule (foo X Z) (pair X A) (let Z (+ A 1))) (query foo)"
    p.Add "(fact (parent tom bob)) (fact (parent bob liz)) (query parent)"
    p.Add "(fact (parent tom bob)) (fact (parent bob liz)) (rule (grandparent X Z) (parent X Y) (parent Y Z)) (query grandparent)"
    p.Add "(fact (parent tom bob)) (query (parent tom bob))"
    p.Add "(fact (parent tom bob)) (rule (kid X) (parnet tom X)) (query kid)"
    p.Add "(fact (person ""Bob"")) (query (person ""bob""))"
    p.Add "(fact (person alice 25)) (fact (person bob 15)) (fact (banned bob)) (rule (adult_allowed X) (person X Age) (> Age 18) (not (banned X))) (query adult_allowed)"
    p.Add "(fact (person alice)) (fact (amount alice ""10"")) (rule (bad X S) (person X) (sum S (amount X ""10""))) (query bad)"
    p.Add "(fact (person alice)) (fact (amount2 alice 1 2)) (rule (bad X S) (person X) (sum S (amount2 X V W))) (query bad)"
    p.Add "(fact (person alice)) (fact (person bob)) (fact (person carol)) (fact (amount alice 10)) (fact (amount alice 15)) (fact (amount bob 7)) (rule (total X S) (person X) (sum S (amount X V))) (query total)"
    p.Add "(fact (person alice)) (fact (person bob)) (fact (person carol)) (fact (banned bob)) (rule (allowed X) (person X) (not (banned X))) (query allowed)"
    p.Add "(fact (person alice)) (fact (person bob)) (fact (person carol)) (fact (sale alice widget)) (fact (sale alice gadget)) (fact (sale bob widget)) (rule (salescount X N) (person X) (count N (sale X Y))) (query salescount)"
    p.Add "(fact (person alice)) (fact (sale alice widget)) (rule (bad X N) (person X) (count N (sale X Y)) (count N (sale X Z))) (query bad)"
    p.Add "(fact (person alice)) (fact (sale alice widget)) (rule (bad X) (person X) (count n (sale X Y))) (query bad)"
    p.Add "(fact (person alice)) (rule (bad X N) (person X) (count N)) (query bad)"
    p.Add "(fact (person tom)) (fact (amount tom 10)) (rule (total X S) (person X) (sum S (amont X V))) (query total)"
    p.Add "(fact (person tom)) (fact (banned tom)) (rule (ok X) (person X) (not (bannd X))) (query ok)"
    p.Add "(fact (person tom)) (fact (banned tom)) (rule (ok X) (person X) (not (banned X))) (query ok)"
    p.Add "(fact (person tom)) (fact (sale tom widget)) (rule (sales X N) (person X) (count N (sael X Y))) (query sales)"
    p.Add "(fact (person tom)) (rule (ok X) (person X) (not (leave X X))) (query ok)"
    p.Add "(fact (product ""Banana"")) (fact (product ""Cyan"")) (fact (product ""Apple"")) (rule (has-an N) (product N) (text-contains N ""an"")) (query has-an)"
    p.Add "(fact (reports_to alice bob)) (fact (reports_to bob carol)) (fact (reports_to carol dave)) (rule (indirect_report X Y) (reports_to X Y)) (rule (indirect_report X Y) (reports_to X Z) (indirect_report Z Y)) (query indirect_report)"
    p.Add "(fact (score alice 9)) (fact (score bob 10)) (rule (high_scorer X) (score X N) (> N 8)) (query high_scorer)"
    p.Add "(fact (tag apple)) (fact (tag banana)) (fact (tag cherry)) (rule (not_banana X) (tag X) (<> X banana)) (query not_banana)"
    p.Add "(fact (text-contains ""a"" ""b"")) (query text-contains)"
    p.Add "(fact (val 2.5)) (rule (rnd S) (val A) (let S (round A))) (query rnd)"
    p.Add "(fact (val -4)) (rule (absd S) (val A) (let S (abs A))) (query absd)"
    p.Add "(fact (widget a b)) (rule (bad X) (widget (foo X))) (query bad)"
    p.Add "(headless extra) (fact (p a)) (query p)"
    p.Add "(headless) (fact (p a)) (query (not (p b)))"
    p.Add "(headless) (fact (p a)) (query (p a))"
    p.Add "(headless) (fact (parent tom bob)) (fact (parent bob liz)) (query parent)"
    p.Add "(headless) (fact (thing a)) (rule (nothing_here X) (thing X) (thing none)) (query nothing_here)"
    p.Add "(headless) (query p)"
    p.Add "(headless) (query personhosttest1)"
    p.Add "(headless) (query vladatalogaliastest)"
    p.Add "(headless) (query vlaspilldup)"
    p.Add "(headless) (query vlaspillnums)"
    p.Add "(headless) (query vlaspillpart)"
    p.Add "(headless) (query vlaspillplain)"
    p.Add "(headless) (query vlaspillsched)"
    p.Add "(headless) (query widehosttest1)"
    p.Add "(headless) (rule (indirect X Y) (reportstohosttest1 X Y)) (rule (indirect X Y) (reportstohosttest1 X Z) (indirect Z Y)) (query indirect)"
    p.Add "(headless) (rule (ok X) (personhosttest1 X) (not (emptydataloghosttest1 X))) (query ok)"
    p.Add "(headless) (rule (rich X) (staffinghosttest1 (name X) (salary S)) (> S 80000)) (query rich)"
    p.Add "(query (not (leave X Y)))"
    p.Add "(query (not (staff (name ""Ann"") (level 3))))"
    p.Add "(query (staff (name ""Ann"") (level 3)))"
    p.Add "(query (sub-atom ""abc"" 0 1 A ""a""))"
    p.Add "(query p)"
    p.Add "(query pp)"
    p.Add "(rule (allkeyed N S D) (staffing (name N) (salary S) (dept D))) (query allkeyed)"
    p.Add "(rule (bad X S) (staffing (name X) (bogus S))) (query bad)"
    p.Add "(rule (bad X Y) (staffing (name X) (name Y))) (query bad)"
    p.Add "(rule (bad X Y) (staffing (name X) (salary (foo Y)))) (query bad)"
    p.Add "(rule (bad X) (staffing X (salary S))) (query bad)"
    p.Add "(rule (bad X) (widget (foo X))) (fact (widget a b)) (query bad)"
    p.Add "(rule (can-cover Person Shift) (rota (name Person) (shift Shift))) (query (can-cover ""Bob"" ""Night""))"
    p.Add "(rule (foo X Y) (bar X)) (query foo)"
    p.Add "(rule (forty N) (codes (code C) (name N)) (text-starts-with C 40)) (query forty)"
    p.Add "(rule (half N) (codes (code C) (name N)) (text-starts-with C ""0.5"")) (query half)"
    p.Add "(rule (has-level P L) (staff (name P) (level L))) (query (has-level ""Ann"" 3))"
    p.Add "(rule (has-level P L) (staff (name P) (level L))) (query (has-level ""Bob"" 3))"
    p.Add "(rule (has-level P L) (staff (name P) (level L))) (query (not (has-level ""Ann"" 3)))"
    p.Add "(rule (has-level P L) (staff (name P) (level L))) (query (not (has-level Who 2)))"
    p.Add "(rule (has-level P L) (staff (name P) (level L))) (query (not (has-level Who 3)))"
    p.Add "(rule (highearner X) (staff X S) (> S 80000)) (query highearner)"
    p.Add "(rule (indirect_report X Y) (reports_to X Y)) (rule (indirect_report X Y) (reports_to X Z) (indirect_report Z Y)) (query indirect_report)"
    p.Add "(rule (levels N L) (staff (name N)) (textjoin L "", "" (staff (name N) (level V)))) (query levels)"
    p.Add "(rule (pair X Y) (staffing (name X)) (staffing (name Y))) (query pair)"
    p.Add "(rule (rich X) (staffing (name X) (salary S)) (> S 80000)) (query rich)"
    p.Add "(rule (senior X) (staf (name X) (level L)) (> L 2)) (query senior)"
    p.Add "(rule (staffing (name X)) (staffing (name X))) (query staffing)"
    p.Add "(rule (who P) (vlaspillempty (name P))) (query who)"
    p.Add "(rule (who P) (vlaspillsched (name P))) (query who)"
    p.Add "(rule (who X) (emptydataloghosttest1 X)) (query who)"
    p.Add "(rule (zeros N) (codes (code C) (name N)) (text-starts-with C ""00"")) (query zeros)"
    p.Add "; only a comment"
    Set DatalogParityPrograms = p
End Function

' ---------------------------------------------------------------------
'  OPTIMIZE.1: the six forms, their shapes, and the five result states.
'
'  Everything here is a SPELLING pin. None of the six forms executes in
'  this version, so what is being proved is narrow and load-bearing:
'  that each spelling PARSES (reaching its not-yet refusal rather than a
'  shape refusal), that a wrong shape is refused by name rather than
'  guessed, and that the words of every refusal say what they say. A
'  spelling settled now and refused clearly now is a spelling nobody has
'  to change later (SD-4), and that is this item's actual product.
'
'  THE SECTION 17 BLOCK is the scoping run's own proof. Every rule line
'  of scripts/pareto_logic.txt section 17 that needs a shape - the
'  killer case's count from a Table column, duty-rotation's two-atom
'  group, the wedding's weights from a column, the two per-row choices,
'  the implications - is written out in the settled spellings and
'  required to parse. The three shapes this item's scoping added
'  ((per ...), choose-any, and a variable cost) are each there because
'  one of those lines could not be said without it.
' ---------------------------------------------------------------------

' The words a program's refusal carries, or "(no refusal)" if it did
' not refuse. Err.Number and Err.Description are both read before
' On Error GoTo 0, which clears them.
Private Function OptRefusalOf(ByVal program As String) As String
    Dim d As String
    Dim num As Long
    On Error Resume Next
    Err.Clear
    VLA_Optimize.OptimizeRun program
    num = Err.Number
    d = Err.Description
    On Error GoTo 0
    If num = 0 Then
        OptRefusalOf = "(no refusal)"
    Else
        OptRefusalOf = d
    End If
End Function

Private Sub AssertOptRefusal(ByVal name As String, ByVal program As String, ByVal frag As String)
    Dim d As String
    d = OptRefusalOf(program)
    Report name, InStr(1, d, frag, vbTextCompare) > 0, "got: " & d
End Sub

Private Sub TestOptimizeForms()
    Dim v As Variant
    Dim words As String

    ' --- the five choice spellings, each parsing and refusing ---------
    AssertOptRefusal "optimize forms: choose-exactly parses and is refused by name", _
        "(fact (elig s p)) (choose-exactly 2 (assign S P) (elig S P)) (query elig)", _
        "asks OPTIMIZE to make a choice"
    AssertOptRefusal "optimize forms: choose-at-least parses and is refused by name", _
        "(fact (elig s p)) (choose-at-least 2 (assign S P) (elig S P)) (query elig)", _
        "asks OPTIMIZE to make a choice"
    AssertOptRefusal "optimize forms: choose-at-most parses and is refused by name", _
        "(fact (elig s p)) (choose-at-most 2 (assign S P) (elig S P)) (query elig)", _
        "asks OPTIMIZE to make a choice"
    AssertOptRefusal "optimize forms: choose-between parses and is refused by name", _
        "(fact (elig s p)) (choose-between 2 4 (assign S P) (elig S P)) (query elig)", _
        "asks OPTIMIZE to make a choice"
    AssertOptRefusal "optimize forms: choose-any parses and is refused by name", _
        "(fact (link a b)) (choose-any (on F T) (link F T)) (query link)", _
        "asks OPTIMIZE to make a choice"

    ' A bare (choose ...) is refused BY NAME, listing all five - so the
    ' one word nobody should write teaches the five that work, and
    ' nothing in a formula bar resembles Excel's own CHOOSE.
    AssertOptRefusal "optimize forms: a bare (choose ...) is refused, listing the five", _
        "(fact (elig s p)) (choose 2 (assign S P) (elig S P)) (query elig)", _
        "choose-exactly, choose-at-least, choose-at-most, choose-between, or choose-any"

    ' --- hard rules, both polarities ----------------------------------
    AssertOptRefusal "optimize forms: require parses and is refused by name", _
        "(fact (reviews t p)) (require (senior P) (reviews T P)) (query reviews)", _
        "a rule about every possible answer"
    AssertOptRefusal "optimize forms: a ground require (no body) parses", _
        "(fact (seat a h)) (require (seat ""Ann"" ""Head"")) (query seat)", _
        "a rule about every possible answer"
    AssertOptRefusal "optimize forms: forbid parses and is refused by name", _
        "(fact (prepares t p)) (forbid (prepares T P) (reviews T P)) (query prepares)", _
        "a rule about every possible answer"

    ' --- soft preferences ---------------------------------------------
    AssertOptRefusal "optimize forms: prefer parses and is refused by name", _
        "(fact (wish a b)) (prefer (together A B) (wish A B)) (query wish)", _
        "says what a good answer looks like"
    AssertOptRefusal "optimize forms: avoid parses and is refused by name", _
        "(fact (off p s)) (avoid (assign S P) (off P S)) (query off)", _
        "says what a good answer looks like"
    AssertOptRefusal "optimize forms: prefer takes an optional (cost N)", _
        "(fact (wish a b)) (prefer (together A B) (wish A B) (cost 2)) (query wish)", _
        "says what a good answer looks like"

    ' --- objectives ---------------------------------------------------
    AssertOptRefusal "optimize forms: minimize parses and is refused by name", _
        "(fact (ot p)) (minimize (overtime P) (ot P)) (query ot)", _
        "as small or as large as possible"
    AssertOptRefusal "optimize forms: maximize parses and is refused by name", _
        "(fact (m i)) (maximize (margin I) (m I)) (query m)", _
        "as small or as large as possible"
    AssertOptRefusal "optimize forms: minimise is the same form (the OPTIMISE courtesy)", _
        "(fact (ot p)) (minimise (overtime P) (ot P)) (query ot)", _
        "as small or as large as possible"
    AssertOptRefusal "optimize forms: maximise is the same form", _
        "(fact (m i)) (maximise (margin I) (m I)) (query m)", _
        "as small or as large as possible"

    ' --- the kept schedule --------------------------------------------
    AssertOptRefusal "optimize forms: fewest-changes-from parses and is refused by name", _
        "(fact (p a)) (fewest-changes-from LastMonth) (query p)", _
        "names a schedule to stay close to"

    ' --- the budget, as effort ----------------------------------------
    AssertOptRefusal "optimize forms: (effort quick) parses", _
        "(fact (p a)) (effort quick) (query p)", "sets how much work"
    AssertOptRefusal "optimize forms: (effort normal) parses", _
        "(fact (p a)) (effort normal) (query p)", "sets how much work"
    AssertOptRefusal "optimize forms: (effort thorough) parses", _
        "(fact (p a)) (effort thorough) (query p)", "sets how much work"
    AssertOptRefusal "optimize forms: (effort 50000) parses - a number, for experts", _
        "(fact (p a)) (effort 50000) (query p)", "sets how much work"
    AssertOptRefusal "optimize forms: an unknown effort level is refused by name", _
        "(fact (p a)) (effort sideways) (query p)", "is not an effort level"
    AssertOptRefusal "optimize forms: (effort 0) is refused - no work is not a budget", _
        "(fact (p a)) (effort 0) (query p)", "is not an effort level"
    AssertOptRefusal "optimize forms: an effort of two words is refused", _
        "(fact (p a)) (effort quick thorough) (query p)", "is written wrong"

    ' --- shapes refused rather than guessed ---------------------------
    AssertOptRefusal "optimize forms: choose-exactly with no pool is refused", _
        "(fact (p a)) (choose-exactly 2 (assign S P)) (query p)", "is written wrong"
    AssertOptRefusal "optimize forms: choose-exactly with too many parts is refused", _
        "(fact (p a)) (choose-exactly 2 (assign S P) (elig S P) (per (sh S)) (extra X)) (query p)", _
        "is written wrong"
    AssertOptRefusal "optimize forms: a count that is neither a number nor a name is refused", _
        "(fact (p a)) (choose-exactly ""two"" (assign S P) (elig S P)) (query p)", _
        "'two' is neither a whole number of zero or more nor a name"
    ' "of zero or more" is load-bearing in the words, not decoration:
    ' -1 IS a whole number, and the first draft of this message said
    ' only "a whole number", which would have been wrong about the very
    ' next assertion. The fragment quotes the count too, so a refusal
    ' about the wrong slot cannot pass either of these.
    AssertOptRefusal "optimize forms: a negative count is refused", _
        "(fact (p a)) (choose-exactly -1 (assign S P) (elig S P)) (query p)", _
        "'-1' is neither a whole number of zero or more nor a name"
    AssertOptRefusal "optimize forms: choose-between with its numbers the wrong way round is refused", _
        "(fact (p a)) (choose-between 4 2 (assign S P) (elig S P)) (query p)", _
        "write the smaller number first"
    AssertOptRefusal "optimize forms: a bare word where a row belongs is refused", _
        "(fact (p a)) (choose-exactly 2 assign (elig S P)) (query p)", "is not one"
    AssertOptRefusal "optimize forms: a bare word where the pool belongs is refused", _
        "(fact (p a)) (choose-exactly 2 (assign S P) elig) (query p)", "is not one"
    AssertOptRefusal "optimize forms: a misspelled (per ...) is refused, not ignored", _
        "(fact (p a)) (choose-exactly 2 (assign S P) (elig S P) (pre (sh S))) (query p)", _
        "is written wrong"
    AssertOptRefusal "optimize forms: an empty (per) is refused", _
        "(fact (p a)) (choose-exactly 2 (assign S P) (elig S P) (per)) (query p)", _
        "at least one row to group by"
    AssertOptRefusal "optimize forms: choose-any with a count is refused - it takes none", _
        "(fact (p a)) (choose-any 1 (on F T) (link F T)) (query p)", "is written wrong"
    AssertOptRefusal "optimize forms: a (cost ...) holding a row is refused", _
        "(fact (p a)) (prefer (together A B) (cost (w A))) (query p)", "not a row"
    AssertOptRefusal "optimize forms: a (cost ...) of two things is refused", _
        "(fact (p a)) (prefer (together A B) (cost 1 2)) (query p)", "is written wrong"
    AssertOptRefusal "optimize forms: a (prefer ...) that is only a cost is refused", _
        "(fact (p a)) (prefer (cost 1)) (query p)", "is written wrong"
    AssertOptRefusal "optimize forms: fewest-changes-from needs exactly one name", _
        "(fact (p a)) (fewest-changes-from A B) (query p)", "is written wrong"
    AssertOptRefusal "optimize forms: fewest-changes-from of a row is refused", _
        "(fact (p a)) (fewest-changes-from (t X)) (query p)", "not a row"
    AssertOptRefusal "optimize forms: an unknown top-level form is refused, listing OPTIMIZE's own", _
        "(fact (p a)) (choos 2 (assign S P) (elig S P)) (query p)", _
        "is not an OPTIMIZE form"
    AssertOptRefusal "optimize forms: the unknown-form refusal names choose-any, not DATALOG's three forms", _
        "(fact (p a)) (wibble x) (query p)", "choose-any"

    ' --- section 17's own sentences, in the settled spellings ---------
    '
    ' Each of these is one line of scripts/pareto_logic.txt section 17,
    ' written out. Reaching a not-yet refusal is the pass: it means the
    ' spelling SAID the sentence. A shape refusal here would mean the
    ' spellings cannot say something the corpus asks for, which is the
    ' one thing this item's scoping run was for.

    ' optimize-roster, the killer case: "every shift gets exactly the
    ' people it needs" - the count comes from the Shifts table's own
    ' Need column, which is why (per ...) exists and why a three-slot
    ' choice form could not say this at all.
    AssertOptRefusal "optimize forms (s17): the killer case's count comes from a Table column", _
        "(fact (p a)) (choose-exactly N (assign S P) (elig S P) (per (shifts S N))) (query p)", _
        "asks OPTIMIZE to make a choice"
    ' optimize-duty-rotation: "every month, every person gets exactly
    ' one duty" - a group of two, which is why (per ...) holds a list.
    AssertOptRefusal "optimize forms (s17): a group of two things", _
        "(fact (p a)) (choose-exactly 1 (assign M P D) (duty D) (per (month M) (person P))) (query p)", _
        "asks OPTIMIZE to make a choice"
    ' optimize-toy and optimize-audit-independence: exactly k per group.
    AssertOptRefusal "optimize forms (s17): exactly two of the people on every shift", _
        "(fact (p a)) (choose-exactly 2 (assign S P) (person P) (per (shift S))) (query p)", _
        "asks OPTIMIZE to make a choice"
    ' optimize-roster's needs as a range, the loose variant's shape.
    AssertOptRefusal "optimize forms (s17): between two and four per shift", _
        "(fact (p a)) (choose-between 2 4 (assign S P) (elig S P) (per (shift S))) (query p)", _
        "asks OPTIMIZE to make a choice"
    ' optimize-quote-bike: "each extra is in the quote or not", and
    ' optimize-network-connect: "each link is switched on, or not" - the
    ' two sentences that needed the fifth form.
    AssertOptRefusal "optimize forms (s17): each extra is in the quote or not", _
        "(fact (p a)) (choose-any (include I) (extra I)) (query p)", _
        "asks OPTIMIZE to make a choice"
    AssertOptRefusal "optimize forms (s17): each link is switched on, or not", _
        "(fact (p a)) (choose-any (on F T) (links F T C)) (query p)", _
        "asks OPTIMIZE to make a choice"
    ' optimize-sod-close: "every reviewer is senior" - an implication,
    ' and the reason require is not forbid with a not in it.
    AssertOptRefusal "optimize forms (s17): every reviewer is senior", _
        "(fact (p a)) (require (senior P) (reviews T P)) (query p)", _
        "a rule about every possible answer"
    ' optimize-seat-dinner: "Ann and Bob sit together" - the same
    ' implication shape, earning itself a second time.
    AssertOptRefusal "optimize forms (s17): Ann and Bob sit together", _
        "(fact (p a)) (require (seat ""Bob"" T) (seat ""Ann"" T)) (query p)", _
        "a rule about every possible answer"
    ' optimize-roster-senior, through a derived rule doing the
    ' existential: this is why nothing spells `exists`.
    AssertOptRefusal "optimize forms (s17): every night has a senior, through a derived rule", _
        "(fact (night n)) (rule (covered N) (night N) (assign N P) (senior P)) (require (covered N) (night N)) (query night)", _
        "a rule about every possible answer"
    ' optimize-roster-apart: "Brianna and Tyler never work the same
    ' shift", read from an Apart table.
    AssertOptRefusal "optimize forms (s17): two named people never share a shift", _
        "(fact (p a)) (forbid (assign S A) (assign S B) (apart A B)) (query p)", _
        "a rule about every possible answer"
    ' optimize-roster-apart again: "Brianna asked for Saturday off if we
    ' can" - a request is an objective term, never a rule.
    AssertOptRefusal "optimize forms (s17): a request is a preference, not a rule", _
        "(fact (p a)) (avoid (assign S P) (requests-off P S)) (query p)", _
        "says what a good answer looks like"
    ' optimize-seat-wedding: "the least total weight of wishes broken" -
    ' the weight is a Table column, which is why (cost ...) takes a
    ' variable and not only a number.
    AssertOptRefusal "optimize forms (s17): the wedding's weights come from a Table column", _
        "(fact (p a)) (minimize (split A B) (wishes A B W) (cost W)) (query p)", _
        "as small or as large as possible"
    ' optimize-config-laptop: "the laptop fits the dock, and the dock
    ' fits the monitor".
    AssertOptRefusal "optimize forms (s17): the laptop fits the dock", _
        "(fact (p a)) (require (fits A B) (pick ""laptop"" A) (pick ""dock"" B)) (query p)", _
        "a rule about every possible answer"
    ' optimize-quote-bike: "carbon wheels need the carbon frame", and
    ' "the quote includes the child seat".
    AssertOptRefusal "optimize forms (s17): carbon wheels need the carbon frame", _
        "(fact (p a)) (require (include ""carbon-frame"") (include ""carbon-wheels"")) (query p)", _
        "a rule about every possible answer"
    ' optimize-roster-kept, and OPTIMIZE.7's own case.
    AssertOptRefusal "optimize forms (s17): change last week's roster as little as possible", _
        "(fact (p a)) (fewest-changes-from Kept) (query p)", _
        "names a schedule to stay close to"
    ' optimize-sod-close's objective: "as few preparing roles as
    ' possible go to seniors" - a count of matching rows, which needs no
    ' cost at all (the default of 1).
    AssertOptRefusal "optimize forms (s17): as few preparing roles as possible go to seniors", _
        "(fact (p a)) (minimize (prepares T P) (senior P)) (query p)", _
        "as small or as large as possible"
    ' optimize-quote-bike's "the highest total margin", a maximum whose
    ' weight is a column, and optimize-config-laptop's "lowest total
    ' price" - the same one shape, twice.
    AssertOptRefusal "optimize forms (s17): the highest total margin", _
        "(fact (p a)) (maximize (include I) (catalogue I K Price Margin) (cost Margin)) (query p)", _
        "as small or as large as possible"
    ' Written order is priority: "least overtime first; then the fewest
    ' changes from Kept". Two objectives in one program, and the first
    ' one written is the one refused.
    AssertOptRefusal "optimize forms (s17): two objectives, and written order is priority", _
        "(fact (p a)) (minimize (overtime P) (person P)) (minimize (changed S P) (kept S P)) (query p)", _
        "as small or as large as possible"

    ' --- the five reserved result states, and their words -------------
    Report "optimize states: a refusal's words", _
           VLA_Optimize.OptimizeStatusWords(VLA_Optimize.VLA_OPTIMIZE_REFUSED) = _
           "this program could not be read; the cell says why", _
           "got: " & VLA_Optimize.OptimizeStatusWords(VLA_Optimize.VLA_OPTIMIZE_REFUSED)
    Report "optimize states: no schedule satisfies every rule", _
           VLA_Optimize.OptimizeStatusWords(VLA_Optimize.VLA_OPTIMIZE_NO_SCHEDULE) = _
           "no schedule satisfies every rule", _
           "got: " & VLA_Optimize.OptimizeStatusWords(VLA_Optimize.VLA_OPTIMIZE_NO_SCHEDULE)
    Report "optimize states: none found within the budget, but there may be one", _
           VLA_Optimize.OptimizeStatusWords(VLA_Optimize.VLA_OPTIMIZE_NONE_IN_BUDGET) = _
           "no schedule found within the budget; there may be one", _
           "got: " & VLA_Optimize.OptimizeStatusWords(VLA_Optimize.VLA_OPTIMIZE_NONE_IN_BUDGET)
    ' The name's own honesty rule, in the difference between these two.
    Report "optimize states: best found within the budget, NOT proven best", _
           VLA_Optimize.OptimizeStatusWords(VLA_Optimize.VLA_OPTIMIZE_BEST_IN_BUDGET) = _
           "best found within the budget, not proven best", _
           "got: " & VLA_Optimize.OptimizeStatusWords(VLA_Optimize.VLA_OPTIMIZE_BEST_IN_BUDGET)
    words = VLA_Optimize.OptimizeStatusWords(VLA_Optimize.VLA_OPTIMIZE_PROVEN_BEST)
    Report "optimize states: proven best says WHY nothing was searched", _
           InStr(1, words, "proven best", vbTextCompare) = 1 And _
           InStr(1, words, "makes no choices", vbTextCompare) > 0, "got: " & words
    Report "optimize states: the five states are five different numbers", _
           VLA_Optimize.VLA_OPTIMIZE_REFUSED = 0 And VLA_Optimize.VLA_OPTIMIZE_NO_SCHEDULE = 1 And _
           VLA_Optimize.VLA_OPTIMIZE_NONE_IN_BUDGET = 2 And VLA_Optimize.VLA_OPTIMIZE_BEST_IN_BUDGET = 3 And _
           VLA_Optimize.VLA_OPTIMIZE_PROVEN_BEST = 4, "the reserved numbers moved"
    ' A sixth state announces itself rather than answering blankly.
    Dim stateErr As Long
    On Error Resume Next
    Err.Clear
    VLA_Optimize.OptimizeStatusWords 5
    stateErr = Err.Number
    On Error GoTo 0
    Report "optimize states: a sixth state is refused, not answered blankly", stateErr <> 0, "no error raised"

    ' The three effort levels carry no number yet, on purpose: a number
    ' published before it means anything is a number users tune around.
    Report "optimize effort: the three levels are deliberately unset until OPTIMIZE.3/.6 measure", _
           VLA_Optimize.VLA_OPTIMIZE_WORK_QUICK = 0 And VLA_Optimize.VLA_OPTIMIZE_WORK_NORMAL = 0 And _
           VLA_Optimize.VLA_OPTIMIZE_WORK_THOROUGH = 0, "a work count was set before it was measured"
    Report "optimize effort: the default level is written down now", _
           VLA_Optimize.VLA_OPTIMIZE_EFFORT_DEFAULT = "normal", _
           "got: " & VLA_Optimize.VLA_OPTIMIZE_EFFORT_DEFAULT

    ' --- the worksheet surface, all three names -----------------------
    v = VLA_Optimize.OPTIMIZE("(fact (parent tom bob)) (fact (parent bob liz)) (query parent)")
    Report "optimize cell: a zero-choice program spills headers plus its rows", _
           ResultRowCount(v) = 3 And ResultColCount(v) = 2, _
           "shape " & ResultRowCount(v) & "x" & ResultColCount(v)
    v = VLA_Optimize.OPTIMISE("(fact (parent tom bob)) (fact (parent bob liz)) (query parent)")
    Report "optimize cell: OPTIMISE answers exactly what OPTIMIZE does", _
           ResultRowCount(v) = 3 And ResultColCount(v) = 2, _
           "shape " & ResultRowCount(v) & "x" & ResultColCount(v)
    v = VLA_Optimize.OPTIMIZE_STATUS("(fact (parent tom bob)) (query parent)")
    Report "optimize cell: OPTIMIZE_STATUS says proven best for a zero-choice program", _
           ResultTextStartsWith(v, "proven best"), "got: " & ResultDescribe(v)

    ' A query with no rows: THE HEADER ROW, WITH NOTHING UNDER IT. This
    ' is the "no schedule" answer's shape from the first day, and it
    ' needs no code of its own - it is what a DATALOG query with no rows
    ' already gives. A word in the cell would not spill, so a reader's
    ' N2# would become #REF! (measured live, DATALOG.15 step 12); a word
    ' in a ROW would become data. The words live in OPTIMIZE_STATUS.
    v = VLA_Optimize.OPTIMIZE("(fact (thing a)) (rule (none_here X) (thing X) (thing zzz)) (query none_here)")
    Report "optimize cell: no rows spills the header row and nothing under it", _
           ResultRowCount(v) = 1 And ResultCellIs(v, 1, 1, "X"), _
           "shape " & ResultRowCount(v) & "x" & ResultColCount(v)

    ' A refusal stays #OPTIMIZE! TEXT, where breaking the readers is the
    ' right thing to do - and the prefix names the function the user
    ' called even when the words are the shared engine's.
    v = VLA_Optimize.OPTIMIZE("(fact (elig s p)) (choose-exactly 2 (assign S P) (elig S P)) (query elig)")
    Report "optimize cell: a refusal is readable #OPTIMIZE! text", _
           ResultTextStartsWith(v, "#OPTIMIZE!"), "got: " & ResultDescribe(v)
    v = VLA_Optimize.OPTIMIZE("(fact (p a))")
    Report "optimize cell: a DATALOG refusal reaches the cell under the OPTIMIZE prefix", _
           ResultTextStartsWith(v, "#OPTIMIZE!"), "got: " & ResultDescribe(v)
    v = VLA_Optimize.OPTIMISE("(fact (p a))")
    Report "optimize cell: OPTIMISE refuses under the same prefix, not its own", _
           ResultTextStartsWith(v, "#OPTIMIZE!"), "got: " & ResultDescribe(v)
    v = VLA_Optimize.OPTIMIZE_STATUS("(fact (p a))")
    Report "optimize cell: OPTIMIZE_STATUS refuses as text too", _
           ResultTextStartsWith(v, "#OPTIMIZE!"), "got: " & ResultDescribe(v)
    ' A table argument that is not a range, worded as OPTIMIZE's own
    ' rather than as DATALOG's - the PROLOG/SQL precedent.
    v = VLA_Optimize.OPTIMIZE("(fact (p a)) (query p)", 42)
    Report "optimize cell: a table argument that is not a range names OPTIMIZE, not DATALOG", _
           ResultTextHas(v, "#OPTIMIZE!") And ResultTextHas(v, "every OPTIMIZE table argument"), _
           "got: " & ResultDescribe(v)

    ' `rule` is unchanged from DATALOG, and stays the word (the owner's
    ' call): a rule whose body reads chosen rows means what it says.
    ' What OPTIMIZE.5 refuses is narrower, and its refusal is RESERVED
    ' here with its wording - unreachable until a choice runs, which is
    ' exactly why the words are pinned now.
    Report "optimize forms: a recursive rule over the DATA still answers, as in DATALOG", _
           ParityVerdict("(fact (reports_to alice bob)) (fact (reports_to bob carol)) " & _
                         "(rule (indirect X Y) (reports_to X Y)) " & _
                         "(rule (indirect X Y) (reports_to X Z) (indirect Z Y)) (query indirect)") = "", _
           "a recursive rule over data differed between the engines"
End Sub

' ---------------------------------------------------------------------
'  OPTIMIZE.1: the memo (standing decision 1).
'
'  Three properties, and the third is the one that matters most:
'    1. Two cells asking the same question run ONE search.
'    2. A changed input is a different question, and runs another.
'    3. LOSING THE MEMO CHANGES NO ANSWER. OPTIMIZE.0's probe found
'       that a VBA project reset wipes every module variable, and that
'       after one the next structural change re-ran every UDF in the
'       workbook (C7). So the memo must tolerate being lost at any
'       moment; the pin is to lose it between two calls, which
'       OptimizeMemoClear does exactly.
' ---------------------------------------------------------------------
Private Sub TestOptimizeMemo()
    Dim a As Collection, b As Collection
    Dim runsBefore As Long
    Dim keyA As String, keyB As String
    Dim bases As Object
    Dim arr(1 To 1, 1 To 2) As Variant
    Dim i As Long
    Const prog As String = "(fact (parent tom bob)) (fact (parent bob liz)) (query parent)"

    VLA_Optimize.OptimizeMemoClear
    Report "optimize memo: a cleared memo holds nothing and has run nothing", _
           VLA_Optimize.OptimizeMemoCount() = 0 And VLA_Optimize.OptimizeMemoRuns() = 0, _
           "count " & VLA_Optimize.OptimizeMemoCount() & ", runs " & VLA_Optimize.OptimizeMemoRuns()

    Set a = VLA_Optimize.OptimizeRun(prog)
    Report "optimize memo: the first ask runs one search", VLA_Optimize.OptimizeMemoRuns() = 1, _
           "runs " & VLA_Optimize.OptimizeMemoRuns()
    Set b = VLA_Optimize.OptimizeRun(prog)
    Report "optimize memo: TWO CELLS, ONE SEARCH - the second ask runs none", _
           VLA_Optimize.OptimizeMemoRuns() = 1, "runs " & VLA_Optimize.OptimizeMemoRuns()
    Report "optimize memo: and the second ask gives the same answer", _
           ParityRelationKey(VLA_Runtime.VlaDictGet(a.Item(2), CStr(a.Item(1)))) = _
           ParityRelationKey(VLA_Runtime.VlaDictGet(b.Item(2), CStr(b.Item(1)))), _
           "the memoized answer differed"
    Report "optimize memo: the status comes from the memo too, at no cost", _
           CStr(a.Item(7)) = CStr(b.Item(7)) And VLA_Optimize.OptimizeMemoRuns() = 1, _
           "runs " & VLA_Optimize.OptimizeMemoRuns()

    ' THE RESET. Everything the memo held is gone, as a project reset
    ' leaves it - and the answer is the same, which is the whole
    ' contract: faster, never different.
    VLA_Optimize.OptimizeMemoClear
    Set b = VLA_Optimize.OptimizeRun(prog)
    Report "optimize memo: A RESET BETWEEN TWO CELLS costs a search and changes no answer", _
           VLA_Optimize.OptimizeMemoRuns() = 1 And _
           ParityRelationKey(VLA_Runtime.VlaDictGet(a.Item(2), CStr(a.Item(1)))) = _
           ParityRelationKey(VLA_Runtime.VlaDictGet(b.Item(2), CStr(b.Item(1)))), _
           "the answer changed across a reset"

    ' A different program is a different question.
    runsBefore = VLA_Optimize.OptimizeMemoRuns()
    VLA_Optimize.OptimizeRun "(fact (parent tom bob)) (query parent)"
    Report "optimize memo: a different program runs its own search", _
           VLA_Optimize.OptimizeMemoRuns() = runsBefore + 1, _
           "runs " & VLA_Optimize.OptimizeMemoRuns()

    ' A changed TABLE is a different question, even with identical
    ' rules text - which is the half of the key that a hash of the
    ' rules alone would have missed.
    arr(1, 1) = "alice": arr(1, 2) = "bob"
    Set bases = VLA_Runtime.VlaDictNew()
    VLA_Runtime.VlaDictSet bases, "reports_to", VLA_Relation.RelFromRange(arr)
    keyA = VLA_Optimize.OptimizeMemoKey("(query reports_to)", bases, VLA_Runtime.VlaDictNew())
    arr(1, 2) = "carol"
    Set bases = VLA_Runtime.VlaDictNew()
    VLA_Runtime.VlaDictSet bases, "reports_to", VLA_Relation.RelFromRange(arr)
    keyB = VLA_Optimize.OptimizeMemoKey("(query reports_to)", bases, VLA_Runtime.VlaDictNew())
    Report "optimize memo: one edited cell in one Table is a different key", keyA <> keyB, _
           "the same key for different data"

    ' And the same data is the same key, or the memo would never hit.
    arr(1, 2) = "bob"
    Set bases = VLA_Runtime.VlaDictNew()
    VLA_Runtime.VlaDictSet bases, "reports_to", VLA_Relation.RelFromRange(arr)
    Report "optimize memo: the same rules and the same data give the same key", _
           VLA_Optimize.OptimizeMemoKey("(query reports_to)", bases, VLA_Runtime.VlaDictNew()) = keyA, _
           "the key is not deterministic"

    ' A renamed COLUMN changes what a keyed rule resolves to, so it has
    ' to change the key even though no value moved.
    Dim hdrA As Object, hdrB As Object
    Set hdrA = VLA_Runtime.VlaDictNew()
    VLA_Runtime.VlaDictSet hdrA, "reports_to", TwoColPair("boss", "Boss")
    Set hdrB = VLA_Runtime.VlaDictNew()
    VLA_Runtime.VlaDictSet hdrB, "reports_to", TwoColPair("chief", "Chief")
    Report "optimize memo: a renamed column is a different key, though no value moved", _
           VLA_Optimize.OptimizeMemoKey("(query reports_to)", bases, hdrA) <> _
           VLA_Optimize.OptimizeMemoKey("(query reports_to)", bases, hdrB), _
           "a header rename left the key alone"

    ' A refusal is never memoized: it is cheap, and a stale one would
    ' be a lie.
    VLA_Optimize.OptimizeMemoClear
    For i = 1 To 3
        OptRefusalOf "(fact (p a)) (effort sideways) (query p)"
    Next i
    Report "optimize memo: a refusal is never memoized", VLA_Optimize.OptimizeMemoCount() = 0, _
           "count " & VLA_Optimize.OptimizeMemoCount()

    ' The cap holds, so a workbook of many OPTIMIZE cells cannot grow
    ' the session's memory without bound.
    VLA_Optimize.OptimizeMemoClear
    For i = 1 To 40
        VLA_Optimize.OptimizeRun "(fact (p a" & i & ")) (query p)"
    Next i
    Report "optimize memo: the memo is capped, and holds at most sixteen answers", _
           VLA_Optimize.OptimizeMemoCount() <= 16 And VLA_Optimize.OptimizeMemoRuns() = 40, _
           "count " & VLA_Optimize.OptimizeMemoCount() & ", runs " & VLA_Optimize.OptimizeMemoRuns()
    VLA_Optimize.OptimizeMemoClear
End Sub

' One (folded, original) column-name pair, in the Collection-of-pairs
' shape VLA_Relation.RangeColumnNames returns - the memo key reads it
' the same way the engines do.
Private Function TwoColPair(ByVal folded As String, ByVal original As String) As Collection
    Dim cols As Collection
    Set cols = New Collection
    Dim pair As Collection
    Set pair = New Collection
    pair.Add folded
    pair.Add original
    cols.Add pair
    Set TwoColPair = cols
End Function

' ---------------------------------------------------------------------
'  OPTIMIZE.1: the host pins - a real Table, a real Unicode cell, and
'  decision 1's "one search, many views" on a real sheet.
'
'  Why these need a live workbook at all, when everything above is
'  pure: every real bug this family's own MVP found showed up only
'  through the real =OPTIMIZE(...) path (this module's own header note),
'  and step 1's whole reason for existing is what a real cell's Value2
'  hands the encoder - which no hand-built array can prove.
' ---------------------------------------------------------------------
Private Sub TestOptimizeHostTable()
    Dim prior As Worksheet
    Set prior = ActiveSheet
    Dim q As String
    q = Chr$(34)

    On Error Resume Next
    ActiveWorkbook.Names("VlaOptSpill").Delete
    On Error GoTo 0

    VlaEnsureSheet "VlaOptimizeHostSheet"
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets("VlaOptimizeHostSheet")
    ws.Activate
    Do While ws.ListObjects.Count > 0
        ws.ListObjects(1).Delete
    Loop
    ws.Cells.Clear

    ' A two-edge chain, as TestDatalogHostTable's own fixture has it.
    ws.Range("A1:B1").Value = Array("P", "C")
    ws.Range("A2:B2").Value = Array("alice", "bob")
    ws.Range("A3:B3").Value = Array("bob", "carol")
    Dim loEdges As ListObject
    Set loEdges = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:B3"), , xlYes)
    loEdges.Name = "ReportsToOptTest1"

    ' A Dim rather than a Const: VBA's constant expressions are fussy
    ' about a continued concatenation, and this is not a hot path.
    Dim closure As String
    closure = "(rule (indirect X Y) (reportstoopttest1 X Y))" & _
              " (rule (indirect X Y) (reportstoopttest1 X Z) (indirect Z Y))" & _
              " (query indirect)"

    ' The same live Table, through both engines: the parity claim on the
    ' path a user actually takes.
    Dim viaDatalog As Variant, viaOptimize As Variant, viaOptimise As Variant
    viaDatalog = VLA_Datalog.DATALOG(closure, loEdges.Range)
    viaOptimize = VLA_Optimize.OPTIMIZE(closure, loEdges.Range)
    viaOptimise = VLA_Optimize.OPTIMISE(closure, loEdges.Range)
    Report "optimize host: a recursive rule over a LIVE Table answers 3 rows plus a header", _
           ResultRowCount(viaOptimize) = 4 And ResultColCount(viaOptimize) = 2, _
           "got: " & ResultDescribe(viaOptimize)
    Report "optimize host: OPTIMIZE's live answer has DATALOG's own shape", _
           ResultRowCount(viaOptimize) = ResultRowCount(viaDatalog) And _
           ResultColCount(viaOptimize) = ResultColCount(viaDatalog), _
           "OPTIMIZE " & ResultDescribe(viaOptimize) & " against DATALOG " & ResultDescribe(viaDatalog)
    Report "optimize host: and its header row is the rule's own variables", _
           ResultCellIs(viaOptimize, 1, 1, "X") And ResultCellIs(viaOptimize, 1, 2, "Y"), _
           "got: " & ResultDescribe(viaOptimize)
    Report "optimize host: OPTIMISE answers the same over the same Table", _
           ResultRowCount(viaOptimise) = 4 And ResultColCount(viaOptimise) = 2, _
           "got: " & ResultDescribe(viaOptimise)

    ' TWO CELLS, ONE SEARCH - over real Ranges, which is the case
    ' standing decision 1 is actually about.
    VLA_Optimize.OptimizeMemoClear
    VLA_Optimize.OPTIMIZE closure, loEdges.Range
    Dim runsAfterFirst As Long
    runsAfterFirst = VLA_Optimize.OptimizeMemoRuns()
    Dim statusCell As Variant
    statusCell = VLA_Optimize.OPTIMIZE_STATUS(closure, loEdges.Range)
    Report "optimize host: an answer cell and a status cell over one Table run ONE search", _
           runsAfterFirst = 1 And VLA_Optimize.OptimizeMemoRuns() = 1, _
           "runs " & runsAfterFirst & " then " & VLA_Optimize.OptimizeMemoRuns()
    Report "optimize host: and the status cell says proven best", _
           ResultTextStartsWith(statusCell, "proven best"), "got: " & ResultDescribe(statusCell)

    ' A RESET BETWEEN THE TWO CELLS, which OPTIMIZE.0's probe C7 says
    ' can happen at any moment: it costs a search and changes nothing.
    VLA_Optimize.OptimizeMemoClear
    Dim statusAfterReset As Variant
    statusAfterReset = VLA_Optimize.OPTIMIZE_STATUS(closure, loEdges.Range)
    Report "optimize host: a memo wiped between the two cells changes no answer", _
           ResultTextStartsWith(statusAfterReset, "proven best") And _
           VLA_Optimize.OptimizeMemoRuns() = 1, _
           "got: " & ResultDescribe(statusAfterReset) & ", runs " & VLA_Optimize.OptimizeMemoRuns()

    ' --- step 1's host pin: a real Unicode cell -----------------------
    '
    ' The one thing no hand-built array can prove: what a live cell's
    ' Value2 hands the key encoder. "Zoe" with a diaeresis is the exact
    ' string VlaSha256HexOfAsciiText could not carry - its StrConv goes
    ' through the ANSI code page - and the memo may never hand one
    ' workbook's answer to another workbook's inputs.
    ws.Range("D1").Value = "Name"
    ws.Range("D2").Value = "Zo" & ChrW(&HEB)
    ws.Range("D3").Value = "Bob"
    Dim loNames As ListObject
    Set loNames = ws.ListObjects.Add(xlSrcRange, ws.Range("D1:D3"), , xlYes)
    loNames.Name = "PeopleOptTest1"

    Dim keyUnicode As String, keyUnicodeAgain As String, keyAscii As String
    keyUnicode = LiveTableMemoKey("(query peopleopttest1)", loNames.Range)
    keyUnicodeAgain = LiveTableMemoKey("(query peopleopttest1)", loNames.Range)
    Report "optimize host: a Unicode cell's memo key is the same on two reads", _
           keyUnicode = keyUnicodeAgain And Len(keyUnicode) = 64, _
           "got " & keyUnicode & " then " & keyUnicodeAgain
    ws.Range("D2").Value = "Zoe"
    keyAscii = LiveTableMemoKey("(query peopleopttest1)", loNames.Range)
    Report "optimize host: Zoe with a diaeresis and plain Zoe are DIFFERENT keys", _
           keyUnicode <> keyAscii, "one key for two different workbooks' data"
    ws.Range("D2").Value = "Zo" & ChrW(&HEB)
    Report "optimize host: and putting the diaeresis back gives the first key again", _
           LiveTableMemoKey("(query peopleopttest1)", loNames.Range) = keyUnicode, _
           "the key is not a function of the data alone"

    ' The value itself survives the engine unchanged, so the key is
    ' guarding something the answer really carries.
    Dim uniAnswer As Variant
    uniAnswer = VLA_Optimize.OPTIMIZE("(query peopleopttest1)", loNames.Range)
    Report "optimize host: the Unicode value reaches the spill unchanged", _
           ResultCellIs(uniAnswer, 2, 1, "Zo" & ChrW(&HEB)), "got: " & ResultDescribe(uniAnswer)

    ' --- decision 1: ONE SEARCH, MANY VIEWS, on a real sheet ----------
    '
    ' An OPTIMIZE cell's own spill, named in the workbook's Names, read
    ' back by a second OPTIMIZE question BY ITS COLUMN NAMES. This is
    ' what DATALOG.15 was built ahead of this item for, and it is the
    ' whole shape of the decision: a schedule is searched for once and
    ' read as many times as you like.
    Dim spillCell As Object
    Set spillCell = ws.Range("G1")
    On Error Resume Next
    spillCell.Formula2 = "=OPTIMIZE(" & q & "(rule (pair X Y) (reportstoopttest1 X Y)) (query pair)" & q & ", ReportsToOptTest1)"
    On Error GoTo 0
    ws.Calculate
    Dim spillRange As Object
    On Error Resume Next
    Set spillRange = spillCell.SpillingToRange
    On Error GoTo 0
    If spillRange Is Nothing Then
        Report "optimize host: an OPTIMIZE formula spills (an Excel with dynamic arrays)", False, _
               "SpillingToRange was Nothing"
    Else
        Report "optimize host: an OPTIMIZE formula spills its header row plus its rows", _
               spillRange.Address = "$G$1:$H$3", "got " & spillRange.Address
        ws.Parent.Names.Add Name:="VlaOptSpill", RefersTo:="='VlaOptimizeHostSheet'!$G$1#"
        Dim readBack As Variant
        readBack = VLA_Optimize.OPTIMIZE("(rule (who P) (vlaoptspill (x P))) (query who)", spillRange)
        Report "optimize host: ONE SEARCH, MANY VIEWS - a second question reads that spill by its column names", _
               ResultRowCount(readBack) = 3 And ResultCellIs(readBack, 1, 1, "P"), _
               "got: " & ResultDescribe(readBack)
    End If

    On Error Resume Next
    ActiveWorkbook.Names("VlaOptSpill").Delete
    On Error GoTo 0
    VLA_Optimize.OptimizeMemoClear
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
    prior.Activate
End Sub

' The memo key for a live Range, built exactly as OPTIMIZE builds it -
' the same RelFromRange and RangeColumnNames the engine uses, so the
' key a pin computes is the key a formula would.
Private Function LiveTableMemoKey(ByVal program As String, ByVal rng As Object) As String
    Dim relations As Object
    Dim headerMap As Object
    Dim nm As String
    Dim colsOk As Boolean
    Dim cols As Collection
    Set relations = VLA_Runtime.VlaDictNew()
    Set headerMap = VLA_Runtime.VlaDictNew()
    Dim ok As Boolean, reason As String, detail As String
    nm = VLA_Relation.TableArgResolve(rng, ok, reason, detail)
    If Not ok Then
        LiveTableMemoKey = "(unnamed: " & reason & ")"
        Exit Function
    End If
    VLA_Runtime.VlaDictSet relations, nm, VLA_Relation.RelFromRange(rng)
    Set cols = VLA_Relation.RangeColumnNames(rng, colsOk)
    If colsOk Then VLA_Runtime.VlaDictSet headerMap, nm, cols
    LiveTableMemoKey = VLA_Optimize.OptimizeMemoKey(program, relations, headerMap)
End Function

' Builds a (folded, original) pair the same shape VLA_Relation.
' RangeColumnNames returns - VLA_Sql.SqlRun's own second argument.
Private Function SqlColPair(ByVal folded As String, ByVal original As String) As Collection
    Dim p As New Collection
    p.Add folded
    p.Add original
    Set SqlColPair = p
End Function

' ---------------------------------------------------------------------
'  PROLOG.1: VLA_Unify.UnifyOneWay's own mechanics, entirely pure - no
'  vocabulary, no host. Form literals built by hand, the same way this
'  file's own Datalog/SQL suites build Relation tuples by hand.
' ---------------------------------------------------------------------
Private Sub TestUnify()
    Dim bn As Collection, bv As Collection
    Dim ok As Boolean

    ' Bare slot binds the whole concrete atom.
    Set bn = New Collection: Set bv = New Collection
    ok = VLA_Unify.UnifyOneWay("{x}", "hello", bn, bv)
    Report "unify: bare slot matches and binds", _
           ok And bn.Count = 1 And CollItemIs(bn, 1, "x") And CollItemIs(bv, 1, "hello"), "got ok=" & ok

    ' Bare slot binds a whole LIST subform, not just an atom.
    Dim lst As New Collection
    lst.Add "a": lst.Add "b"
    Set bn = New Collection: Set bv = New Collection
    ok = VLA_Unify.UnifyOneWay("{x}", lst, bn, bv)
    Report "unify: bare slot binds a whole list subform", _
           ok And IsObject(bv.Item(1)) And bv.Item(1).Count = 2, "got ok=" & ok

    ' Literal atom must match text-for-text.
    Set bn = New Collection: Set bv = New Collection
    ok = VLA_Unify.UnifyOneWay("set!", "set!", bn, bv)
    Report "unify: matching literal atom succeeds with no bindings", ok And bn.Count = 0, "got ok=" & ok

    Set bn = New Collection: Set bv = New Collection
    ok = VLA_Unify.UnifyOneWay("set!", "get!", bn, bv)
    Report "unify: mismatched literal atom fails", Not ok, "expected failure"

    ' Glued slot: prefix{name}suffix strips cleanly.
    Set bn = New Collection: Set bv = New Collection
    ok = VLA_Unify.UnifyOneWay("make-{d}", "make-bold", bn, bv)
    Report "unify: glued slot strips its known prefix", _
           ok And CStr(bn.Item(1)) = "d" And CStr(bv.Item(1)) = "bold", "got ok=" & ok

    Set bn = New Collection: Set bv = New Collection
    ok = VLA_Unify.UnifyOneWay("xl{d}", "xldescending", bn, bv)
    Report "unify: glued slot strips its known prefix (xl{d})", ok And CStr(bv.Item(1)) = "descending", "got ok=" & ok

    Set bn = New Collection: Set bv = New Collection
    ok = VLA_Unify.UnifyOneWay("make-{d}", "unmake-bold", bn, bv)
    Report "unify: glued slot fails when the concrete atom's prefix doesn't match", Not ok, "expected failure"

    ' List recursion: same length, every element unifying.
    Dim tmpl As New Collection, conc As New Collection
    tmpl.Add "set!": tmpl.Add "{v}": tmpl.Add "{e}"
    conc.Add "set!": conc.Add "total": conc.Add "5"
    Set bn = New Collection: Set bv = New Collection
    ok = VLA_Unify.UnifyOneWay(tmpl, conc, bn, bv)
    Report "unify: list recursion matches head literal and binds two slots", _
           ok And bn.Count = 2 And CStr(bv.Item(1)) = "total" And CStr(bv.Item(2)) = "5", "got ok=" & ok

    Dim conc2 As New Collection
    conc2.Add "set!": conc2.Add "total"
    Set bn = New Collection: Set bv = New Collection
    ok = VLA_Unify.UnifyOneWay(tmpl, conc2, bn, bv)
    Report "unify: a shorter concrete list fails outright", Not ok, "expected failure"

    ' Repeated variable consistency - PROLOG.1's own real correction:
    ' the SAME slot name appearing twice in one template must bind the
    ' SAME value both times, or the whole match fails, never silently
    ' record the second occurrence and forget the first. Modeled on the
    ' real corpus shape this fix protects (english_expanded.vla's own
    ' "paste-values", (paste-values (range {r}) (range {r}))).
    Dim tmplRep As New Collection, concSame As New Collection, concDiff As New Collection
    tmplRep.Add "paste-values": tmplRep.Add "{r}": tmplRep.Add "{r}"
    concSame.Add "paste-values": concSame.Add "a1:b2": concSame.Add "a1:b2"
    concDiff.Add "paste-values": concDiff.Add "a1:b2": concDiff.Add "c3:d4"

    Set bn = New Collection: Set bv = New Collection
    ok = VLA_Unify.UnifyOneWay(tmplRep, concSame, bn, bv)
    Report "unify: a repeated slot binds when both occurrences agree", ok And bn.Count = 1, "got ok=" & ok

    Set bn = New Collection: Set bv = New Collection
    ok = VLA_Unify.UnifyOneWay(tmplRep, concDiff, bn, bv)
    Report "unify: a repeated slot refuses when the two occurrences disagree (PROLOG.1's own correction, not pre-existing behavior)", _
           Not ok, "expected failure"

    ' The multiple-embedded-slots refusal is a hard raise, not a False
    ' return - a template authoring defect, same discipline this
    ' module's own header names.
    Dim raised As Boolean
    raised = False
    On Error Resume Next
    Err.Clear
    Set bn = New Collection: Set bv = New Collection
    VLA_Unify.UnifyOneWay "make-{a}-{b}", "make-x-y", bn, bv
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "unify: an atom gluing two slots together is refused loudly, not silently mismatched", raised, "no error raised"
End Sub

' ---------------------------------------------------------------------
'  PROLOG.1: G-RENDER's own regression pin, through the real
'  EnglishRenderText entry point - proves UnifyForm's thin-client
'  refactor (VLA_English.bas) didn't change observable behavior for an
'  ordinary rule, AND pins the real correctness fix that rides along:
'  before this pass, rendering (paste-values (range "a1:b2")
'  (range "c3:d4")) - two DIFFERENT ranges - against the "convert range
'  {r:range} to values" rule (which repeats {r} twice) would have
'  silently matched using only the FIRST occurrence, producing a
'  sentence that drops the fact that source and destination differ.
'  Lives here (QUERY AND LOGIC's own suite), not VlaSelfTest, per the
'  owner's own correction: VLA_Unify is PROLOG.1, and this tranche's
'  tests stay together regardless of which shipped feature (G-RENDER)
'  a given piece of substrate also happens to serve.
' ---------------------------------------------------------------------
Private Sub TestGRenderUnify()
    EnglishResetGrammar
    EnglishAddPhrase "make cell {r:cell} {d:bold|italic}", "(make-{d} (range {r}))"
    Dim rendered As String, d As String
    On Error Resume Next
    Err.Clear
    rendered = EnglishRenderText("(make-bold (range ""a1""))")
    d = Err.Description
    On Error GoTo 0
    Report "g-render/unify: a glued-slot rule still round-trips through the thin-client walker", _
           InStr(1, rendered, "bold", vbTextCompare) > 0 And InStr(1, rendered, "a1", vbTextCompare) > 0, _
           "got: '" & rendered & "' err: " & d

    EnglishResetGrammar
    EnglishAddPhrase "convert range {r:range} to values", "(paste-values (range {r}) (range {r}))"

    On Error Resume Next
    Err.Clear
    rendered = EnglishRenderText("(paste-values (range ""b2:b9"") (range ""b2:b9""))")
    d = Err.Description
    On Error GoTo 0
    Report "g-render/unify: a repeated-slot rule renders when both occurrences genuinely agree", _
           InStr(1, rendered, "b2:b9", vbTextCompare) > 0, "got: '" & rendered & "' err: " & d

    Dim raised As Boolean
    raised = False
    On Error Resume Next
    Err.Clear
    rendered = EnglishRenderText("(paste-values (range ""a1:b2"") (range ""c3:d4""))")
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "g-render/unify: a repeated-slot rule refuses when the two occurrences genuinely differ, rather than silently dropping the mismatch", _
           raised, "expected a refusal (no matching rule), got: '" & rendered & "'"
    EnglishResetGrammar
End Sub

' ---------------------------------------------------------------------
'  PROLOG.2: VLA_Unify.UnifyTwoWay - real unification, variables on
'  either side, occurs check refused by name. Entirely pure, hand-built
'  form literals, no parser or vocabulary needed - the unifier "is
'  never handed a token stream" (this section's own PROLOG paragraph).
'  A shared envN/envT pair is threaded across MULTIPLE calls in several
'  cases below, on purpose: PROLOG.4's own SLD resolution will thread
'  one environment across a whole chain of subgoal unifications the
'  same way, so proving that threading works correctly here is part of
'  what this item is for, not an incidental test convenience.
' ---------------------------------------------------------------------
Private Sub TestUnifyTwoWay()
    Dim envN As Collection, envT As Collection
    Dim ok As Boolean

    ' Two matching ground constants unify with no new bindings.
    Set envN = New Collection: Set envT = New Collection
    ok = VLA_Unify.UnifyTwoWay("a", "a", envN, envT)
    Report "unify2: matching ground constants succeed with no bindings", ok And envN.Count = 0, "got ok=" & ok

    Set envN = New Collection: Set envT = New Collection
    ok = VLA_Unify.UnifyTwoWay("a", "b", envN, envT)
    Report "unify2: mismatched ground constants fail", Not ok, "expected failure"

    ' A free variable unifies with a ground constant and binds it.
    Set envN = New Collection: Set envT = New Collection
    ok = VLA_Unify.UnifyTwoWay("X", "a", envN, envT)
    Report "unify2: a free variable binds to a ground constant", _
           ok And envN.Count = 1 And CollItemIs(envN, 1, "X") And CollItemIs(envT, 1, "a"), "got ok=" & ok

    ' The SAME variable, dereferenced, must agree with itself and
    ' disagree with a conflicting constant - proves EnvWalkInto actually
    ' resolves an existing binding rather than re-binding blindly.
    ok = VLA_Unify.UnifyTwoWay("X", "a", envN, envT)
    Report "unify2: re-unifying a bound variable with its own value succeeds (dereferenced)", ok And envN.Count = 1, "got ok=" & ok

    ok = VLA_Unify.UnifyTwoWay("X", "b", envN, envT)
    Report "unify2: re-unifying a bound variable with a DIFFERENT value fails", Not ok, "expected failure"

    ' Two free variables unify with each other (var-var binding), then
    ' binding EITHER one to a ground value must resolve BOTH - the
    ' chain-following case a one-way walker (PROLOG.1) has no shape for
    ' at all.
    Set envN = New Collection: Set envT = New Collection
    ok = VLA_Unify.UnifyTwoWay("X", "Y", envN, envT)
    Report "unify2: two free variables unify with each other", ok, "got ok=" & ok

    ok = VLA_Unify.UnifyTwoWay("Y", "5", envN, envT)
    Report "unify2: binding the SECOND half of a var-var chain succeeds", ok, "got ok=" & ok

    ok = VLA_Unify.UnifyTwoWay("X", "5", envN, envT)
    Report "unify2: the FIRST half of the chain now resolves to the same value transitively", ok, "got ok=" & ok

    ok = VLA_Unify.UnifyTwoWay("X", "6", envN, envT)
    Report "unify2: the chain correctly refuses a conflicting value too", Not ok, "expected failure"

    ' A free variable binds to a whole compound term.
    Dim compoundTerm As New Collection
    compoundTerm.Add "f": compoundTerm.Add "a": compoundTerm.Add "b"
    Set envN = New Collection: Set envT = New Collection
    ok = VLA_Unify.UnifyTwoWay("X", compoundTerm, envN, envT)
    Report "unify2: a free variable binds to a whole compound term", _
           ok And IsObject(envT.Item(1)) And envT.Item(1).Count = 3, "got ok=" & ok

    ' Structural unification of two compound terms, one carrying a
    ' variable - (f X b) against (f a b) binds X to "a".
    Dim tmplTerm As New Collection, groundTerm As New Collection
    tmplTerm.Add "f": tmplTerm.Add "X": tmplTerm.Add "b"
    groundTerm.Add "f": groundTerm.Add "a": groundTerm.Add "b"
    Set envN = New Collection: Set envT = New Collection
    ok = VLA_Unify.UnifyTwoWay(tmplTerm, groundTerm, envN, envT)
    Report "unify2: a variable inside a compound term binds correctly", ok, "got ok=" & ok
    ok = VLA_Unify.UnifyTwoWay("X", "a", envN, envT)
    Report "unify2: the variable bound inside that compound term dereferences correctly afterward", ok, "got ok=" & ok

    ' Arity mismatch fails outright, no raise.
    Dim shortTerm As New Collection
    shortTerm.Add "f": shortTerm.Add "X"
    Set envN = New Collection: Set envT = New Collection
    ok = VLA_Unify.UnifyTwoWay(shortTerm, groundTerm, envN, envT)
    Report "unify2: an arity mismatch between two compound terms fails outright", Not ok, "expected failure"

    ' Both sides carrying a variable in the SAME position, simultaneously
    ' - the case that proves this is genuinely two-way, not one-way run
    ' twice: (f X) unify (f Y), then X and Y must be linked such that
    ' binding one resolves the other.
    Dim leftVar As New Collection, rightVar As New Collection
    leftVar.Add "f": leftVar.Add "X"
    rightVar.Add "f": rightVar.Add "Y"
    Set envN = New Collection: Set envT = New Collection
    ok = VLA_Unify.UnifyTwoWay(leftVar, rightVar, envN, envT)
    Report "unify2: both sides carrying a variable in the same position unify", ok, "got ok=" & ok
    ok = VLA_Unify.UnifyTwoWay("X", "9", envN, envT)
    Report "unify2: binding X afterward succeeds", ok, "got ok=" & ok
    ok = VLA_Unify.UnifyTwoWay("Y", "9", envN, envT)
    Report "unify2: Y - linked to X, never touched directly - now agrees with X's own value", ok, "got ok=" & ok
    ok = VLA_Unify.UnifyTwoWay("Y", "10", envN, envT)
    Report "unify2: Y still refuses a value that conflicts with what X was bound to", Not ok, "expected failure"

    ' Occurs check: a variable can never bind to a term containing
    ' itself - refused by name, this project's own already-decided
    ' divergence from most real Prolog implementations (never skipped
    ' for speed).
    Dim cyclicTerm As New Collection
    cyclicTerm.Add "f": cyclicTerm.Add "X"
    Dim raised As Boolean
    raised = False
    Set envN = New Collection: Set envT = New Collection
    On Error Resume Next
    Err.Clear
    VLA_Unify.UnifyTwoWay "X", cyclicTerm, envN, envT
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "unify2: a variable unifying with a term that contains itself is refused (occurs check)", raised, "no error raised"

    ' Occurs check must not false-positive when the SAME variable name
    ' merely appears elsewhere but not inside the term actually being
    ' bound to it.
    Dim otherVarTerm As New Collection
    otherVarTerm.Add "f": otherVarTerm.Add "Y"
    Set envN = New Collection: Set envT = New Collection
    raised = False
    On Error Resume Next
    Err.Clear
    ok = VLA_Unify.UnifyTwoWay("X", otherVarTerm, envN, envT)
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "unify2: occurs check does not false-positive on a DIFFERENT variable's own name", ok And Not raised, "got ok=" & ok & " raised=" & raised
End Sub

' ---------------------------------------------------------------------
'  PROLOG.3: VLA_Prolog.PROLOG - the real worksheet function, called
'  directly (entirely pure - no ParamArray table argument in any of
'  these cases needs a live workbook at all). Covers ground facts and
'  conjunctive queries, cross-conjunct binding intersection, multiple
'  solutions in fact-authored order, the pure-yes/no boolean-scalar
'  degenerate shape, a compound-term argument round-tripping through
'  VlaWriteForm, bag semantics (a duplicate fact produces a duplicate
'  row), and every named parse-time refusal.
' ---------------------------------------------------------------------
Private Sub TestProlog()
    Dim result As Variant

    ' Single conjunct, one free variable, one solution.
    result = VLA_Prolog.PROLOG("(fact (parent tom bob)) (fact (parent bob liz)) (query (parent tom X))")
    Report "prolog: single-conjunct query returns a header row plus one solution row", _
           ResultRowCount(result) = 2 And ResultColCount(result) = 1 And ResultCellIs(result, 1, 1, "X") And ResultCellIs(result, 2, 1, "bob"), _
           "got shape/contents mismatch"

    ' Two conjuncts sharing a variable - the whole point of a
    ' conjunctive query - must intersect bindings across both.
    result = VLA_Prolog.PROLOG("(fact (parent tom bob)) (fact (parent bob liz)) (query (parent tom Y) (parent Y Z))")
    Report "prolog: two conjuncts sharing a variable intersect bindings correctly", _
           ResultRowCount(result) = 2 And ResultColCount(result) = 2 And ResultCellIs(result, 2, 1, "bob") And ResultCellIs(result, 2, 2, "liz"), _
           "got: Y=" & ResultDescribe(result) & " Z=" & ResultDescribe(result)

    ' Multiple solutions, in fact-authored order (a bag, not a set).
    result = VLA_Prolog.PROLOG("(fact (parent tom bob)) (fact (parent tom ann)) (query (parent tom X))")
    Report "prolog: multiple matching facts produce multiple solution rows, in authored order", _
           ResultRowCount(result) = 3 And ResultCellIs(result, 2, 1, "bob") And ResultCellIs(result, 3, 1, "ann"), _
           "got " & ResultRowCount(result) - 1 & " rows"

    ' Bag semantics: the SAME fact authored twice produces two identical
    ' rows - real Prolog's own behavior, not deduped the way DATALOG's
    ' fixpoint-derived relations are.
    result = VLA_Prolog.PROLOG("(fact (parent tom bob)) (fact (parent tom bob)) (query (parent tom X))")
    Report "prolog: a fact authored twice produces two identical solution rows (bag, not set)", _
           ResultRowCount(result) = 3 And ResultCellIs(result, 2, 1, "bob") And ResultCellIs(result, 3, 1, "bob"), _
           "got " & ResultRowCount(result) - 1 & " rows"

    ' A query with no free variables at all collapses to a boolean
    ' scalar - TRUE when at least one solution exists, FALSE otherwise.
    result = VLA_Prolog.PROLOG("(fact (parent tom bob)) (query (parent tom bob))")
    Report "prolog: a fully-ground query with no free variables returns TRUE, not an array", _
           ResultBoolIs(result, True), "got " & TypeName(result) & " " & result

    result = VLA_Prolog.PROLOG("(fact (parent tom bob)) (query (parent tom ann))")
    Report "prolog: a fully-ground query that matches nothing returns FALSE", _
           ResultBoolIs(result, False), "got " & TypeName(result) & " " & result

    ' A compound-term argument (the boundary DATALOG stays clear of) -
    ' bound and rendered back through VLA.VlaWriteForm.
    result = VLA_Prolog.PROLOG("(fact (likes tom (color red))) (query (likes tom X))")
    Report "prolog: a variable bound to a compound term renders via VlaWriteForm", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, "(color red)"), "got: " & ResultDescribe(result)

    ' A predicate with zero arguments (nullary) is legal here - unlike
    ' DATALOG, this engine never sizes an array off arity.
    result = VLA_Prolog.PROLOG("(fact (raining)) (query (raining))")
    Report "prolog: a nullary (zero-argument) predicate is legal and answers TRUE", _
           ResultBoolIs(result, True), "got " & TypeName(result) & " " & result

    ' Refusals, all parse-time (or, for the one worksheet-signature
    ' check, before parsing even starts). PROLOG (the worksheet
    ' function) catches its own RaiseMsg internally and returns
    ' "#PROLOG! <rendered message text>" - Err.Description holds the
    ' RENDERED template (SubstituteSlots' own output, VLA_Messages.bas),
    ' never the raw message id itself, so these check real wording
    ' pulled from each message's own template text, not an id string
    ' that would never actually appear in the cell.
    Report "prolog: a variable inside a (fact ...) is refused - facts must be ground", _
           InStr(1, CStr(VLA_Prolog.PROLOG("(fact (parent tom X)) (query (parent tom Y))")), "facts must be fully ground", vbTextCompare) > 0, _
           "got: " & VLA_Prolog.PROLOG("(fact (parent tom X)) (query (parent tom Y))")

    Report "prolog: the same predicate used with two different arities is refused", _
           InStr(1, CStr(VLA_Prolog.PROLOG("(fact (p a)) (fact (p a b)) (query (p X))")), "must have the same arity", vbTextCompare) > 0, _
           "got: " & VLA_Prolog.PROLOG("(fact (p a)) (fact (p a b)) (query (p X))")

    Report "prolog: a program with no (query ...) is refused, not silently answering nothing", _
           InStr(1, CStr(VLA_Prolog.PROLOG("(fact (p a))")), "no (query ...)", vbTextCompare) > 0, _
           "got: " & VLA_Prolog.PROLOG("(fact (p a))")

    Report "prolog: two (query ...) forms in one program is refused as ambiguous", _
           InStr(1, CStr(VLA_Prolog.PROLOG("(fact (p a)) (query (p X)) (query (p Y))")), "(query ...) forms", vbTextCompare) > 0, _
           "got: " & VLA_Prolog.PROLOG("(fact (p a)) (query (p X)) (query (p Y))")

    ' PROLOG.6 shipped: tables are real now - a non-Range table argument
    ' (a plain Long, not a cell reference at all) is refused by name as
    ' not-a-range, TableArgResolve's own shared category code, rather
    ' than the PROLOG.3-era "not yet supported" wording this test
    ' originally pinned.
    Report "prolog: a non-Range table argument is refused by name, not silently ignored (PROLOG.6)", _
           InStr(1, CStr(VLA_Prolog.PROLOG("(fact (p a)) (query (p X))", 1)), "must be a cell range", vbTextCompare) > 0, _
           "got: " & VLA_Prolog.PROLOG("(fact (p a)) (query (p X))", 1)
End Sub

' ---------------------------------------------------------------------
'  PROLOG.4: VLA_Prolog.PROLOG - (rule head body...) forms,
'  unification-driven SLD resolution with backtracking, no cut. Covers:
'  a single non-recursive rule (both its boolean-scalar and free-variable
'  shapes); a predicate defined by a FACT and a RULE together, both
'  firing in written order (the concrete case PROLOG.3's own fact-only
'  dict had no answer for); real recursion - the classic ancestor/parent
'  transitive closure over a 3-fact chain, asserted against every
'  solution's own value AND order (not just a row count), the shape most
'  likely to silently mis-derive if clause-variable freshening were ever
'  wrong (two overlapping recursive invocations sharing one variable by
'  accident would either lose solutions or bind something to the wrong
'  value, not merely miscount rows); the malformed-rule-shape refusal;
'  and the resolution-step ceiling, deliberately exercised for REAL - a
'  genuinely non-terminating rule run to actual exhaustion, unlike
'  SQL.7/DATALOG's own round-ceiling tests, which skip exhausting their
'  own (far larger, far more expensive-per-round) limits and trust the
'  ceiling check's own structural simplicity instead. This one is cheap
'  enough (each PROLOG resolution step is one comparison plus a shallow
'  unify, not a full relational round) to actually run to the ceiling,
'  and doing so is the ONLY way to empirically learn whether
'  PROLOG_MAX_STEPS's own conservative value is already too high for
'  VBA's real native call stack - confirmed live, not hypothetical: the
'  first value tried (1000) genuinely overflowed on this exact test,
'  caught and lowered to 200 (see VLA_Prolog.bas's own declarations-
'  section comment for the incident). Two more pins added the same
'  session, owner-requested, promoting the owner's own live-verification
'  handoff examples into permanent regression coverage: a CYCLIC graph's
'  own free-variable reachability query (the textbook non-termination
'  shape every intro Prolog course teaches - real Prolog has genuinely
'  infinite solutions here too, not an artifact of this engine's own
'  budget), proving the ceiling generalizes beyond the single-clause
'  (loop X):-(loop X) shape above to a real two-clause recursive
'  predicate whose non-termination comes from the DATA, not the rule;
'  and occurs-check reached through a real freshened rule head (a
'  variable appearing both bare and wrapped in a compound term at two
'  argument positions, unified against a query sharing one free variable
'  across both), proving the check survives the whole PROLOG.4 pipeline
'  rather than only the hand-built primitive TestUnifyTwoWay (PROLOG.2)
'  already covers directly.
' ---------------------------------------------------------------------
Private Sub TestPrologRules()
    Dim result As Variant

    ' A single non-recursive rule over facts already in scope.
    result = VLA_Prolog.PROLOG( _
        "(fact (parent tom bob)) (fact (parent bob liz)) " & _
        "(rule (grandparent X Z) (parent X Y) (parent Y Z)) " & _
        "(query (grandparent tom liz))")
    Report "prolog.4: a non-recursive rule proves a fully-ground query TRUE", _
           ResultBoolIs(result, True), "got " & TypeName(result) & " " & result

    result = VLA_Prolog.PROLOG( _
        "(fact (parent tom bob)) (fact (parent bob liz)) " & _
        "(rule (grandparent X Z) (parent X Y) (parent Y Z)) " & _
        "(query (grandparent tom Z))")
    Report "prolog.4: a non-recursive rule's own free variable resolves through its body", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, "liz"), "got: " & ResultDescribe(result)

    ' A predicate defined by a FACT and a RULE together - both fire, in
    ' written order, with no special-casing needed anywhere in the
    ' solver (this module's own header has why).
    result = VLA_Prolog.PROLOG( _
        "(fact (likes tom pizza)) (fact (foodie ann)) " & _
        "(rule (likes X sushi) (foodie X)) " & _
        "(query (likes X Y))")
    Report "prolog.4: a predicate defined by both a fact and a rule tries both, in written order", _
           ResultRowCount(result) = 3 And ResultCellIs(result, 2, 1, "tom") And ResultCellIs(result, 2, 2, "pizza") _
           And ResultCellIs(result, 3, 1, "ann") And ResultCellIs(result, 3, 2, "sushi"), _
           "got " & (ResultRowCount(result) - 1) & " rows"

    ' Real recursion: the classic ancestor/parent transitive closure over
    ' a 3-fact chain (tom -> bob -> liz -> ann). Asserted against every
    ' solution's own value AND order, not just a row count - the
    ' clause-freshening machinery is exactly what a wrong answer here
    ' would implicate, since ancestor/2 invokes itself recursively and
    ' each invocation's own X/Y/Z must never collide with its parent
    ' call's.
    result = VLA_Prolog.PROLOG( _
        "(fact (parent tom bob)) (fact (parent bob liz)) (fact (parent liz ann)) " & _
        "(rule (ancestor X Y) (parent X Y)) " & _
        "(rule (ancestor X Y) (parent X Z) (ancestor Z Y)) " & _
        "(query (ancestor tom Y))")
    Report "prolog.4: recursive ancestor/parent finds every generation, in derivation order", _
           ResultRowCount(result) = 4 And ResultCellIs(result, 2, 1, "bob") And ResultCellIs(result, 3, 1, "liz") And ResultCellIs(result, 4, 1, "ann"), _
           "got " & (ResultRowCount(result) - 1) & " rows: " & JoinColumn(result, 1)

    ' A negative case over the SAME recursive program - tom is not his
    ' own ancestor, and the search must terminate (not loop forever)
    ' finding out.
    result = VLA_Prolog.PROLOG( _
        "(fact (parent tom bob)) (fact (parent bob liz)) (fact (parent liz ann)) " & _
        "(rule (ancestor X Y) (parent X Y)) " & _
        "(rule (ancestor X Y) (parent X Z) (ancestor Z Y)) " & _
        "(query (ancestor tom tom))")
    Report "prolog.4: a fully-ground recursive query that doesn't hold terminates and returns FALSE", _
           ResultBoolIs(result, False), "got " & TypeName(result) & " " & result

    ' Refusals.
    Report "prolog.4: a (rule ...) with no body at all is refused - that's just a fact", _
           InStr(1, CStr(VLA_Prolog.PROLOG("(rule (p X)) (query (p X))")), "just a fact", vbTextCompare) > 0, _
           "got: " & VLA_Prolog.PROLOG("(rule (p X)) (query (p X))")

    ' The resolution-step ceiling, exercised for real: (loop X) :- (loop
    ' X) has no base case and can never terminate on its own. A clean
    ' "#PROLOG! ... resolution steps ..." refusal here is this item's
    ' own termination guarantee actually firing; a raw VBA "Out of stack
    ' space" crash instead would mean PROLOG_MAX_STEPS (VLA_Prolog.bas's
    ' own module-level Const, 200 as of this writing, lowered live from
    ' an initial 1000 that genuinely did overflow VBA's native call stack
    ' on this exact test - see VLA_Prolog.bas's own declarations-section
    ' comment for the full incident) is already too high for this
    ' machine's real native call-stack depth and needs lowering further.
    Dim ceilingResult As String
    ceilingResult = CStr(VLA_Prolog.PROLOG("(rule (loop X) (loop X)) (query (loop a))"))
    ' PROLOG.28: re-pointed from the retired step ceiling's text. A rule that
    ' calls itself grows DEPTH, so it is the depth refusal that answers now -
    ' and never the work one, which would blame a big Table.
    Report "prolog.4: a genuinely non-terminating rule is refused by the depth ceiling, not left to hang or crash", _
           InStr(1, ceilingResult, "rules deep", vbTextCompare) > 0 And InStr(1, ceilingResult, "facts and rules", vbTextCompare) = 0, "got: " & ceilingResult

    ' A cyclic graph's own naive transitive closure - the textbook
    ' Prolog gotcha every intro course teaches: a free-variable query
    ' over a cycle (a->b->c->a) has genuinely infinitely many solutions
    ' in REAL Prolog too (every path length is a distinct derivation,
    ' even though the endpoint value only ever cycles through three
    ' values) - not an artifact of this engine's own step-budget choice.
    ' A second, independent proof that the ceiling generalizes beyond the
    ' single-clause (loop X):-(loop X) shape above: this one recurses
    ' through TWO clauses (a base case that can succeed, and a recursive
    ' case) and only fails to terminate because the DATA, not the rule
    ' shape, contains a cycle.
    Dim cycleResult As String
    cycleResult = CStr(VLA_Prolog.PROLOG( _
        "(fact (edge a b)) (fact (edge b c)) (fact (edge c a)) " & _
        "(rule (path X Y) (edge X Y)) " & _
        "(rule (path X Y) (edge X Z) (path Z Y)) " & _
        "(query (path a Y))"))
    ' PROLOG.28: data that loops back on itself is the depth refusal's own
    ' second named cause.
    Report "prolog.4: a cyclic graph's own free-variable reachability query is refused by the depth ceiling, not left to hang - a real, not just single-clause, non-termination shape", _
           InStr(1, cycleResult, "rules deep", vbTextCompare) > 0 And InStr(1, cycleResult, "facts and rules", vbTextCompare) = 0, "got: " & cycleResult

    ' Occurs-check, reached through real clause freshening rather than a
    ' hand-built form literal (TestUnifyTwoWay's own pin, PROLOG.2,
    ' covers the primitive directly) - proving the check survives the
    ' whole PROLOG.4 pipeline: a rule head repeating the SAME variable
    ' bare in one argument position and wrapped in a compound term in
    ' another ((mirror X (box X))), queried with both positions unified
    ' to one shared free variable ((mirror Y Y)) so the two occurrences
    ' collide during unification, not before it.
    Dim occursResult As String
    occursResult = CStr(VLA_Prolog.PROLOG( _
        "(fact (anything ok)) " & _
        "(rule (mirror X (box X)) (anything ok)) " & _
        "(query (mirror Y Y))"))
    Report "prolog.4: occurs-check survives real clause freshening, not just the hand-built primitive", _
           InStr(1, occursResult, "contains itself", vbTextCompare) > 0, "got: " & occursResult

    ' REGRESSION (found while building PROLOG.5.3, not by any test until
    ' now): a query's own free variable resolving to a compound term that
    ' ITSELF still contains an unresolved nested variable - a rule head
    ' repeating a variable both bare (X) and wrapped in a compound
    ' argument ((box X)), with the wrapped copy bound at head-unification
    ' time and the bare copy bound SEPARATELY, later, by a body goal
    ' (num X). Before ResolveTermDeep (VLA_Prolog.bas), the OLD shallow
    ' EnvWalkInto-only base case would have rendered Z as the raw
    ' freshened placeholder text "(box X#<n>)" instead of the real value
    ' "(box 7)", since it only ever chased a bare variable-to-variable
    ' chain and never re-derefed a nested position inside whatever it
    ' landed on.
    result = VLA_Prolog.PROLOG( _
        "(fact (num 7)) (rule (r X (box X)) (num X)) (query (r Y Z))")
    Report "prolog.4 REGRESSION: a free variable resolving to a compound term with a still-unresolved nested variable renders the real value, not the raw variable name", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, "7") And ResultCellIs(result, 2, 2, "(box 7)"), _
           "got: Y=" & ResultDescribe(result) & " Z=" & ResultDescribe(result)
End Sub

' ---------------------------------------------------------------------
'  PROLOG.5.1: VLA_Prolog.PROLOG - `is`/arithmetic. Covers: a rule
'  computing a fresh variable via a rule body chain (fact -> is);
'  nested arithmetic evaluated bottom-up in a pure query, the shape
'  DATALOG's own flat (let Z (op X Y)) structurally cannot express;
'  `is` checking an ALREADY-bound target via ordinary unification, both
'  the matching and the mismatching case - real Prolog's own `is/2`
'  semantics, free from reusing UnifyTwoWay rather than a bespoke
'  bind-only mechanism; two chained `is` calls in ONE rule body, proving
'  the first call's own result is visible to the second through the
'  same threaded env; and every named refusal - an unbound variable
'  reaching evaluation, a non-numeric ground value (reached through a
'  rule, not a hand-typed literal - the runtime-not-just-static-shape
'  case), divide-by-zero, an unrecognized operator and a wrong-arity
'  operand count (both parse-time), a malformed (is ...) shape, and
'  reserved predicate names (is/not/findall/! all tried). PROLOG.5.4
'  (VLA_Tests_Query.TestPrologCut, below) is where cut itself is now
'  tested; this file's own bare-! smoke check here only proves cut is no
'  longer a not-yet-supported refusal at parse time.
' ---------------------------------------------------------------------
Private Sub TestPrologArithmetic()
    Dim result As Variant

    ' A rule computing a fresh variable through a fact -> is chain.
    result = VLA_Prolog.PROLOG("(fact (base 10)) (rule (double X Y) (base X) (is Y (* X 2))) (query (double X Y))")
    Report "prolog.5.1: is computes a fresh variable through a rule body chain", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, "10") And ResultCellIs(result, 2, 2, "20"), _
           "got: X=" & ResultDescribe(result) & " Y=" & ResultDescribe(result)

    ' Nested arithmetic, evaluated bottom-up in a pure query - the shape
    ' DATALOG's own flat (let Z (op X Y)) structurally can't express at
    ' all, since DATALOG forbids compound terms entirely.
    result = VLA_Prolog.PROLOG("(query (is X (+ (* 2 3) 4)))")
    Report "prolog.5.1: nested arithmetic evaluates bottom-up", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, "10"), "got: " & ResultDescribe(result)

    ' is checking an ALREADY-bound target via ordinary unification - real
    ' Prolog's own is/2 semantics, free from reusing UnifyTwoWay rather
    ' than DATALOG's own bind-only (let ...), which requires a brand-new
    ' variable and cannot express this at all.
    result = VLA_Prolog.PROLOG("(query (is 10 (+ 4 6)))")
    Report "prolog.5.1: is against an already-ground target succeeds when the value matches", _
           ResultBoolIs(result, True), "got " & TypeName(result) & " " & result

    result = VLA_Prolog.PROLOG("(query (is 10 (+ 4 5)))")
    Report "prolog.5.1: is against an already-ground target fails when the value doesn't match", _
           ResultBoolIs(result, False), "got " & TypeName(result) & " " & result

    ' Two chained is calls in ONE rule body - the first call's own result
    ' must be visible to the second through the same threaded env.
    result = VLA_Prolog.PROLOG("(rule (twice-plus-one X Y) (is Temp (* X 2)) (is Y (+ Temp 1))) (query (twice-plus-one 5 Y))")
    Report "prolog.5.1: two chained is calls in one rule body share the same threaded bindings", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, "11"), "got: " & ResultDescribe(result)

    ' Refusals.
    Dim r As String

    r = CStr(VLA_Prolog.PROLOG("(query (is X (+ Y 1)))"))
    Report "prolog.5.1: an unbound variable reaching evaluation is refused", _
           InStr(1, r, "unbound variable", vbTextCompare) > 0, "got: " & r

    r = CStr(VLA_Prolog.PROLOG("(fact (label hello)) (rule (compute X) (label Y) (is X (+ Y 1))) (query (compute X))"))
    Report "prolog.5.1: a non-numeric ground value reached through a rule (not a hand-typed literal) is refused", _
           InStr(1, r, "isn't one", vbTextCompare) > 0, "got: " & r

    r = CStr(VLA_Prolog.PROLOG("(query (is X (/ 5 0)))"))
    Report "prolog.5.1: divide-by-zero is refused", _
           InStr(1, r, "divide by zero", vbTextCompare) > 0, "got: " & r

    ' PROLOG.17 RE-POINTED THIS, and the re-point is the point. It read
    ' `(mod 5 2)` and was correct for four items - until PROLOG.17 made
    ' `mod` a real operator, at which point it asserted that a WORKING
    ' operator does not work. Nothing about it looked stale: the expected
    ' text is still perfectly emittable, and only the INPUT had changed
    ' meaning underneath it. That class is not what rule G of
    ' check_test_assertion_safety.ps1 covers (rule G is about a RENDERING
    ' the writer can no longer emit), so PROLOG.17 added rule E to
    ' tools/check_prolog_arith_operators.ps1 to catch it mechanically -
    ' run RED against this very line before it was changed.
    '
    ' `%` is the replacement because it is a spelling a spreadsheet user
    ' plausibly reaches for and this engine deliberately does NOT have.
    r = CStr(VLA_Prolog.PROLOG("(query (is X (% 5 2)))"))
    Report "prolog.5.1: an unrecognized arithmetic operator is refused at parse time", _
           InStr(1, r, "isn't an arithmetic operator", vbTextCompare) > 0, "got: " & r
    ' ...and the twin that makes the re-point honest rather than a pin
    ' that quietly vanished: the operator it USED to name now computes,
    ' and the VALUE is asserted, not merely the success.
    result = VLA_Prolog.PROLOG("(query (is X (mod 5 2)))")
    Report "prolog.17: ...while (mod 5 2), which this test used to cite as unrecognized, now gives exactly 1", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "1"), "got: " & ResultDescribe(result)

    r = CStr(VLA_Prolog.PROLOG("(query (is X (+ 1 2 3)))"))
    Report "prolog.5.1: an arithmetic operator with the wrong number of operands is refused", _
           InStr(1, r, "exactly two operands", vbTextCompare) > 0, "got: " & r

    r = CStr(VLA_Prolog.PROLOG("(query (is X))"))
    Report "prolog.5.1: a malformed (is ...) shape is refused", _
           InStr(1, r, "exactly two arguments", vbTextCompare) > 0, "got: " & r

    r = CStr(VLA_Prolog.PROLOG("(fact (is a b)) (query (p X))"))
    Report "prolog.5.1: 'is' is refused as a predicate name in a (fact ...)", _
           InStr(1, r, "reserved word", vbTextCompare) > 0, "got: " & r

    r = CStr(VLA_Prolog.PROLOG("(rule (not X) (is X 1)) (query (p X))"))
    Report "prolog.5.1: 'not' is refused as a predicate name in a (rule ...) - reserved for PROLOG.5.2", _
           InStr(1, r, "reserved word", vbTextCompare) > 0, "got: " & r

    r = CStr(VLA_Prolog.PROLOG("(fact (findall a)) (query (p X))"))
    Report "prolog.5.1: 'findall' is refused as a predicate name - reserved for PROLOG.5.3", _
           InStr(1, r, "reserved word", vbTextCompare) > 0, "got: " & r

    ' PROLOG.5.4 shipped: a bare ! is now a legal, if here-trivial, cut -
    ' full scoping/pruning coverage lives in TestPrologCut, below; this
    ' is only a smoke check that it's no longer refused at parse time.
    result = VLA_Prolog.PROLOG("(fact (p a)) (query (p X) !)")
    Report "prolog.5.1: a bare ! (cut) is accepted, not refused, from PROLOG.5.4 onward", _
           ResultCol1Is(result, "a"), "got: " & ResultDescribe(result)

    ' ---- Coverage-completion pass, owner-requested: every branch of the
    ' new PROLOG.5.1 code touched by at least one call, not just the
    ' shapes the earlier tests happened to exercise. Two of these
    ' (marked below) are regression pins for the exact two bugs this
    ' item's own build found and fixed by code review, before either
    ' one had ever actually been exercised by a real call - the most
    ' important gap this pass closes, not just a count-padding exercise.

    ' All four operators, standalone and clean (earlier tests only ever
    ' exercised + and * directly). ResultCol1Is/ResultDescribe (below
    ' TestPrologArithmetic, this file's own new helpers) both check
    ' IsArray via a nested If, never "IsArray(result) And
    ' CStr(result(2,1))=..." in one expression, AND never concatenate the
    ' whole array `result` itself into a string (`"x" & result` on an
    ' ARRAY raises its own type mismatch, distinct from indexing one) -
    ' both are the exact class of bug this session's own ancestor-test
    ' crash and this item's own ValidateBodyItem fix already caught
    ' elsewhere, nearly reintroduced right here, twice, while writing
    ' these very regression pins.
    result = VLA_Prolog.PROLOG("(query (is X (+ 3 4)))")
    Report "prolog.5.1: + standalone", ResultCol1Is(result, "7"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (is X (- 10 3)))")
    Report "prolog.5.1: - standalone", ResultCol1Is(result, "7"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (is X (* 6 7)))")
    Report "prolog.5.1: * standalone", ResultCol1Is(result, "42"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (is X (/ 20 4)))")
    Report "prolog.5.1: / standalone", ResultCol1Is(result, "5"), "got: " & ResultDescribe(result)

    ' Decimal operands and negative literals - "-5" is one token (the
    ' reader's own symbol scan has no terminator between '-' and a
    ' digit), never a unary-minus OPERATOR application (unary (- X) is
    ' refused as wrong-arity, since ComputeArithmetic's own "-" is
    ' binary-only - ValidateArithExpr's own wrong-arity check, exercised
    ' above).
    result = VLA_Prolog.PROLOG("(query (is X (+ 2.5 1.5)))")
    Report "prolog.5.1: decimal operands", ResultCol1Is(result, "4"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (is X (/ 5 2)))")
    Report "prolog.5.1: a non-integer result renders as a real decimal", ResultCol1Is(result, "2.5"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (is X (+ -5 3)))")
    Report "prolog.5.1: a negative literal ('-5', one token) evaluates correctly", ResultCol1Is(result, "-2"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (is X (* -2 -3)))")
    Report "prolog.5.1: two negative literals multiply to a positive result", ResultCol1Is(result, "6"), "got: " & ResultDescribe(result)

    ' `is` used as a computed FILTER across real backtracking - distinct
    ' from every earlier test, which only ever applied `is` to a single,
    ' already-determined binding; here three candidate facts are tried in
    ' turn and only one satisfies the is-check, proving `is` composes
    ' with multi-candidate search rather than only ever appearing after
    ' the search has already narrowed to one row.
    result = VLA_Prolog.PROLOG( _
        "(fact (score alice 90)) (fact (score bob 80)) (fact (score carol 95)) " & _
        "(rule (doubled-is-180 X) (score X S) (is 180 (* S 2))) " & _
        "(query (doubled-is-180 X))")
    Dim doubledOk As Boolean
    If IsArray(result) Then doubledOk = (ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, "alice"))
    Report "prolog.5.1: is composes with backtracking as a computed filter over multiple candidates", _
           doubledOk, "got: " & ResultDescribe(result)

    ' Quoted-string operands - a quoted numeric string evaluates like any
    ' other numeric atom; a quoted NON-numeric string starting with an
    ' uppercase letter ("Hello) is the precise regression for the exact
    ' bug this item's own build found and fixed: checking IsVarAtom on
    ' the POST-LeafText text would have misread "Hello as an unbound
    ' variable named Hello, rather than correctly refusing it as
    ' non-numeric - asserted here by checking BOTH that the right message
    ' fires AND that the wrong one doesn't.
    ' This one EXPECTS SUCCESS (a real spilled array, X=5), not a
    ' refusal string - reuses ResultCol1Is/ResultDescribe rather than
    ' r = CStr(PROLOG(...)), which would itself have crashed here
    ' (CStr on a whole array raises its own type mismatch, the SAME
    ' class of near-miss just caught and fixed three tests above).
    result = VLA_Prolog.PROLOG("(query (is X " & Chr$(34) & "5" & Chr$(34) & "))")
    Report "prolog.5.1: a quoted numeric string evaluates correctly", ResultCol1Is(result, "5"), "got: " & ResultDescribe(result)

    r = CStr(VLA_Prolog.PROLOG("(query (is X " & Chr$(34) & "Hello" & Chr$(34) & "))"))
    Report "prolog.5.1: a quoted non-numeric string starting uppercase is refused as non-numeric, NOT misread as an unbound variable (the exact quote-vs-variable bug this item found and fixed)", _
           InStr(1, r, "isn't one", vbTextCompare) > 0 And InStr(1, r, "unbound variable", vbTextCompare) = 0, "got: " & r

    ' REGRESSION: a runtime-bound compound term inside an is-expression -
    ' the exact second bug this item's own build found and fixed.
    ' ValidateArithExpr only ever checks the expression's own WRITTEN
    ' shape at parse time; Y here is a bare variable (needs no shape
    ' validation at all), and only turns out to be bound to an arbitrary
    ' compound term - never an arithmetic sub-expression - once solving
    ' actually reaches it. Before the fix, EvalArithTerm's own compound
    ' branch would have touched lst.Item(1)/(2)/(3) blind and risked a
    ' raw crash instead of this worded refusal; three distinct malformed
    ' shapes below exercise all three of the runtime re-check's own
    ' guards, not just one.
    r = CStr(VLA_Prolog.PROLOG("(fact (thing (color red))) (rule (compute X) (thing Y) (is X (+ Y 1))) (query (compute X))"))
    Report "prolog.5.1 REGRESSION: a variable bound to a 2-item compound term (wrong arity) inside is is refused, not crashed", _
           InStr(1, r, "isn't one", vbTextCompare) > 0, "got: " & r

    r = CStr(VLA_Prolog.PROLOG("(fact (thing ((a) b c))) (rule (compute X) (thing Y) (is X (+ Y 1))) (query (compute X))"))
    Report "prolog.5.1 REGRESSION: a variable bound to a 3-item term with a NESTED first element inside is is refused, not crashed", _
           InStr(1, r, "isn't one", vbTextCompare) > 0, "got: " & r

    ' PROLOG.17 RE-POINTED THE INPUT `mod` -> `%`, AND THIS ONE REACHED A
    ' LIVE RUN AND KILLED IT. Recorded in full because the failure mode is
    ' the interesting part, not the fix.
    '
    ' This test exercises EvalArithTerm's THIRD runtime guard: a variable
    ' that dereferences to a 3-item term whose operator symbol is not an
    ' operator at all. It used `(mod 5 2)` and was correct for four items.
    ' The moment PROLOG.17 made `mod` real, `(mod 5 2)` stopped being an
    ' unrecognized symbol and became a VALID SUB-EXPRESSION: it computed
    ' 1, `(+ 1 1)` computed 2, the query SUCCEEDED, and PROLOG returned a
    ' SPILLED ARRAY. `CStr(anArray)` then raised a type mismatch, which
    ' KILLED THE WHOLE RUN instead of failing one assertion - so every
    ' test after this line went unreported.
    '
    ' Two lessons, both mechanized rather than remembered. (1) `CStr(...)`
    ' around a PROLOG call is safe ONLY while the query genuinely refuses;
    ' when the input stops refusing it stops being a failing assertion and
    ' becomes a crash. (2) The tell was in this test's own DESCRIPTION -
    ' the word "unrecognized" - not in the message it asserts, which is
    ' the not-numeric refusal ("isn't one") and is still perfectly
    ' emittable. Rule E of tools/check_prolog_arith_operators.ps1 keyed
    ' only on the refusal TEXT and so could not see this; it now also
    ' triggers on that word, and is narrowed to assertions genuinely about
    ' the refusal string so a POSITIVE twin mentioning the history is not
    ' flagged for mentioning it.
    '
    ' `%` is the replacement for the same reason as elsewhere in this
    ' file: a spelling a spreadsheet user plausibly reaches for and this
    ' engine deliberately does not have.
    r = CStr(VLA_Prolog.PROLOG("(fact (thing (% 5 2))) (rule (compute X) (thing Y) (is X (+ Y 1))) (query (compute X))"))
    Report "prolog.5.1/17 REGRESSION: a variable bound to a 3-item term with an unrecognized operator symbol inside is is refused, not crashed", _
           InStr(1, r, "isn't one", vbTextCompare) > 0, "got: " & r

    ' REGRESSION: the non-short-circuit `And` fix in ValidateBodyItem - a
    ' genuinely empty () rule body item must fall through to the
    ' pre-existing "empty predicate form" refusal, not raise a raw
    ' "Subscript out of range" the instant lst.Item(1) is touched.
    r = CStr(VLA_Prolog.PROLOG("(rule (p X) ()) (query (p X))"))
    Report "prolog.5.1 REGRESSION: an empty () rule body item is refused cleanly, not crashed", _
           InStr(1, r, "empty ()", vbTextCompare) > 0, "got: " & r

    ' A bare, non-"!" atom as a rule body item - ValidateBodyItem's own
    ' "not an is-form, not a cut" fall-through to the ordinary
    ' atom-not-a-list refusal.
    r = CStr(VLA_Prolog.PROLOG("(rule (p X) foo) (query (p X))"))
    Report "prolog.5.1: a bare non-! atom as a rule body item falls through to the ordinary atom-not-a-list refusal", _
           InStr(1, r, "predicate form", vbTextCompare) > 0, "got: " & r

    ' An empty () operand inside an is-expression - ValidateArithExpr's
    ' own lst.Count < 1 guard, checked before touching lst.Item(1).
    ' PROLOG.17 RE-POINTED THE EXPECTED MESSAGE, not the input. Arity is
    ' now PER-OPERATOR, and `()` has no operator to look an arity up for,
    ' so "needs exactly two operands" was both unanswerable and never
    ' quite true: an empty form is not a binary operator with the wrong
    ' number of arguments, it is a form with no operator at all. The
    ' guard being tested - lst.Count < 1, checked BEFORE any Item(1)
    ' access - is unchanged, and so is what this line exists to prove:
    ' that the empty operand is REFUSED rather than crashing on a
    ' subscript.
    '
    ' Worth recording how this was caught: by walking the change, NOT by
    ' a check. Rule E of tools/check_prolog_arith_operators.ps1 reads
    ' only assertions on the unknown-operator text; this one pinned the
    ' WRONG-ARITY text for an input whose refusal moved to a different
    ' message, which is a third class again and is not mechanically
    ' covered by rule E or by rule G.
    r = CStr(VLA_Prolog.PROLOG("(query (is X (+ () 3)))"))
    Report "prolog.5.1/17: an empty () arithmetic operand is refused as having no operator, not crashed", _
           InStr(1, r, "isn't an arithmetic operator", vbTextCompare) > 0, "got: " & r
    Report "prolog.17: ...and the refusal names the empty form '()' itself", _
           InStr(1, r, "'()'", vbTextCompare) > 0, "got: " & r

    ' A nested form sitting in the OPERATOR position - ValidateArithExpr's
    ' own IsObject(lst.Item(1)) guard.
    r = CStr(VLA_Prolog.PROLOG("(query (is X ((+ 1 2) 3 4)))"))
    Report "prolog.5.1: a nested form in the operator position is refused as an unknown operator, not crashed", _
           InStr(1, r, "isn't an arithmetic operator", vbTextCompare) > 0, "got: " & r

    ' '!' reserved as a FACT predicate name specifically (the fact-site
    ' branch of IsReservedPredicateName - 'is'/'not'/'findall' above only
    ' ever exercised the fact/rule-head call sites for the OTHER three).
    r = CStr(VLA_Prolog.PROLOG("(fact (! a)) (query (p X))"))
    Report "prolog.5.1: '!' is refused as a predicate name - reserved for PROLOG.5.4", _
           InStr(1, r, "reserved word", vbTextCompare) > 0, "got: " & r
End Sub

' ---------------------------------------------------------------------
'  PROLOG.7: VLA_Prolog.PROLOG - the six comparison operators as goals.
'  Covers: each of the six succeeding AND failing as a ground query;
'  every boundary case that would survive a wrong spelling-translation
'  (=< and >= at equality, =:= and =\= as NUMERIC equality); the
'  threshold filter the README's own staffing example wanted and could
'  not write; a comparison never contributing an output column; nested
'  arithmetic on either side; a comparison inside a rule body (freshened
'  per invocation) and inside `not`; every refusal naming the form the
'  user actually wrote rather than `(is ...)`; and all six refused as
'  predicate names.
'
'  DISCRIMINATION NOTE, and it decided how several of these are written.
'  An unknown predicate in SolveGoalList is a SILENT dead end, not an
'  error - so if this whole item were deleted, `(> S 80000)` would simply
'  yield no solutions. A bare "this comparison fails" test would then
'  still pass against no implementation at all, proving nothing. Every
'  failure case below is therefore either a ground query asserting
'  Boolean FALSE next to its own TRUE twin (deletion collapses both to
'  False, so the pair discriminates), or a filter asserting a strict,
'  named, NON-EMPTY subset (deletion yields an empty one).
' ---------------------------------------------------------------------
Private Sub TestPrologComparison()
    Dim result As Variant
    Dim r As String

    ' ---- each operator, succeeding then failing. The True of each pair
    ' is the discriminator (deleting the dispatch arm turns it False);
    ' the False proves the operator is not vacuously succeeding.
    result = VLA_Prolog.PROLOG("(query (< 1 2))")
    Report "prolog.7: (< 1 2) succeeds", ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (< 2 1))")
    Report "prolog.7: (< 2 1) fails", ResultBoolIs(result, False), "got: " & ResultDescribe(result)

    result = VLA_Prolog.PROLOG("(query (> 2 1))")
    Report "prolog.7: (> 2 1) succeeds", ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (> 1 2))")
    Report "prolog.7: (> 1 2) fails", ResultBoolIs(result, False), "got: " & ResultDescribe(result)

    result = VLA_Prolog.PROLOG("(query (=< 1 2))")
    Report "prolog.7: (=< 1 2) succeeds", ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (=< 2 1))")
    Report "prolog.7: (=< 2 1) fails", ResultBoolIs(result, False), "got: " & ResultDescribe(result)

    result = VLA_Prolog.PROLOG("(query (>= 2 1))")
    Report "prolog.7: (>= 2 1) succeeds", ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (>= 1 2))")
    Report "prolog.7: (>= 1 2) fails", ResultBoolIs(result, False), "got: " & ResultDescribe(result)

    result = VLA_Prolog.PROLOG("(query (=:= 2 2))")
    Report "prolog.7: (=:= 2 2) succeeds", ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (=:= 2 3))")
    Report "prolog.7: (=:= 2 3) fails", ResultBoolIs(result, False), "got: " & ResultDescribe(result)

    result = VLA_Prolog.PROLOG("(query (=\= 2 3))")
    Report "prolog.7: (=\= 2 3) succeeds", ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (=\= 2 2))")
    Report "prolog.7: (=\= 2 2) fails", ResultBoolIs(result, False), "got: " & ResultDescribe(result)

    ' ---- the boundary cases a wrong spelling-translation would survive.
    ' =< must be <= and not <; >= must be >= and not >. Without these,
    ' mapping =< to "<" would pass every test above.
    result = VLA_Prolog.PROLOG("(query (=< 2 2))")
    Report "prolog.7: =< is inclusive at equality (it is <=, not <)", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (>= 2 2))")
    Report "prolog.7: >= is inclusive at equality (it is >=, not >)", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)

    ' =:= is NUMERIC equality, so two different spellings of one number
    ' are equal - which text comparison would get wrong.
    result = VLA_Prolog.PROLOG("(query (=:= 2.0 2))")
    Report "prolog.7: =:= compares numerically, so 2.0 =:= 2", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)

    ' PROLOG.8 re-points this pin rather than retiring it. It was planted
    ' to catch =:= leaking into `=` while `=` was still unimplemented, and
    ' it asserted FALSE because an unknown predicate is a silent dead end.
    ' `=` is real now, so the ASSERTION flips to TRUE - but the thing being
    ' guarded has not changed, and the case below it is the guard that now
    ' does the real work: `=` is UNIFICATION, so it compares terms
    ' structurally and 2.0 is not the same term as 2, where =:= (numeric)
    ' says they are equal. If the two ever became aliases, that second
    ' assertion is what breaks.
    result = VLA_Prolog.PROLOG("(query (= 1 1))")
    Report "prolog.8: `=` unifies two identical ground terms", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (= 2.0 2))")
    Report "prolog.7/8: `=` is NOT =:= - unification is structural, so 2.0 does not unify with 2", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)

    ' `<=` is not a Prolog spelling at all (real Prolog writes =<), so it
    ' must stay an ordinary unknown predicate rather than be accepted.
    ' PROLOG.22 re-pointed this pin, and it is a stronger pin for it: an
    ' unknown predicate used to FAIL, so `1 <= 2` answered FALSE - a
    ' confidently wrong answer to a true comparison. It now refuses as a
    ' predicate nothing defines, which is still "not accepted".
    result = VLA_Prolog.PROLOG("(query (<= 1 2))")
    r = ResultDescribe(result)
    Report "prolog.7/22: `<=` is not a Prolog operator and is not accepted as one - it refuses as undefined rather than answering FALSE", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "'<=' is called by this query", vbTextCompare) > 0, "got: " & r

    ' ---- the motivating case: a THRESHOLD filter, which PROLOG.6's own
    ' README example had to route around. Strict non-empty subset, so
    ' deleting the dispatch arm empties it rather than merely reordering.
    result = VLA_Prolog.PROLOG("(fact (emp alice 90000)) (fact (emp bob 70000)) (query (emp N S) (> S 80000))")
    Report "prolog.7: a comparison filters a real backtracking search to a strict subset", _
           IsArray(result) And ResultRowCount(result) = 2 And ResultCol1Is(result, "alice"), _
           "got: " & ResultDescribe(result)

    ' A comparison NEVER binds, so it contributes no output column: the
    ' query above has exactly two, N and S, not a third for the goal.
    Report "prolog.7: a comparison goal contributes no output column (it never binds)", _
           IsArray(result) And ResultColCount(result) = 2, _
           "got columns: " & ResultColCount(result)

    ' The complement, proving the filter is a real test and not a
    ' constant: the same program with the comparison inverted selects the
    ' OTHER employee, so neither row is being dropped for another reason.
    result = VLA_Prolog.PROLOG("(fact (emp alice 90000)) (fact (emp bob 70000)) (query (emp N S) (< S 80000))")
    Report "prolog.7: inverting the comparison selects the complementary row", _
           IsArray(result) And ResultRowCount(result) = 2 And ResultCol1Is(result, "bob"), _
           "got: " & ResultDescribe(result)

    ' ---- nested arithmetic on either side, proving EvalArithTerm's own
    ' recursion is reached through a comparison and not only through is.
    result = VLA_Prolog.PROLOG("(query (> (+ 2 3) 4))")
    Report "prolog.7: a nested arithmetic expression evaluates on the left side", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (> 10 (* 2 3)))")
    Report "prolog.7: a nested arithmetic expression evaluates on the right side", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)

    ' ---- a comparison inside a RULE body, so the goal is freshened per
    ' invocation before it is dispatched (FreshenTerm keeps position 1 -
    ' the operator - verbatim and rewrites only the argument variables).
    result = VLA_Prolog.PROLOG("(fact (emp alice 90000)) (fact (emp bob 70000)) (rule (rich N) (emp N S) (> S 80000)) (query (rich N))")
    Report "prolog.7: a comparison in a rule body survives freshening and filters correctly", _
           IsArray(result) And ResultRowCount(result) = 2 And ResultCol1Is(result, "alice"), _
           "got: " & ResultDescribe(result)

    ' ---- a comparison nested inside `not`, proving the dispatch is
    ' reached through SolveIsolated's own bounded sub-call too.
    result = VLA_Prolog.PROLOG("(fact (emp alice 90000)) (fact (emp bob 70000)) (query (emp N S) (not (> S 80000)))")
    Report "prolog.7: a comparison composes with `not`", _
           IsArray(result) And ResultRowCount(result) = 2 And ResultCol1Is(result, "bob"), _
           "got: " & ResultDescribe(result)

    ' ---- refusals. Each asserts BOTH that the right reason is given AND
    ' that the named form is the one the user actually wrote - the second
    ' half is the whole point of this item's {form} rewrite, and would
    ' pass vacuously if only the reason were checked.
    r = CStr(VLA_Prolog.PROLOG("(query (> X 1))"))
    Report "prolog.7: an unbound variable in a comparison is refused", _
           InStr(1, r, "unbound variable", vbTextCompare) > 0, "got: " & r
    Report "prolog.7: that refusal names (> ...), NOT (is ...) - the form the user actually wrote", _
           InStr(1, r, "(> ...)", vbTextCompare) > 0 And InStr(1, r, "(is ...)", vbTextCompare) = 0, "got: " & r

    r = CStr(VLA_Prolog.PROLOG("(fact (label hello)) (query (label Y) (=< Y 1))"))
    Report "prolog.7: a non-numeric value reaching a comparison is refused", _
           InStr(1, r, "isn't one", vbTextCompare) > 0, "got: " & r
    Report "prolog.7: that refusal names (=< ...), not (is ...)", _
           InStr(1, r, "(=< ...)", vbTextCompare) > 0 And InStr(1, r, "(is ...)", vbTextCompare) = 0, "got: " & r

    r = CStr(VLA_Prolog.PROLOG("(query (>= 1 (/ 1 0)))"))
    Report "prolog.7: divide-by-zero inside a comparison operand is refused", _
           InStr(1, r, "divide by zero", vbTextCompare) > 0, "got: " & r
    Report "prolog.7: that refusal names (>= ...), not (is ...)", _
           InStr(1, r, "(>= ...)", vbTextCompare) > 0 And InStr(1, r, "(is ...)", vbTextCompare) = 0, "got: " & r

    ' PROLOG.17 re-pointed `mod` to `%` here for the reason recorded at
    ' the sibling assertion in TestPrologArithmetic above - `mod` became a
    ' real operator and this line would otherwise have claimed it was not.
    ' The form-attribution half of the pair is untouched and is what this
    ' assertion actually exists for.
    r = CStr(VLA_Prolog.PROLOG("(query (< 1 (% 5 2)))"))
    Report "prolog.7: an unrecognized operator inside a comparison operand is refused at parse time", _
           InStr(1, r, "isn't an arithmetic operator", vbTextCompare) > 0, "got: " & r
    Report "prolog.7: that refusal names (< ...), not (is ...)", _
           InStr(1, r, "(< ...)", vbTextCompare) > 0 And InStr(1, r, "(is ...)", vbTextCompare) = 0, "got: " & r
    ' ...and the operator that USED to be the example now works inside the
    ' very same comparison form, asserted as a Boolean because a
    ' comparison binds nothing.
    result = VLA_Prolog.PROLOG("(query (< 0 (mod 5 2)))")
    Report "prolog.17: ...while (< 0 (mod 5 2)) now succeeds - mod computes 1 inside a comparison operand too", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)

    r = CStr(VLA_Prolog.PROLOG("(query (=\= 1 (+ 1 2 3)))"))
    Report "prolog.7: a wrong operand count inside a comparison operand is refused", _
           InStr(1, r, "exactly two operands", vbTextCompare) > 0, "got: " & r
    Report "prolog.7: that refusal names (=\= ...), not (is ...)", _
           InStr(1, r, "(=\= ...)", vbTextCompare) > 0 And InStr(1, r, "(is ...)", vbTextCompare) = 0, "got: " & r

    ' Malformed comparison shapes - too few and too many arguments. The
    ' one-argument case also pins that the arity check runs BEFORE any
    ' Item(2)/Item(3) access, which would otherwise raise a raw
    ' "Subscript out of range" instead of a worded refusal.
    r = CStr(VLA_Prolog.PROLOG("(query (> 1))"))
    Report "prolog.7: a one-argument comparison is refused by name, not by a subscript crash", _
           InStr(1, r, "exactly two arguments", vbTextCompare) > 0, "got: " & r
    r = CStr(VLA_Prolog.PROLOG("(query (> 1 2 3))"))
    Report "prolog.7: a three-argument comparison is refused", _
           InStr(1, r, "exactly two arguments", vbTextCompare) > 0, "got: " & r

    ' ---- all six refused as user-defined predicate names, the same
    ' forward-reservation rule is/not/findall/! already follow.
    Dim opName As Variant
    For Each opName In Array("<", ">", "=<", ">=", "=:=", "=\=")
        r = CStr(VLA_Prolog.PROLOG("(fact (" & opName & " a b)) (query (p X))"))
        Report "prolog.7: '" & opName & "' is refused as a predicate name in a (fact ...)", _
               InStr(1, r, "reserved word", vbTextCompare) > 0, "got: " & r
    Next opName

    ' The refusal must also LIST the six, not just the original four -
    ' a message still enumerating is/not/findall/! would be quietly wrong
    ' about which names it had just refused.
    r = CStr(VLA_Prolog.PROLOG("(fact (> a b)) (query (p X))"))
    Report "prolog.7: the reserved-word refusal names the comparison operators among the reserved set", _
           InStr(1, r, "=:=", vbTextCompare) > 0, "got: " & r
End Sub

' ---------------------------------------------------------------------
'  PROLOG.8: VLA_Prolog.PROLOG - the four term-matching goals, `=`, `\=`,
'  `==` and `\==`, as ordinary goals in their own right.
'
'  THE DISCRIMINATION PROBLEM, inherited verbatim from PROLOG.7's own
'  header and worth restating because it shapes every case below. An
'  unknown predicate in SolveGoalList is a SILENT dead end, not an error -
'  so a query naming an unimplemented goal simply yields no solutions, and
'  a test asserting "this fails" would pass against NO implementation at
'  all. Every failure case here is therefore twinned with a ground case
'  asserting TRUE, or written as a strict NON-EMPTY subset of a real
'  backtracking search. Deleting any part of the dispatch must break
'  something, not quietly satisfy it.
'
'  THE PAIR THAT CARRIES THE MOST WEIGHT is `(= X 1)` against `(== X 1)`:
'  the same shape, opposite answers, and they can only both be right if
'  `=` binds a free variable and `==` refuses to. Two more - `(= X 1)
'  (== X 1)` and `(= X 1) (= Y 1) (== X Y)` - pin the DEREFERENCE, the
'  step that separates VLA_Unify.TermsIdentical from the FormsEqual it
'  otherwise resembles: a written-form compare answers False to both.
'
'  Also covered: the full ground truth table for all four; `=` binding
'  forward into a later goal, backward as a filter over real
'  backtracking, into a compound term, through a variable-to-variable
'  chain, and inside a rule body across freshening; `\=` as a filter
'  contributing no output column; the phantom-column skip (a variable
'  appearing ONLY inside a non-binding goal must not surface as a query
'  column - PROLOG.5.2's own finding, reached by a new route);
'  structural-not-numeric equality for both `=` and `==`, which is what
'  keeps them distinct from `=:=`; the occurs-check decision (`\=`
'  inherits `=`'s refusal, `==`/`\==` never reach a bind and so answer
'  ordinarily); an operand that is NOT arithmetic-validated, proving
'  these four take arbitrary terms where a comparison takes numbers;
'  every named shape refusal naming the form the user actually wrote -
'  never `(is ...)` and never a desugared `(not ...)`, which is the whole
'  reason the roadmap's proposed `\=`-as-`(not (= ...))` desugaring was
'  rejected; and all four refused as user-defined predicate names.
' ---------------------------------------------------------------------
Private Sub TestPrologUnification()
    Dim result As Variant
    Dim r As String

    ' ---- the ground truth table. Each FALSE sits beside its own TRUE
    ' twin, so neither an absent dispatch (everything fails) nor an
    ' always-true one can satisfy the pair.
    result = VLA_Prolog.PROLOG("(query (= bob bob))")
    Report "prolog.8: (= bob bob) succeeds", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (= bob ann))")
    Report "prolog.8: (= bob ann) fails", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)

    result = VLA_Prolog.PROLOG("(query (\= bob ann))")
    Report "prolog.8: (\= bob ann) succeeds - the two do not unify", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (\= bob bob))")
    Report "prolog.8: (\= bob bob) fails - they do unify", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)

    result = VLA_Prolog.PROLOG("(query (== bob bob))")
    Report "prolog.8: (== bob bob) succeeds", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (== bob ann))")
    Report "prolog.8: (== bob ann) fails", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)

    result = VLA_Prolog.PROLOG("(query (\== bob ann))")
    Report "prolog.8: (\== bob ann) succeeds", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (\== bob bob))")
    Report "prolog.8: (\== bob bob) fails", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)

    ' ---- THE HEADLINE: `=` BINDS. An unknown predicate yields a Boolean
    ' FALSE, so asserting a spilled ARRAY with a real value in it is what
    ' separates a working dispatch from no dispatch at all.
    result = VLA_Prolog.PROLOG("(query (= X 1))")
    Report "prolog.8: (= X 1) BINDS X and spills it as a real output column", _
           ResultRowCount(result) = 2 And ResultColCount(result) = 1 _
           And ResultCellIs(result, 1, 1, "X") And ResultCol1Is(result, "1"), _
           "got: " & ResultDescribe(result)

    ' ---- ...and `==` does NOT. Identical shape, opposite answer. This
    ' pair is the single most discriminating test in this Sub: it can only
    ' come out right if the two operators differ in exactly the way the
    ' item exists to provide. `==` also contributes no column at all, so
    ' the result collapses to a bare Boolean rather than a spill.
    result = VLA_Prolog.PROLOG("(query (== X 1))")
    Report "prolog.8: (== X 1) FAILS - `==` never binds, so a free X is not identical to 1", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)

    ' ---- the DEREFERENCE, at the top node: once `=` has bound X, `==`
    ' must compare what X now MEANS, not the variable atom as written. A
    ' written-form compare (VLA_Unify's own FormsEqual, which this
    ' otherwise resembles) answers False here.
    result = VLA_Prolog.PROLOG("(query (= X 1) (== X 1))")
    Report "prolog.8: `==` sees through a binding `=` already made (dereference, not written form)", _
           ResultCol1Is(result, "1"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (= X 1) (\== X 1))")
    Report "prolog.8: ...and its twin (\== X 1) correctly finds no solution once X is bound to 1", _
           ResultRowCount(result) = 1, "got: " & ResultDescribe(result)

    ' ---- the DEREFERENCE, at a NESTED node: two variables separately
    ' bound to the same value are identical. Both are still bare atoms as
    ' written, so this is the second case a written-form compare fails.
    result = VLA_Prolog.PROLOG("(query (= X 1) (= Y 1) (== X Y))")
    Report "prolog.8: two variables bound to the same value ARE identical", _
           ResultColCount(result) = 2 _
           And ResultCellIs(result, 2, 1, "1") And ResultCellIs(result, 2, 2, "1"), _
           "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (= X 1) (= Y 2) (== X Y))")
    Report "prolog.8: ...and two bound to DIFFERENT values are not", _
           ResultRowCount(result) = 1, "got: " & ResultDescribe(result)

    ' ---- "the same variable" - real Prolog's own use for `==`, and the
    ' case that needs no binding at all to answer correctly.
    result = VLA_Prolog.PROLOG("(query (== X X))")
    Report "prolog.8: (== X X) succeeds - a free variable is identical to itself", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (== X Y))")
    Report "prolog.8: (== X Y) fails - two DISTINCT free variables are not the same variable", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)

    ' ---- `=` binding FORWARD into a later goal, and BACKWARD as a filter
    ' over a real backtracking search. The second is a strict, non-empty
    ' subset of two facts, so an always-succeeding `=` would return two
    ' rows and an absent one would return none.
    result = VLA_Prolog.PROLOG("(fact (p 1)) (fact (p 2)) (query (= X 2) (p X))")
    Report "prolog.8: a binding made by `=` is visible to a LATER goal", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "2"), _
           "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(fact (p 1)) (fact (p 2)) (query (p X) (= X 2))")
    Report "prolog.8: `=` filters a real backtracking search to a strict subset", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "2"), _
           "got: " & ResultDescribe(result)

    ' ---- unification is STRUCTURAL and TWO-WAY: it reaches inside a
    ' compound term and binds a variable found there, which is the whole
    ' capability `(= X (f Y))` was scoped for.
    result = VLA_Prolog.PROLOG("(query (= X (f 1)))")
    Report "prolog.8: `=` binds a variable to a whole COMPOUND term", _
           ResultCol1Is(result, "(f 1)"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (= (f X) (f 7)))")
    Report "prolog.8: `=` unifies INTO a compound term, binding X from inside it", _
           ResultCol1Is(result, "7"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (= (f 1) (g 1)))")
    Report "prolog.8: two compounds with different heads do not unify", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (== (f 1 2) (f 1 2)))")
    Report "prolog.8: (== (f 1 2) (f 1 2)) succeeds - structural identity recurses", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (== (f 1 2) (f 1)))")
    Report "prolog.8: ...and differing arity is not identical", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)

    ' ---- a variable-to-variable chain: X = Y first, then Y bound by a
    ' fact, must resolve X through the chain (EnvWalkInto's own union-find
    ' walk, exercised through the new arm).
    result = VLA_Prolog.PROLOG("(fact (p 5)) (query (= X Y) (p Y))")
    Report "prolog.8: `=` chains variable to variable, so a later binding resolves both", _
           ResultColCount(result) = 2 _
           And ResultCellIs(result, 2, 1, "5") And ResultCellIs(result, 2, 2, "5"), _
           "got: " & ResultDescribe(result)

    ' ---- STRUCTURAL, NEVER NUMERIC - what keeps all four distinct from
    ' PROLOG.7's `=:=`, asserted here beside the =:= that DOES say equal.
    result = VLA_Prolog.PROLOG("(query (== 2.0 2))")
    Report "prolog.8: (== 2.0 2) fails - `==` is structural, unlike =:=", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (=:= 2.0 2))")
    Report "prolog.8: ...while (=:= 2.0 2) still succeeds - the two are not aliases", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)

    ' ---- A QUOTED STRING IS NOT A BARE SYMBOL, even with identical
    ' letters. Found live, on PROLOG.8's own first table-backed test:
    ' `TableCellToTerm` stores a TEXT cell as Chr$(34) & value (the
    ' string-literal marker LeafText strips for display), while a NUMERIC
    ' cell becomes a bare number. So `(\= D eng)` against a table whose
    ' Dept cell reads "eng" succeeds on EVERY row, and `(== D eng)` fails
    ' on every row - both correct for the operands, both surprising, and
    ' the query has to write "eng" quoted to mean the cell's own value.
    ' That convention is PROLOG.6's, deliberate (it is what stops a
    ' capitalized text cell being read as a variable), and these two pin
    ' it at the unification level where the pure suite can reach it -
    ' TestPrologHostTable already depends on it, but only incidentally,
    ' by quoting "Alice" without a test saying why it must.
    ' PROLOG.10 RE-POINTED THIS PIN rather than retiring it, the same move
    ' PROLOG.8 made to PROLOG.7's own `(= 1 1)` tripwire. It was written
    ' to record that a quoted string and a bare symbol are DIFFERENT
    ' TERMS, and they still are - `(== "eng" "eng")` below is untouched,
    ' and nothing about unification changed. What changed is that saying
    ' so silently was itself the defect: this exact comparison is the
    ' confusable case PROLOG.10 exists to stop being silent, so the
    ' assertion moves from "answers False" to "says why".
    r = CStr(VLA_Prolog.PROLOG("(query (= " & Chr$(34) & "eng" & Chr$(34) & " eng))"))
    Report "prolog.8/10: a quoted string still does not unify with a bare symbol - and now REFUSES rather than answering a silent False", _
           InStr(1, r, "different things", vbTextCompare) > 0, "got: " & r
    Report "prolog.8/10: ...naming the shared text and calling the bare side a NAME", _
           InStr(1, r, "the name eng", vbTextCompare) > 0, "got: " & r
    result = VLA_Prolog.PROLOG("(query (== " & Chr$(34) & "eng" & Chr$(34) & " " & Chr$(34) & "eng" & Chr$(34) & "))")
    Report "prolog.8: ...while two quoted strings with the same text ARE identical", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)

    ' ---- `\=` as a filter, and the column it must NOT contribute.
    result = VLA_Prolog.PROLOG("(fact (p 1)) (fact (p 2)) (query (p X) (\= X 1))")
    Report "prolog.8: `\=` filters a real backtracking search to a strict subset", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "2"), _
           "got: " & ResultDescribe(result)
    Report "prolog.8: `\=` contributes no output column (it binds nothing outward)", _
           ResultColCount(result) = 1, _
           "got columns: " & ResultDescribe(result)

    ' ---- THE PHANTOM-COLUMN SKIP. Y appears ONLY inside a non-binding
    ' goal, and that goal SUCCEEDS - so without CollectVars' own PROLOG.8
    ' skip, Y would be collected as an output column and then resolve to
    ' nothing, rendering the literal atom name "Y" into the sheet as
    ' though it were a value the query had found. Exactly one column, and
    ' it is X's.
    result = VLA_Prolog.PROLOG("(fact (p 7)) (query (p X) (\== Y bob))")
    Report "prolog.8: a variable appearing ONLY in a non-binding goal is not a phantom output column", _
           ResultColCount(result) = 1 And ResultCellIs(result, 1, 1, "X") _
           And ResultCol1Is(result, "7"), _
           "got: " & ResultDescribe(result)
    ' ...and the deliberate exception: `=` DOES bind outward, so a
    ' variable appearing only there IS a real column. Same shape as the
    ' case above, opposite expectation - the fork must land on the
    ' binding line, not on "is it one of the four".
    result = VLA_Prolog.PROLOG("(fact (p 7)) (query (p X) (= Y bob))")
    Report "prolog.8: ...but a variable bound by `=` alone IS a real column - the fork is binding, not family", _
           ResultColCount(result) = 2 _
           And ResultCellIs(result, 2, 1, "7") And ResultCellIs(result, 2, 2, "bob"), _
           "got: " & ResultDescribe(result)

    ' ---- inside a RULE BODY, surviving per-invocation freshening the way
    ' every other rule-body variable already does.
    result = VLA_Prolog.PROLOG("(fact (p 1)) (fact (p 2)) (rule (q X) (p X) (\= X 1)) (query (q Y))")
    Report "prolog.8: `\=` in a rule body survives freshening and filters correctly", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "2"), _
           "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(rule (r X) (= X 7)) (query (r Y))")
    Report "prolog.8: a binding `=` makes in a rule body reaches the caller's own variable", _
           ResultCol1Is(result, "7"), "got: " & ResultDescribe(result)

    ' ---- composes with `not`, whose isolated sub-proof must see the new
    ' arm exactly as it sees is/findall.
    result = VLA_Prolog.PROLOG("(query (not (= bob ann)))")
    Report "prolog.8: `=` composes with `not`", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (not (= bob bob)))")
    Report "prolog.8: ...and its twin correctly fails", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)

    ' ---- AN OPERAND IS AN ARBITRARY TERM, NOT AN ARITHMETIC EXPRESSION.
    ' This is the one behaviour that separates ValidateBodyItem's new arm
    ' from the comparison arm directly above it: `foo` is not one of the
    ' frozen +/-/*// operators, so if these operands were put through
    ' ValidateArithExpr the way a comparison's are, this correct program
    ' would be refused at parse time.
    result = VLA_Prolog.PROLOG("(query (= X (foo 1 2)))")
    Report "prolog.8: an operand is an arbitrary TERM - never arithmetic-validated the way a comparison's is", _
           ResultCol1Is(result, "(foo 1 2)"), "got: " & ResultDescribe(result)

    ' ---- THE OCCURS CHECK, decided rather than inherited by accident.
    ' `\=` raises exactly what `=` raises: the refusal is about the TERM
    ' being unrepresentable, which is equally true whichever operator
    ' encloses it.
    r = CStr(VLA_Prolog.PROLOG("(query (= X (f X)))"))
    Report "prolog.8: (= X (f X)) is refused by the occurs check", _
           InStr(1, r, "contains itself", vbTextCompare) > 0, "got: " & r
    r = CStr(VLA_Prolog.PROLOG("(query (\= X (f X)))"))
    Report "prolog.8: (\= X (f X)) inherits that SAME refusal - the decided behaviour, not a silent True", _
           InStr(1, r, "contains itself", vbTextCompare) > 0, "got: " & r
    ' `==`/`\==` never reach a bind, so they never occurs-check. A refusal
    ' would come back as a String, so asserting a real Boolean is what
    ' distinguishes "answered ordinarily" from "raised".
    result = VLA_Prolog.PROLOG("(query (== X (f X)))")
    Report "prolog.8: (== X (f X)) is an ordinary False - `==` never binds, so it never occurs-checks", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (\== X (f X)))")
    Report "prolog.8: ...and (\== X (f X)) an ordinary True - the asymmetry with `\=` is deliberate", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)

    ' ---- shape refusals. The one-argument case also pins that the arity
    ' check runs BEFORE any Item(2)/Item(3) access, which would otherwise
    ' raise a raw "Subscript out of range" instead of a worded refusal.
    r = CStr(VLA_Prolog.PROLOG("(query (= 1))"))
    Report "prolog.8: a one-argument `=` is refused by name, not by a subscript crash", _
           InStr(1, r, "exactly two arguments", vbTextCompare) > 0, "got: " & r
    r = CStr(VLA_Prolog.PROLOG("(query (== 1 2 3))"))
    Report "prolog.8: a three-argument `==` is refused", _
           InStr(1, r, "exactly two arguments", vbTextCompare) > 0, "got: " & r

    ' THE MISATTRIBUTION GUARD, and the reason the roadmap's proposed
    ' `\=`-as-`(not (= X Y))` desugaring was rejected: a user who wrote
    ' `(\= ...)` must be told about `(\= ...)`. A parse-time desugaring
    ' would have made this refusal name `(not ...)` or `(= ...)` - a form
    ' they never typed - which is precisely what PROLOG.7 spent an item
    ' and tools/check_prolog_form_attribution.ps1 removing.
    r = CStr(VLA_Prolog.PROLOG("(query (\= 1))"))
    Report "prolog.8: that refusal names (\= ...) - not (= ...), not (not ...), not (is ...)", _
           InStr(1, r, "(\= ...)", vbTextCompare) > 0 _
           And InStr(1, r, "(not ...)", vbTextCompare) = 0 _
           And InStr(1, r, "(is ...)", vbTextCompare) = 0, "got: " & r
    r = CStr(VLA_Prolog.PROLOG("(query (\== 1))"))
    Report "prolog.8: ...and (\== ...) names itself too, distinctly from its near-twin", _
           InStr(1, r, "(\== ...)", vbTextCompare) > 0, "got: " & r

    ' The term-matching refusal must talk about TERMS, not the NUMBERS a
    ' comparison's own same-arity refusal talks about - the reason the two
    ' ids were kept separate rather than folded into one.
    r = CStr(VLA_Prolog.PROLOG("(query (= 1))"))
    Report "prolog.8: the shape refusal says `terms`, where a comparison's says `numbers`", _
           InStr(1, r, "terms", vbTextCompare) > 0, "got: " & r

    ' ---- all four refused as user-defined predicate names, the same
    ' forward-reservation rule is/not/findall/! and the six comparisons
    ' already follow.
    Dim opName As Variant
    For Each opName In Array("=", "\=", "==", "\==")
        r = CStr(VLA_Prolog.PROLOG("(fact (" & opName & " a b)) (query (p X))"))
        Report "prolog.8: '" & opName & "' is refused as a predicate name in a (fact ...)", _
               InStr(1, r, "reserved word", vbTextCompare) > 0, "got: " & r
    Next opName

    ' The refusal must also LIST the four, not just the ten it named
    ' before - a message still enumerating only is/not/findall/! and the
    ' comparisons would be quietly wrong about which names it had just
    ' refused. Pinned mechanically as well, in both directions, by
    ' tools/check_prolog_reserved_names.ps1.
    r = CStr(VLA_Prolog.PROLOG("(fact (== a b)) (query (p X))"))
    Report "prolog.8: the reserved-word refusal names the term-matching operators among the reserved set", _
           InStr(1, r, "\==", vbTextCompare) > 0, "got: " & r
End Sub

' ---------------------------------------------------------------------
'  PROLOG.9: VLA_Prolog.PROLOG - the six ISO type-test goals, written
'  var?, nonvar?, atom?, number?, atomic? and compound? - plus the six
'  bare ISO spellings, reserved and dispatched to a refusal that names
'  the question-mark form rather than failing silently.
'
'  THE DISCRIMINATION PROBLEM THIS SUB IS BUILT AROUND. An unknown
'  predicate in SolveGoalList is a SILENT dead end, so a query asserting
'  "this goal fails" passes against NO implementation at all - and half
'  of these six are naturally written as goals expected to fail, which
'  makes them the worst family in the module for that trap. Every FALSE
'  below therefore sits beside its own TRUE twin over the same predicate,
'  so neither an absent dispatch (everything fails) nor an always-true
'  one can satisfy the pair. Delete SolveTypeTest and the TRUE half of
'  every pair goes red.
'
'  The classifier was also checked BEFORE import, PROLOG.8's own
'  precedent: a transliteration of SolveTypeTest/LeafIsNumberTerm run
'  over 25 term shapes x 6 predicates, plus four coherence laws per
'  shape. These tests pin the answers that transliteration predicted.
'
'  PROLOG.11's own six assertions live at the END of this Sub rather
'  than in one of their own: they are about CollectVars mistaking a
'  goal-shaped DATA term for a goal, and the terms that shape most
'  plausibly is a type test, so they read against these. Both halves of
'  that pin are here together - the two that recover a dropped column,
'  and the four that hold the goal-position skips in place - because a
'  fix satisfying either half alone is a different bug.
' ---------------------------------------------------------------------
Private Sub TestPrologTypeTests()
    Dim result As Variant
    Dim r As String

    ' ---- the ground truth table, each FALSE beside its own TRUE twin.
    result = VLA_Prolog.PROLOG("(query (var? X))")
    Report "prolog.9: (var? X) succeeds - X is free", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (var? bob))")
    Report "prolog.9: (var? bob) fails - a ground atom is not a variable", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)

    result = VLA_Prolog.PROLOG("(query (nonvar? bob))")
    Report "prolog.9: (nonvar? bob) succeeds", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (nonvar? X))")
    Report "prolog.9: (nonvar? X) fails - X is free", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)

    result = VLA_Prolog.PROLOG("(query (atom? bob))")
    Report "prolog.9: (atom? bob) succeeds", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (atom? 42))")
    Report "prolog.9: (atom? 42) fails - 42 is a number, not an atom", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)

    result = VLA_Prolog.PROLOG("(query (number? 42))")
    Report "prolog.9: (number? 42) succeeds", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (number? bob))")
    Report "prolog.9: (number? bob) fails", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)

    result = VLA_Prolog.PROLOG("(query (atomic? bob))")
    Report "prolog.9: (atomic? bob) succeeds - an atom is atomic", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (atomic? 42))")
    Report "prolog.9: (atomic? 42) succeeds - a number is atomic too", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (atomic? (f a)))")
    Report "prolog.9: (atomic? (f a)) fails - a compound term is not atomic", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)

    result = VLA_Prolog.PROLOG("(query (compound? (f a)))")
    Report "prolog.9: (compound? (f a)) succeeds", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (compound? bob))")
    Report "prolog.9: (compound? bob) fails", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (compound? 42))")
    Report "prolog.9: (compound? 42) fails", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)

    ' ---- THE PHANTOM COLUMN. `(query (var? X))` SUCCEEDS, so without
    ' CollectVars' own PROLOG.9 skip X would be collected as an output
    ' column, resolve to nothing, and render the literal text "X" into
    ' the cell as though it were a value the query had found. A Boolean
    ' result is the proof the skip is in place: any array here at all is
    ' the bug. PROLOG.5.2 found this shape with `not`, PROLOG.8 with
    ' `\==`; this is the third route to it.
    result = VLA_Prolog.PROLOG("(query (var? X))")
    Report "prolog.9: a type test contributes NO output column - (var? X) is a bare Boolean, never a phantom column headed X", _
           Not IsArray(result), "got: " & ResultDescribe(result)

    ' ---- THE DEREFERENCE, and the single most discriminating pair here.
    ' Classification must ask what X MEANS, not what it is written as.
    ' Without EnvWalkInto, SolveTypeTest sees the atom "X", IsVarAtom
    ' answers True, and `number` answers False for a term already known
    ' to be 1 - so this pair comes out right only if the walk happens.
    result = VLA_Prolog.PROLOG("(query (= X 1) (number? X))")
    Report "prolog.9: `number` sees through a binding `=` already made (dereference, not written form)", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "1"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (= X 1) (var? X))")
    Report "prolog.9: ...and its twin (var? X) correctly finds nothing once X is bound", _
           ResultRowCount(result) = 1, "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (= X (f a)) (compound? X))")
    Report "prolog.9: `compound` sees a compound term X was bound to", _
           ResultRowCount(result) = 2, "got: " & ResultDescribe(result)

    ' ---- THE PROLOG.10 FAULT LINE, pinned as PROLOG.9's own LOCAL
    ' reading and nothing more. A source string literal carries the
    ' reader's leading-Chr$(34) marker exactly as a TEXT cell does, so
    ' these four are the pure-test equivalent of classifying a text cell.
    ' PROLOG.9 reads the marker as identity-bearing: a marked leaf is an
    ' ATOM, never a number, whatever its text says. That is compatible
    ' with PROLOG.10's options A and D and is REVERSED by its option B -
    ' if B is ever adopted, these two `"42"` rows are the tests it flips,
    ' and they are meant to be found by whoever adopts it.
    result = VLA_Prolog.PROLOG("(query (number? ""42""))")
    Report "prolog.9/10: (number? ""42"") FAILS - the quoted-string marker is part of the term, so a text 42 is not the number 42", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (atom? ""42""))")
    Report "prolog.9/10: ...and (atom? ""42"") SUCCEEDS - it is an atom, which is the same judgement seen from the other side", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (atomic? ""42""))")
    Report "prolog.9/10: (atomic? ""42"") succeeds - both readings agree it is atomic; only the atom/number split is at stake", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (atom? ""eng""))")
    Report "prolog.9/10: (atom? ""eng"") succeeds - a marked string is an atom, capitalisation or not", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (var? ""Hello""))")
    Report "prolog.9/10: (var? ""Hello"") FAILS - a capitalised STRING is not an unbound variable (the marker is read on the RAW text)", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)

    ' ---- the coherence law the transliteration asserted 100 times, run
    ' once here against the real engine: every atomic term is an atom or
    ' a number, never both and never neither.
    result = VLA_Prolog.PROLOG("(query (atomic? bob) (atom? bob) (nonvar? bob))")
    Report "prolog.9: bob is atomic AND an atom AND nonvar, all three together", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (atomic? 42) (number? 42) (nonvar? 42))")
    Report "prolog.9: 42 is atomic AND a number AND nonvar, all three together", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)

    ' ---- the shared shape refusal, and the form attribution that keeps
    ' it honest across all six.
    r = CStr(VLA_Prolog.PROLOG("(query (atom? X Y))"))
    Report "prolog.9: (atom? X Y) is refused - a type test takes exactly one term", _
           InStr(1, r, "exactly one argument", vbTextCompare) > 0, "got: " & r
    Report "prolog.9: the type-test shape refusal names the form the user WROTE, not (is ...)", _
           InStr(1, r, "(atom? ...)", vbTextCompare) > 0 And InStr(1, r, "(is ...)", vbTextCompare) = 0, "got: " & r
    r = CStr(VLA_Prolog.PROLOG("(query (compound? X Y))"))
    Report "prolog.9: ...and the same one id names (compound? ...) when that is what was written", _
           InStr(1, r, "(compound? ...)", vbTextCompare) > 0, "got: " & r

    ' ---- all six refused as user-defined predicate names, the same
    ' forward-reservation rule every reserved name already follows.
    Dim tName As Variant
    For Each tName In Array("var?", "nonvar?", "atom?", "number?", "atomic?", "compound?")
        r = CStr(VLA_Prolog.PROLOG("(fact (" & tName & " a)) (query (p X))"))
        Report "prolog.9: '" & tName & "' is refused as a predicate name in a (fact ...)", _
               InStr(1, r, "reserved word", vbTextCompare) > 0, "got: " & r
    Next tName

    r = CStr(VLA_Prolog.PROLOG("(fact (atom? a)) (query (p X))"))
    Report "prolog.9: the reserved-word refusal LISTS the six type tests among the reserved set", _
           InStr(1, r, "nonvar?", vbTextCompare) > 0 And InStr(1, r, "atomic?", vbTextCompare) > 0, "got: " & r

    ' ---- THE BARE ISO SPELLINGS. This engine writes a type test with a
    ' trailing question mark - VLA's own macro layer already spells its
    ' predicates null?/eq?/equal? beside car/cdr/cons/list - but a Prolog
    ' author's first instinct is `(atom X)`. Reserved AND dispatched, so
    ' that instinct meets a refusal naming the right spelling instead of
    ' the SILENT dead end an unknown predicate would be.
    '
    ' This is the one family here where the "assert FALSE proves nothing"
    ' trap would be total: leave these six unreserved and `(query (atom
    ' X))` yields a bare Boolean FALSE, which is exactly what a test
    ' asserting failure would have accepted. So every assertion below
    ' checks the REFUSAL TEXT, which no absent implementation can produce.
    Dim isoName As Variant
    For Each isoName In Array("var", "nonvar", "atom", "number", "atomic", "compound")
        r = CStr(VLA_Prolog.PROLOG("(query (" & isoName & " bob))"))
        Report "prolog.9: the bare ISO '" & isoName & "' is REFUSED with guidance, never silently failed", _
               InStr(1, r, "question mark", vbTextCompare) > 0, "got: " & r
        Report "prolog.9: ...and that refusal names both the form written and '" & isoName & "?' to write instead", _
               InStr(1, r, "(" & isoName & " ...)", vbTextCompare) > 0 _
               And InStr(1, r, "(" & isoName & "? ...)", vbTextCompare) > 0, "got: " & r
    Next isoName

    ' The ISO names are RESERVED as well as dispatched. Without that, a
    ' user could define `(fact (atom a))` and have the guidance arm above
    ' - which sits above the clauseDict lookup - silently shadow their own
    ' facts. That is the "dispatched but not reserved" defect
    ' tools/check_prolog_reserved_names.ps1's rule D exists to catch.
    r = CStr(VLA_Prolog.PROLOG("(fact (atom a)) (query (p X))"))
    Report "prolog.9: the bare ISO 'atom' is also RESERVED, so it can never be defined and then silently shadowed", _
           InStr(1, r, "reserved word", vbTextCompare) > 0, "got: " & r
    Report "prolog.9: the reserved-word refusal lists the six ISO spellings too, distinct from the six question-mark forms", _
           InStr(1, r, "ISO spellings", vbTextCompare) > 0, "got: " & r

    ' ---- THE WORKED EXAMPLE docs/RELEASES.md PRINTS. Every behaviour it
    ' relies on was pinned individually above, but the COMPOSITION was
    ' not, and a worked example in release notes is a promise to a reader
    ' who will paste it verbatim. Pinned here so the notes cannot drift
    ' from the engine.
    result = VLA_Prolog.PROLOG("(rule (halved N H) (number? N) (is H (/ N 2))) (query (between 1 5 X) (halved X Y))")
    Report "prolog.9: the RELEASES.md worked example spills five rows and two columns", _
           ResultRowCount(result) = 6 And ResultColCount(result) = 2, "got: " & ResultDescribe(result)
    Report "prolog.9: ...headed X and Y, with 1 halving to 0.5 and 5 to 2.5", _
           ResultCellIs(result, 1, 1, "X") And ResultCellIs(result, 1, 2, "Y") _
           And ResultCellIs(result, 2, 1, "1") And ResultCellIs(result, 2, 2, "0.5") _
           And ResultCellIs(result, 6, 1, "5") And ResultCellIs(result, 6, 2, "2.5"), _
           "got: " & ResultDescribe(result)

    ' ---- and the CLAIM the notes make ABOUT that example: the guard is
    ' what makes the rule safe to call with anything. This pair is the
    ' discriminating one - the guarded rule SKIPS a non-numeric row and
    ' keeps going, while the identical rule without the guard stops the
    ' whole query with an arithmetic refusal. Remove `(number? N)` from
    ' the first and it produces the second's refusal instead of rows.
    result = VLA_Prolog.PROLOG("(fact (thing eng)) (fact (thing 4)) (rule (halved N H) (number? N) (is H (/ N 2))) (query (thing T) (halved T Y))")
    Report "prolog.9: the guard lets a non-numeric row FAIL TO MATCH rather than stop the query - only 4 survives", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, "4") And ResultCellIs(result, 2, 2, "2"), _
           "got: " & ResultDescribe(result)
    r = CStr(VLA_Prolog.PROLOG("(fact (thing eng)) (fact (thing 4)) (rule (halved N H) (is H (/ N 2))) (query (thing T) (halved T Y))"))
    Report "prolog.9: ...and WITHOUT the guard the same query dies on 'eng', which is what the guard is for", _
           InStr(1, r, "expected a number", vbTextCompare) > 0 And InStr(1, r, "eng", vbTextCompare) > 0, "got: " & r

    ' ---- PROLOG.11: a goal-shaped name used as DATA. CollectVars' own
    ' skip-shapes are statements about GOALS, and used to fire at every
    ' nesting depth - so a compound term whose functor happened to be
    ' spelled `not` or `atom?` had its variables dropped from the output
    ' columns even though they genuinely bind against a stored fact.
    '
    ' Both assertions below fail against the pre-PROLOG.11 code by
    ' reporting ONE column where two are due, which is the whole defect:
    ' a query that silently returns fewer columns than it found. Asserting
    ' the column COUNT and the recovered VALUE together is what makes them
    ' discriminating - a dropped column is not an error, just a quieter
    ' wrong answer.
    result = VLA_Prolog.PROLOG("(fact (holds a (atom? bob))) (query (holds A (atom? W)))")
    Report "prolog.11: a nested `(atom? W)` used as DATA keeps its column - W binds to bob", _
           ResultColCount(result) = 2 And ResultRowCount(result) = 2 _
           And ResultCellIs(result, 2, 1, "a") And ResultCellIs(result, 2, 2, "bob"), _
           "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(fact (holds b (not bob))) (query (holds A (not W)))")
    Report "prolog.11: ...and so does a nested `(not W)`, the shape that was always reachable and always wrong", _
           ResultColCount(result) = 2 And ResultRowCount(result) = 2 _
           And ResultCellIs(result, 2, 1, "b") And ResultCellIs(result, 2, 2, "bob"), _
           "got: " & ResultDescribe(result)

    ' ---- and the REGRESSION half: the same skip-shapes must still fire
    ' at the top of a query conjunct, where the term really is a goal.
    ' A fix that simply deleted the skips would pass the two above and
    ' fail all four below, so they are the other half of the same pin.
    result = VLA_Prolog.PROLOG("(query (var? X))")
    Report "prolog.11: at GOAL position `(var? X)` still contributes no column - a bare Boolean", _
           Not IsArray(result), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (\== X Y))")
    Report "prolog.11: at GOAL position `(\== X Y)` still contributes no column", _
           Not IsArray(result), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(fact (foo 1)) (query (not (foo 2)))")
    Report "prolog.11: at GOAL position `(not (foo 2))` still contributes no column", _
           Not IsArray(result), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(fact (foo 1)) (fact (foo 2)) (query (findall X (foo X) Bag))")
    Report "prolog.11: at GOAL position `findall` still contributes ONLY its Bag, never its Template or Goal", _
           ResultColCount(result) = 1 And ResultCellIs(result, 1, 1, "Bag"), "got: " & ResultDescribe(result)
End Sub

' ---------------------------------------------------------------------
'  PROLOG.15: VLA_Prolog.PROLOG - the REST of the ISO type-test family.
'  `callable?`, `is-list?` and `ground?` join TypeTestKindFor; `integer?`
'  and `float?` are reserved and REFUSED, because this engine has only
'  Doubles and there is nothing for them to be true of.
'
'  A separate Sub from TestPrologTypeTests rather than an extension of
'  it, on PROLOG.13's own precedent: one Sub per item keeps a failing
'  name saying which item broke.
'
'  THE DISCRIMINATION TRAP GOVERNS EVERY ASSERTION HERE, and it is
'  sharper for this item than for PROLOG.9's. An unknown predicate in
'  SolveGoalList is a SILENT dead end - zero rows, no explanation - so a
'  test asserting "this goal fails" passes against NO IMPLEMENTATION AT
'  ALL. Four of these five predicates are naturally written as goals
'  expected to fail. So every failure assertion below is either a GROUND
'  query asserting FALSE beside its own TRUE twin (both Booleans, and an
'  absent implementation cannot produce the TRUE), or a filter asserting
'  a strict non-empty subset, or an assertion on REFUSAL TEXT, which no
'  absent implementation can produce either.
'
'  Booleans go through ResultBoolIs, never `VarType(r) = vbBoolean And
'  r = True`: VBA's And does not short-circuit, so a FAILING assertion
'  written that way kills the run instead of reporting it. PROLOG.20
'  measured 37 of those already in this suite; this item adds none.
' ---------------------------------------------------------------------
' =====================================================================
'  PROLOG.17 - ARITHMETIC BREADTH
'
'  THE DISCRIMINATION PROBLEM IS UNUSUALLY SHARP HERE, and it shapes
'  every assertion below. An unknown predicate fails SILENTLY in this
'  engine, so a test asserting "this goal fails" passes against NO
'  implementation at all. Arithmetic is worse than that: a WRONG
'  operator produces a plausible NUMBER rather than a failure. So every
'  assertion here pins the VALUE, and wherever two readings of an
'  operator are both defensible, the twin's value DIFFERS under the
'  other reading.
'
'  The clearest case is `mod` on negative operands, which is why it has
'  the most rows. ISO's `mod` FLOORS (the result takes the sign of the
'  DIVISOR) and ISO's `rem` TRUNCATES (sign of the DIVIDEND). VBA's own
'  `Mod` operator truncates - so an implementation that reached for the
'  obvious VBA operator would ship `rem` under the name `mod` and be
'  wrong ONLY on mixed-sign operands, silently, in three engines. Both
'  ship here so a user never has to guess which one a bare `mod` meant.
'
'  A query with a free variable ALWAYS spills, so the `is` rows assert
'  through ResultRowCount/ResultCol1Is; a comparison binds nothing and
'  collapses to a bare Boolean. Both shapes appear below deliberately.
' =====================================================================
Private Sub TestPrologArithmeticBreadth()
    Dim result As Variant
    Dim r As String

    ' =================================================================
    '  mod - FLOORED, and the four sign cases
    ' =================================================================

    result = VLA_Prolog.PROLOG("(query (is X (mod 7 3)))")
    Report "prolog.17: (mod 7 3) is 1", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "1"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (is X (mod 6 3)))")
    Report "prolog.17: (mod 6 3) is 0 - an exact division leaves nothing", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "0"), "got: " & ResultDescribe(result)

    ' ---- THE TWO ROWS THAT DECIDE WHICH `mod` SHIPPED. Under the
    ' truncating reading these are -1 and 1; under the floored reading
    ' they are 2 and -2. Nothing else in the suite distinguishes them.
    result = VLA_Prolog.PROLOG("(query (is X (mod -7 3)))")
    Report "prolog.17: (mod -7 3) is 2 - FLOORED, so the result takes the sign of the DIVISOR (a truncating mod would give -1)", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "2"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (is X (mod 7 -3)))")
    Report "prolog.17: (mod 7 -3) is -2 - same law, other sign (a truncating mod would give 1)", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "-2"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (is X (mod -7 -3)))")
    Report "prolog.17: (mod -7 -3) is -1 - mod and rem AGREE when the signs match", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "-1"), "got: " & ResultDescribe(result)

    ' =================================================================
    '  rem - TRUNCATING, and it must DISAGREE with mod
    ' =================================================================

    result = VLA_Prolog.PROLOG("(query (is X (rem 7 3)))")
    Report "prolog.17: (rem 7 3) is 1 - same as mod when both operands are positive", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "1"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (is X (rem -7 3)))")
    Report "prolog.17: (rem -7 3) is -1, where (mod -7 3) is 2 - the two are NOT aliases", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "-1"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (is X (rem 7 -3)))")
    Report "prolog.17: (rem 7 -3) is 1, where (mod 7 -3) is -2 - sign of the DIVIDEND, not the divisor", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "1"), "got: " & ResultDescribe(result)

    ' ---- both refuse a zero divisor by name rather than returning 0
    r = CStr(VLA_Prolog.PROLOG("(query (is X (mod 5 0)))"))
    Report "prolog.17: (mod 5 0) is refused as a division by zero, never silently 0", _
           InStr(1, r, "divide by zero", vbTextCompare) > 0, "got: " & r
    r = CStr(VLA_Prolog.PROLOG("(query (is X (rem 5 0)))"))
    Report "prolog.17: (rem 5 0) likewise", _
           InStr(1, r, "divide by zero", vbTextCompare) > 0, "got: " & r

    ' =================================================================
    '  // - integer division, TRUNCATING toward zero
    ' =================================================================

    result = VLA_Prolog.PROLOG("(query (is X (// 7 2)))")
    Report "prolog.17: (// 7 2) is 3 - integer division, and NOT the 3.5 that / gives", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "3"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (is X (/ 7 2)))")
    Report "prolog.17: ...while (/ 7 2) is still 3.5 - adding // did not quietly make / integer division", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "3.5"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (is X (// -7 2)))")
    Report "prolog.17: (// -7 2) is -3 - truncates TOWARD ZERO, so not the -4 that flooring would give", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "-3"), "got: " & ResultDescribe(result)
    r = CStr(VLA_Prolog.PROLOG("(query (is X (// 5 0)))"))
    Report "prolog.17: (// 5 0) is refused as a division by zero", _
           InStr(1, r, "divide by zero", vbTextCompare) > 0, "got: " & r

    ' =================================================================
    '  min / max
    ' =================================================================

    result = VLA_Prolog.PROLOG("(query (is X (min 3 7)))")
    Report "prolog.17: (min 3 7) is 3", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "3"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (is X (max 3 7)))")
    Report "prolog.17: (max 3 7) is 7 - the twin, so min and max cannot both be the same function", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "7"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (is X (min -3 -7)))")
    Report "prolog.17: (min -3 -7) is -7 - smaller means more negative, not smaller in magnitude", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "-7"), "got: " & ResultDescribe(result)

    ' =================================================================
    '  ** - power, and the two edges VBA would have raised on
    ' =================================================================

    result = VLA_Prolog.PROLOG("(query (is X (** 2 10)))")
    Report "prolog.17: (** 2 10) is 1024", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "1024"), "got: " & ResultDescribe(result)
    ' ---- a fractional exponent on a POSITIVE base, which proves the
    ' domain guard below does not over-refuse. Asserted as BOUNDS rather
    ' than as the exact rendering "3": whether VBA's own `^` returns
    ' 9 ^ 0.5 bit-exactly as 3 or one ulp under it is a property of the C
    ' runtime's pow, not of this engine, and nothing here can run VBA to
    ' find out. Pinning the exact string would make this test a bet on
    ' that; pinning the bounds tests what the item actually claims.
    result = VLA_Prolog.PROLOG("(query (< 2.999 (** 9 0.5)) (> 3.001 (** 9 0.5)))")
    Report "prolog.17: (** 9 0.5) is 3 to within a rounding step - a fractional exponent is a root", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (is X (** 2 0)))")
    Report "prolog.17: (** 2 0) is 1", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "1"), "got: " & ResultDescribe(result)
    ' ---- a negative base to a FRACTIONAL power has no real answer.
    ' Refused by name; VBA's own ^ would have raised a raw runtime error
    ' out of a module whose raw-raise ceiling is ZERO.
    r = CStr(VLA_Prolog.PROLOG("(query (is X (** -8 0.5)))"))
    Report "prolog.17: (** -8 0.5) is refused - a negative base to a fractional power is not a real number", _
           InStr(1, r, "isn't a real number", vbTextCompare) > 0, "got: " & r
    result = VLA_Prolog.PROLOG("(query (is X (** -8 2)))")
    Report "prolog.17: ...but (** -8 2) is 64 - a negative base to a WHOLE power is fine, so the guard is not just refusing negatives", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "64"), "got: " & ResultDescribe(result)
    r = CStr(VLA_Prolog.PROLOG("(query (is X (** 2 10000)))"))
    Report "prolog.17: (** 2 10000) is refused as too large, never a raw VBA overflow escaping the shared substrate", _
           InStr(1, r, "larger than any number", vbTextCompare) > 0, "got: " & r

    ' =================================================================
    '  the UNARY family - the arity change is the real work in this item
    ' =================================================================

    result = VLA_Prolog.PROLOG("(query (is X (abs -3.5)))")
    Report "prolog.17: (abs -3.5) is 3.5 - a UNARY operator, which no operator could be before this item", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "3.5"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (is X (sign -9)))")
    Report "prolog.17: (sign -9) is -1", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "-1"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (is X (sign 0)))")
    Report "prolog.17: (sign 0) is 0 - the twin that separates sign from a mere negativity test", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "0"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (is X (sqrt 9)))")
    Report "prolog.17: (sqrt 9) is 3", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "3"), "got: " & ResultDescribe(result)
    r = CStr(VLA_Prolog.PROLOG("(query (is X (sqrt -1)))"))
    Report "prolog.17: (sqrt -1) is refused, not NaN and not a raw VBA error", _
           InStr(1, r, "isn't a real number", vbTextCompare) > 0, "got: " & r

    ' ---- THE ROUNDING FAMILY, whose four members agree on 3 and part
    ' company everywhere else. Each row below would be WRONG under at
    ' least one of the other three readings, which is the only way to
    ' pin four functions that are so easily confused for one another.
    result = VLA_Prolog.PROLOG("(query (is X (truncate 3.7)))")
    Report "prolog.17: (truncate 3.7) is 3", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "3"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (is X (truncate -3.7)))")
    Report "prolog.17: (truncate -3.7) is -3 - toward ZERO (floor would give -4)", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "-3"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (is X (floor -3.7)))")
    Report "prolog.17: (floor -3.7) is -4 - toward MINUS INFINITY, the twin that separates floor from truncate", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "-4"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (is X (ceiling 3.2)))")
    Report "prolog.17: (ceiling 3.2) is 4", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "4"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (is X (ceiling -3.2)))")
    Report "prolog.17: (ceiling -3.2) is -3 - ceiling of a negative moves TOWARD zero", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "-3"), "got: " & ResultDescribe(result)

    ' ---- ROUND IS HALF AWAY FROM ZERO, and this is the single most
    ' likely quiet defect in the unary family: VBA's own Round() rounds
    ' half to EVEN, so VBA would answer 2 here and 4 on the next row.
    ' Both rows are needed - banker's rounding gets 3.5 RIGHT by
    ' coincidence, so a suite that tested only 3.5 would not notice.
    result = VLA_Prolog.PROLOG("(query (is X (round 2.5)))")
    Report "prolog.17: (round 2.5) is 3 - HALF AWAY FROM ZERO (VBA's own Round would give 2)", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "3"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (is X (round 3.5)))")
    Report "prolog.17: (round 3.5) is 4 - which banker's rounding also gives, so the row above is the discriminating one", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "4"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (is X (round -2.5)))")
    Report "prolog.17: (round -2.5) is -3 - away from zero in the negative direction too", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "-3"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (is X (round 2.4)))")
    Report "prolog.17: (round 2.4) is 2 - and ordinary rounding still rounds down", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "2"), "got: " & ResultDescribe(result)

    ' =================================================================
    '  ARITY IS NOW PER-OPERATOR - both directions
    ' =================================================================

    r = CStr(VLA_Prolog.PROLOG("(query (is X (abs 1 2)))"))
    Report "prolog.17: a unary operator given two operands is refused, and asks for ONE", _
           InStr(1, r, "exactly one operand", vbTextCompare) > 0, "got: " & r
    r = CStr(VLA_Prolog.PROLOG("(query (is X (mod 5)))"))
    Report "prolog.17: ...and a binary operator given one is refused, and asks for TWO", _
           InStr(1, r, "exactly two operands", vbTextCompare) > 0, "got: " & r
    ' ---- the worked example must match the arity being complained about,
    ' or the advice contradicts the complaint.
    r = CStr(VLA_Prolog.PROLOG("(query (is X (sqrt 1 2)))"))
    Report "prolog.17: the unary arity refusal shows a UNARY example, not (+ X Y)", _
           InStr(1, r, "(abs X)", vbTextCompare) > 0 And InStr(1, r, "(+ X Y)", vbTextCompare) = 0, "got: " & r

    ' =================================================================
    '  the unknown-operator refusal no longer claims only four exist
    ' =================================================================

    r = CStr(VLA_Prolog.PROLOG("(query (is X (% 5 2)))"))
    Report "prolog.17: the unknown-operator refusal advertises mod, not 'only +, -, *, and /'", _
           InStr(1, r, "mod", vbTextCompare) > 0 And InStr(1, r, "only +, -, *, and /", vbTextCompare) = 0, "got: " & r

    ' =================================================================
    '  THE SHARED SUBSTRATE, reached through a COMPARISON rather than
    '  through `is` - the whole reason PROLOG.7 made these texts
    '  {form}-templated. A new operator must work in every form that
    '  evaluates an arithmetic expression, not only in the one it was
    '  developed against.
    ' =================================================================

    result = VLA_Prolog.PROLOG("(query (=:= (mod -7 3) 2))")
    Report "prolog.17: mod works inside a comparison operand, with the same floored answer", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (=:= (abs -4) 4))")
    Report "prolog.17: a UNARY operator works inside a comparison operand too", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    r = CStr(VLA_Prolog.PROLOG("(query (> 1 (sqrt -1)))"))
    Report "prolog.17: a domain refusal raised inside a comparison names (> ...), not (is ...)", _
           InStr(1, r, "(> ...)", vbTextCompare) > 0 And InStr(1, r, "(is ...)", vbTextCompare) = 0, "got: " & r
    r = CStr(VLA_Prolog.PROLOG("(query (between 1 (abs -3 -4) X))"))
    Report "prolog.17: ...and a unary arity refusal inside (between ...) names that form", _
           InStr(1, r, "(between ...)", vbTextCompare) > 0 And InStr(1, r, "(is ...)", vbTextCompare) = 0, "got: " & r

    ' =================================================================
    '  NESTING, and the interaction with PROLOG.17's own whole?
    ' =================================================================

    result = VLA_Prolog.PROLOG("(query (is X (abs (- 3 10))))")
    Report "prolog.17: a unary operator nests over a binary one - (abs (- 3 10)) is 7", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "7"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (is X (max (mod 7 3) (mod -7 3))))")
    Report "prolog.17: ...and a binary over two unaries - (max 1 2) is 2, which only holds if mod FLOORS", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "2"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (is X (// 7 2)) (whole? X))")
    Report "prolog.17: (// 7 2) yields a WHOLE number, so whole? succeeds on it", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "3"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (is X (/ 7 2)) (whole? X))")
    Report "prolog.17: ...and its twin (/ 7 2) does not, so whole? yields no solution row", _
           ResultRowCount(result) = 1, "got: " & ResultDescribe(result)
End Sub

Private Sub TestPrologTypeTestsRest()
    Dim result As Variant
    Dim r As String

    ' =================================================================
    '  ground?
    ' =================================================================

    ' ---- the ground truth pair, both Booleans, neither producible by an
    ' absent implementation (the TRUE half is the discriminating one).
    result = VLA_Prolog.PROLOG("(query (ground? bob))")
    Report "prolog.15: (ground? bob) succeeds - a ground atom has no variable in it", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (ground? X))")
    Report "prolog.15: ...and its twin (ground? X) fails - X is free", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (ground? 42))")
    Report "prolog.15: (ground? 42) succeeds - a number is ground", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (ground? nil))")
    Report "prolog.15: (ground? nil) succeeds - the empty list is an atom, and an atom is ground", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)

    ' ---- THE DEREFERENCE, and the single most discriminating pair in
    ' this item. The roadmap entry said TermHasVariable "already computes
    ' exactly this". It does not: it takes NO ENVIRONMENT, so it computes
    ' ground-AS-WRITTEN rather than ground-AS-MEANT, and would answer
    ' FALSE here about a term already known to be 1. This is PROLOG.9's
    ' own `number? sees through a binding = already made` pin, one item
    ' later, about the function that item cited.
    result = VLA_Prolog.PROLOG("(query (= X 1) (ground? X))")
    Report "prolog.15: `ground?` sees through a binding `=` already made (dereference, not written form)", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "1"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (= X Y) (ground? X))")
    Report "prolog.15: ...and its twin, X bound only to another FREE variable, correctly finds nothing", _
           ResultRowCount(result) = 1, "got: " & ResultDescribe(result)

    ' ---- DEEP, not just top-level. A one-step EnvWalkInto answers the
    ' pair above correctly and this one wrongly: it would see the object
    ' `(f Y)` and stop. Both halves are needed, because either alone is
    ' satisfiable by an implementation that is wrong about the other.
    result = VLA_Prolog.PROLOG("(query (ground? (f a)))")
    Report "prolog.15: (ground? (f a)) succeeds - a compound of ground parts is ground", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (ground? (f Y)))")
    Report "prolog.15: ...and (ground? (f Y)) fails - a variable ANYWHERE inside is enough", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (= Y b) (ground? (f Y)))")
    Report "prolog.15: ...and once Y is bound the SAME term is ground - the walk dereferences at every node, not just the top", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "b"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (ground? (f (g (h Z)))))")
    Report "prolog.15: ...and a variable three levels down is still found", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)

    ' ---- POSITION 1, which is the SECOND independent reason
    ' TermHasVariable is the wrong substrate and one no entry named.
    ' That function walks `2 To Count`, skipping position 1 on this
    ' module's "position 1 is a functor" convention - true of terms as
    ' PARSED, false of an arbitrary DATA argument. `(Z a)` is writable
    ' and has a variable in position 1, and ResolveTermDeep (which walks
    ' `1 To Count`) SUBSTITUTES it at render time. A ground? carrying the
    ' carve-out would call this term ground and then print it with the Z
    ' showing.
    '
    ' The pair is the whole pin: lowercase `z` is an ATOM and must stay
    ' ground, so an implementation that simply refused every 2-element
    ' compound would fail the twin.
    result = VLA_Prolog.PROLOG("(query (ground? (Z a)))")
    Report "prolog.15: (ground? (Z a)) FAILS - a variable in FUNCTOR position is still a variable", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (ground? (z a)))")
    Report "prolog.15: ...and its twin (ground? (z a)) succeeds - lowercase z is an atom, so the term really is ground", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)

    ' ---- a real list, and a real findall bag, are ground when their
    ' elements are. Run through the engine rather than written literally,
    ' so this pins the bag and not just the atom nil.
    result = VLA_Prolog.PROLOG("(fact (p 1)) (fact (p 2)) (query (findall X (p X) B) (ground? B))")
    Report "prolog.15: a findall bag of ground elements is ground", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "(list 1 2)"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (ground? (cons a (cons W nil))))")
    Report "prolog.15: ...and a list with a free variable in it is NOT ground", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)

    ' ---- ground? as the GUARD it exists to be: a strict, non-empty
    ' subset. This is the filter shape, and it is discriminating in a way
    ' a bare FALSE is not - an absent implementation gives ZERO rows,
    ' never two of three.
    '
    ' The non-ground row comes from a RULE HEAD carrying a variable its
    ' body never binds, not from a fact: facts must be ground
    ' (prolog-fact-has-variable), so `(fact (thing (f Q)))` would be
    ' REFUSED and this test would measure the refusal instead of the
    ' filter. Found while writing it.
    Dim prog As String
    prog = "(fact (thing a)) (fact (thing 7)) (fact (anchor yes)) " & _
           "(rule (item X) (thing X)) (rule (item (f Y)) (anchor yes)) "
    result = VLA_Prolog.PROLOG(prog & "(query (item T))")
    Report "prolog.15: the unfiltered twin finds THREE items, one of them not ground", _
           ResultRowCount(result) = 4, "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG(prog & "(query (item T) (ground? T))")
    Report "prolog.15: ...and ground? filters it to a strict non-empty subset - a and 7 survive, (f Y) does not", _
           ResultRowCount(result) = 3 And ResultCellIs(result, 2, 1, "a") And ResultCellIs(result, 3, 1, "7"), _
           "got: " & ResultDescribe(result)

    ' =================================================================
    '  is-list?
    ' =================================================================

    result = VLA_Prolog.PROLOG("(query (is-list? nil))")
    Report "prolog.15: (is-list? nil) succeeds - the empty list is a list", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (is-list? (cons a nil)))")
    Report "prolog.15: (is-list? (cons a nil)) succeeds - a one-cell chain terminating in nil", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (is-list? (list a b c)))")
    Report "prolog.15: ...and so does the same list written with PROLOG.21's own (list ...) sugar", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)

    ' ---- THE TERMINATOR, and the pair that proves it is checked rather
    ' than assumed. `(cons a b)` is an improper list - the same term
    ' PROLOG.21 proved must not CONTRACT to (list a) - and it must not
    ' answer True here either, for the same reason.
    result = VLA_Prolog.PROLOG("(query (is-list? (cons a b)))")
    Report "prolog.15: (is-list? (cons a b)) FAILS - an improper list ends in b, not nil", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (is-list? bob))")
    Report "prolog.15: (is-list? bob) fails - a bare atom that is not nil is no list", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (is-list? 42))")
    Report "prolog.15: (is-list? 42) fails", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (is-list? X))")
    Report "prolog.15: (is-list? X) fails - a free variable is not a list, and asking does not bind it to one", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (is-list? (f a)))")
    Report "prolog.15: (is-list? (f a)) fails - an ordinary compound is not a cons cell", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)

    ' ---- THE PER-STEP DEREFERENCE, and the most important assertion in
    ' this item. PROLOG.21's IsProperConsListTerm takes NO env by design:
    ' it runs at DISPLAY time, after ResolveTermDeep. A GOAL runs before
    ' display, and a list's tail is routinely a variable bound to the
    ' rest of the chain - so `is-list?` goes through PROLOG.13's
    ' ListTermToItems instead.
    '
    ' Built on the display-time walker this pair reads 1 row and 1 row.
    ' Built correctly it reads 2 and 1. The twin - a tail bound to
    ' something that is NOT a list - is what stops the first passing on
    ' an implementation that simply answers True for any cons cell.
    result = VLA_Prolog.PROLOG("(query (= T (cons b nil)) (is-list? (cons a T)))")
    Report "prolog.15: `is-list?` dereferences the TAIL at every step - a chain whose tail is a BOUND variable is still a list", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "(list b)"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (= T zzz) (is-list? (cons a T)))")
    Report "prolog.15: ...and its twin, a tail bound to something that is NOT a list, correctly finds nothing", _
           ResultRowCount(result) = 1, "got: " & ResultDescribe(result)

    ' ---- A PARTIAL LIST ANSWERS, WHERE A LIST GOAL RAISES, and the
    ' asymmetry is decided rather than inherited. PROLOG.13 refuses
    ' `(cons a T)` by name because a goal that must USE a list cannot
    ' proceed without one; a type test only ASKS, so False is the correct
    ' answer and is ISO is_list/1's own. The refusing twin is what makes
    ' this pair discriminating: it proves the engine really does
    ' distinguish the two cases rather than answering False everywhere.
    result = VLA_Prolog.PROLOG("(query (is-list? (cons a T)))")
    Report "prolog.15: a PARTIAL list ANSWERS False - a type test never raises", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)
    r = CStr(VLA_Prolog.PROLOG("(query (length (cons a T) N))"))
    Report "prolog.15: ...while PROLOG.13's (length ...) still REFUSES the identical term, which is the asymmetry", _
           InStr(1, r, "#PROLOG!", vbTextCompare) > 0, "got: " & r

    ' ---- a real findall bag is a real list, empty one included. Run
    ' through the engine, never written literally, so this pins the
    ' representation and not the spelling of nil.
    result = VLA_Prolog.PROLOG("(fact (p 1)) (fact (p 2)) (query (findall X (p X) B) (is-list? B))")
    Report "prolog.15: a findall bag is a list", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "(list 1 2)"), "got: " & ResultDescribe(result)
    ' PROLOG.22: the empty bag comes from a DEFINED predicate with no row to
    ' match - p has no row whose first column is 2. The pin used to ask an
    ' undefined q, which now refuses before any bag is built.
    result = VLA_Prolog.PROLOG("(fact (p 1 a)) (query (findall X (p 2 X) B) (is-list? B))")
    Report "prolog.15: ...and so is an EMPTY findall bag, which is the atom nil", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "nil"), "got: " & ResultDescribe(result)

    ' ---- is-list? as a GUARD over a table: a strict, non-empty subset.
    result = VLA_Prolog.PROLOG("(fact (v (list a b))) (fact (v plain)) (fact (v (cons x y))) (query (v L) (is-list? L))")
    Report "prolog.15: is-list? filters to a strict non-empty subset - the proper list survives, the atom and the improper cell do not", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, "(list a b)"), "got: " & ResultDescribe(result)

    ' =================================================================
    '  callable?
    ' =================================================================

    ' ---- ISO callable/1 is "an atom or a compound term". The
    ' discriminating case is the NUMBER: an implementation that simply
    ' reused `nonvar?` passes every other assertion here and fails this
    ' one, which is why the pair sits first.
    result = VLA_Prolog.PROLOG("(query (callable? bob))")
    Report "prolog.15: (callable? bob) succeeds - an atom is callable", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (callable? 42))")
    Report "prolog.15: (callable? 42) FAILS - a number is nonvar but NOT callable, which is what makes it more than nonvar?", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (nonvar? 42))")
    Report "prolog.15: ...and its twin (nonvar? 42) still SUCCEEDS, so the two are genuinely different questions", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (callable? (f a)))")
    Report "prolog.15: (callable? (f a)) succeeds - a compound term is callable", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (callable? X))")
    Report "prolog.15: (callable? X) fails - a free variable is not callable", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (callable? nil))")
    Report "prolog.15: (callable? nil) succeeds - the empty list is an atom", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (= X (f a)) (callable? X))")
    Report "prolog.15: `callable?` sees through a binding `=` already made", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "(f a)"), "got: " & ResultDescribe(result)

    ' ---- THE PROLOG.10 FAULT LINE, one item on. `callable?` reads the
    ' quoted-string marker through the SAME LeafIsNumberTerm `atom?` and
    ' `number?` already ask, so PROLOG.9's local reading is still written
    ' down exactly once and cannot drift. A marked "42 is an atom, so it
    ' is callable - which is the same judgement seen from a third side.
    ' If PROLOG.10's option B is ever adopted, this row flips with the
    ' two labelled prolog.9/10 and is meant to be found alongside them.
    result = VLA_Prolog.PROLOG("(query (callable? ""42""))")
    Report "prolog.15/10: (callable? ""42"") succeeds - a marked leaf is an atom, so it is callable, the same reading as (atom? ""42"")", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)

    ' =================================================================
    '  the coherence laws, run live rather than argued
    ' =================================================================

    ' The transliterated matrix asserted these on all 28 shapes before
    ' import. These run the two corners that are easiest to get wrong
    ' against the real engine: nil, which is an atom AND a list AND
    ' callable AND ground all at once, and a cons cell, which is a
    ' compound and a list and neither atomic nor callable-by-accident.
    result = VLA_Prolog.PROLOG("(query (atom? nil) (atomic? nil) (is-list? nil) (callable? nil) (ground? nil))")
    Report "prolog.15: nil is an atom AND atomic AND a list AND callable AND ground, all five together", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (compound? (list a b)) (is-list? (list a b)) (callable? (list a b)) (ground? (list a b)))")
    Report "prolog.15: a list is compound AND a list AND callable AND ground, all four together", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (atomic? (list a b)))")
    Report "prolog.15: ...and its twin, a list is NOT atomic, so the four above are not all trivially true", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)

    ' =================================================================
    '  integer? / float? - RESERVED AND REFUSED
    ' =================================================================

    ' This engine has only Doubles. NumberToTerm is Trim$(Str$(v)), so
    ' 3.0 and 3 are the SAME ground atom "3" and no function looking at a
    ' term can recover which was meant; the only implementable `integer?`
    ' is "whole-valued", a question about a VALUE under the name of a
    ' question about a TYPE, and ISO's own float(3.0) is TRUE where that
    ' reading is FALSE. PROLOG.17 owned the decision.
    '
    ' PROLOG.17 HAS NOW TAKEN IT, and these tests are UNCHANGED because
    ' of which way it went: Doubles-only is PERMANENT, so `integer?` and
    ' `float?` stay reserved and refused for good rather than becoming
    ' ordinary type tests. The whole-valued question ships instead under
    ' a name that promises no type - `whole?`, tested below - which is
    ' precisely the branch PROLOG.15 wrote down for this outcome. The one
    ' thing that DID change is the refusal's advice: it now names
    ' `(whole? X)` as well as `(number? X)`, and the assertion that the
    ' advice is TRUE is directly below.
    '
    ' EVERY assertion here checks REFUSAL TEXT, and that is the whole
    ' design: if these four names were simply left unreserved,
    ' `(query (integer? 42))` would answer a bare Boolean FALSE - which
    ' is exactly what an assertion of failure would have accepted. Only
    ' the text distinguishes "refused on purpose" from "never built".
    Dim numName As Variant
    For Each numName In Array("integer?", "float?", "integer", "float")
        r = CStr(VLA_Prolog.PROLOG("(query (" & numName & " 42))"))
        Report "prolog.15: '" & numName & "' is REFUSED with an explanation, never silently failed", _
               InStr(1, r, "one kind of number", vbTextCompare) > 0, "got: " & r
        Report "prolog.15: ...and that refusal names the form '" & numName & "' the user actually wrote", _
               InStr(1, r, "(" & numName & " ...)", vbTextCompare) > 0, "got: " & r
        r = CStr(VLA_Prolog.PROLOG("(fact (" & numName & " a)) (query (p X))"))
        Report "prolog.15: ...and '" & numName & "' is RESERVED too, so it can never be defined and then silently shadowed", _
               InStr(1, r, "reserved word", vbTextCompare) > 0, "got: " & r
    Next numName

    ' ---- the bare ISO spellings land on the SAME refusal in ONE hop,
    ' rather than being bounced through the question-mark refusal first
    ' and then told the question-mark form does not work either.
    r = CStr(VLA_Prolog.PROLOG("(query (integer 42))"))
    Report "prolog.15: the bare `integer` reaches the number-type refusal directly, not the question-mark one", _
           InStr(1, r, "one kind of number", vbTextCompare) > 0 _
           And InStr(1, r, "isn't how PROLOG spells", vbTextCompare) = 0, "got: " & r
    Report "prolog.15: ...and it still teaches the question-mark spelling (integer? ...) in the same message", _
           InStr(1, r, "(integer? ...)", vbTextCompare) > 0, "got: " & r

    ' ---- and the test that DOES exist is named as the alternative, so
    ' the refusal is advice rather than only a wall.
    Report "prolog.15: the number-type refusal points at (number? X), the test that does exist", _
           InStr(1, r, "(number? X)", vbTextCompare) > 0, "got: " & r
    result = VLA_Prolog.PROLOG("(query (number? 42))")
    Report "prolog.15: ...and that advice is true - (number? 42) really does succeed", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)

    ' =================================================================
    '  PROLOG.17: whole? - the whole-valued question, under a name that
    '  promises no type
    ' =================================================================

    ' EVERY PAIR HERE IS A DISCRIMINATING ONE, because `whole?` sits in a
    ' place where two wrong implementations both look plausible: an alias
    ' of `number?` (too wide) and an alias of `atom?`'s negation (wider
    ' still). A test asserting only that `(whole? 3)` succeeds passes
    ' against BOTH. The FALSE half beside each TRUE half is what makes
    ' the assertion mean anything - and since an unknown predicate fails
    ' SILENTLY, the TRUE half is also what proves the predicate exists at
    ' all.

    ' ---- THE NARROWING PAIR: the single most discriminating assertion
    ' in this item. An implementation that simply reused `number?` passes
    ' the first two and fails the third.
    result = VLA_Prolog.PROLOG("(query (whole? 3))")
    Report "prolog.17: (whole? 3) succeeds", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (number? 3.5))")
    Report "prolog.17: (number? 3.5) succeeds - 3.5 IS a number", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (whole? 3.5))")
    Report "prolog.17: ...but (whole? 3.5) FAILS - whole? is a strict NARROWING of number?, not an alias of it", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)

    ' ---- THE DECISION ITSELF, made visible, and stated CAREFULLY. It is
    ' tempting to say "3.0 and 3 are the same ground atom here" - that is
    ' PROLOG.15's sentence and it is true only of a NumberToTerm-COMPUTED
    ' value. A source literal `3.0` is a DIFFERENT term from `3`, which
    ' this suite already pins: `(= 2.0 2)` and `(== 2.0 2)` are both
    ' FALSE. Checked before this comment was written, because the wrong
    ' version of it reads plausibly.
    '
    ' That makes the pair below SHARPER, not weaker. `whole?` answers True
    ' for BOTH spellings although they are not the same term - which is
    ' exactly what it means for it to ask about a VALUE rather than about
    ' a type or about term identity. It is also why the name is not
    ' `integer?`: ISO's float(3.0) is TRUE, so a name promising a TYPE
    ' would be wrong on this very line.
    result = VLA_Prolog.PROLOG("(query (whole? 3.0))")
    Report "prolog.17: (whole? 3.0) succeeds - it asks about the VALUE, and 3.0 has no fractional part", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (== 3.0 3))")
    Report "prolog.17: ...even though (== 3.0 3) FAILS - the two are different TERMS, and whole? is not a question about terms", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)

    ' ---- negatives, where a floor/truncate confusion would show. Int()
    ' floors and Fix() truncates and they DISAGREE on -3.5, but they
    ' agree wherever the value is already whole, so both halves must hold.
    result = VLA_Prolog.PROLOG("(query (whole? -3))")
    Report "prolog.17: (whole? -3) succeeds - a negative whole number is whole", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (whole? -3.5))")
    Report "prolog.17: ...and (whole? -3.5) FAILS - rounding a negative toward zero must not make it whole", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (whole? 0))")
    Report "prolog.17: (whole? 0) succeeds - zero is whole", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)

    ' ---- the three NON-numbers, each against the TRUE half above.
    result = VLA_Prolog.PROLOG("(query (whole? bob))")
    Report "prolog.17: (whole? bob) fails - an atom is not a number", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (whole? X))")
    Report "prolog.17: (whole? X) fails - an unbound variable is not a number (and collapses to a BOOLEAN, since CollectVars skips a type test's argument)", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (whole? (f a)))")
    Report "prolog.17: (whole? (f a)) fails - a compound term is not a number", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)

    ' ---- THE QUOTED-STRING MARKER, the reading `whole?` must SHARE with
    ' `number?` rather than invent its own of. A text cell reading 3 is
    ' not the number 3, and `whole?` asks LeafIsNumberTerm rather than
    ' IsInvariantNumericString precisely so it cannot drift on this.
    result = VLA_Prolog.PROLOG("(query (whole? " & Chr$(34) & "3" & Chr$(34) & "))")
    Report "prolog.17: (whole? " & Chr$(34) & "3" & Chr$(34) & ") FAILS - a quoted string is an atom, exactly as number? already reads it", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)

    ' ---- DEREFERENCE through a binding `=` already made. Note the shape
    ' change: `=` genuinely binds, so X IS collected as an output column
    ' and the result SPILLS - where every assertion above collapsed to a
    ' bare Boolean because a type test's own argument is skipped. The
    ' failing twin spills a HEADER-ONLY array (one row), not a Boolean.
    result = VLA_Prolog.PROLOG("(query (= X 4) (whole? X))")
    Report "prolog.17: `whole?` sees through a binding `=` already made", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "4"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (= X 4.5) (whole? X))")
    Report "prolog.17: ...and its twin (= X 4.5) (whole? X) yields NO solution row", _
           ResultRowCount(result) = 1, "got: " & ResultDescribe(result)

    ' ---- A TYPE TEST DOES NOT EVALUATE ARITHMETIC, pinned as a pair so
    ' the boundary cannot move silently: `(/ 6 2)` handed to `whole?` is
    ' a COMPOUND TERM, not the number 3.
    result = VLA_Prolog.PROLOG("(query (whole? (/ 6 2)))")
    Report "prolog.17: (whole? (/ 6 2)) FAILS - whole? asks about a term, it never evaluates one", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)

    ' ---- ...and the same expression through `is`, which DOES evaluate.
    ' THE VALUE IS ASSERTED, not merely the success: a wrong operator here
    ' would produce a plausible NUMBER, so the twin's value must differ.
    ' This pair is also where the integer decision is visible from the
    ' outside - 7/2 is 3.5 and NOT 3, because this engine has one kind of
    ' number and division never silently became integer division.
    result = VLA_Prolog.PROLOG("(query (is X (/ 6 2)) (whole? X))")
    Report "prolog.17: (is X (/ 6 2)) gives exactly 3, and (whole? X) then succeeds", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "3"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (is X (/ 7 2)))")
    Report "prolog.17: (is X (/ 7 2)) gives exactly 3.5 - division is NOT integer division here", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "3.5"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (is X (/ 7 2)) (whole? X))")
    Report "prolog.17: ...so (whole? X) yields NO solution row for it", _
           ResultRowCount(result) = 1, "got: " & ResultDescribe(result)

    ' ---- RESERVED as well as dispatched, the discipline every name in
    ' this module follows: a name the solver acts on but the parser does
    ' not reserve is one a user can define and then have silently shadowed.
    r = CStr(VLA_Prolog.PROLOG("(fact (whole? a)) (query (p X))"))
    Report "prolog.17: `whole?` is RESERVED, so it can never be defined and then silently shadowed", _
           InStr(1, r, "reserved word", vbTextCompare) > 0, "got: " & r
    Report "prolog.17: ...and the refusal enumerates it by name among the type tests", _
           InStr(1, r, "whole?", vbTextCompare) > 0, "got: " & r

    ' ---- THE REFUSAL'S NEW ADVICE IS TRUE, which is the whole point of
    ' shipping `whole?` in the same item that closes `integer?`. The
    ' message tells a user to ask (whole? X) instead; this asserts that
    ' the thing it recommends actually answers.
    r = CStr(VLA_Prolog.PROLOG("(query (integer? 42))"))
    Report "prolog.17: the number-type refusal now points at (whole? X), not only (number? X)", _
           InStr(1, r, "(whole? X)", vbTextCompare) > 0, "got: " & r
    result = VLA_Prolog.PROLOG("(query (whole? 42))")
    Report "prolog.17: ...and that advice is true - (whole? 42) really does succeed", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)

    ' =================================================================
    '  reservation and the bare ISO spellings, for the three that OPTIMIZE
    ' =================================================================

    Dim tName As Variant
    For Each tName In Array("callable?", "is-list?", "ground?")
        r = CStr(VLA_Prolog.PROLOG("(fact (" & tName & " a)) (query (p X))"))
        Report "prolog.15: '" & tName & "' is refused as a predicate name in a (fact ...)", _
               InStr(1, r, "reserved word", vbTextCompare) > 0, "got: " & r
    Next tName

    ' ---- THE BARE ISO SPELLINGS. `is_list` is the one name in the
    ' family whose ISO spelling is not simply the house spelling minus
    ' the question mark: ISO writes an UNDERSCORE where this engine
    ' writes a hyphen. TypeTestIsoSpellingFor maps from what a Prolog
    ' author actually types, so it maps from `is_list` and points at
    ' `is-list?` - and this pin is what holds that decision.
    ' Two Variant arrays assigned to locals first, never Array(...)(k)
    ' inline: VBA's own indexing of a function result is a shape this
    ' suite does not use anywhere else, and a test module that fails to
    ' COMPILE takes every other assertion down with it.
    Dim isoNames As Variant, wantNames As Variant
    Dim isoName As Variant, wantName As Variant
    Dim k As Long
    isoNames = Array("callable", "is_list", "ground")
    wantNames = Array("callable?", "is-list?", "ground?")
    For k = 0 To 2
        isoName = isoNames(k)
        wantName = wantNames(k)
        r = CStr(VLA_Prolog.PROLOG("(query (" & isoName & " bob))"))
        Report "prolog.15: the bare ISO '" & isoName & "' is REFUSED with guidance, never silently failed", _
               InStr(1, r, "question mark", vbTextCompare) > 0, "got: " & r
        Report "prolog.15: ...and that refusal names both (" & isoName & " ...) and (" & wantName & " ...) to write instead", _
               InStr(1, r, "(" & isoName & " ...)", vbTextCompare) > 0 _
               And InStr(1, r, "(" & wantName & " ...)", vbTextCompare) > 0, "got: " & r
        r = CStr(VLA_Prolog.PROLOG("(fact (" & isoName & " a)) (query (p X))"))
        Report "prolog.15: ...and the bare ISO '" & isoName & "' is RESERVED as well as dispatched", _
               InStr(1, r, "reserved word", vbTextCompare) > 0, "got: " & r
    Next k

    ' ---- THE NEAR-MISS SPELLINGS, and this REVERSES a PROLOG.15
    ' decision on purpose rather than by drift.
    '
    ' PROLOG.15 declined to reserve `is-list` and pinned it as usable,
    ' on the grounds that admitting it would make
    ' TypeTestIsoSpellingFor's contract "every near-miss anyone might
    ' type" - a rule with no edge. That objection was right about that
    ' TABLE and wrong about the CLASS: the near-misses are GENERATED, by
    ' two rules over the reserved set (swap hyphen for underscore; drop a
    ' trailing question mark), so the class is closed and countable.
    ' Enumerated mechanically over all 44 names it yields exactly three
    ' that were not already reserved - and one of them, `is_list?`, is a
    ' spelling nobody had thought of. That enumeration is the edge the
    ' rule was missing, and it is why AliasSpellingFor exists.
    '
    ' This test is RE-POINTED rather than deleted, so the reversal is
    ' visible in the diff instead of being a pin that quietly vanished.
    Dim aliasName As Variant, aliasWant As Variant
    Dim aliasNames As Variant, aliasWants As Variant
    Dim ai As Long
    aliasNames = Array("is-list", "is_list?", "sum_list")
    aliasWants = Array("is-list?", "is-list?", "sum-list")
    For ai = 0 To 2
        aliasName = aliasNames(ai)
        aliasWant = aliasWants(ai)
        r = CStr(VLA_Prolog.PROLOG("(query (" & aliasName & " nil))"))
        Report "prolog.15-alias: the near-miss '" & aliasName & "' is REFUSED with guidance, never silently failed", _
               InStr(1, r, "isn't how PROLOG spells this", vbTextCompare) > 0, "got: " & r
        Report "prolog.15-alias: ...and names both (" & aliasName & " ...) and the (" & aliasWant & " ...) to write instead", _
               InStr(1, r, "(" & aliasName & " ...)", vbTextCompare) > 0 _
               And InStr(1, r, "(" & aliasWant & " ...)", vbTextCompare) > 0, "got: " & r
        r = CStr(VLA_Prolog.PROLOG("(fact (" & aliasName & " a)) (query (p X))"))
        Report "prolog.15-alias: ...and '" & aliasName & "' is RESERVED as well as dispatched", _
               InStr(1, r, "reserved word", vbTextCompare) > 0, "got: " & r
    Next ai

    ' ---- the message teaches the two RULES, not the predicate, because
    ' the three misses do not share a family: two are type tests and one
    ' is a list goal. This is the assertion that would fail if someone
    ' rewrote the text to talk about type tests.
    r = CStr(VLA_Prolog.PROLOG("(query (sum_list nil N))"))
    Report "prolog.15-alias: the refusal explains the hyphen rule and the question-mark rule, not the predicate", _
           InStr(1, r, "hyphen rather than an underscore", vbTextCompare) > 0 _
           And InStr(1, r, "question mark", vbTextCompare) > 0, "got: " & r

    ' ---- and the DISCRIMINATING twin: the real spellings still work, so
    ' the refusals above are about the near-miss and not about the
    ' predicate having broken.
    result = VLA_Prolog.PROLOG("(query (is-list? nil))")
    Report "prolog.15-alias: ...while the real `is-list?` still succeeds", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (sum-list (list 1 2 3) N))")
    Report "prolog.15-alias: ...and the real `sum-list` still totals 6", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "6"), "got: " & ResultDescribe(result)

    ' ---- `nth0`/`nth1` are a DIFFERENT class and deliberately absent.
    ' They are numbered variants rather than a hyphen/underscore slip,
    ' and are not derivable by the two rules above. `nth1` would be a
    ' true alias of this engine's one-based `nth`, but `nth0` is a
    ' different predicate - zero-based - so pointing it at `nth` would be
    ' a confidently wrong answer. Pinned as still-definable so the
    ' decision is visible, exactly as `is-list` was before this item
    ' reversed it.
    result = VLA_Prolog.PROLOG("(fact (nth1 a)) (query (nth1 X))")
    Report "prolog.15-alias: `nth1` is NOT reserved - a numbered variant is a different class, filed rather than folded in", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "a"), "got: " & ResultDescribe(result)

    ' ---- the reserved-word refusal must enumerate what it refuses from.
    r = CStr(VLA_Prolog.PROLOG("(fact (ground? a)) (query (p X))"))
    Report "prolog.15: the reserved-word refusal LISTS the three new type tests", _
           InStr(1, r, "callable?", vbTextCompare) > 0 _
           And InStr(1, r, "is-list?", vbTextCompare) > 0 _
           And InStr(1, r, "ground?", vbTextCompare) > 0, "got: " & r
    Report "prolog.15: ...and their bare ISO spellings, is_list with its underscore", _
           InStr(1, r, "is_list", vbTextCompare) > 0, "got: " & r
    Report "prolog.15: ...and the four number-type names it reserves in order to refuse them", _
           InStr(1, r, "number-type names", vbTextCompare) > 0 _
           And InStr(1, r, "integer?", vbTextCompare) > 0 _
           And InStr(1, r, "float", vbTextCompare) > 0, "got: " & r
    Report "prolog.15-alias: ...and the three alias spellings, is_list? among them", _
           InStr(1, r, "alias spellings", vbTextCompare) > 0 _
           And InStr(1, r, "is_list?", vbTextCompare) > 0 _
           And InStr(1, r, "sum_list", vbTextCompare) > 0, "got: " & r

    ' =================================================================
    '  the phantom column, a fourth and fifth and sixth time
    ' =================================================================

    ' `(query (ground? X))` SUCCEEDS in the sense that it returns an
    ' answer, and `(query (is-list? X))` and `(query (callable? X))` do
    ' too. Without CollectVars' 2-long skip each would spill a column
    ' headed X containing the literal text "X". The three new tests
    ' inherit that skip for free BECAUSE it asks TypeTestKindFor rather
    ' than naming the six PROLOG.9 shipped - which is the whole reason
    ' that skip was written as a table lookup, and these three are the
    ' first evidence it was written right.
    result = VLA_Prolog.PROLOG("(query (ground? X))")
    Report "prolog.15: (ground? X) contributes NO output column - a bare Boolean, never a phantom column headed X", _
           Not IsArray(result), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (is-list? X))")
    Report "prolog.15: ...and neither does (is-list? X)", _
           Not IsArray(result), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (callable? X))")
    Report "prolog.15: ...and neither does (callable? X)", _
           Not IsArray(result), "got: " & ResultDescribe(result)

    ' ---- and PROLOG.11's other half: at DATA position the same shape
    ' must still keep its column. A fix that deleted the skip would pass
    ' the three above only by breaking this one.
    result = VLA_Prolog.PROLOG("(fact (holds a (ground? bob))) (query (holds A (ground? W)))")
    Report "prolog.11/15: a nested `(ground? W)` used as DATA keeps its column - W binds to bob", _
           ResultColCount(result) = 2 And ResultRowCount(result) = 2 _
           And ResultCellIs(result, 2, 1, "a") And ResultCellIs(result, 2, 2, "bob"), _
           "got: " & ResultDescribe(result)

    ' =================================================================
    '  the shared shape refusal still covers the new arms
    ' =================================================================

    r = CStr(VLA_Prolog.PROLOG("(query (ground? X Y))"))
    Report "prolog.15: (ground? X Y) is refused - a type test takes exactly one term", _
           InStr(1, r, "exactly one argument", vbTextCompare) > 0, "got: " & r
    Report "prolog.15: ...and that one shared refusal names (ground? ...), the form the user WROTE", _
           InStr(1, r, "(ground? ...)", vbTextCompare) > 0 And InStr(1, r, "(is ...)", vbTextCompare) = 0, "got: " & r
    r = CStr(VLA_Prolog.PROLOG("(query (is-list? A B C))"))
    Report "prolog.15: ...and names (is-list? ...) when that is what was written", _
           InStr(1, r, "(is-list? ...)", vbTextCompare) > 0, "got: " & r

    ' =================================================================
    '  the guard these exist to be, composed
    ' =================================================================

    ' The point of a type test is to make a rule safe to call with
    ' anything. This is PROLOG.9's own RELEASES.md pair, in the new
    ' family: the guarded rule SKIPS the row it cannot handle and keeps
    ' going, and the unguarded twin stops the whole query. Without the
    ' second half the first proves only that something returned rows.
    result = VLA_Prolog.PROLOG("(fact (box (list 1 2 3))) (fact (box plain)) (rule (total B N) (is-list? B) (sum-list B N)) (query (box B) (total B N))")
    Report "prolog.15: (is-list? B) guards sum-list - the proper list totals 6 and the atom row is skipped, not fatal", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 2, "6"), "got: " & ResultDescribe(result)
    r = CStr(VLA_Prolog.PROLOG("(fact (box (list 1 2 3))) (fact (box plain)) (rule (total B N) (sum-list B N)) (query (box B) (total B N))"))
    Report "prolog.15: ...and WITHOUT the guard the identical query dies on `plain`, which is what the guard is for", _
           InStr(1, r, "#PROLOG!", vbTextCompare) > 0, "got: " & r
End Sub

' ---------------------------------------------------------------------
'  PROLOG.10: VLA_Prolog.PROLOG - the quoted-string marker adjudicated.
'  A text cell reading eng is the term `"eng`, and a bare eng written in
'  a query is a DIFFERENT term. That does not change here. What changes
'  is that COMPARING the two stops being silent, because `(\= D eng)`
'  succeeding on every row of a table is a confidently wrong answer.
'
'  Scoped to the four explicit comparison goals and NOT to clause
'  matching - see RefuseIfQuotedVersusBare's own header for the four
'  reasons. The tests below pin BOTH halves of that scope, because the
'  scope is the decision: the ones that must refuse, and the ones that
'  must stay silent.
' ---------------------------------------------------------------------
Private Sub TestPrologQuotedVersusBare()
    Dim result As Variant
    Dim r As String

    ' ---- THE REPORTED DEFECT, all four operators. Each of these used to
    ' answer silently; each now says why. `\=` and `\==` are the ones
    ' that were actively wrong - they answered TRUE on every row.
    r = CStr(VLA_Prolog.PROLOG("(query (= " & Chr$(34) & "eng" & Chr$(34) & " eng))"))
    Report "prolog.10: `=` refuses a quoted-versus-bare comparison instead of answering False", _
           InStr(1, r, "different things", vbTextCompare) > 0, "got: " & r
    r = CStr(VLA_Prolog.PROLOG("(query (\= " & Chr$(34) & "eng" & Chr$(34) & " eng))"))
    Report "prolog.10: `\=` refuses too - this is the one that answered TRUE on every row of a table", _
           InStr(1, r, "different things", vbTextCompare) > 0, "got: " & r
    r = CStr(VLA_Prolog.PROLOG("(query (== " & Chr$(34) & "eng" & Chr$(34) & " eng))"))
    Report "prolog.10: `==` refuses", _
           InStr(1, r, "different things", vbTextCompare) > 0, "got: " & r
    r = CStr(VLA_Prolog.PROLOG("(query (\== " & Chr$(34) & "eng" & Chr$(34) & " eng))"))
    Report "prolog.10: `\==` refuses - the other one that answered TRUE on every row", _
           InStr(1, r, "different things", vbTextCompare) > 0, "got: " & r

    ' ---- THE OPEN SUB-QUESTION, answered in the wording rather than
    ' papered over: the same near-miss happens between a TEXT 42 and a
    ' NUMERIC 42, where calling 42 a "name" would be wrong.
    r = CStr(VLA_Prolog.PROLOG("(query (= " & Chr$(34) & "42" & Chr$(34) & " 42))"))
    Report "prolog.10: a TEXT 42 against a NUMERIC 42 refuses, and calls the bare side a NUMBER, not a name", _
           InStr(1, r, "the number 42", vbTextCompare) > 0, "got: " & r

    ' ---- NARROW BY CONSTRUCTION. Two genuinely different values must
    ' still compare silently, whichever side carries a marker. Without
    ' these the refusal could be firing on every mismatch and the tests
    ' above would not notice.
    result = VLA_Prolog.PROLOG("(query (= " & Chr$(34) & "eng" & Chr$(34) & " " & Chr$(34) & "sales" & Chr$(34) & "))")
    Report "prolog.10: two QUOTED strings that genuinely differ still answer False, no refusal", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (= eng sales))")
    Report "prolog.10: two BARE symbols that genuinely differ still answer False, no refusal", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(fact (p 1)) (fact (p 2)) (query (p X) (\= X 1))")
    Report "prolog.10: an ordinary non-marker mismatch still filters silently - `\=` keeps working", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "2"), "got: " & ResultDescribe(result)

    ' ---- THE DIAGNOSIS MUST NOT MIS-BLAME. `(f "eng")` versus `(g eng)`
    ' contains a marker-only pair AND a real functor difference; the
    ' reason they do not unify is f versus g, so blaming the marker would
    ' be exactly the confidently-wrong-diagnosis this refusal exists to
    ' remove, reintroduced one level up. Class 2 beats class 1.
    result = VLA_Prolog.PROLOG("(query (= (f " & Chr$(34) & "eng" & Chr$(34) & ") (g eng)))")
    Report "prolog.10: a marker difference alongside a REAL difference is not blamed on the marker - silent False", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)
    r = CStr(VLA_Prolog.PROLOG("(query (= (f " & Chr$(34) & "eng" & Chr$(34) & ") (f eng)))"))
    Report "prolog.10: ...but the SAME functor with a marker-only argument does refuse, nested", _
           InStr(1, r, "different things", vbTextCompare) > 0, "got: " & r

    ' ---- THE SCOPE DECISION, pinned. Clause matching must stay SILENT,
    ' and this is the half that would break if the refusal were moved
    ' into UnifyTwoWay: adding a non-matching fact would turn a working
    ' query into an error, which no definite-clause program may do.
    result = VLA_Prolog.PROLOG("(fact (color " & Chr$(34) & "red" & Chr$(34) & ")) (fact (color red)) (query (color red))")
    Report "prolog.10: MONOTONICITY - a near-miss against another clause never aborts a query that has a real match", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    ' PROLOG.12 RE-POINTED THIS PIN. It was written to record that the
    ' plain-query half was deliberately left open - a query that merely
    ' finds nothing was silent, because zero rows is a correct answer.
    ' PROLOG.12 closes that half post hoc, so this exact query now
    ' explains itself. The assertion moves rather than being deleted,
    ' because the case it guards is the same one; only the verdict on it
    ' changed, and the reasoning is in PROLOG.12's own entry.
    '
    ' The SILENT half of the scope decision has not gone away - it moves
    ' to the case below, where there is no near-miss to report. Keeping a
    ' silent case is what stops the diagnosis quietly firing on every
    ' empty result.
    r = CStr(VLA_Prolog.PROLOG("(fact (emp ann " & Chr$(34) & "eng" & Chr$(34) & ")) (query (emp N eng))"))
    Report "prolog.10/12: a plain query that finds nothing now says WHY, post hoc, instead of an unexplained empty result", _
           InStr(1, r, "found no rows at all", vbTextCompare) > 0, "got: " & r
    result = VLA_Prolog.PROLOG("(fact (emp ann " & Chr$(34) & "eng" & Chr$(34) & ")) (query (emp N sales))")
    Report "prolog.10/12: ...but an empty result with NO near-miss is still silent - the header-only shape, not a refusal", _
           ResultRowCount(result) = 1 And ResultColCount(result) = 1, "got: " & ResultDescribe(result)

    ' ---- and the convention itself is UNCHANGED. Only the silence went.
    result = VLA_Prolog.PROLOG("(query (== " & Chr$(34) & "eng" & Chr$(34) & " " & Chr$(34) & "eng" & Chr$(34) & "))")
    Report "prolog.10: two quoted strings with the same text are still identical - the semantics did not move", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(fact (emp ann " & Chr$(34) & "eng" & Chr$(34) & ")) (query (emp N " & Chr$(34) & "eng" & Chr$(34) & "))")
    Report "prolog.10: and the documented workaround still works - quoting the query matches the text cell", _
           ResultCol1Is(result, "ann"), "got: " & ResultDescribe(result)

    ' ---- PROLOG.12: the post-hoc half. It runs ONLY when the whole
    ' query found nothing, so it can never break a query that works - the
    ' monotonicity property PROLOG.10 refused to give up.
    r = CStr(VLA_Prolog.PROLOG("(fact (emp ann " & Chr$(34) & "42" & Chr$(34) & ")) (query (emp N 42))"))
    Report "prolog.12: the numeric variant is diagnosed too, and says NUMBER rather than name", _
           InStr(1, r, "found no rows at all", vbTextCompare) > 0 _
           And InStr(1, r, "the number 42", vbTextCompare) > 0, "got: " & r

    ' MONOTONICITY, the property that made the post-hoc placement the only
    ' acceptable one: a query that finds a real answer is never touched,
    ' however many near-misses sit beside it. Same knowledge base as the
    ' refusing case above, one matching fact added.
    result = VLA_Prolog.PROLOG("(fact (emp ann " & Chr$(34) & "eng" & Chr$(34) & ")) (fact (emp bob eng)) (query (emp N eng))")
    Report "prolog.12: adding a fact that MATCHES silences the diagnosis entirely - it only ever runs on a total miss", _
           ResultCol1Is(result, "bob"), "got: " & ResultDescribe(result)

    ' ---- THE TWO KNOWN LIMITS, pinned as behaviour so they cannot drift
    ' into being quietly fixed or quietly widened. Both need the near-miss
    ' recorded DURING solving, which is the threading this design exists
    ' to avoid, so both are silent by construction rather than by
    ' accident.
    result = VLA_Prolog.PROLOG("(rule (p X) (q X)) (fact (q " & Chr$(34) & "eng" & Chr$(34) & ")) (query (p eng))")
    Report "prolog.12: LIMIT - a near-miss reachable only through a RULE BODY is not diagnosed, and stays silent", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(fact (dept " & Chr$(34) & "eng" & Chr$(34) & ")) (fact (emp ann eng)) (query (dept D) (emp N D))")
    Report "prolog.12: LIMIT - a near-miss that only appears AFTER an earlier conjunct binds is not diagnosed either", _
           ResultRowCount(result) = 1 And ResultColCount(result) = 2, "got: " & ResultDescribe(result)
End Sub

' ---------------------------------------------------------------------
'  PROLOG.9: VLA_Prolog.PROLOG - `between/3`, the FIRST generator in this
'  engine's dispatch. Every other arm is deterministic; this one binds
'  its third argument to each value in turn and backtracks, so its tests
'  are about enumeration, modes, bounding and CUT rather than about a
'  single answer.
'
'  Discrimination, as above: the headline generating test asserts a
'  spilled ARRAY with real values in it, which no absent dispatch can
'  produce (an unknown predicate yields a bare Boolean False), and every
'  refusal test asserts the refusal's own wording rather than merely
'  that something went wrong.
' ---------------------------------------------------------------------
Private Sub TestPrologBetween()
    Dim result As Variant
    Dim r As String

    ' ---- THE HEADLINE: it ENUMERATES. An unknown predicate would give a
    ' bare Boolean False here, so asserting three real rows in order is
    ' what separates a working generator from no dispatch at all.
    result = VLA_Prolog.PROLOG("(query (between 1 3 X))")
    Report "prolog.9: (between 1 3 X) generates three solutions, one column", _
           ResultRowCount(result) = 4 And ResultColCount(result) = 1, "got: " & ResultDescribe(result)
    Report "prolog.9: ...and they are 1, 2, 3 in ascending order under the header X", _
           ResultCellIs(result, 1, 1, "X") And ResultCellIs(result, 2, 1, "1") _
           And ResultCellIs(result, 3, 1, "2") And ResultCellIs(result, 4, 1, "3"), _
           "got: " & ResultDescribe(result)

    ' ---- TEST MODE: X already bound. Semi-deterministic, binds nothing,
    ' enumerates nothing - real Prolog's own second mode.
    result = VLA_Prolog.PROLOG("(query (between 1 10 5))")
    Report "prolog.9: (between 1 10 5) succeeds - a bound third argument is a TEST", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (between 1 10 50))")
    Report "prolog.9: (between 1 10 50) fails - out of range", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (between 1 10 1))")
    Report "prolog.9: (between 1 10 1) succeeds - the low bound is INCLUSIVE", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (between 1 10 10))")
    Report "prolog.9: (between 1 10 10) succeeds - the high bound is INCLUSIVE", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (between 1 10 11))")
    Report "prolog.9: (between 1 10 11) fails - one past the high bound", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)

    ' ---- an EMPTY range is not an error. `(between 1 N X)` with N bound
    ' to 0 must yield no rows rather than stopping the query. Free
    ' variables but zero solutions spills the header-only shape.
    result = VLA_Prolog.PROLOG("(query (between 5 1 X))")
    Report "prolog.9: (between 5 1 X) yields NO solutions and does not refuse - an empty range is not an error", _
           ResultRowCount(result) = 1, "got: " & ResultDescribe(result)

    ' ---- THE MODE ASYMMETRY, deliberate and easy to get wrong: the same
    ' enormous range is refused when GENERATING and succeeds when
    ' TESTING, because testing enumerates nothing and so has no range to
    ' bound. A range check written in the wrong place breaks exactly one
    ' of this pair.
    result = VLA_Prolog.PROLOG("(query (between 1 1000000 5))")
    Report "prolog.9: (between 1 1000000 5) SUCCEEDS - test mode enumerates nothing, so a huge range is free", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    r = CStr(VLA_Prolog.PROLOG("(query (between 1 1000000 X))"))
    Report "prolog.9: ...but (between 1 1000000 X) is REFUSED rather than hanging or grinding to the step ceiling", _
           InStr(1, r, "would generate", vbTextCompare) > 0, "got: " & r
    Report "prolog.9: the range refusal names the range and the ceiling, and does NOT blame a runaway rule", _
           InStr(1, r, "1000000", vbTextCompare) > 0 And InStr(1, r, "base case", vbTextCompare) = 0, "got: " & r

    ' ---- the ceiling boundary, both sides. The dispatch charges one
    ' unit of work for the goal itself before a single value is generated,
    ' so the largest range that fits is PROLOG_MAX_WORK - 1. One more must
    ' produce the RANGE refusal, never the work-ceiling one - that
    ' off-by-one is the whole reason the up-front check exists.
    ' PROLOG.28: re-pointed from the old edge, 119 and 120, which the step
    ' ceiling set; the work budget moved it to 99,999 and 100,000.
    ' A CUT after the generator, and this is not a shortcut: what this
    ' assertion is for is that the UP-FRONT range check passes at 99,999,
    ' and the cut stops the enumeration after the first value, so the check
    ' is pinned for a few units of work instead of a hundred thousand. The
    ' first draft asked for every value, and hung Excel hard enough to
    ' crash it mid-suite (owner's live run, 2026-09-11) - a test may pin an
    ' edge, but never by making the suite pay the whole budget.
    result = VLA_Prolog.PROLOG("(query (between 1 99999 X) !)")
    Report "prolog.9: a range of 99,999 passes the up-front check - the largest that fits, cut after the first", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, "1"), "got: " & ResultDescribe(result)
    r = ResultDescribe(VLA_Prolog.PROLOG("(query (between 1 100000 X))"))
    Report "prolog.9: (between 1 100000 X) is one too many, and says so as a RANGE problem not a budget one", _
           InStr(1, r, "would generate", vbTextCompare) > 0 And InStr(1, r, "facts and rules", vbTextCompare) = 0, "got: " & r

    ' ---- the bounds are arithmetic EXPRESSIONS, the same ones (is ...)
    ' and the six comparisons accept, evaluated by the same walker.
    result = VLA_Prolog.PROLOG("(query (between 1 (+ 1 2) X))")
    Report "prolog.9: a bound may be an arithmetic expression - (between 1 (+ 1 2) X) generates three", _
           ResultRowCount(result) = 4, "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (is N 3) (between 1 N X))")
    Report "prolog.9: a bound may be a variable an earlier conjunct bound", _
           ResultRowCount(result) = 4 And ResultColCount(result) = 2, "got: " & ResultDescribe(result)

    ' ---- the refusals, each asserted on its own wording.
    r = CStr(VLA_Prolog.PROLOG("(query (between L 3 X))"))
    Report "prolog.9: an UNBOUND bound is refused by name", _
           InStr(1, r, "unbound variable", vbTextCompare) > 0, "got: " & r
    Report "prolog.9: ...and that inherited arithmetic refusal names (between ...), never (is ...)", _
           InStr(1, r, "(between ...)", vbTextCompare) > 0 And InStr(1, r, "(is ...)", vbTextCompare) = 0, "got: " & r
    r = CStr(VLA_Prolog.PROLOG("(query (between 1 3.5 X))"))
    Report "prolog.9: a fractional bound is refused - between counts in whole numbers", _
           InStr(1, r, "whole numbers", vbTextCompare) > 0, "got: " & r
    r = CStr(VLA_Prolog.PROLOG("(query (between 1 10 2.5))"))
    Report "prolog.9: a fractional third argument is REFUSED, not silently failed", _
           InStr(1, r, "whole numbers", vbTextCompare) > 0, "got: " & r
    r = CStr(VLA_Prolog.PROLOG("(query (between 1 10 bob))"))
    Report "prolog.9: a non-numeric third argument is refused by name", _
           InStr(1, r, "isn't one", vbTextCompare) > 0, "got: " & r
    r = CStr(VLA_Prolog.PROLOG("(query (between 1 10 ""5""))"))
    Report "prolog.9/10: a TEXT ""5"" is refused too - between reads the quoted-string marker the same way `number?` does", _
           InStr(1, r, "text cell", vbTextCompare) > 0, "got: " & r
    r = CStr(VLA_Prolog.PROLOG("(query (between 1 3))"))
    Report "prolog.9: (between 1 3) is refused - it needs exactly three arguments", _
           InStr(1, r, "exactly three arguments", vbTextCompare) > 0, "got: " & r

    ' ---- it BINDS OUTWARD, unlike every other goal PROLOG.9 adds. The
    ' generated value must be the same TEXT a numeric fact holds, or this
    ' filter silently matches nothing - a strict non-empty subset, so it
    ' cannot pass by everything failing OR by everything succeeding.
    result = VLA_Prolog.PROLOG("(fact (foo 2)) (fact (foo 5)) (query (between 1 3 X) (foo X))")
    Report "prolog.9: a generated value matches a stored numeric fact - exactly one of 1,2,3 is a foo", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "2"), "got: " & ResultDescribe(result)

    ' ---- CUT, which the roadmap entry does not mention at all.
    ' `between`'s loop sits between a cut's own origin and its firing
    ' site, so it must stop generating the moment the signal comes back.
    ' Without that one line the cut silently does nothing and this pair
    ' reads 6 and 6 instead of 2 and 6.
    result = VLA_Prolog.PROLOG("(rule (p X) (between 1 5 X) !) (query (p X))")
    Report "prolog.9: `!` after a between goal PRUNES the generator - exactly one solution, not five", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "1"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(rule (p X) (between 1 5 X)) (query (p X))")
    Report "prolog.9: ...and the identical rule WITHOUT the cut still generates all five", _
           ResultRowCount(result) = 6, "got: " & ResultDescribe(result)

    ' ---- reserved, like every other dispatched name.
    r = CStr(VLA_Prolog.PROLOG("(fact (between a b c)) (query (p X))"))
    Report "prolog.9: 'between' is refused as a predicate name in a (fact ...)", _
           InStr(1, r, "reserved word", vbTextCompare) > 0, "got: " & r
    Report "prolog.9: the reserved-word refusal names 'between' among the reserved set", _
           InStr(1, r, "between", vbTextCompare) > 0, "got: " & r
End Sub

' ---------------------------------------------------------------------
'  PROLOG.5.2: VLA_Prolog.PROLOG - `not` (negation-as-failure). Covers:
'  a directly-ground goal both succeeding and failing to be negated;
'  `not` used as a computed FILTER over real backtracking with an
'  already-bound variable (the correct, intended NAF usage pattern - a
'  variable bound by an EARLIER conjunct, tested by a later negated
'  one); a rule body's own `not` over a fresh existential variable,
'  freshened per candidate the same way any other rule-body variable
'  already is; the real design fork this item's own scoping named -
'  hand-traced, not assumed - proving a variable appearing ONLY inside a
'  negated goal never becomes a phantom output column; the shared
'  PROLOG_MAX_STEPS budget, run to real exhaustion inside a negated goal
'  rather than trusted to inherit the ceiling; an error raised while
'  proving a negated goal propagating all the way out instead of being
'  swallowed as mere failure; `not` composing with `is`; nested (double)
'  negation both ways; and every named (not ...) shape refusal.
' ---------------------------------------------------------------------
Private Sub TestPrologNegation()
    Dim result As Variant
    Dim r As String

    ' A directly-ground goal, negated: succeeds (no matching fact).
    result = VLA_Prolog.PROLOG("(fact (parent tom bob)) (query (not (parent tom liz)))")
    Report "prolog.5.2: not succeeds when the negated ground goal has no solution", _
           ResultBoolIs(result, True), "got " & TypeName(result) & " " & result

    ' A directly-ground goal, negated: fails (a matching fact exists).
    result = VLA_Prolog.PROLOG("(fact (parent tom bob)) (query (not (parent tom bob)))")
    Report "prolog.5.2: not fails when the negated ground goal has a solution", _
           ResultBoolIs(result, False), "got " & TypeName(result) & " " & result

    ' The correct, intended NAF usage pattern: X is bound by an EARLIER,
    ' ordinary conjunct (real backtracking over three candidates), then
    ' tested by a LATER negated one - alice and carol are employed, bob
    ' is not, so only bob should survive the filter.
    result = VLA_Prolog.PROLOG( _
        "(fact (person alice)) (fact (person bob)) (fact (person carol)) " & _
        "(fact (employed alice)) (fact (employed carol)) " & _
        "(query (person X) (not (employed X)))")
    ' Nested If, never "ResultCol1Is(result, ...) And UBound(result, 1) = 2"
    ' in one expression - VBA's And does not short-circuit, so UBound
    ' would still be touched, and still crash with a raw type mismatch on
    ' a non-array result, even though ResultCol1Is itself is IsArray-safe
    ' (this module's own already-documented class of trap, TestPrologArithmetic's
    ' own doubledOk precedent applied here).
    Dim employedFilterOk As Boolean
    If IsArray(result) Then employedFilterOk = (UBound(result, 1) = 2 And ResultCol1Is(result, "bob"))
    Report "prolog.5.2: not composes with backtracking as a computed filter over an already-bound variable", _
           employedFilterOk, "got: " & ResultDescribe(result)

    ' not inside a RULE body, over a fresh existential variable (P) that
    ' must be freshened per candidate exactly like any other rule-body
    ' variable (FreshenTerm's own existing generic recursion, unchanged)
    ' - the classic "orphan" shape: tom and liz have no parent fact, bob
    ' does, asserted against every solution's own value AND order.
    result = VLA_Prolog.PROLOG( _
        "(fact (person tom)) (fact (person bob)) (fact (person liz)) " & _
        "(fact (parent tom bob)) " & _
        "(rule (orphan X) (person X) (not (parent P X))) " & _
        "(query (orphan X))")
    Report "prolog.5.2: not inside a rule body filters correctly with its own freshened existential variable", _
           ResultRowCount(result) = 3 And ResultCellIs(result, 2, 1, "tom") And ResultCellIs(result, 3, 1, "liz"), _
           "got " & (ResultRowCount(result) - 1) & " rows: " & JoinColumn(result, 1)

    ' The real design fork, hand-traced rather than assumed: a variable
    ' appearing ONLY inside a negated goal (no foo row has 2 in its first
    ' column, so the negation succeeds) must never surface as a phantom
    ' one-column "X" output - it must collapse to the same boolean
    ' scalar shape as any other query with no real free variables.
    ' PROLOG.22 re-pointed this pin: it used to negate a foo that did not
    ' exist at all, which is exactly the confidently wrong TRUE that item
    ' removed - `(not (undefined X))` now refuses (TestPrologUnknownPredicate).
    result = VLA_Prolog.PROLOG("(fact (foo 1 a)) (query (not (foo 2 X)))")
    Report "prolog.5.2: a variable appearing only inside a negated goal is never collected as a free output column", _
           ResultBoolIs(result, True), "got " & TypeName(result) & " " & result

    ' The shared ceiling, run to real exhaustion inside a negated goal -
    ' proving the query has ONE depth, the frames of `not`'s isolated
    ' sub-search counted on top of the caller's, not a fresh allowance
    ' handed to every `not` (PROLOG.28; before it, the one step budget).
    r = CStr(VLA_Prolog.PROLOG("(rule (loop X) (loop X)) (query (not (loop a)))"))
    Report "prolog.5.2: a genuinely non-terminating negated goal is refused by the shared depth ceiling, not left to hang or crash", _
           InStr(1, r, "rules deep", vbTextCompare) > 0 And InStr(1, r, "facts and rules", vbTextCompare) = 0, "got: " & r

    ' An error (not a mere failure) raised while proving a negated goal
    ' must propagate all the way out of not, exactly as if the same goal
    ' had been proved un-negated - the other half of "what must never
    ' leak back out": negation-as-failure catches proof FAILURE, never a
    ' raised error.
    r = CStr(VLA_Prolog.PROLOG("(query (not (is X (/ 5 0))))"))
    Report "prolog.5.2: an error raised while proving a negated goal propagates out of not, rather than being swallowed as mere failure", _
           InStr(1, r, "divide by zero", vbTextCompare) > 0, "got: " & r

    ' not composing with is (the non-error case): (is 5 (+ 2 2)) fails
    ' (5 <> 4), so its negation succeeds.
    result = VLA_Prolog.PROLOG("(query (not (is 5 (+ 2 2))))")
    Report "prolog.5.2: not composes with a failing is goal", _
           ResultBoolIs(result, True), "got " & TypeName(result) & " " & result

    ' Nested (double) negation, both ways - proving the recursive
    ' dispatch survives nesting with no special-casing needed anywhere.
    result = VLA_Prolog.PROLOG("(fact (p a)) (query (not (not (p a))))")
    Report "prolog.5.2: double negation of a goal that holds is TRUE", _
           ResultBoolIs(result, True), "got " & TypeName(result) & " " & result

    result = VLA_Prolog.PROLOG("(fact (p a)) (query (not (not (p b))))")
    Report "prolog.5.2: double negation of a goal that doesn't hold is FALSE", _
           ResultBoolIs(result, False), "got " & TypeName(result) & " " & result

    ' Refusals.
    r = CStr(VLA_Prolog.PROLOG("(query (not))"))
    Report "prolog.5.2: (not) with zero arguments is refused", _
           InStr(1, r, "exactly one argument", vbTextCompare) > 0, "got: " & r

    r = CStr(VLA_Prolog.PROLOG("(fact (foo a)) (fact (bar b)) (query (not (foo a) (bar b)))"))
    Report "prolog.5.2: (not ...) with two arguments is refused", _
           InStr(1, r, "exactly one argument", vbTextCompare) > 0, "got: " & r

    r = CStr(VLA_Prolog.PROLOG("(query (not foo))"))
    Report "prolog.5.2: a bare non-! atom as not's own goal reuses the ordinary atom-not-a-list refusal", _
           InStr(1, r, "predicate form", vbTextCompare) > 0, "got: " & r

    ' PROLOG.5.4: a bare ! is legal syntax everywhere a goal is expected,
    ' including inside not - real Prolog's own opaque-cut rule means it
    ' is confined to SolveIsolated's own sub-search (TestPrologCut, below,
    ' proves the SCOPE of that opacity directly); here it just behaves as
    ' an ordinary trivially-succeeding goal, so (not !) is FALSE.
    result = VLA_Prolog.PROLOG("(fact (p a)) (query (not !))")
    Report "prolog.5.2: a bare ! as not's own goal is legal (PROLOG.5.4) and trivially succeeds, so (not !) is FALSE", _
           ResultBoolIs(result, False), "got " & TypeName(result) & " " & result
End Sub

' ---------------------------------------------------------------------
'  PROLOG.13: VLA_Prolog.PROLOG - LIST TERMS and the six list goals.
'  Covers: the cons representation itself, destructured by ordinary
'  unification and in a RULE HEAD with no list goal involved at all
'  (which is the claim that cons inherits UnifyTwoWay and FreshenTerm for
'  free, run against the real engine rather than trusted); a literal list
'  stored in a fact and read back; each of length/member/nth/append/
'  reverse/sum-list in every mode it has, each failure case beside its
'  own TRUE twin; the composition the whole item exists for - measuring
'  and summing a findall bag, including a 100-element one that no
'  rule-based length could reach inside the step budget; the per-step
'  dereference that makes a chained tail a proper list; the partial-list
'  refusal beside the bound twin that makes it discriminating; cut
'  pruning a list generator; and every named refusal.
'
'  TWO SHAPE RULES THIS SUB OBEYS THROUGHOUT, because getting either
'  wrong turns a failing assertion into a dead run. A query with ANY free
'  variable ALWAYS spills - as a header-only array when it finds nothing
'  - and only a query with NO free variables collapses to a Boolean. So
'  "found nothing" is written as ResultRowCount = 1 where a variable is
'  in play and as result = False where none is. And every check goes
'  through the guarded helpers rather than an IsArray/UBound expression,
'  because VBA's And does not short-circuit.
' ---------------------------------------------------------------------
Private Sub TestPrologLists()
    Dim result As Variant
    Dim r As String

    ' ---- THE REPRESENTATION, before any goal touches it. A cons cell is
    ' an ordinary compound term, so `=` splits one with no new machinery
    ' whatsoever - this is the property that decided the representation,
    ' and it is asserted here against the engine rather than reasoned
    ' about in a header.
    result = VLA_Prolog.PROLOG("(query (= (cons H T) (cons a (cons b nil))))")
    Report "prolog.13: a cons cell destructures by ordinary unification - H is the head, T the tail", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, "a") And ResultCellIs(result, 2, 2, "(list b)"), _
           "got: " & ResultDescribe(result)

    ' ...and in a RULE HEAD, which is the sharper half of the same claim:
    ' the head is FRESHENED before it is matched, so this passes only if
    ' FreshenTerm treats a cons cell's position 1 as a functor and its
    ' 2..N as ordinary arguments. The body is a trivially-true fact
    ' lookup because a (rule ...) needs at least one body goal.
    result = VLA_Prolog.PROLOG("(fact (ok yes)) (rule (firstof (cons H T) H) (ok yes)) (query (firstof (cons x (cons y nil)) F))")
    Report "prolog.13: a cons cell in a RULE HEAD unifies and freshens correctly - no list goal involved", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "x"), "got: " & ResultDescribe(result)

    ' A literal list stored in a FACT, read back and walked. Facts must be
    ' ground, and a list of ground elements is ground - TermHasVariable
    ' walks into every cons cell and finds none.
    result = VLA_Prolog.PROLOG("(fact (colors (cons red (cons green nil)))) (query (colors L) (member X L))")
    Report "prolog.13: a literal list written in a (fact ...) is stored, retrieved and walked", _
           ResultRowCount(result) = 3 And ResultCellIs(result, 2, 2, "red") And ResultCellIs(result, 3, 2, "green"), _
           "got: " & ResultDescribe(result)

    ' ---- LENGTH.
    result = VLA_Prolog.PROLOG("(query (length (cons a (cons b (cons c nil))) N))")
    Report "prolog.13: (length L N) counts a three-element list", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "3"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (length nil N))")
    Report "prolog.13: (length nil N) is 0 - the empty list has a length, it is not an error", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "0"), "got: " & ResultDescribe(result)
    ' The ground pair. Neither of these has a free variable, so both are
    ' Booleans, and the FALSE one is only meaningful beside the TRUE one -
    ' an unknown predicate fails silently, so a lone "this fails" pin
    ' passes against no implementation at all.
    result = VLA_Prolog.PROLOG("(query (length (cons a nil) 1))")
    Report "prolog.13: (length L N) TESTS a length when N is already bound", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (length (cons a nil) 2))")
    Report "prolog.13: ...and its twin, a wrong length, fails", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)

    ' ---- MEMBER, in both modes.
    result = VLA_Prolog.PROLOG("(query (member X (cons a (cons b nil))))")
    Report "prolog.13: (member X L) with X free ENUMERATES the list, one solution per element", _
           ResultRowCount(result) = 3 And ResultCellIs(result, 2, 1, "a") And ResultCellIs(result, 3, 1, "b"), _
           "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (member b (cons a (cons b nil))))")
    Report "prolog.13: (member X L) with X bound TESTS membership", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (member z (cons a (cons b nil))))")
    Report "prolog.13: ...and its twin, an element that is not there, fails", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (member a nil))")
    Report "prolog.13: nothing is a member of the empty list - an ordinary no, not an error", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)

    ' A FILTER, asserting a strict non-empty SUBSET. Three elements in,
    ' two out - so this cannot pass against an engine that enumerates
    ' everything, nor against one that enumerates nothing.
    result = VLA_Prolog.PROLOG("(query (member X (cons a (cons b (cons c nil)))) (\== X b))")
    Report "prolog.13: member composes with a filter - a strict, non-empty subset of the list", _
           ResultRowCount(result) = 3 And ResultCellIs(result, 2, 1, "a") And ResultCellIs(result, 3, 1, "c"), _
           "got: " & ResultDescribe(result)

    ' ---- NTH, one-based, in both modes.
    result = VLA_Prolog.PROLOG("(query (nth 2 (cons a (cons b (cons c nil))) X))")
    Report "prolog.13: (nth N L X) picks the Nth element, counting from 1", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "b"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (nth N (cons a (cons b nil)) X))")
    Report "prolog.13: (nth N L X) with N free enumerates position and element together", _
           ResultRowCount(result) = 3 And ResultCellIs(result, 2, 1, "1") And ResultCellIs(result, 2, 2, "a") _
           And ResultCellIs(result, 3, 1, "2"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (nth 1 (cons a nil) a))")
    Report "prolog.13: (nth N L X) fully ground TESTS a position", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    ' The three ways an index can name no position all FAIL rather than
    ' refuse, and that is one uniform rule rather than three cases: the
    ' positions of a list are 1..Count, and an index outside that set is a
    ' correct negative answer about the list, exactly as ISO nth1/3 has
    ' it. Refusing would make (nth N L X) unusable as a test.
    result = VLA_Prolog.PROLOG("(query (nth 2 (cons a nil) a))")
    Report "prolog.13: ...and an index past the end fails", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (nth 0 (cons a nil) a))")
    Report "prolog.13: ...and 0 fails, because nth counts from 1", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (nth 1.5 (cons a nil) a))")
    Report "prolog.13: ...and a fractional index fails - it names no position, and is never rounded to one", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)
    ' An index far past what a Long can hold. It must FAIL like any other
    ' out-of-range index, not crash: the range is bounded in Double
    ' arithmetic before the conversion, so CLng never sees this number.
    ' Without that ordering this is a raw, unworded overflow.
    result = VLA_Prolog.PROLOG("(query (nth 99999999999 (cons a nil) a))")
    Report "prolog.13: ...and an index too large for a Long fails rather than overflowing into a raw error", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)

    ' ---- APPEND, both modes.
    result = VLA_Prolog.PROLOG("(query (append (cons a nil) (cons b nil) C))")
    Report "prolog.13: (append A B C) joins two lists", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "(list a b)"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (append nil (cons a nil) C))")
    Report "prolog.13: nil is append's identity on the left", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "(list a)"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (append (cons a nil) nil C))")
    Report "prolog.13: ...and on the right", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "(list a)"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (append (cons a nil) (cons b nil) (cons a (cons b nil))))")
    Report "prolog.13: (append A B C) fully ground CHECKS a join", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (append (cons a nil) (cons b nil) (cons b (cons a nil))))")
    Report "prolog.13: ...and its twin, a join in the wrong order, fails", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)

    ' The SPLIT mode - A free, C known. A list of two has three splits,
    ' and both ends are included, which is what distinguishes a real split
    ' generator from one that quietly skips the empty prefix.
    result = VLA_Prolog.PROLOG("(query (append A B (cons a (cons b nil))))")
    Report "prolog.13: (append A B C) with A free enumerates every way to SPLIT C - both ends included", _
           ResultRowCount(result) = 4 And ResultCellIs(result, 2, 1, "nil") _
           And ResultCellIs(result, 2, 2, "(list a b)") _
           And ResultCellIs(result, 3, 1, "(list a)") And ResultCellIs(result, 3, 2, "(list b)") _
           And ResultCellIs(result, 4, 2, "nil"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (append A B nil))")
    Report "prolog.13: ...and the empty list has exactly one split, nil and nil", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, "nil") And ResultCellIs(result, 2, 2, "nil"), _
           "got: " & ResultDescribe(result)

    ' THE MODES THAT DECIDED HOW BRANCH 2 IS KEYED. The split branch tests
    ' whether C is a list, NOT whether A is free - so a KNOWN A with a
    ' free B is "drop this prefix", and a known B with a free A is "drop
    ' this suffix". The first version keyed it on "A is free" and refused
    ' both of these, with a proper A being exactly what disqualified it.
    result = VLA_Prolog.PROLOG("(query (append nil B (cons a nil)))")
    Report "prolog.13: (append A B C) with A KNOWN and B free drops a prefix - here the empty one", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "(list a)"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (append (cons a nil) B (cons a (cons b nil))))")
    Report "prolog.13: ...and a real one - (append (a) B (a b)) leaves B as (b)", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "(list b)"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (append A (cons b nil) (cons a (cons b nil))))")
    Report "prolog.13: ...and the mirror, a known SUFFIX leaving A as the prefix", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "(list a)"), "got: " & ResultDescribe(result)
    ' The discriminating twin for all three: a prefix that is NOT there
    ' finds nothing. B is free, so this spills header-only rather than
    ' collapsing to a Boolean.
    result = VLA_Prolog.PROLOG("(query (append (cons a nil) B (cons b (cons c nil))))")
    Report "prolog.13: ...and a prefix the list does not have finds nothing, rather than inventing a split", _
           ResultRowCount(result) = 1, "got: " & ResultDescribe(result)

    ' ---- REVERSE.
    result = VLA_Prolog.PROLOG("(query (reverse (cons a (cons b (cons c nil))) R))")
    Report "prolog.13: (reverse L R) reverses a list", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "(list c b a)"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (reverse nil R))")
    Report "prolog.13: reversing the empty list gives the empty list", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "nil"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (reverse (cons a (cons b nil)) (cons b (cons a nil))))")
    Report "prolog.13: (reverse L R) fully ground CHECKS a reversal", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (reverse (cons a (cons b nil)) (cons a (cons b nil))))")
    Report "prolog.13: ...and its twin, an unreversed list, fails", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)

    ' ---- SUM-LIST.
    result = VLA_Prolog.PROLOG("(query (sum-list (cons 1 (cons 2 (cons 3 nil))) N))")
    Report "prolog.13: (sum-list L N) adds a list of numbers", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "6"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (sum-list nil N))")
    Report "prolog.13: the empty list sums to 0 - the identity, not a special case", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "0"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (sum-list (cons 1 (cons 2 nil)) 3))")
    Report "prolog.13: (sum-list L N) fully ground CHECKS a total", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (sum-list (cons 1 (cons 2 nil)) 4))")
    Report "prolog.13: ...and its twin, a wrong total, fails", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)
    ' Elements go through the SAME EvalArithTerm `is` and the comparisons
    ' use, so an element may itself be an expression.
    result = VLA_Prolog.PROLOG("(query (sum-list (cons (+ 1 2) (cons 4 nil)) N))")
    Report "prolog.13: a sum-list element may be an arithmetic expression, the same walker `is` uses", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "7"), "got: " & ResultDescribe(result)

    ' ---- THE DEREFERENCE, at every step of the chain rather than only at
    ' the top. The pair is what makes it discriminating: the FIRST refuses
    ' because T is still free and (cons a T) is a partial list; the SECOND
    ' is the identical term with T bound, and must be a proper list of
    ' one. A walk that dereferenced only its argument would refuse both.
    r = CStr(VLA_Prolog.PROLOG("(query (length (cons a T) N))"))
    Report "prolog.13: a PARTIAL list - (cons a T) with T unbound - is refused by name, not answered", _
           InStr(1, r, "needs a list", vbTextCompare) > 0, "got: " & r
    result = VLA_Prolog.PROLOG("(query (= T nil) (length (cons a T) N))")
    Report "prolog.13: ...and the same term with T BOUND is a proper list - the tail is dereferenced at every step", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 2, "1"), "got: " & ResultDescribe(result)

    ' ---- THE COMPOSITION THE WHOLE ITEM EXISTS FOR: findall's bag is no
    ' longer a dead end.
    result = VLA_Prolog.PROLOG("(fact (color red)) (fact (color green)) (fact (color blue)) (query (findall X (color X) Bag) (length Bag N))")
    Report "prolog.13: a findall bag can be MEASURED - findall stops being a terminal operation", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 2, "3"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(fact (sale 10)) (fact (sale 20)) (fact (sale 30)) (query (findall X (sale X) Bag) (sum-list Bag Total))")
    Report "prolog.13: ...and SUMMED, which is the derived-total idiom a spreadsheet actually wants", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 2, "60"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(fact (color red)) (fact (color green)) (query (findall X (color X) Bag) (member green Bag))")
    Report "prolog.13: ...and searched, with member over a bag the query itself derived", _
           ResultRowCount(result) = 2, "got: " & ResultDescribe(result)

    ' THE BUDGET CLAIM, run rather than argued. A hundred-element bag:
    ' findall's own dispatch charges 1, the between goal 1, its hundred
    ' values 100, and length exactly 1 more however long the list is -
    ' 103 against a ceiling of 120. A length written as a recursive Prolog
    ' RULE would charge one step per element on top of the harvest and
    ' could never finish this, which is why the six goals are native.
    result = VLA_Prolog.PROLOG("(query (findall X (between 1 100 X) Bag) (length Bag N))")
    Report "prolog.13: a HUNDRED-element bag can be measured inside the step budget - one step, not one per element", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 2, "100"), "got: " & ResultDescribe(result)

    ' ---- CUT. member's loop sits between a cut's own origin and its
    ' firing site, so it must stop generating the moment the signal comes
    ' back and must never absorb it - SolveBetween's own rule, which this
    ' pair holds for a list generator. Without that line the two read 3
    ' and 3 instead of 1 and 3.
    result = VLA_Prolog.PROLOG("(rule (p X) (member X (cons a (cons b (cons c nil)))) !) (query (p X))")
    Report "prolog.13: `!` after a member goal PRUNES the generator - exactly one solution, not three", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "a"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(rule (p X) (member X (cons a (cons b (cons c nil))))) (query (p X))")
    Report "prolog.13: ...and the identical rule WITHOUT the cut still generates all three", _
           ResultRowCount(result) = 4, "got: " & ResultDescribe(result)

    ' ---- REFUSALS. Every one is visible in a cell as "#PROLOG! ..." so
    ' none of these needs a live workbook to observe.
    r = CStr(VLA_Prolog.PROLOG("(query (length foo N))"))
    Report "prolog.13: a non-list argument is refused by name, and the refusal TEACHES the cons spelling", _
           InStr(1, r, "needs a list", vbTextCompare) > 0 And InStr(1, r, "(cons a (cons b nil))", vbTextCompare) > 0, "got: " & r
    r = CStr(VLA_Prolog.PROLOG("(query (member X (pair a b)))"))
    Report "prolog.13: ...and a compound term that is not a cons cell is refused the same way", _
           InStr(1, r, "needs a list", vbTextCompare) > 0, "got: " & r
    r = CStr(VLA_Prolog.PROLOG("(query (append A B C))"))
    Report "prolog.13: (append A B C) with everything free is refused - there is nothing to split", _
           InStr(1, r, "needs a list", vbTextCompare) > 0, "got: " & r
    r = CStr(VLA_Prolog.PROLOG("(query (sum-list (cons 1 (cons b nil)) N))"))
    Report "prolog.13: a non-numeric element in sum-list reuses the shared arithmetic refusal, naming (sum-list ...)", _
           InStr(1, r, "(sum-list ...)", vbTextCompare) > 0 And InStr(1, r, "isn't one", vbTextCompare) > 0, "got: " & r
    r = CStr(VLA_Prolog.PROLOG("(query (nth foo (cons a nil) X))"))
    Report "prolog.13: a non-numeric INDEX is refused, naming (nth ...) - unlike an out-of-range one, which fails", _
           InStr(1, r, "(nth ...)", vbTextCompare) > 0, "got: " & r

    ' The shared SHAPE refusal, and the two arities it has to tell apart -
    ' one raise site serving six forms of two different shapes, so both
    ' the form and the count come from the caller.
    r = CStr(VLA_Prolog.PROLOG("(query (length (cons a nil)))"))
    Report "prolog.13: (length ...) with one argument is refused, naming its own form and count", _
           InStr(1, r, "(length ...)", vbTextCompare) > 0 And InStr(1, r, "exactly 2 arguments", vbTextCompare) > 0, "got: " & r
    r = CStr(VLA_Prolog.PROLOG("(query (nth 1 (cons a nil)))"))
    Report "prolog.13: ...and (nth ...) with two is refused with the OTHER count, from the same message", _
           InStr(1, r, "(nth ...)", vbTextCompare) > 0 And InStr(1, r, "exactly 3 arguments", vbTextCompare) > 0, "got: " & r

    ' Reserved, like every other dispatched name - forward declaration, so
    ' a knowledge base written today cannot be broken by a goal that ships
    ' later.
    Dim listName As Variant
    For Each listName In Array("length", "member", "nth", "append", "reverse", "sum-list")
        r = CStr(VLA_Prolog.PROLOG("(fact (" & listName & " a b c)) (query (p X))"))
        Report "prolog.13: '" & listName & "' is refused as a predicate name in a (fact ...)", _
               InStr(1, r, "reserved word", vbTextCompare) > 0, "got: " & r
    Next listName

    ' ...and the refusal must ENUMERATE them, not merely refuse. A message
    ' still listing only the sets that existed before this item would be
    ' quietly wrong about what a user may name a predicate.
    r = CStr(VLA_Prolog.PROLOG("(fact (member a b c)) (query (p X))"))
    Report "prolog.13: ...and the refusal names the six list goals, so the advice can never be wrong", _
           InStr(1, r, "six list goals", vbTextCompare) > 0 And InStr(1, r, "sum-list", vbTextCompare) > 0, "got: " & r

    ' `cons` and `nil` are deliberately NOT reserved - they are data, a
    ' functor and an atom, and this pins that decision rather than leaving
    ' it to prose. A user may still define a predicate called cons.
    result = VLA_Prolog.PROLOG("(fact (cons a b)) (query (cons a Y))")
    Report "prolog.13: `cons` is NOT reserved - it is a functor, not a goal, so a user may still define it", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "b"), "got: " & ResultDescribe(result)

    ' ================= PROLOG.21: the (list ...) shorthand =============
    ' NOT a new representation - cons cells are unchanged, and every pin
    ' above still passes. What changes is the two ends: how a list is
    ' READ and how it is WRITTEN. Both, together, because either alone
    ' would be a display that lies.

    ' ---- READING. The sugar and the cons spelling must produce the SAME
    ' TERM, not merely similar output - so this asserts it by UNIFYING
    ' one against the other, which is the strongest available statement
    ' of "same term" in this engine.
    result = VLA_Prolog.PROLOG("(query (= (list a b c) (cons a (cons b (cons c nil)))))")
    Report "prolog.21: (list a b c) and the cons chain are the SAME TERM, proved by unifying them", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (= (list) nil))")
    Report "prolog.21: ...and an empty (list) is the atom nil", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    ' The discriminating twin: sugar that expanded WRONG would still
    ' unify with something, so a pin that only ever succeeds proves
    ' little. This one must FAIL.
    result = VLA_Prolog.PROLOG("(query (= (list a b) (cons a (cons b (cons c nil)))))")
    Report "prolog.21: ...and a list of two is NOT the chain of three - the sugar is not merely 'some list'", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)

    ' ---- WRITING. The rendering is where the owner feels this item.
    result = VLA_Prolog.PROLOG("(fact (color red)) (fact (color green)) (fact (color blue)) (query (findall X (color X) Bag))")
    Report "prolog.21: a findall bag RENDERS as (list ...), not as a cons chain", _
           ResultCol1Is(result, "(list red green blue)"), "got: " & ResultDescribe(result)
    ' PROLOG.22: a defined predicate with no matching row, not an undefined
    ' `nothing`, which now refuses before a bag exists to render.
    result = VLA_Prolog.PROLOG("(fact (p 1 a)) (query (findall X (p 2 X) Bag))")
    Report "prolog.21: ...and an empty bag still renders as the atom nil, not an empty (list)", _
           ResultCol1Is(result, "nil"), "got: " & ResultDescribe(result)

    ' ---- ONLY A PROPER LIST CONTRACTS, and this pair is the sharpest in
    ' the Sub. An improper cons cell must print as itself: if the
    ' terminator check were dropped, (cons a b) would print as (list a)
    ' and read back as (cons a nil), SILENTLY DISCARDING the b. Proved
    ' load-bearing by mutation before import; held here against the real
    ' engine.
    result = VLA_Prolog.PROLOG("(query (= X (cons a b)))")
    Report "prolog.21: an IMPROPER (cons a b) prints as itself - contraction never invents a list", _
           ResultCol1Is(result, "(cons a b)"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (= X (cons a nil)))")
    Report "prolog.21: ...while the proper one beside it does contract", _
           ResultCol1Is(result, "(list a)"), "got: " & ResultDescribe(result)
    ' A list nested inside an ordinary compound term still contracts, and
    ' a list OF improper cells keeps its elements honest.
    result = VLA_Prolog.PROLOG("(query (= X (pair (list a b) z)))")
    Report "prolog.21: a list nested inside an ordinary compound term contracts too", _
           ResultCol1Is(result, "(pair (list a b) z)"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (= X (list (cons a b))))")
    Report "prolog.21: a LIST OF improper cells contracts the list and leaves the cells alone", _
           ResultCol1Is(result, "(list (cons a b))"), "got: " & ResultDescribe(result)

    ' ---- THE LAW, RUN LIVE. Everything above tests one direction. This
    ' tests the property that makes (list ...) a shorthand rather than a
    ' display: the text a cell shows must READ BACK as the same term. The
    ' rendered output of one query is spliced into the SOURCE of a
    ' second, and the second asserts the round trip by unification. If
    ' rendering and reading ever disagreed, this is the pin that catches
    ' it - and no amount of one-directional testing would.
    Dim renderedBag As String
    result = VLA_Prolog.PROLOG("(fact (color red)) (fact (color green)) (query (findall X (color X) Bag))")
    If ResultRowCount(result) = 2 Then renderedBag = ResultCellText(result, 2, 1)
    Report "prolog.21: the rendered form is re-readable at all (a non-empty string came back)", _
           Len(renderedBag) > 0, "got: '" & renderedBag & "'"
    result = VLA_Prolog.PROLOG( _
        "(fact (color red)) (fact (color green)) " & _
        "(query (findall X (color X) Bag) (= Bag " & renderedBag & "))")
    Report "prolog.21: THE ROUND TRIP - a bag's own rendered text, read back, unifies with the bag itself", _
           ResultRowCount(result) = 2, "got: " & ResultDescribe(result)
    ' ...and its discriminating twin, so the round trip is not passing
    ' because `=` succeeds against anything: the same text with one
    ' element changed must NOT unify.
    result = VLA_Prolog.PROLOG( _
        "(fact (color red)) (fact (color green)) " & _
        "(query (findall X (color X) Bag) (= Bag " & Replace(renderedBag, "green", "blue") & "))")
    Report "prolog.21: ...and the same text with one element changed does NOT unify", _
           ResultRowCount(result) = 1, "got: " & ResultDescribe(result)

    ' ---- THE GOAL/DATA SPLIT, which is what lets `list` be reserved AND
    ' dispatched. In a DATA position the sugar fires; in a GOAL position
    ' nothing expands it and it reaches the solver, where it is refused.
    r = CStr(VLA_Prolog.PROLOG("(query (list a b))"))
    Report "prolog.21: (list ...) written where a GOAL belongs is refused by name, not silently failed", _
           InStr(1, r, "not something PROLOG can prove", vbTextCompare) > 0, "got: " & r
    r = CStr(VLA_Prolog.PROLOG("(fact (list a b)) (query (p X))"))
    Report "prolog.21: ...and `list` is refused as a predicate name in a (fact ...)", _
           InStr(1, r, "reserved word", vbTextCompare) > 0, "got: " & r
    ' The nested goal positions. `not`'s argument and findall's Goal are
    ' the ONLY ones, so these two hold the flag in the direction that
    ' would otherwise eat a goal.
    r = CStr(VLA_Prolog.PROLOG("(query (not (list a b)))"))
    Report "prolog.21: `not`'s own argument is a GOAL, so a list there is refused rather than expanded", _
           InStr(1, r, "not something PROLOG can prove", vbTextCompare) > 0, "got: " & r
    r = CStr(VLA_Prolog.PROLOG("(query (findall X (list a) Bag))"))
    Report "prolog.21: ...and so is findall's own Goal", _
           InStr(1, r, "not something PROLOG can prove", vbTextCompare) > 0, "got: " & r
    ' ...and the other direction: findall's Template and Bag are DATA, so
    ' a list in either must expand normally. Without that, this query
    ' would refuse instead of answering.
    result = VLA_Prolog.PROLOG("(fact (p a 1)) (fact (p b 2)) (query (findall (list X Y) (p X Y) Bag))")
    Report "prolog.21: findall's TEMPLATE is data - a list there expands, giving a list of lists", _
           ResultCol1Is(result, "(list (list a 1) (list b 2))"), "got: " & ResultDescribe(result)

    ' ---- SUGAR IS ACCEPTED WHEREVER A TERM GOES, not just in queries.
    result = VLA_Prolog.PROLOG("(fact (colors (list red green blue))) (query (colors L) (length L N))")
    Report "prolog.21: a (list ...) written in a FACT is stored as a real list and measures 3", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 2, "3"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(fact (ok yes)) (rule (heads (list H T) H) (ok yes)) (query (heads (list x y) F))")
    Report "prolog.21: ...and in a RULE HEAD, where it expands before the head is ever matched", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "x"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (member b (list a b c)))")
    Report "prolog.21: ...and as an argument to a list goal, which is where it will mostly be typed", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (append (list a) (list b) C))")
    Report "prolog.21: ...on both sides of an append, rendering the join in the same spelling", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "(list a b)"), "got: " & ResultDescribe(result)

    ' A NESTED list written as sugar - the recursion, in both directions
    ' at once: written nested, expanded nested, contracted nested.
    result = VLA_Prolog.PROLOG("(query (= X (list (list a b) (list c))))")
    Report "prolog.21: nested sugar expands and contracts recursively, round-tripping to itself", _
           ResultCol1Is(result, "(list (list a b) (list c))"), "got: " & ResultDescribe(result)
End Sub

' ---------------------------------------------------------------------
'  PROLOG.5.3: VLA_Prolog.PROLOG - `findall`. Covers: a bare-variable
'  Template harvesting every solution into a list, in derivation order;
'  zero solutions collapsing to an empty list; a COMPOUND Template,
'  proving SubstituteTemplate's own reconstruction (not just a flat
'  value list); findall composing with an outer already-bound variable
'  across real backtracking (the "group by" idiom); an already-bound Bag
'  argument genuinely checked via unification, both matching and
'  mismatching; the shared step ceiling and the error-propagates-out
'  case; the real design fork this item's own build found (a Template
'  containing a literal `(not X)`-shaped sub-term as data still has its
'  own variable collected correctly); and every named shape refusal.
' ---------------------------------------------------------------------
Private Sub TestPrologFindall()
    Dim result As Variant
    Dim r As String

    ' A bare-variable Template, every solution harvested into a list, in
    ' derivation (candidate) order.
    ' PROLOG.13 RE-POINTS THIS PIN, and the three below it, and the two
    ' at the end of this Sub. The bag is a CONS CHAIN now, not a headless
    ' Collection - `(cons red (cons green (cons blue nil)))` where this
    ' used to read `(red green blue)`. Re-pointed rather than deleted,
    ' the fourth marker pin to move (PROLOG.8 moved PROLOG.7's,
    ' PROLOG.10 moved PROLOG.8's, PROLOG.12 moved PROLOG.10's): the
    ' behaviour being pinned - every solution harvested, in derivation
    ' order - has not changed at all, only how the resulting list is
    ' spelled, so deleting the pin would lose a property that still
    ' holds. See VLA_Prolog.bas's own PROLOG.13 header for why the
    ' representation changed and what the old headless shape could not do.
    result = VLA_Prolog.PROLOG("(fact (color red)) (fact (color green)) (fact (color blue)) (query (findall X (color X) Bag))")
    Report "prolog.5.3/13: findall harvests every solution into a list, in derivation order", _
           ResultCol1Is(result, "(list red green blue)"), "got: " & ResultDescribe(result)

    ' Zero solutions collapses to an empty list, never an error - real
    ' findall's own signature behavior (unlike bagof/setof). PROLOG.13:
    ' the empty list is now the ATOM `nil`, where it used to be a
    ' zero-length Collection rendering as `()`. That single change is
    ' what lets the ISO type-test edge case below answer correctly.
    ' PROLOG.22: "no solutions" means a DEFINED goal that matches nothing -
    ' p has no row whose first column is 2. This pin and the three below
    ' used to ask an undefined `nonexistent`, and "not an error" is no
    ' longer true of that: an undefined predicate now refuses, pinned in
    ' TestPrologUnknownPredicate.
    result = VLA_Prolog.PROLOG("(fact (p 1 a)) (query (findall X (p 2 X) Bag))")
    Report "prolog.5.3/13: findall over a goal with no solutions binds Bag to the empty list nil, not an error", _
           ResultCol1Is(result, "nil"), "got: " & ResultDescribe(result)

    ' A COMPOUND Template - proves SubstituteTemplate's own
    ' reconstruction of Template's structure per solution, not just a
    ' flat list of Template's own variable values.
    result = VLA_Prolog.PROLOG( _
        "(fact (person alice 30)) (fact (person bob 25)) " & _
        "(query (findall (pair Name Age) (person Name Age) Bag))")
    Report "prolog.5.3/13: a compound Template is reconstructed per solution, not flattened", _
           ResultCol1Is(result, "(list (pair alice 30) (pair bob 25))"), "got: " & ResultDescribe(result)

    ' findall composing with an OUTER already-bound variable across real
    ' backtracking - the classic "group by" idiom: for each department
    ' (bound by an earlier, ordinary conjunct), harvest its own
    ' employees into a separate list.
    result = VLA_Prolog.PROLOG( _
        "(fact (dept eng)) (fact (dept sales)) " & _
        "(fact (employee alice eng)) (fact (employee bob eng)) (fact (employee carol sales)) " & _
        "(query (dept D) (findall E (employee E D) Bag))")
    ' PROLOG.13 re-points the two bag spellings here AND rewrites the
    ' assertion itself onto the guarded helpers. The original built one
    ' combined expression - UBound(result, 1) = 3 And CStr(result(2, 1))
    ' = ... - and VBA's And does not short-circuit, so on a HEADER-ONLY
    ' result (free variables, zero solutions, UBound 1) every one of
    ' those indexes still evaluated and the run would have died with
    ' "Subscript out of range" at the exact moment it was about to report
    ' the failure. It never fired because the test passed; a failing
    ' assertion written that way kills the run instead of reporting it,
    ' which is what ResultRowCount/ResultCellIs exist to prevent.
    Dim groupByOk As Boolean
    groupByOk = (ResultRowCount(result) = 3)
    If groupByOk Then groupByOk = ResultCellIs(result, 2, 1, "eng")
    If groupByOk Then groupByOk = ResultCellIs(result, 2, 2, "(list alice bob)")
    If groupByOk Then groupByOk = ResultCellIs(result, 3, 1, "sales")
    If groupByOk Then groupByOk = ResultCellIs(result, 3, 2, "(list carol)")
    Report "prolog.5.3/13: findall composes with an outer already-bound variable across real backtracking (group-by)", _
           groupByOk, "got: " & ResultDescribe(result)

    ' An already-bound Bag argument is genuinely CHECKED via unification,
    ' not just always bound fresh - both the matching and mismatching
    ' case.
    ' PROLOG.13: both re-pointed at cons spellings. This pair is the one
    ' that proves the representation change is real rather than cosmetic
    ' - the Bag argument is now WRITTEN as a cons chain in the program
    ' text and unified against the harvested one, so a bag and a
    ' hand-written list have to be the same term for the first pin to
    ' pass, and the old headless `(red green)` would fail both. It is
    ' also the round trip the old shape could never do: `(red green)`
    ' written down reads back as a term whose functor is `red`.
    result = VLA_Prolog.PROLOG("(fact (color red)) (fact (color green)) (query (findall X (color X) (cons red (cons green nil))))")
    Report "prolog.5.3/13: an already-bound Bag matching the harvested list succeeds", _
           ResultBoolIs(result, True), "got " & TypeName(result) & " " & result

    result = VLA_Prolog.PROLOG("(fact (color red)) (fact (color green)) (query (findall X (color X) (cons green (cons red nil))))")
    Report "prolog.5.3/13: an already-bound Bag NOT matching the harvested list (wrong order) fails", _
           ResultBoolIs(result, False), "got " & TypeName(result) & " " & result

    ' The shared ceiling, run to real exhaustion inside findall's own Goal -
    ' the identical sub-call microscope `not` already proves this for.
    r = CStr(VLA_Prolog.PROLOG("(rule (loop X) (loop X)) (query (findall X (loop a) Bag))"))
    Report "prolog.5.3: a genuinely non-terminating goal inside findall is refused by the shared depth ceiling, not left to hang or crash", _
           InStr(1, r, "rules deep", vbTextCompare) > 0 And InStr(1, r, "facts and rules", vbTextCompare) = 0, "got: " & r

    ' An error (not a mere failure) raised while proving findall's own
    ' Goal must propagate all the way out - findall's own half of "what
    ' must never leak back out."
    r = CStr(VLA_Prolog.PROLOG("(query (findall X (is X (/ 5 0)) Bag))"))
    Report "prolog.5.3: an error raised while proving findall's own goal propagates out, rather than being swallowed as an empty bag", _
           InStr(1, r, "divide by zero", vbTextCompare) > 0, "got: " & r

    ' The real design fork this item's own build found: a Template that
    ' itself contains a literal `(not X)`-shaped sub-term, as ordinary
    ' DATA (building a list of negation-shaped terms - a legitimate
    ' real-Prolog idiom), must still have X collected correctly - proving
    ' CollectTemplateVars' own deliberately non-goal-aware recursion
    ' (reusing goal-aware CollectVars here instead would have silently
    ' skipped X, since (not X) also happens to match CollectVars' own
    ' negation skip-shape check).
    result = VLA_Prolog.PROLOG("(fact (thing a)) (query (findall (not X) (thing X) Bag))")
    Report "prolog.5.3/13: a Template containing a literal (not X)-shaped sub-term still has its own variable collected correctly", _
           ResultCol1Is(result, "(list (not a))"), "got: " & ResultDescribe(result)

    ' Refusals.
    r = CStr(VLA_Prolog.PROLOG("(fact (foo a)) (query (findall X (foo X)))"))
    Report "prolog.5.3: (findall ...) with only two arguments (missing Bag) is refused", _
           InStr(1, r, "exactly three arguments", vbTextCompare) > 0, "got: " & r

    r = CStr(VLA_Prolog.PROLOG("(fact (foo a)) (query (findall X (foo X) Bag Extra))"))
    Report "prolog.5.3: (findall ...) with four arguments is refused", _
           InStr(1, r, "exactly three arguments", vbTextCompare) > 0, "got: " & r

    r = CStr(VLA_Prolog.PROLOG("(query (findall X foo Bag))"))
    Report "prolog.5.3: a bare non-! atom as findall's own goal reuses the ordinary atom-not-a-list refusal", _
           InStr(1, r, "predicate form", vbTextCompare) > 0, "got: " & r

    ' PROLOG.5.4: a bare ! is legal as findall's own Goal too - it
    ' trivially succeeds exactly once, so findall harvests exactly one
    ' Template instantiation. A ground Template (not a variable) keeps
    ' this case simple; TestPrologCut, below, is where cut's own opacity
    ' inside findall's isolated sub-search is actually proven.
    result = VLA_Prolog.PROLOG("(query (findall done ! Bag))")
    Report "prolog.5.3/13: a bare ! as findall's own goal is legal (PROLOG.5.4) and harvests exactly one Template instantiation", _
           ResultCol1Is(result, "(list done)"), "got: " & ResultDescribe(result)

    ' ---- PROLOG.13: THE ISO EDGE CASE PROLOG.9 RECORDED AND NEVER
    ' PINNED. PROLOG.9's own header claimed `(compound EmptyBag)` was
    ' "pinned rather than special-cased" as True. It was not pinned at
    ' all - no test in this module ever combined findall with a type
    ' test, so the claim had never once been executed. It was also the
    ' non-ISO answer, forced by representation: an empty bag was a
    ' zero-length Collection, and SolveTypeTest classifies by IsObject.
    '
    ' PROLOG.13's nil is an ATOM, so the same unchanged SolveTypeTest now
    ' gives ISO's answer. Not one line of it moved. These four run the
    ' bag through a REAL findall rather than writing nil literally,
    ' deliberately - a literal would prove only that `nil` is an atom,
    ' which was never in doubt; what needs pinning is that findall's own
    ' empty bag IS that atom.
    ' NOT a bare Boolean, and the distinction is the one that caught the
    ' previous session three times: `Bag` is a FREE variable, and a query
    ' with a free variable ALWAYS spills - as a HEADER-ONLY array when it
    ' finds nothing. Only a query with no free variables at all collapses
    ' to a Boolean. So "this goal failed" is spelled here as a row count
    ' of 1 (the header alone), never as result = False.
    ' PROLOG.22: the empty bag comes from a defined predicate with no
    ' matching row, as in TestPrologFindall's own empty-bag pin.
    result = VLA_Prolog.PROLOG("(fact (p 1 a)) (query (findall X (p 2 X) Bag) (compound? Bag))")
    Report "prolog.13: (compound? EmptyBag) is now FALSE - the ISO answer, where PROLOG.9's representation forced True", _
           ResultRowCount(result) = 1, "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(fact (p 1 a)) (query (findall X (p 2 X) Bag) (atomic? Bag))")
    Report "prolog.13: ...and (atomic? EmptyBag) is TRUE - the same judgement seen from the other side", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "nil"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(fact (p 1 a)) (query (findall X (p 2 X) Bag) (atom? Bag))")
    Report "prolog.13: ...and (atom? EmptyBag) is TRUE - nil is an ordinary atom, not a number and not a variable", _
           ResultRowCount(result) = 2, "got: " & ResultDescribe(result)

    ' THE DISCRIMINATING TWIN, without which the three above would pass
    ' against an engine that answered False to every type test: a
    ' NON-empty bag is still compound, because it really is a cons cell.
    result = VLA_Prolog.PROLOG("(fact (color red)) (query (findall X (color X) Bag) (compound? Bag))")
    Report "prolog.13: ...while a NON-empty bag IS compound - the twin that makes the three above discriminating", _
           ResultRowCount(result) = 2 And ResultCol1Is(result, "(list red)"), "got: " & ResultDescribe(result)
End Sub

' ---------------------------------------------------------------------
'  PROLOG.5.4: VLA_Prolog.PROLOG - cut (`!`). Covers: cut committing to
'  the FIRST matching clause of the same predicate call it appears in
'  (no further candidates tried, even though more would otherwise
'  match); the flagship proof this item's own design required - cut
'  ALSO prunes EARLIER goals in the SAME clause body, not just the
'  current predicate's own remaining clauses, hand-traced before being
'  written (a two-predicate rule body that would otherwise have four
'  solutions collapses to exactly one); the opposite boundary, equally
'  load-bearing - cut does NOT escape past its own clause's own
'  selection to prune a caller's later, unrelated choice point, and does
'  NOT affect goals to its own RIGHT in the same body, which keep their
'  full normal backtracking; cut's own opacity across both `not`'s and
'  `findall`'s isolated sub-search (real Prolog's own rule), each proven
'  by a query where a LEAKED signal would visibly prune an unrelated
'  outer choice point and a correctly-opaque one would not; multiple
'  cuts in one clause body composing without double-firing or
'  corrupting each other's own barrier; and the shared PROLOG_MAX_STEPS
'  budget re-proven against the identical genuinely-non-terminating rule
'  every other PROLOG.5.x item already re-proves it against.
' ---------------------------------------------------------------------
Private Sub TestPrologCut()
    Dim result As Variant
    Dim r As String

    ' Cut commits to the FIRST matching clause - p has three facts, but
    ' only the first (a) is ever tried; b and c are never reached. Checks
    ' the ROW COUNT explicitly, not just row 2's own value (ResultCol1Is
    ' alone would still pass if cut failed to prune and b/c leaked in as
    ' extra rows 3/4 - the exact gap a weaker assertion here would miss).
    result = VLA_Prolog.PROLOG("(fact (p a)) (fact (p b)) (fact (p c)) (query (p X) !)")
    Dim commitOk As Boolean
    If IsArray(result) Then commitOk = (ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, "a"))
    Report "prolog.5.4: cut commits to the first matching clause, pruning every later candidate of the same predicate", _
           commitOk, "got: " & ResultDescribe(result)

    ' The flagship proof: cut prunes EARLIER goals in the SAME clause
    ' body too, not just the predicate it's textually attached to.
    ' Without the cut, (p X)(q Y) would have FOUR solutions (a/1, a/2,
    ' b/1, b/2) - hand-traced before writing this assertion: the cut
    ' commits to the FIRST successful path through the whole body,
    ' pruning both q's own remaining candidate (q 2) AND p's own
    ' remaining candidate (p b), leaving exactly ONE solution.
    result = VLA_Prolog.PROLOG( _
        "(fact (p a)) (fact (p b)) (fact (q 1)) (fact (q 2)) " & _
        "(rule (test X Y) (p X) (q Y) !) " & _
        "(query (test X Y))")
    Dim scopeOk As Boolean
    If IsArray(result) Then scopeOk = (ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, "a") And ResultCellIs(result, 2, 2, "1"))
    Report "prolog.5.4: cut prunes EARLIER goals in the same clause body, not just its own predicate's remaining clauses (4 possible solutions collapse to 1)", _
           scopeOk, "got: " & ResultDescribe(result)

    ' The opposite boundary: cut does NOT escape past its own clause's
    ' own selection to prune a caller's LATER, unrelated choice point,
    ' and does NOT affect goals to its own RIGHT in the same body (r's
    ' own two facts both still appear). q has two p-facts to choose from
    ' but only the first (a) survives its own internal cut; r, called
    ' AFTER q returns, keeps its full normal backtracking over both of
    ' its own facts.
    result = VLA_Prolog.PROLOG( _
        "(fact (p a)) (fact (p b)) (rule (q X) (p X) !) (fact (r 1)) (fact (r 2)) " & _
        "(query (q X) (r Y))")
    Dim boundaryOk As Boolean
    If IsArray(result) Then
        boundaryOk = (ResultRowCount(result) = 3 And ResultCellIs(result, 2, 1, "a") And ResultCellIs(result, 2, 2, "1") _
                      And ResultCellIs(result, 3, 1, "a") And ResultCellIs(result, 3, 2, "2"))
    End If
    Report "prolog.5.4: cut does not escape its own clause to prune a caller's later choice point, and goals to cut's own right keep full backtracking", _
           boundaryOk, "got: " & ResultDescribe(result)

    ' Cut's own opacity across `not`'s isolated sub-search (real Prolog's
    ' own rule): blocked/0 reaches a cut (pruning thing's own second
    ' fact) and then FAILS outright ((thing 3) matches neither thing
    ' fact), so blocked/0 has zero solutions and (not (blocked)) is always
    ' TRUE, regardless of Y. If the internal cut leaked out of the isolated
    ' sub-search, the outer p(Y) loop would wrongly stop after Y=a; a
    ' correctly-opaque cut leaves it fully backtracking over both facts.
    ' PROLOG.22 re-pointed the failing goal: it was an undefined
    ' (nonexistent Z), which now refuses the program before solving.
    result = VLA_Prolog.PROLOG( _
        "(fact (p a)) (fact (p b)) (fact (thing 1)) (fact (thing 2)) " & _
        "(rule (blocked) (thing W) ! (thing 3)) " & _
        "(query (p Y) (not (blocked)))")
    Dim notOpaqueOk As Boolean
    If IsArray(result) Then notOpaqueOk = (ResultRowCount(result) = 3 And ResultCellIs(result, 2, 1, "a") And ResultCellIs(result, 3, 1, "b"))
    Report "prolog.5.4: a cut fired inside not's own isolated goal never leaks out to prune the outer query's own unrelated choice point", _
           notOpaqueOk, "got: " & ResultDescribe(result)

    ' The identical opacity proof for findall's own isolated sub-search:
    ' grab/1 reaches a cut after its first thing-fact, so grab(Z) always
    ' harvests exactly one element, regardless of Y - if the internal cut
    ' leaked, the outer p(Y) loop would wrongly stop after Y=a.
    '
    ' PROLOG.13 re-points the two bag spellings - `(1)` became
    ' `(cons 1 nil)` - and rewrites the assertion onto the guarded
    ' helpers while it is here. The original built one combined
    ' expression, and VBA's And does not short-circuit, so the moment
    ' this test genuinely FAILED - a leaked cut gives 1 solution, not 2 -
    ' `result(3, 1)` would still have been evaluated and raised
    ' "Subscript out of range", killing the run instead of reporting the
    ' leak it exists to catch. It never fired only because it passed.
    result = VLA_Prolog.PROLOG( _
        "(fact (p a)) (fact (p b)) (fact (thing 1)) (fact (thing 2)) " & _
        "(rule (grab X) (thing X) !) " & _
        "(query (p Y) (findall Z (grab Z) Bag))")
    Dim findallOpaqueOk As Boolean
    findallOpaqueOk = (ResultRowCount(result) = 3)
    If findallOpaqueOk Then findallOpaqueOk = ResultCellIs(result, 2, 1, "a")
    If findallOpaqueOk Then findallOpaqueOk = ResultCellIs(result, 2, 2, "(list 1)")
    If findallOpaqueOk Then findallOpaqueOk = ResultCellIs(result, 3, 1, "b")
    If findallOpaqueOk Then findallOpaqueOk = ResultCellIs(result, 3, 2, "(list 1)")
    Report "prolog.5.4/13: a cut fired inside findall's own isolated goal never leaks out to prune the outer query's own unrelated choice point", _
           findallOpaqueOk, "got: " & ResultDescribe(result)

    ' Multiple cuts in ONE clause body compose without double-firing or
    ' corrupting each other's own barrier (both share the SAME
    ' per-invocation suffix, freshened identically) - only a(1)/b(10)
    ' ever appear; a(2) and b(20) are both pruned.
    result = VLA_Prolog.PROLOG( _
        "(fact (a 1)) (fact (a 2)) (fact (b 10)) (fact (b 20)) " & _
        "(rule (r X Y) (a X) ! (b Y) !) " & _
        "(query (r X Y))")
    Dim multiCutOk As Boolean
    If IsArray(result) Then multiCutOk = (ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, "1") And ResultCellIs(result, 2, 2, "10"))
    Report "prolog.5.4: two cuts in one clause body compose correctly, sharing one barrier, without double-firing", _
           multiCutOk, "got: " & ResultDescribe(result)

    ' The shared step ceiling, re-proven with the two new cut-signal
    ' parameters now threaded through every recursive frame - the
    ' identical genuinely non-terminating rule PROLOG.4/5.2/5.3 each
    ' already run to real exhaustion, with no cut anywhere in it, so this
    ' isolates whether cut's own new machinery alone changed the ceiling.
    r = CStr(VLA_Prolog.PROLOG("(rule (loop X) (loop X)) (query (loop a))"))
    Report "prolog.5.4: a genuinely non-terminating rule (no cut involved) is still refused cleanly by the shared depth ceiling with cut's new parameters threaded through every frame", _
           InStr(1, r, "rules deep", vbTextCompare) > 0 And InStr(1, r, "facts and rules", vbTextCompare) = 0, "got: " & r
End Sub

' ---------------------------------------------------------------------
'  PROLOG.6: VLA_Prolog.PrologRun - named-column (keyed-atom) desugaring.
'  Pure - calls PrologRun directly with a hand-built headerMap (no live
'  Table needed at all, DATALOG.5's own TestDatalogKeyedAtoms precedent),
'  so table-sourced FACT INJECTION itself (which genuinely needs a live
'  Range) is TestPrologHostTable's own job, below. Covers: a fully-keyed
'  query resolving correctly, asserted against both value AND derivation
'  order; a NON-table-sourced predicate's own compound-term arguments
'  left completely untouched (the real ambiguity this item's own header
'  names - a 2-element-list argument is only ever a keyed pair against a
'  KNOWN table-sourced predicate, never assumed); every named refusal
'  (unknown column, a column keyed twice, inconsistent keying); a
'  RULE body composing full keying with real backtracking; and two
'  DISCRIMINATING nested-desugaring proofs (inside `not` and inside
'  `findall`) - each built so a desugaring bug that silently failed to
'  fire would flip the assertion's own expected TRUTH VALUE, not just
'  produce an incidentally-matching result.
' ---------------------------------------------------------------------
' ---------------------------------------------------------------------
'  PROLOG.14: VLA_Prolog.PROLOG - disjunction `(or ...)` and if-then-else
'  `(if ...)`. Covers: both branches of a disjunction contributing; the
'  case where the FIRST branch fails and the second carries the answer
'  (an unknown predicate fails SILENTLY, so a disjunction test whose
'  first branch already succeeds proves nothing about the second);
'  neither branch succeeding; cut transparency through both forms;
'  if-then-else committing to its condition's first solution; the else
'  running only when the condition never succeeded - never when the
'  condition succeeded and the then-branch failed; both `if` arities;
'  the phantom-column rule (a variable is an output column only if every
'  success path binds it); the ISO spellings `->` and `\+` reaching a
'  refusal that names the house form; every shape refusal and every
'  reserved-name refusal; and the PROLOG.5.4-era cut-signal defect this
'  item found and repaired, with the control that isolates it.
' ---------------------------------------------------------------------
Private Sub TestPrologControl()
    Dim result As Variant
    Dim r As String
    Dim facts As String
    facts = "(fact (p 1)) (fact (p 2)) (fact (q a)) (fact (q b)) "

    ' ---- DISJUNCTION -------------------------------------------------
    ' Both branches are explored, in written order, and the continuation
    ' runs under each. Four solutions from two two-fact predicates.
    result = VLA_Prolog.PROLOG(facts & "(rule (both X) (or (p X) (q X))) (query (both X))")
    Dim bothOk As Boolean
    If IsArray(result) Then
        bothOk = (ResultRowCount(result) = 5 And ResultCellIs(result, 2, 1, "1") And ResultCellIs(result, 3, 1, "2") _
                  And ResultCellIs(result, 4, 1, "a") And ResultCellIs(result, 5, 1, "b"))
    End If
    Report "prolog.14: (or ...) explores every branch, in written order - two two-fact predicates give four solutions", _
           bothOk, "got: " & ResultDescribe(result)

    ' THE DISCRIMINATION CASE. A disjunction whose FIRST branch already
    ' succeeds would pass against an implementation that never looked at
    ' the second one at all. This is the shape that cannot: the first
    ' branch matches no fact - q holds a and b, never c - so every row
    ' below comes from the second. PROLOG.22 re-pointed that first branch:
    ' it was an undefined (nosuch X), and an unknown predicate no longer
    ' fails silently - it refuses the program before solving.
    result = VLA_Prolog.PROLOG(facts & "(rule (f X) (or (q c) (p X))) (query (f X))")
    Dim secondOk As Boolean
    If IsArray(result) Then
        secondOk = (ResultRowCount(result) = 3 And ResultCellIs(result, 2, 1, "1") And ResultCellIs(result, 3, 1, "2"))
    End If
    Report "prolog.14: when the FIRST branch fails, the second carries the whole answer", _
           secondOk, "got: " & ResultDescribe(result)

    ' ...and its twin, so the test above cannot be passing because the
    ' rule matches everything: with BOTH branches matching nothing there
    ' are no rows at all, and a query with a free variable still spills
    ' its header row, so this is 1 rather than 0. (Both branches were
    ' undefined predicates before PROLOG.22; see the pin above.)
    result = VLA_Prolog.PROLOG(facts & "(rule (g X) (or (q c) (p 9))) (query (g X))")
    Report "prolog.14: ...and with BOTH branches failing there are no rows, only the header", _
           ResultRowCount(result) = 1, "got: " & ResultDescribe(result)

    ' Cut is TRANSPARENT through a disjunction: the bare ! written as the
    ' first branch cuts the ENCLOSING clause, so p's own second fact is
    ' pruned and the second branch is never tried. Without transparency
    ' this would be four rows (X = 1 and 2, each with Y = a and b); the
    ' row count is asserted explicitly because that is the whole
    ' difference.
    result = VLA_Prolog.PROLOG(facts & "(rule (cutor X Y) (p X) (or ! (q Y)) (q Y)) (query (cutor X Y))")
    Dim cutOrOk As Boolean
    If IsArray(result) Then
        cutOrOk = (ResultRowCount(result) = 3 And ResultCellIs(result, 2, 1, "1") And ResultCellIs(result, 2, 2, "a") _
                   And ResultCellIs(result, 3, 1, "1") And ResultCellIs(result, 3, 2, "b"))
    End If
    Report "prolog.14: a bare ! inside a branch cuts the enclosing clause - cut is transparent through (or ...)", _
           cutOrOk, "got: " & ResultDescribe(result)

    ' ---- THE PHANTOM COLUMN ------------------------------------------
    ' The defect this item's own design exists to avoid, and PROLOG.5.2's
    ' finding reached by a fourth route. X is bound only by the first
    ' branch and Y only by the second, so on any given solution one of
    ' them is free - collecting either would render its own raw name into
    ' a spilled cell as though it were a value. Neither is a column, so
    ' the query has no free variables at all and collapses to a BOOLEAN.
    result = VLA_Prolog.PROLOG("(fact (p 1)) (fact (q a)) (query (or (p X) (q Y)))")
    Report "prolog.14: (or (p X) (q Y)) binds X on one branch and Y on the other, so NEITHER is an output column - a boolean, not two phantom columns", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)

    ' ...and the twin that proves the rule is not simply "never collect
    ' from a disjunction": a variable EVERY branch binds is a real column
    ' and does spill, one row per branch that succeeds.
    result = VLA_Prolog.PROLOG("(fact (p 1)) (fact (q a)) (query (or (p X) (q X)))")
    Dim sharedOk As Boolean
    If IsArray(result) Then
        sharedOk = (ResultRowCount(result) = 3 And ResultColCount(result) = 1 _
                    And ResultCellIs(result, 2, 1, "1") And ResultCellIs(result, 3, 1, "a"))
    End If
    Report "prolog.14: ...but a variable EVERY branch binds is a real column and spills", _
           sharedOk, "got: " & ResultDescribe(result)

    ' A branch is a GOAL, not data, so `not`'s own skip still applies
    ' inside one: the negated branch binds nothing, so X is not a column
    ' and this is a boolean. Collected as data it would put X back.
    result = VLA_Prolog.PROLOG("(fact (p 1)) (fact (q a)) (query (or (not (p X)) (q X)))")
    Report "prolog.14: a branch is collected as a GOAL - not's own skip still applies inside one, so X is no column here", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)

    ' ---- IF-THEN-ELSE ------------------------------------------------
    ' The condition COMMITS to its first solution: p has two facts but
    ' only p(1) is ever used, so this is two rows and not four. The
    ' else-goal never runs here; PROLOG.22 made it one that matches nothing,
    ' (q c), where it was an undefined (nope Y) - an unknown predicate now
    ' refuses the program even on a branch the data never reaches.
    result = VLA_Prolog.PROLOG(facts & "(rule (ite X Y) (if (p X) (q Y) (q c))) (query (ite X Y))")
    Dim iteOk As Boolean
    If IsArray(result) Then
        iteOk = (ResultRowCount(result) = 3 And ResultCellIs(result, 2, 1, "1") And ResultCellIs(result, 2, 2, "a") _
                 And ResultCellIs(result, 3, 1, "1") And ResultCellIs(result, 3, 2, "b"))
    End If
    Report "prolog.14: (if ...) commits to the condition's FIRST solution - p's two facts give two rows, not four", _
           iteOk, "got: " & ResultDescribe(result)

    ' THE DECISIVE PAIR, and the one most easily got backwards. Both are
    ' GROUND queries, so both are booleans and neither can pass by
    ' spilling something. The condition succeeds and the then-goal FAILS:
    ' the whole form fails, and the else-goal - which would have
    ' succeeded - must NOT run. The then-goal is (q 1), which matches no
    ' fact; before PROLOG.22 it was an undefined (nosuch 1), in both.
    result = VLA_Prolog.PROLOG("(fact (p 1)) (fact (q a)) (query (if (p 1) (q 1) (q a)))")
    Report "prolog.14: condition succeeds, then-goal fails - the whole (if ...) FAILS and the else-goal never runs", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)

    result = VLA_Prolog.PROLOG("(fact (p 1)) (fact (q a)) (query (if (p 9) (q 1) (q a)))")
    Report "prolog.14: ...and its twin - when the condition never succeeds, the SAME else-goal does run and the form is TRUE", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)

    ' The two-argument form. No else-goal, so a failed condition simply
    ' fails; asserted FALSE beside its own TRUE twin so neither can be
    ' passing for want of an implementation.
    result = VLA_Prolog.PROLOG("(fact (p 1)) (query (if (p 9) (p 1)))")
    Report "prolog.14: (if C T) with no else - a failed condition just fails", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)

    result = VLA_Prolog.PROLOG("(fact (p 1)) (query (if (p 1) (p 1)))")
    Report "prolog.14: ...and succeeds when the condition does", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)

    ' The condition's own bindings survive into the then-goal - which is
    ' what makes the form worth having - and the column rule is visible
    ' in the same query: W and M are bound only on the then-path, T on
    ' both, so T alone is a column. Committing to link(a b) is also why
    ' there is ONE row: without the commit, link(c d) would give a second.
    ' The else-goal must still mention T and nothing else, since it is the
    ' second path the column rule intersects; PROLOG.22 made it (tag none
    ' T), which matches no fact, where it was an undefined (nope T).
    result = VLA_Prolog.PROLOG( _
        "(fact (link a b)) (fact (link c d)) (fact (tag b yes)) (fact (tag d no)) " & _
        "(query (if (link W M) (tag M T) (tag none T)))")
    Dim bindOk As Boolean
    If IsArray(result) Then
        bindOk = (ResultRowCount(result) = 2 And ResultColCount(result) = 1 And ResultCellIs(result, 2, 1, "yes"))
    End If
    Report "prolog.14: the condition's bindings reach the then-goal, and only the variable BOTH paths bind is a column", _
           bindOk, "got: " & ResultDescribe(result)

    ' The else-less form has only one success path, so nothing is
    ' intersected away and the condition's own variable IS a column -
    ' two columns here where the three-argument form above gave one.
    result = VLA_Prolog.PROLOG("(fact (p 1)) (fact (q a)) (query (if (p X) (q Y)))")
    Dim twoColOk As Boolean
    If IsArray(result) Then
        twoColOk = (ResultRowCount(result) = 2 And ResultColCount(result) = 2 _
                    And ResultCellIs(result, 2, 1, "1") And ResultCellIs(result, 2, 2, "a"))
    End If
    Report "prolog.14: (if C T) has one success path, so the condition's own variable is a column too", _
           twoColOk, "got: " & ResultDescribe(result)

    ' ---- THE CUT SIGNAL, which if-then-else is built on ---------------
    ' A PROLOG.5.4-era defect this item found and repaired: a cut inside
    ' a CALLED predicate used to overwrite the caller's own cut while it
    ' was still travelling outward, so the caller's ! silently pruned
    ' nothing. Real Prolog answers 1 alone here.
    result = VLA_Prolog.PROLOG("(rule (a X) (g X) !) (fact (a 9)) (rule (g 1) !) (fact (g 2)) (query (a X))")
    Dim nestedCutOk As Boolean
    If IsArray(result) Then nestedCutOk = (ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, "1"))
    Report "prolog.14: a cut inside a called predicate no longer swallows the caller's own cut - (a X) answers 1 alone, not 1 and 9", _
           nestedCutOk, "got: " & ResultDescribe(result)

    ' The control that isolates it: the SAME program with the callee's
    ' own cut removed, which was correct before this repair and after.
    ' The two must agree; before the repair they did not.
    result = VLA_Prolog.PROLOG("(rule (a X) (g X) !) (fact (a 9)) (fact (g 1)) (fact (g 2)) (query (a X))")
    Dim controlCutOk As Boolean
    If IsArray(result) Then controlCutOk = (ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, "1"))
    Report "prolog.14: ...and the control, the same program without the callee's cut, answers 1 as it always did", _
           controlCutOk, "got: " & ResultDescribe(result)

    ' The same signal, seen from the other side: a ! written to the RIGHT
    ' of an (if ...) in the same body must still prune that clause, so
    ' the second (m ...) fact is never reached. If the if-then-else
    ' absorbed the signal instead of letting it pass, this would be two
    ' rows.
    ' (q c), never reached, was an undefined (nope Y) before PROLOG.22.
    result = VLA_Prolog.PROLOG(facts & "(rule (m X Y) (if (p X) (q Y) (q c)) !) (fact (m 9 9)) (query (m X Y))")
    Dim rightCutOk As Boolean
    If IsArray(result) Then
        rightCutOk = (ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, "1") And ResultCellIs(result, 2, 2, "a"))
    End If
    Report "prolog.14: a ! to the RIGHT of an (if ...) still prunes its own clause - the form passes the signal on rather than absorbing it", _
           rightCutOk, "got: " & ResultDescribe(result)

    ' ---- A BRANCH IS A GOAL POSITION AT PARSE TIME TOO ----------------
    ' List sugar must NOT fire inside a branch: a (list ...) there is in
    ' goal position, where a list is not something that can be proved.
    ' Expanded as data it would become a cons chain, quietly become an
    ' unknown predicate, and fail in silence instead of saying so.
    r = CStr(VLA_Prolog.PROLOG("(fact (p 1)) (query (or (list a b) (p 1)))"))
    Report "prolog.14: a (list ...) inside a branch is still a GOAL - it reaches the 'not something PROLOG can prove' refusal, not the sugar", _
           InStr(1, r, "not something PROLOG can prove", vbTextCompare) > 0, "got: " & r

    ' ---- SHAPE REFUSALS ----------------------------------------------
    r = CStr(VLA_Prolog.PROLOG("(fact (p 1)) (query (or (p 1)))"))
    Report "prolog.14: (or ...) with a single branch is refused by name", _
           InStr(1, r, "at least two goals", vbTextCompare) > 0, "got: " & r

    r = CStr(VLA_Prolog.PROLOG("(fact (p 1)) (query (if (p 1)))"))
    Report "prolog.14: (if ...) with only a condition is refused by name", _
           InStr(1, r, "condition and a then-goal", vbTextCompare) > 0, "got: " & r

    r = CStr(VLA_Prolog.PROLOG("(fact (p 1)) (query (if (p 1) (p 1) (p 1) (p 1)))"))
    Report "prolog.14: (if ...) with four arguments is refused - there is no else-if chain", _
           InStr(1, r, "condition and a then-goal", vbTextCompare) > 0, "got: " & r

    ' ---- THE ISO SPELLINGS -------------------------------------------
    ' `->` and `\+` are what a Prolog author types first. Below the
    ' clause lookup an unknown predicate is a SILENT dead end, so both
    ' are reserved and dispatched to a refusal that names the house form.
    r = CStr(VLA_Prolog.PROLOG("(fact (p 1)) (query (-> (p 1) (p 1) (p 1)))"))
    Report "prolog.14: ISO's (-> ...) is refused with the spelling this engine uses", _
           InStr(1, r, "(if ...)", vbTextCompare) > 0, "got: " & r

    ' The `;` guidance rides on this refusal, because `;` can never be
    ' reserved - VLA.Tokenize eats it as a comment before any reader sees
    ' it - so this is the only place an author reaching for ISO
    ' if-then-else can be told what it does here.
    Report "prolog.14: ...and the same refusal says what a ';' does here, since ';' can never be a reserved name", _
           InStr(1, r, "comment", vbTextCompare) > 0, "got: " & r

    r = CStr(VLA_Prolog.PROLOG("(fact (p 1)) (query (\+ (p 1)))"))
    Report "prolog.14: ISO's (\+ ...) is refused and points at (not ...) - it would otherwise FAIL silently, which reads as a successful negation", _
           InStr(1, r, "(not ...)", vbTextCompare) > 0, "got: " & r

    ' ---- RESERVED AS PREDICATE NAMES ---------------------------------
    ' A name the solver acts on but the parser does not reserve is a
    ' predicate a user can define and have silently shadowed.
    r = CStr(VLA_Prolog.PROLOG("(fact (or a b)) (query (p X))"))
    Report "prolog.14: 'or' is refused as a predicate name in a (fact ...)", _
           InStr(1, r, "reserved word", vbTextCompare) > 0, "got: " & r

    r = CStr(VLA_Prolog.PROLOG("(rule (if X) (p X)) (query (p X))"))
    Report "prolog.14: 'if' is refused as a predicate name in a (rule ...)", _
           InStr(1, r, "reserved word", vbTextCompare) > 0, "got: " & r

    r = CStr(VLA_Prolog.PROLOG("(fact (-> a b)) (query (p X))"))
    Report "prolog.14: '->' is refused as a predicate name, so it cannot be defined and then shadowed", _
           InStr(1, r, "reserved word", vbTextCompare) > 0, "got: " & r

    r = CStr(VLA_Prolog.PROLOG("(fact (\+ a)) (query (p X))"))
    Report "prolog.14: '\+' is refused as a predicate name on the same terms", _
           InStr(1, r, "reserved word", vbTextCompare) > 0, "got: " & r

    ' ---- KEYED TABLE ATOMS INSIDE A BRANCH ---------------------------
    ' Desugaring must recurse into a branch exactly as it does into
    ' `not`'s goal: un-desugared, a keyed atom simply would not match,
    ' which is a silent wrong answer rather than a refusal.
    Dim q As String
    q = Chr$(34)
    Dim headerMap As Object
    Set headerMap = VLA_Runtime.VlaDictNew()
    Dim cols As New Collection
    Dim cName As New Collection: cName.Add "name": cName.Add "Name": cols.Add cName
    Dim cDept As New Collection: cDept.Add "dept": cDept.Add "Dept": cols.Add cDept
    VLA_Runtime.VlaDictSet headerMap, "staffing", cols
    Dim clauseDict As Object
    Set clauseDict = VLA_Runtime.VlaDictNew()
    result = VLA_Prolog.PrologRun( _
        "(fact (staffing " & q & "alice" & q & " " & q & "eng" & q & ")) " & _
        "(fact (staffing " & q & "bob" & q & " " & q & "ops" & q & ")) " & _
        "(query (or (staffing (dept " & q & "zzz" & q & ") (name Name)) (staffing (dept " & q & "ops" & q & ") (name Name))))", _
        clauseDict, headerMap)
    Dim keyedOk As Boolean
    If IsArray(result) Then keyedOk = (ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, "bob"))
    Report "prolog.14: a keyed table atom inside a branch is desugared - and the FIRST branch matches nothing, so the row comes from the second", _
           keyedOk, "got: " & ResultDescribe(result)
End Sub

' PROLOG.18: the text goals, part one - the one spelling of a number,
' the alias class rule E re-derived, atom-length, and atom-concat with
' the marking decision it carries. Pure: no live workbook.
'
' THE DISCRIMINATION RULE, stated because this family is exposed to it:
' an unknown predicate fails SILENTLY, so a test asserting that a text
' goal fails would pass against no implementation at all. Every failure
' asserted here is a ground query answering FALSE beside its own TRUE
' twin, a strict non-empty subset, or an assertion on refusal TEXT.
'
' THE SHAPE RULE: a query with a free variable always spills, a header
' row even with no solutions; only a ground query collapses to a Boolean.
'
' Refusals are asserted through ResultTextStartsWith and ResultDescribe,
' never `CStr(VLA_Prolog.PROLOG(...))`. That idiom is safe only while the
' query really refuses: a regression that made one of these SPILL would
' turn a failing assertion into a type mismatch that kills the run,
' PROLOG.17's own incident. The phantom-column tests are exactly the ones
' whose failure mode is a spill.
Private Sub TestPrologText()
    Dim result As Variant
    Dim r As String
    Dim q As String
    q = Chr$(34)

    ' ---- ONE SPELLING FOR A NUMBER ------------------------------------
    ' sum-list's total goes through NumberToTerm, the Str$ path, which
    ' dropped a fraction's leading zero. Now it reads 0.5, as `is` does.
    result = VLA_Prolog.PROLOG("(query (sum-list (list 0.25 0.25) S))")
    Report "prolog.18: a fraction from the Str$ path reads 0.5, never .5", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, "0.5"), "got: " & ResultDescribe(result)

    ' THE WRONG ANSWER it caused, and the decisive pin: under ".5" this
    ' was FALSE - a confidently wrong "no" - because ".5" and the literal
    ' 0.5 are different atoms. Beside its twin, so it cannot pass for want
    ' of an implementation.
    result = VLA_Prolog.PROLOG("(query (sum-list (list 0.25 0.25) 0.5))")
    Report "prolog.18: (sum-list (list 0.25 0.25) 0.5) is TRUE - the total now matches the literal 0.5", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (sum-list (list 0.25 0.25) 0.75))")
    Report "prolog.18: ...and its twin against 0.75 is FALSE", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)

    ' The two paths agree term for term. S and T are free, so this spills;
    ' had the two spellings differed, == would fail and leave a header only.
    result = VLA_Prolog.PROLOG("(query (sum-list (list 0.25 0.25) S) (is T (+ 0.25 0.25)) (== S T))")
    Report "prolog.18: the sum-list path and the is path spell 0.25 + 0.25 identically, so (== S T) holds", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, "0.5") And ResultCellIs(result, 2, 2, "0.5"), _
           "got: " & ResultDescribe(result)

    result = VLA_Prolog.PROLOG("(query (sum-list (list -0.25 -0.25) S))")
    Report "prolog.18: a negative fraction reads -0.5, never -.5", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, "-0.5"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (is X (- 0 0.5)))")
    Report "prolog.18: ...and the is path, now routed through the same function, spells it identically", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, "-0.5"), "got: " & ResultDescribe(result)

    ' ---- THE ALIAS CLASS, re-derived --------------------------------
    ' Rule E of check_prolog_reserved_names.ps1 derives the near-miss
    ' spellings from the reserved set on every run; it found `whole`
    ' missing, a silent unknown predicate since PROLOG.17.
    result = VLA_Prolog.PROLOG("(query (whole 3))")
    r = ResultDescribe(result)
    Report "prolog.18: (whole 3) is refused and pointed at (whole? ...) - it used to fail silently", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "(whole? ...)", vbTextCompare) > 0, "got: " & r
    result = VLA_Prolog.PROLOG("(fact (whole 3)) (query (p X))")
    r = ResultDescribe(result)
    Report "prolog.18: ...and 'whole' is reserved, so it cannot be defined and then silently shadowed", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "reserved word", vbTextCompare) > 0, "got: " & r

    ' The ISO/SWI spellings a Prolog author types first, one per shape.
    result = VLA_Prolog.PROLOG("(query (atom_length abc N))")
    r = ResultDescribe(result)
    Report "prolog.18: ISO's atom_length is refused and pointed at (atom-length ...)", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "(atom-length ...)", vbTextCompare) > 0, "got: " & r
    result = VLA_Prolog.PROLOG("(query (atomic_list_concat L x axb))")
    r = ResultDescribe(result)
    Report "prolog.18: SWI's atomic_list_concat is refused and pointed at (atomic-list-concat ...)", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "(atomic-list-concat ...)", vbTextCompare) > 0, "got: " & r
    result = VLA_Prolog.PROLOG("(fact (atom-length a b)) (query (p X))")
    r = ResultDescribe(result)
    Report "prolog.18: 'atom-length' is refused as a predicate name, so it cannot be defined and then shadowed", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "reserved word", vbTextCompare) > 0, "got: " & r

    ' ---- ATOM-LENGTH -------------------------------------------------
    result = VLA_Prolog.PROLOG("(query (atom-length " & q & "hello" & q & " N))")
    Report "prolog.18: (atom-length " & q & "hello" & q & " N) counts 5", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, "5"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (atom-length hello 5))")
    Report "prolog.18: (atom-length hello 5) is TRUE - a bare name is text too", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (atom-length hello 4))")
    Report "prolog.18: ...and its twin (atom-length hello 4) is FALSE", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)
    ' A number's text is its CANONICAL rendering, so 0.50 is three
    ' characters, as ISO counts it - the number, not the way it was typed.
    result = VLA_Prolog.PROLOG("(query (atom-length 0.50 N))")
    Report "prolog.18: (atom-length 0.50 N) is 3 - a number is measured as it prints, 0.5", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, "3"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (atom-length " & q & q & " N))")
    Report "prolog.18: empty text has length 0", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, "0"), "got: " & ResultDescribe(result)

    ' THE PHANTOM COLUMN. Text goals bind, so CollectVars descends into
    ' them and X is collected; what stops a column headed X holding "X" is
    ' that the goal refuses by name when its input is free. Its failure
    ' mode is a SPILL, which is why this is not a CStr(...) assertion.
    result = VLA_Prolog.PROLOG("(query (atom-length X N))")
    r = ResultDescribe(result)
    Report "prolog.18: THE PHANTOM COLUMN - (atom-length X N) with X free refuses by name rather than spilling X", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "can't run yet", vbTextCompare) > 0, "got: " & r
    result = VLA_Prolog.PROLOG("(query (atom-length (f a) N))")
    r = ResultDescribe(result)
    Report "prolog.18: a compound term is not text, and is refused by name", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "compound term", vbTextCompare) > 0, "got: " & r
    result = VLA_Prolog.PROLOG("(query (atom-length abc " & q & "3" & q & "))")
    r = ResultDescribe(result)
    Report "prolog.18: a count given as TEXT is refused - a text 3 is not the number 3 (between's own rule)", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "isn't one", vbTextCompare) > 0, "got: " & r

    ' A character outside the Basic Multilingual Plane is two UTF-16 units:
    ' Excel's LEN says 2, Prolog says 1. The goals that count refuse...
    Dim emoji As String
    emoji = ChrW$(55357) & ChrW$(56832)            ' U+1F600, stored as two units
    result = VLA_Prolog.PROLOG("(query (atom-length " & q & "a" & emoji & q & " N))")
    r = ResultDescribe(result)
    Report "prolog.18: atom-length refuses text holding an emoji rather than choosing between Excel's count and Prolog's", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "stores as two", vbTextCompare) > 0, "got: " & r
    ' ...and a goal that does not count passes it through untouched.
    result = VLA_Prolog.PROLOG("(query (upcase-atom " & q & "a" & emoji & q & " U))")
    Report "prolog.18: ...while upcase-atom, which counts nothing, passes the same emoji through", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, "A" & emoji), "got: " & ResultDescribe(result)

    ' ---- ATOM-CONCAT, and THE MARKING DECISION --------------------------
    result = VLA_Prolog.PROLOG("(query (atom-concat ab c X))")
    Report "prolog.18: (atom-concat ab c X) joins to abc", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, "abc"), "got: " & ResultDescribe(result)

    ' The result is TEXT - marked, exactly as a text cell is. This FAILS
    ' under "mirror the inputs" and under "mark only when a bare atom would
    ' be misread", both of which would hand back the bare name abc - and
    ' (== abc "abc") would then REFUSE, so no row could spill.
    '
    ' A SPILL, not a Boolean: X is free and atom-concat binds it, so X is an
    ' output column - the result-shape rule this Sub's own header states.
    ' The first version asserted ResultBoolIs here and failed on the live
    ' run against a correct engine, which spilled exactly X = abc.
    result = VLA_Prolog.PROLOG("(query (atom-concat ab c X) (== X " & q & "abc" & q & "))")
    Report "prolog.18: MARKING - the result of a text goal is TEXT, identical to the quoted " & q & "abc" & q, _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, "abc"), "got: " & ResultDescribe(result)
    ' ...so comparing it with the bare NAME abc is PROLOG.10's confusable
    ' case, and refuses rather than answering a silent FALSE.
    result = VLA_Prolog.PROLOG("(query (atom-concat ab c X) (== X abc))")
    r = ResultDescribe(result)
    Report "prolog.18: ...and comparing it with the bare name abc refuses through PROLOG.10's own message", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "different things", vbTextCompare) > 0, "got: " & r

    ' A BOUND argument is read as its TEXT, so the relation holds with all
    ' three bound. This FAILS under "compare a bound result by unification".
    result = VLA_Prolog.PROLOG("(query (atom-concat ab c abc))")
    Report "prolog.18: TEST MODE - (atom-concat ab c abc) is TRUE: every bound argument is read as text", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (atom-concat ab c abd))")
    Report "prolog.18: ...and its twin (atom-concat ab c abd) is FALSE", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)

    result = VLA_Prolog.PROLOG("(query (atom-concat " & q & "Item " & q & " 42 X))")
    Report "prolog.18: a number joins as its text - Item 42", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, "Item 42"), "got: " & ResultDescribe(result)
    ' Joining two numbers makes TEXT, ISO's answer - under "mirror the
    ' inputs" it would be the NUMBER 42.
    result = VLA_Prolog.PROLOG("(query (atom-concat 4 2 X) (atom? X))")
    Report "prolog.18: MARKING - (atom-concat 4 2 X) is text: atom? is TRUE", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, "42"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (atom-concat 4 2 X) (number? X))")
    Report "prolog.18: ...and number? is FALSE - header only", _
           ResultRowCount(result) = 1, "got: " & ResultDescribe(result)
    ' THE VARIABLE TRAP: a bare BOB would be read as an unbound VARIABLE.
    result = VLA_Prolog.PROLOG("(query (upcase-atom bob U) (atom? U))")
    Report "prolog.18: MARKING - (upcase-atom bob U) gives TEXT, never a bare BOB that would read as a variable", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, "BOB"), "got: " & ResultDescribe(result)

    ' Taking a text apart by a known prefix, a known suffix, or every split.
    result = VLA_Prolog.PROLOG("(query (atom-concat " & q & "ID-" & q & " Rest " & q & "ID-42" & q & "))")
    Report "prolog.18: a known PREFIX - (atom-concat " & q & "ID-" & q & " Rest " & q & "ID-42" & q & ") gives Rest = 42", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, "42"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (atom-concat " & q & "XX-" & q & " Rest " & q & "ID-42" & q & "))")
    Report "prolog.18: ...a prefix the text does not have finds nothing - header only", _
           ResultRowCount(result) = 1, "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (atom-concat Stem " & q & ".xlsx" & q & " " & q & "report.xlsx" & q & "))")
    Report "prolog.18: a known SUFFIX - Stem = report", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, "report"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (atom-concat A B " & q & "abc" & q & "))")
    Report "prolog.18: every split of abc, both ends included, in order - four rows", _
           ResultRowCount(result) = 5 And ResultCellIs(result, 2, 1, "") And ResultCellIs(result, 2, 2, "abc") _
           And ResultCellIs(result, 5, 1, "abc") And ResultCellIs(result, 5, 2, ""), "got: " & ResultDescribe(result)

    ' THE BUDGET, at the exact edge - between's rule and its off-by-one.
    ' 118 characters have 119 splits: 1 unit for the goal + 119 = 120, and
    ' they fit. One character more needed 120 and was refused BY NAME up
    ' front, never at the ceiling a split later.
    ' PROLOG.28 moved that edge with the work budget: a text of 99,999
    ' characters has 100,000 splits, one past what a query may try, and is
    ' the refusing side now. 118 keeps its assertion because it still fits,
    ' and TestPrologBudgets pins 119 - the OLD edge - answering.
    ' This is the assertion the item's own exposure sweep missed: it swept
    ' between, length and sub-atom and forgot that atom-concat's split has
    ' an up-front bound of its own. The live run found it (owner, 2026-09-11).
    result = VLA_Prolog.PROLOG("(query (atom-concat A B " & q & String$(118, "a") & q & "))")
    Report "prolog.18: BOUNDARY - 118 characters split every way inside the budget (119 rows)", _
           ResultRowCount(result) = 120, "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (atom-concat A B " & q & String$(99999, "a") & q & "))")
    r = ResultDescribe(result)
    Report "prolog.18: BOUNDARY - 99,999 characters are refused by name up front, not at the work ceiling", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "ways", vbTextCompare) > 0, "got: " & r
    ' A known prefix on a LONG text is answered directly, never by trying
    ' every split - the deterministic modes are there for exactly this.
    result = VLA_Prolog.PROLOG("(query (atom-concat " & q & "ID-" & q & " Rest " & q & "ID-" & String$(200, "a") & q & "))")
    Report "prolog.18: a known prefix on a 203-character text is answered directly, not refused", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, String$(200, "a")), "got: " & ResultDescribe(result)

    result = VLA_Prolog.PROLOG("(query (atom-concat A B C))")
    r = ResultDescribe(result)
    Report "prolog.18: THE PHANTOM COLUMN - atom-concat with nothing bound refuses by name", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "can't run yet", vbTextCompare) > 0, "got: " & r

    ' A generator obeys the cut: stop on it, never absorb it.
    result = VLA_Prolog.PROLOG("(rule (firstsplit A) (atom-concat A B " & q & "abc" & q & ") !) (query (firstsplit A))")
    Report "prolog.18: ! after atom-concat's generating mode prunes it to the first split", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, ""), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(rule (firstsplit A) (atom-concat A B " & q & "abc" & q & ")) (query (firstsplit A))")
    Report "prolog.18: ...and without the cut, all four", _
           ResultRowCount(result) = 5, "got: " & ResultDescribe(result)
End Sub

' PROLOG.18: the text goals, part two - sub-atom, atom-number, case, and
' atomic-list-concat, plus the list goals that now name their text twin.
' The same three rules as part one's header: discriminating failures,
' spill-or-Boolean shape, and refusals asserted without CStr(...).
Private Sub TestPrologTextParts()
    Dim result As Variant
    Dim r As String
    Dim q As String
    q = Chr$(34)

    ' ---- SUB-ATOM ----------------------------------------------------
    result = VLA_Prolog.PROLOG("(query (sub-atom " & q & "Frazaro" & q & " 2 3 After Part))")
    Report "prolog.18: (sub-atom " & q & "Frazaro" & q & " 2 3 After Part) - After 2, Part aza", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, "2") And ResultCellIs(result, 2, 2, "aza"), _
           "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (sub-atom " & q & "a,b,c" & q & " B 1 A " & q & "," & q & "))")
    Report "prolog.18: a bound Part finds every occurrence - the commas at 1 and 3", _
           ResultRowCount(result) = 3 And ResultCellIs(result, 2, 1, "1") And ResultCellIs(result, 3, 1, "3"), _
           "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (sub-atom " & q & "banana" & q & " B L A " & q & "ana" & q & "))")
    Report "prolog.18: ...overlapping ones included - ana in banana at 1 and 3", _
           ResultRowCount(result) = 3 And ResultCellIs(result, 2, 1, "1") And ResultCellIs(result, 3, 1, "3"), _
           "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (sub-atom hello 0 1 4 h))")
    Report "prolog.18: (sub-atom hello 0 1 4 h) is TRUE", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (sub-atom hello 0 1 4 e))")
    Report "prolog.18: ...and its twin with e is FALSE", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (sub-atom ab B L A S))")
    Report "prolog.18: with only the text bound, every piece of ab - six, Before then Length ascending", _
           ResultRowCount(result) = 7 And ResultCellIs(result, 2, 1, "0") And ResultCellIs(result, 2, 2, "0") _
           And ResultCellIs(result, 4, 2, "2") And ResultCellIs(result, 4, 4, "ab") And ResultCellIs(result, 7, 1, "2"), _
           "got: " & ResultDescribe(result)

    ' THE BUDGET. All-free is (n+1)(n+2)/2 pieces. The edge used to be 13
    ' characters (105 fit) and 14 (120 refused); PROLOG.28's work budget
    ' moved it to 445 (99,681 fit) and 446 (100,128 refused by name before
    ' one is made). The refusing side is re-pointed here; the fitting side
    ' at 445 would spill 99,681 rows, so thirteen stays as the small case
    ' and TestPrologBudgets pins fourteen, the old refused edge, answering.
    result = VLA_Prolog.PROLOG("(query (sub-atom " & q & String$(13, "a") & q & " B L A S))")
    Report "prolog.18: 13 characters, all free: all 105 pieces", _
           ResultRowCount(result) = 106, "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (sub-atom " & q & String$(446, "a") & q & " B L A S))")
    r = ResultDescribe(result)
    Report "prolog.18: BOUNDARY - 446 characters, all free: 100,128 pieces are refused up front, by name", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "ways", vbTextCompare) > 0, "got: " & r
    ' A bound Before narrows the occurrences BEFORE they are counted, so a
    ' long text with many matches still answers.
    result = VLA_Prolog.PROLOG("(query (sub-atom " & q & String$(200, "a") & q & " 5 L A " & q & "a" & q & "))")
    Report "prolog.18: a bound position narrows a long text's matches before the count - one row, not a refusal", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, "1") And ResultCellIs(result, 2, 2, "194"), _
           "got: " & ResultDescribe(result)

    ' Positions are 0..n: a value outside that set FAILS (nth's rule), a
    ' non-number is refused.
    result = VLA_Prolog.PROLOG("(query (sub-atom hello 1.5 1 A S))")
    Report "prolog.18: a fractional position is no position - header only", _
           ResultRowCount(result) = 1, "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (sub-atom hello 1 1 A S))")
    Report "prolog.18: ...its twin at position 1 finds e", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 2, "e"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (sub-atom hello 99999999999 L A S))")
    Report "prolog.18: a huge position fails cleanly - it never reaches CLng to overflow", _
           ResultRowCount(result) = 1, "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (sub-atom hello " & q & "1" & q & " 1 A S))")
    r = ResultDescribe(result)
    Report "prolog.18: a position given as TEXT is refused - it is not a position at all", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "isn't one", vbTextCompare) > 0, "got: " & r

    result = VLA_Prolog.PROLOG("(rule (firstchar C) (sub-atom hello B 1 A C) !) (query (firstchar C))")
    Report "prolog.18: ! after sub-atom prunes it to the first character", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, "h"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(rule (firstchar C) (sub-atom hello B 1 A C)) (query (firstchar C))")
    Report "prolog.18: ...and without the cut, all five", _
           ResultRowCount(result) = 6, "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (sub-atom X B L A S))")
    r = ResultDescribe(result)
    Report "prolog.18: THE PHANTOM COLUMN - sub-atom with its text free refuses by name", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "can't run yet", vbTextCompare) > 0, "got: " & r

    ' ---- ATOM-NUMBER - text and number cross on purpose ----------------
    result = VLA_Prolog.PROLOG("(query (atom-number " & q & "42" & q & " N) (number? N))")
    Report "prolog.18: (atom-number " & q & "42" & q & " N) gives the NUMBER 42 - number? holds", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, "42"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (atom-number " & q & "0.50" & q & " N))")
    Report "prolog.18: ...in canonical form - " & q & "0.50" & q & " gives 0.5", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, "0.5"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (atom-number " & q & "abc" & q & " N))")
    Report "prolog.18: text that spells no number FAILS - a correct negative answer, header only", _
           ResultRowCount(result) = 1, "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (atom-number A 7) (atom? A))")
    Report "prolog.18: run backwards, (atom-number A 7) gives TEXT - atom? holds", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, "7"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (atom-number " & q & "3.5" & q & " N) (is M (* N 2)))")
    Report "prolog.18: the number it gives does arithmetic - 3.5 doubled is 7", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 2, "7"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (atom-number A N))")
    r = ResultDescribe(result)
    Report "prolog.18: THE PHANTOM COLUMN - atom-number with both free refuses by name", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "can't run yet", vbTextCompare) > 0, "got: " & r
    result = VLA_Prolog.PROLOG("(query (atom-number A " & q & "7" & q & "))")
    r = ResultDescribe(result)
    Report "prolog.18: a number given as TEXT is refused - a text 7 is not the number 7", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "isn't one", vbTextCompare) > 0, "got: " & r

    ' ---- CASE, which obeys R6 and PROLOG.19 at once ---------------------
    ' Built with ChrW$ so this module stays ASCII: no codepage decides
    ' what these tests contain.
    result = VLA_Prolog.PROLOG("(query (upcase-atom " & q & "hello World" & q & " U))")
    Report "prolog.18: upcase-atom", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, "HELLO WORLD"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (downcase-atom " & q & ChrW$(193) & "RBOL " & ChrW$(209) & "AND" & ChrW$(218) & q & " L))")
    Report "prolog.18: downcase-atom moves Spanish accents and the n-tilde - not ASCII-only, unlike Fold", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, ChrW$(225) & "rbol " & ChrW$(241) & "and" & ChrW$(250)), _
           "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (downcase-atom " & q & ChrW$(321) & ChrW$(211) & "D" & ChrW$(377) & q & " L))")
    Report "prolog.18: ...and Latin Extended-A, a Polish place name", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, ChrW$(322) & ChrW$(243) & "d" & ChrW$(378)), _
           "got: " & ResultDescribe(result)
    ' Locale-free: I is i on every machine, including a Turkish one.
    result = VLA_Prolog.PROLOG("(query (downcase-atom " & q & "INDIGO" & q & " L))")
    Report "prolog.18: I downcases to the plain i, never the Turkish dotless one - the table follows no locale", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, "indigo"), "got: " & ResultDescribe(result)
    ' The contested characters refuse, in their contested direction only.
    result = VLA_Prolog.PROLOG("(query (upcase-atom " & q & ChrW$(305) & q & " U))")
    r = ResultDescribe(result)
    Report "prolog.18: the dotless i on UPCASE is refused - Unicode and the invariant table disagree, so any answer is a guess", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "U+0131", vbTextCompare) > 0, "got: " & r
    result = VLA_Prolog.PROLOG("(query (downcase-atom " & q & ChrW$(305) & q & " L))")
    Report "prolog.18: ...its uncontested direction is fine - it downcases to itself", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, ChrW$(305)), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (upcase-atom " & q & ChrW$(945) & ChrW$(946) & q & " U))")
    r = ResultDescribe(result)
    Report "prolog.18: a Greek letter is refused by name, never passed through unchanged", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "U+03B1", vbTextCompare) > 0, "got: " & r
    result = VLA_Prolog.PROLOG("(query (upcase-atom " & q & "5" & ChrW$(8364) & " abc" & q & " U))")
    Report "prolog.18: a character with no case - the euro sign - passes through", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, "5" & ChrW$(8364) & " ABC"), "got: " & ResultDescribe(result)
    ' Above U+7FFF AscW comes back negative; a check written against the
    ' unsigned value would let this full-width letter through.
    result = VLA_Prolog.PROLOG("(query (upcase-atom " & q & "x" & ChrW$(65345) & q & " U))")
    r = ResultDescribe(result)
    Report "prolog.18: a full-width letter (above U+7FFF, negative from AscW) is still refused", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "U+FF41", vbTextCompare) > 0, "got: " & r
    result = VLA_Prolog.PROLOG("(query (downcase-atom " & q & "ENG" & q & " eng))")
    Report "prolog.18: (downcase-atom " & q & "ENG" & q & " eng) is TRUE - the bound result is read as text", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (downcase-atom " & q & "ENG" & q & " sales))")
    Report "prolog.18: ...and against sales is FALSE", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (upcase-atom X " & q & "ABC" & q & "))")
    r = ResultDescribe(result)
    Report "prolog.18: THE PHANTOM COLUMN - upcase-atom with its input free refuses by name", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "can't run yet", vbTextCompare) > 0, "got: " & r

    ' ---- ATOMIC-LIST-CONCAT - join, and split -------------------------
    result = VLA_Prolog.PROLOG("(query (atomic-list-concat P " & q & "," & q & " " & q & "a,b,c" & q & "))")
    Report "prolog.18: SPLIT a,b,c on the comma - a list of three pieces of text", _
           ResultCol1Is(result, "(list " & q & "a" & q & " " & q & "b" & q & " " & q & "c" & q & ")"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (atomic-list-concat P " & q & "," & q & " " & q & "a,b,c" & q & ") (length P N))")
    Report "prolog.18: ...an ordinary list, so length measures it - 3", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 2, "3"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (atomic-list-concat P " & q & ", " & q & " " & q & "red, green" & q & ") (member X P))")
    Report "prolog.18: ...a two-character separator, and member walking the pieces", _
           ResultRowCount(result) = 3 And ResultCellIs(result, 2, 2, "red") And ResultCellIs(result, 3, 2, "green"), _
           "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (atomic-list-concat (list a b c) " & q & "-" & q & " X))")
    Report "prolog.18: JOIN (list a b c) with a hyphen - a-b-c", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, "a-b-c"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (atomic-list-concat (list 1 0.50 x) " & q & "/" & q & " X))")
    Report "prolog.18: ...numbers join by their canonical text - 1/0.5/x", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, "1/0.5/x"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (atomic-list-concat (list First Second) " & q & " " & q & " " & q & "Ada Lovelace" & q & "))")
    Report "prolog.18: a list of free variables destructures a name - First Ada, Second Lovelace", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, "Ada") And ResultCellIs(result, 2, 2, "Lovelace"), _
           "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (atomic-list-concat (list A B) " & q & "," & q & " " & q & "a,b,c" & q & "))")
    Report "prolog.18: ...two variables against three pieces finds nothing - header only", _
           ResultRowCount(result) = 1, "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (atomic-list-concat (list a Y) " & q & "," & q & " " & q & "a,b" & q & "))")
    Report "prolog.18: ...a KNOWN bare element is matched by its text, so (list a Y) takes a,b apart - Y is b", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, "b"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (atomic-list-concat P " & q & "," & q & " " & q & q & ") (length P N))")
    Report "prolog.18: splitting empty text gives ONE empty piece, SWI's answer", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 2, "1"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (atomic-list-concat P " & q & q & " " & q & "ab" & q & "))")
    r = ResultDescribe(result)
    Report "prolog.18: SPLITTING on an empty separator is refused - there is no one way to cut at every nothing", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "empty separator", vbTextCompare) > 0, "got: " & r
    result = VLA_Prolog.PROLOG("(query (atomic-list-concat (list a b) " & q & q & " X))")
    Report "prolog.18: ...while JOINING with one is fine - ab", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, "ab"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (atomic-list-concat (cons a T) " & q & "," & q & " " & q & "a,b" & q & "))")
    r = ResultDescribe(result)
    Report "prolog.18: a PARTIAL list is refused by name, PROLOG.13's rule", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "needs a list", vbTextCompare) > 0, "got: " & r
    result = VLA_Prolog.PROLOG("(query (atomic-list-concat L S W))")
    r = ResultDescribe(result)
    Report "prolog.18: THE PHANTOM COLUMN - atomic-list-concat with nothing bound refuses by name", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "can't run yet", vbTextCompare) > 0, "got: " & r

    ' ---- THE LIST GOALS, HANDED TEXT ---------------------------------
    ' The first thing a user types. It is still refused - text is not a
    ' list - but the refusal now names the goal they wanted.
    result = VLA_Prolog.PROLOG("(query (length " & q & "hello" & q & " N))")
    r = ResultDescribe(result)
    Report "prolog.18: (length " & q & "hello" & q & " N) is refused, and the refusal names (atom-length ...)", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "(atom-length ...)", vbTextCompare) > 0 _
           And InStr(1, r, "needs a list", vbTextCompare) > 0, "got: " & r
    result = VLA_Prolog.PROLOG("(query (append " & q & "ab" & q & " " & q & "c" & q & " X))")
    r = ResultDescribe(result)
    Report "prolog.18: (append " & q & "ab" & q & " " & q & "c" & q & " X) names (atom-concat ...)", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "(atom-concat ...)", vbTextCompare) > 0, "got: " & r
    ' The twin: a NUMBER is not text, so it keeps the plain list refusal.
    result = VLA_Prolog.PROLOG("(query (length 42 N))")
    r = ResultDescribe(result)
    Report "prolog.18: ...but (length 42 N) - a number, not text - keeps the plain refusal and names no text goal", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "needs a list", vbTextCompare) > 0 _
           And InStr(1, r, "atom-length", vbTextCompare) = 0, "got: " & r
End Sub

' PROLOG.30: DATALOG.11's three text tests, the same names and meaning here,
' so a rule cell written with one reads alike in both engines. Each is a TEST
' - it binds nothing and answers once - where sub-atom GENERATES: measured by
' G-PROLOG slice 5's scoping, sub-atom used as "contains" listed Banana twice.
Private Sub TestPrologTextTests()
    Dim result As Variant
    Dim r As String
    Dim q As String
    q = Chr$(34)
    Dim fruit As String
    fruit = "(fact (p " & q & "Banana" & q & ")) (fact (p " & q & "Cyan" & q & ")) (fact (p " & q & "Apple" & q & ")) "
    Dim accts As String
    accts = "(fact (acct " & q & "GL-4010" & q & ")) (fact (acct " & q & "GL-5010" & q & ")) "

    result = VLA_Prolog.PROLOG("(query (text-starts-with " & q & "GL-4010" & q & " " & q & "GL-4" & q & "))")
    Report "prolog.30: text-starts-with GL-4010 GL-4 is TRUE", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (text-starts-with " & q & "GL-4010" & q & " " & q & "gl-4" & q & "))")
    Report "prolog.30: ...and against gl-4 is FALSE - case is exact", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (text-ends-with " & q & "GL-4010" & q & " " & q & "10" & q & "))")
    Report "prolog.30: text-ends-with GL-4010 10 is TRUE", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (text-ends-with abc b))")
    Report "prolog.30: ...and abc does not end with b - FALSE, no refusal", _
           ResultBoolIs(result, False), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG(fruit & "(rule (has-an X) (p X) (text-contains X " & q & "an" & q & ")) (query (has-an X))")
    Report "prolog.30: text-contains answers once per value however often the part occurs - Banana, then Cyan", _
           ResultRowCount(result) = 3 And ResultCellIs(result, 2, 1, "Banana") And ResultCellIs(result, 3, 1, "Cyan"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG(fruit & "(rule (has-an X) (p X) (sub-atom X B L A " & q & "an" & q & ")) (query (has-an X))")
    Report "prolog.30: ...where sub-atom finds every occurrence, and lists Banana twice", _
           ResultRowCount(result) = 4, "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (text-starts-with 4010 40))")
    Report "prolog.30: a number is read as its text - 4010 starts with 40", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (text-starts-with 0.50 " & q & "0.5" & q & "))")
    Report "prolog.30: ...its canonical text - 0.50 reads 0.5", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (text-contains abc " & q & q & "))")
    Report "prolog.30: every text contains the empty text", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG(accts & "(rule (other C) (acct C) (not (text-starts-with C " & q & "GL-4" & q & "))) (query (other C))")
    Report "prolog.30: under not it answers inverted - GL-5010 alone", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, "GL-5010"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(fact (pair " & q & "GL-4010" & q & " " & q & "GL-4" & q & ")) (fact (pair " & q & "GL-5010" & q & " " & q & "GL-4" & q & ")) (rule (ok C) (pair C P) (text-starts-with C P)) (query (ok C))")
    Report "prolog.30: both arguments may be variables bound earlier", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, "GL-4010"), "got: " & ResultDescribe(result)

    result = VLA_Prolog.PROLOG("(query (text-contains X " & q & "a" & q & "))")
    r = ResultDescribe(result)
    Report "prolog.30: THE PHANTOM COLUMN - a free text is refused by name, never spilled", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "can't run yet", vbTextCompare) > 0, "got: " & r
    result = VLA_Prolog.PROLOG("(query (text-contains (f a) a))")
    r = ResultDescribe(result)
    Report "prolog.30: a compound term is not text", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "works on text", vbTextCompare) > 0, "got: " & r
    result = VLA_Prolog.PROLOG("(query (text-contains abc))")
    r = ResultDescribe(result)
    Report "prolog.30: one argument is refused, listing the text tests' shapes", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "(text-contains Text Part)", vbTextCompare) > 0, "got: " & r
    result = VLA_Prolog.PROLOG("(fact (text-contains a b)) (query (text-contains a b))")
    r = ResultDescribe(result)
    Report "prolog.30: text-contains is reserved - no fact may define it", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "is a reserved word in PROLOG", vbTextCompare) > 0, "got: " & r
    result = VLA_Prolog.PROLOG("(query (text_starts_with ab a))")
    r = ResultDescribe(result)
    Report "prolog.30: the underscore spelling points at text-starts-with", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "write (text-starts-with ...) instead", vbTextCompare) > 0, "got: " & r
End Sub

' PROLOG.19: the IMPURE goals, refused for good - assert, write, random,
' read and the rest of ImpureGoalKindFor - and the arithmetic functions
' random, random_float and cputime. Pure: no live workbook.
'
' THE DISCRIMINATION RULE bites hardest here, because every name below
' used to be an UNKNOWN PREDICATE, and an unknown predicate fails
' silently: a test asserting only that `(query (write hello))` does not
' succeed would pass against no implementation at all. Every assertion is
' on REFUSAL TEXT - the family's own phrase and the form written - or on
' the result a working query produces.
'
' THE SHAPE RULE: `(query (write X))` has a free variable, so an
' unrefused one SPILLS - a header headed X with nothing under it. That is
' what the loops below would receive if a name fell out of the table, so
' each goal assertion is also that name's phantom-column pin. Refusals go
' through ResultTextStartsWith and ResultDescribe, which answer False on
' an array rather than raising - never CStr(VLA_Prolog.PROLOG(...)),
' PROLOG.17's incident.
'
' The five name lists are the table, family by family, hyphen twins
' included; the transliteration that checked this Sub before import held
' their union equal to ImpureGoalKindFor's own Case arms.
Private Sub TestPrologImpure()
    Dim result As Variant
    Dim r As String
    Dim nm As Variant
    Dim q As String
    q = Chr$(34)

    ' ---- EVERY NAME: refused as a goal with its family's reason and the
    ' form written, and reserved, so it cannot be defined and then refused
    ' on every call to the definition.
    For Each nm In Array("assert", "asserta", "assertz", "retract", "retractall", "abolish")
        result = VLA_Prolog.PROLOG("(query (" & nm & " X))")
        r = ResultDescribe(result)
        Report "prolog.19: (" & nm & " X) is refused for good as a DATABASE change - never failed silently, never spilled", _
               ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "never changes its own facts and rules", vbTextCompare) > 0 _
               And InStr(1, r, "(" & nm & " ...)", vbTextCompare) > 0, "got: " & r
        result = VLA_Prolog.PROLOG("(fact (" & nm & " a)) (query (p X))")
        r = ResultDescribe(result)
        Report "prolog.19: ...and '" & nm & "' is reserved, so it cannot be defined", _
               ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "reserved word", vbTextCompare) > 0 _
               And InStr(1, r, "'" & nm & "'", vbTextCompare) > 0, "got: " & r
    Next nm
    For Each nm In Array("write", "writeln", "print", "nl", "format", "writeq", "write_canonical", "write-canonical", "write_term", "write-term")
        result = VLA_Prolog.PROLOG("(query (" & nm & " X))")
        r = ResultDescribe(result)
        Report "prolog.19: (" & nm & " X) is refused for good as OUTPUT - never failed silently, never spilled", _
               ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "nowhere to print", vbTextCompare) > 0 _
               And InStr(1, r, "(" & nm & " ...)", vbTextCompare) > 0, "got: " & r
        result = VLA_Prolog.PROLOG("(fact (" & nm & " a)) (query (p X))")
        r = ResultDescribe(result)
        Report "prolog.19: ...and '" & nm & "' is reserved, so it cannot be defined", _
               ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "reserved word", vbTextCompare) > 0 _
               And InStr(1, r, "'" & nm & "'", vbTextCompare) > 0, "got: " & r
    Next nm
    For Each nm In Array("b_setval", "b-setval", "b_getval", "b-getval", "nb_setval", "nb-setval", "nb_getval", "nb-getval", "gensym")
        result = VLA_Prolog.PROLOG("(query (" & nm & " X))")
        r = ResultDescribe(result)
        Report "prolog.19: (" & nm & " X) is refused for good as kept STATE - never failed silently, never spilled", _
               ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "keeps no values between goals", vbTextCompare) > 0 _
               And InStr(1, r, "(" & nm & " ...)", vbTextCompare) > 0, "got: " & r
        result = VLA_Prolog.PROLOG("(fact (" & nm & " a)) (query (p X))")
        r = ResultDescribe(result)
        Report "prolog.19: ...and '" & nm & "' is reserved, so it cannot be defined", _
               ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "reserved word", vbTextCompare) > 0 _
               And InStr(1, r, "'" & nm & "'", vbTextCompare) > 0, "got: " & r
    Next nm
    For Each nm In Array("random", "random_between", "random-between", "random_member", "random-member", "random_permutation", "random-permutation", "get_time", "get-time")
        result = VLA_Prolog.PROLOG("(query (" & nm & " X))")
        r = ResultDescribe(result)
        Report "prolog.19: (" & nm & " X) is refused for good as VOLATILE - never failed silently, never spilled", _
               ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "every time Excel recalculates", vbTextCompare) > 0 _
               And InStr(1, r, "(" & nm & " ...)", vbTextCompare) > 0, "got: " & r
        result = VLA_Prolog.PROLOG("(fact (" & nm & " a)) (query (p X))")
        r = ResultDescribe(result)
        Report "prolog.19: ...and '" & nm & "' is reserved, so it cannot be defined", _
               ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "reserved word", vbTextCompare) > 0 _
               And InStr(1, r, "'" & nm & "'", vbTextCompare) > 0, "got: " & r
    Next nm
    For Each nm In Array("read", "read_term", "read-term", "consult", "halt")
        result = VLA_Prolog.PROLOG("(query (" & nm & " X))")
        r = ResultDescribe(result)
        Report "prolog.19: (" & nm & " X) is refused for good as reaching OUTSIDE - never failed silently, never spilled", _
               ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "reads nothing but", vbTextCompare) > 0 _
               And InStr(1, r, "(" & nm & " ...)", vbTextCompare) > 0, "got: " & r
        result = VLA_Prolog.PROLOG("(fact (" & nm & " a)) (query (p X))")
        r = ResultDescribe(result)
        Report "prolog.19: ...and '" & nm & "' is reserved, so it cannot be defined", _
               ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "reserved word", vbTextCompare) > 0 _
               And InStr(1, r, "'" & nm & "'", vbTextCompare) > 0, "got: " & r
    Next nm

    ' ---- THE LINE EVERYONE TYPES WHILE DEBUGGING. Before this item it
    ' spilled a header headed X with nothing under it - "no match" - for a
    ' p that plainly holds.
    result = VLA_Prolog.PROLOG("(fact (p 1)) (query (p X) (write X))")
    r = ResultDescribe(result)
    Report "prolog.19: (p X) then (write X) is refused, where it used to spill an empty column headed X", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "nowhere to print", vbTextCompare) > 0, "got: " & r

    ' ---- BARE nl AND halt, as Prolog writes them. A bare atom never
    ' reaches the solver, so ValidateBodyItem refuses these by name rather
    ' than calling them "not a predicate form" first.
    result = VLA_Prolog.PROLOG("(fact (p 1)) (query (p X) nl)")
    r = ResultDescribe(result)
    Report "prolog.19: a bare nl is refused as OUTPUT, shown as written", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "nl is refused for good", vbTextCompare) > 0 _
           And InStr(1, r, "nowhere to print", vbTextCompare) > 0, "got: " & r
    result = VLA_Prolog.PROLOG("(query halt)")
    r = ResultDescribe(result)
    Report "prolog.19: a bare halt is refused - there is no session to end", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "halt is refused for good", vbTextCompare) > 0 _
           And InStr(1, r, "no session", vbTextCompare) > 0, "got: " & r
    ' The twin: an ordinary bare atom keeps its own refusal.
    result = VLA_Prolog.PROLOG("(query foo)")
    r = ResultDescribe(result)
    Report "prolog.19: ...while an ordinary bare atom is still told it is not a predicate form", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "expected a predicate form", vbTextCompare) > 0, "got: " & r
    result = VLA_Prolog.PROLOG("(rule (show X) (p X) nl) (fact (p 1)) (query (show X))")
    r = ResultDescribe(result)
    Report "prolog.19: a bare nl in a rule body is refused too", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "nl is refused for good", vbTextCompare) > 0, "got: " & r

    ' ---- TWO ARITIES OF ONE IMPURE NAME are not blamed as an arity
    ' mismatch: the refusal is the one about the goal, which never runs at
    ' any arity. Beside the twin that shows arity checking still works.
    result = VLA_Prolog.PROLOG("(query (format " & q & "a" & q & ") (format " & q & "b" & q & " X))")
    r = ResultDescribe(result)
    Report "prolog.19: format/1 beside format/2 is refused as OUTPUT, not as an arity mismatch", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "nowhere to print", vbTextCompare) > 0 _
           And InStr(1, r, "argument(s)", vbTextCompare) = 0, "got: " & r
    result = VLA_Prolog.PROLOG("(fact (foo a)) (query (foo X) (foo X Y))")
    r = ResultDescribe(result)
    Report "prolog.19: ...while an ordinary predicate used at two arities is still refused as a mismatch", _
           ResultTextStartsWith(result, "#PROLOG!") _
           And InStr(1, r, "used with 1 argument(s) in one place and 2 in another", vbTextCompare) > 0, "got: " & r

    ' ---- REFUSED WHEN REACHED, like every refusing family: a rule whose
    ' impure goal runs is refused, one that is never called does not
    ' poison a query that works.
    result = VLA_Prolog.PROLOG("(rule (show X) (p X) (write X)) (fact (p 1)) (query (show X))")
    r = ResultDescribe(result)
    Report "prolog.19: a rule that reaches (write X) is refused", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "nowhere to print", vbTextCompare) > 0, "got: " & r
    result = VLA_Prolog.PROLOG("(rule (log X) (write X)) (fact (p 1)) (query (p X))")
    Report "prolog.19: ...while a rule holding (write X) that is never called leaves a working query alone", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, "1"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(fact (p 1)) (query (not (assert (p 2))))")
    r = ResultDescribe(result)
    Report "prolog.19: (assert ...) inside (not ...) is refused, not negated into a success", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "never changes its own facts", vbTextCompare) > 0, "got: " & r
    result = VLA_Prolog.PROLOG("(fact (p 1)) (query (or (p X) (write X)))")
    r = ResultDescribe(result)
    Report "prolog.19: a reachable (write X) branch refuses the query even after another branch answered", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "nowhere to print", vbTextCompare) > 0, "got: " & r

    ' ---- ARITHMETIC: the clock and the generator inside an expression.
    ' Before this item `random` was "not an operator PROLOG recognizes" -
    ' loud, but reading as "not yet". Each beside the twin that keeps its
    ' old refusal, so the route cannot be swallowing every unknown name.
    result = VLA_Prolog.PROLOG("(query (is X (random 10)))")
    r = ResultDescribe(result)
    Report "prolog.19: (random 10) in arithmetic is refused for good, pointing at Excel's RAND()", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "(random ...) is refused for good", vbTextCompare) > 0 _
           And InStr(1, r, "RAND()", vbTextCompare) > 0, "got: " & r
    result = VLA_Prolog.PROLOG("(query (is X random_float))")
    r = ResultDescribe(result)
    Report "prolog.19: a bare random_float in arithmetic is refused by name, not called a non-number", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "random_float is refused for good", vbTextCompare) > 0, "got: " & r
    result = VLA_Prolog.PROLOG("(query (is X (random_float)))")
    r = ResultDescribe(result)
    Report "prolog.19: ...and so is (random_float) written as a form", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "(random_float ...) is refused for good", vbTextCompare) > 0, "got: " & r
    result = VLA_Prolog.PROLOG("(query (is X cputime))")
    r = ResultDescribe(result)
    Report "prolog.19: a bare cputime in arithmetic is refused by name", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "cputime is refused for good", vbTextCompare) > 0, "got: " & r
    result = VLA_Prolog.PROLOG("(query (> (random 10) 5))")
    r = ResultDescribe(result)
    Report "prolog.19: (random 10) inside a comparison is refused the same way", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "(random ...) is refused for good", vbTextCompare) > 0, "got: " & r
    result = VLA_Prolog.PROLOG("(query (is X (foo 10)))")
    r = ResultDescribe(result)
    Report "prolog.19: ...while an unknown operator is still called unknown", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "'foo' isn't an arithmetic operator", vbTextCompare) > 0, "got: " & r
    result = VLA_Prolog.PROLOG("(query (is X (+ 1 2)))")
    Report "prolog.19: ...and a known operator still computes", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, "3"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (is X foo))")
    r = ResultDescribe(result)
    Report "prolog.19: ...and an ordinary bare name in arithmetic is still called a non-number, not refused for good", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "'foo'", vbTextCompare) > 0 _
           And InStr(1, r, "refused for good", vbTextCompare) = 0, "got: " & r

    ' ---- NAMES DELIBERATELY LEFT FREE, so the decision is visible: `tab`
    ' and `flag` are ordinary words in a workbook's own knowledge base, and
    ' `recorded` belongs to the recorded database, a family not reserved.
    result = VLA_Prolog.PROLOG("(fact (tab sheet1)) (query (tab X))")
    Report "prolog.19: `tab` is NOT reserved - a sheet tab is an ordinary thing to have facts about", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, "sheet1"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(fact (flag order17)) (query (flag X))")
    Report "prolog.19: `flag` is NOT reserved - a flagged order is an ordinary thing to have facts about", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, "order17"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(fact (recorded sale1)) (query (recorded X))")
    Report "prolog.19: `recorded` is NOT reserved - the recorded database is outside the reserved families", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, "sale1"), "got: " & ResultDescribe(result)

    ' ---- the reserved-word refusal enumerates the new family, twins too.
    result = VLA_Prolog.PROLOG("(fact (gensym a)) (query (p X))")
    r = ResultDescribe(result)
    Report "prolog.19: the reserved-word refusal lists the impure goals, underscore and hyphen spellings both", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "impure goals", vbTextCompare) > 0 _
           And InStr(1, r, "write_canonical", vbTextCompare) > 0 And InStr(1, r, "get-time", vbTextCompare) > 0 _
           And InStr(1, r, "retractall", vbTextCompare) > 0, "got: " & r

    ' ---- the head word is folded like every other goal's.
    result = VLA_Prolog.PROLOG("(query (WRITE hello))")
    r = ResultDescribe(result)
    Report "prolog.19: (WRITE hello) is refused as (write ...)", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "(write ...) is refused for good", vbTextCompare) > 0, "got: " & r

    ' ---- a hyphen twin lands on its family's refusal in ONE hop, not on
    ' the alias refusal pointing at a spelling that refuses too.
    result = VLA_Prolog.PROLOG("(query (get-time T))")
    r = ResultDescribe(result)
    Report "prolog.19: (get-time T) reaches the volatile refusal directly, not the alias one", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "every time Excel recalculates", vbTextCompare) > 0 _
           And InStr(1, r, "isn't how PROLOG spells", vbTextCompare) = 0, "got: " & r
End Sub

' PROLOG.24: a REFUSED name records no arity. Pure: no live workbook.
'
' Every name below is reserved in order to be REFUSED - a bare ISO
' spelling, a number-type name, an alias spelling, a control spelling,
' `list` in goal position, cut in parentheses - and each program writes it
' at TWO arities. Before this item ValidateBodyItem sent every one of them
' into TermPredName, which recorded the first arity and refused the second
' as prolog-arity-mismatch: blaming the author's arities for a goal that is
' refused on sight whatever its shape, and never showing the refusal that
' says why. Each assertion is on the name's OWN refusal text AND on the
' absence of the arity text, so it fails the moment the arity refusal
' comes back.
'
' THE LISTS ARE THE CLASS, family by family: the 28 reserved names that
' reached TermPredName, as tools/check_prolog_reserved_names.ps1's rule G
' derived them from ValidateBodyItem's own arms before the fix (the
' transliteration that checked this Sub before import held the union of
' the lists below equal to that derivation). The impure goals are not
' here: PROLOG.19 gave them this treatment first, and its format/1 beside
' format/2 pin in TestPrologImpure is now carried by the same arm.
'
' THE SHAPE RULE: every query below has free variables, so an unrefused
' one SPILLS. Refusals go through ResultTextStartsWith and ResultDescribe.
Private Sub TestPrologRefusedArity()
    Dim result As Variant
    Dim r As String
    Dim nm As Variant

    For Each nm In Array("var", "nonvar", "atom", "number", "atomic", "compound", "callable", "is_list", "ground")
        result = VLA_Prolog.PROLOG("(query (" & nm & " X) (" & nm & " X Y))")
        r = ResultDescribe(result)
        Report "prolog.24: (" & nm & " X) beside (" & nm & " X Y) gets its ISO-spelling refusal, not an arity mismatch", _
               ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "(" & nm & " ...) isn't how PROLOG spells this type test", vbTextCompare) > 0 _
               And InStr(1, r, "same arity", vbTextCompare) = 0, "got: " & r
    Next nm
    For Each nm In Array("integer?", "float?", "integer", "float")
        result = VLA_Prolog.PROLOG("(query (" & nm & " X) (" & nm & " X Y))")
        r = ResultDescribe(result)
        Report "prolog.24: (" & nm & " X) beside (" & nm & " X Y) gets its number-type refusal, not an arity mismatch", _
               ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "(" & nm & " ...) isn't available: PROLOG has one kind of number", vbTextCompare) > 0 _
               And InStr(1, r, "same arity", vbTextCompare) = 0, "got: " & r
    Next nm
    For Each nm In Array("is-list", "is_list?", "sum_list", "whole", "atom_length", "atom_concat", "sub_atom", "atom_number", "upcase_atom", "downcase_atom", "atomic_list_concat")
        result = VLA_Prolog.PROLOG("(query (" & nm & " X) (" & nm & " X Y))")
        r = ResultDescribe(result)
        Report "prolog.24: (" & nm & " X) beside (" & nm & " X Y) gets its alias-spelling refusal, not an arity mismatch", _
               ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "(" & nm & " ...) isn't how PROLOG spells this - write", vbTextCompare) > 0 _
               And InStr(1, r, "same arity", vbTextCompare) = 0, "got: " & r
    Next nm
    For Each nm In Array("->", "\+")
        result = VLA_Prolog.PROLOG("(query (" & nm & " X) (" & nm & " X Y))")
        r = ResultDescribe(result)
        Report "prolog.24: (" & nm & " X) beside (" & nm & " X Y) gets its control-spelling refusal, not an arity mismatch", _
               ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "(" & nm & " ...) is how real Prolog spells this", vbTextCompare) > 0 _
               And InStr(1, r, "same arity", vbTextCompare) = 0, "got: " & r
    Next nm
    result = VLA_Prolog.PROLOG("(query (list X) (list X Y))")
    r = ResultDescribe(result)
    Report "prolog.24: (list X) beside (list X Y) in goal position gets 'a list is not a goal', not an arity mismatch", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "builds a list, it is not something PROLOG can prove", vbTextCompare) > 0 _
           And InStr(1, r, "same arity", vbTextCompare) = 0, "got: " & r
    result = VLA_Prolog.PROLOG("(query (! X) (! X Y))")
    r = ResultDescribe(result)
    Report "prolog.24: (! X) beside (! X Y) gets the cut refusal, not an arity mismatch", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "cut is written on its own", vbTextCompare) > 0 _
           And InStr(1, r, "same arity", vbTextCompare) = 0, "got: " & r

    ' ---- CUT IN PARENTHESES, at one arity - the silent case the
    ' derivation found. `(!)` was a compound goal named "!" that nothing
    ' dispatched: it fell through to the clauseDict lookup and failed.
    result = VLA_Prolog.PROLOG("(query (!))")
    r = ResultDescribe(result)
    Report "prolog.24: (query (!)) is refused by name - it used to answer FALSE", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "cut is written on its own", vbTextCompare) > 0, "got: " & r
    result = VLA_Prolog.PROLOG("(fact (p 1)) (rule (first X) (p X) (!)) (query (first X))")
    r = ResultDescribe(result)
    Report "prolog.24: (!) in a rule body is refused by name - the rule used to answer nothing, silently", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "cut is written on its own", vbTextCompare) > 0, "got: " & r
    ' ...and its control: the same rule with a BARE ! answers its one row,
    ' so the refusal is about the parentheses and nothing else.
    result = VLA_Prolog.PROLOG("(fact (p 1)) (rule (first X) (p X) !) (query (first X))")
    Report "prolog.24: ...while the same rule with a bare ! answers its one row", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, "1"), "got: " & ResultDescribe(result)

    ' ---- THE ARM IS FOR REFUSED NAMES ONLY. A predicate a program can
    ' define still has its arity held to one - the catch-all must not be
    ' read as "skip every name".
    result = VLA_Prolog.PROLOG("(fact (p 1)) (query (p X) (p X Y))")
    r = ResultDescribe(result)
    Report "prolog.24: an ordinary predicate used at two arities is still refused as an arity mismatch", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "must have the same arity", vbTextCompare) > 0, "got: " & r
End Sub

' PROLOG.22: AN UNKNOWN PREDICATE REFUSES - statically, over the
' predicates the query can reach, before a goal is solved. Pure: no live
' workbook (a table argument with no rows is TestPrologHostTable's).
'
' Every case below used to FAIL SILENTLY, and three were confidently wrong
' rather than merely empty: `not` over an undefined goal said TRUE, findall
' over one built an empty bag that `length` then counted, and a disjunction
' with a misspelled branch answered from the other branch as though
' nothing were wrong. Each is asserted on refusal text NAMING the
' predicate, so none can pass by failing.
'
' The decision is pinned in both of its directions: a rule the query never
' calls is NOT examined (reachable, not whole-program), and a branch the
' data never reaches IS (static, not raise-when-reached). The roadmap entry
' has the options and why this one.
Private Sub TestPrologUnknownPredicate()
    Dim result As Variant
    Dim r As String
    Dim q As String
    q = Chr$(34)

    ' ---- THE PLAIN CASE. It used to spill a header with nothing under it.
    result = VLA_Prolog.PROLOG("(fact (parent tom bob)) (query (parnet tom X))")
    r = ResultDescribe(result)
    Report "prolog.22: a misspelled predicate refuses by name, where it used to spill an empty column", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "'parnet' is called by this query", vbTextCompare) > 0, "got: " & r

    ' ---- THE CONFIDENTLY WRONG TRUE.
    result = VLA_Prolog.PROLOG("(fact (parent tom bob)) (query (not (parnet tom bob)))")
    r = ResultDescribe(result)
    Report "prolog.22: (not ...) over an undefined predicate refuses - it used to answer TRUE", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "'parnet' is called by this query", vbTextCompare) > 0, "got: " & r

    ' ---- THE EMPTY BAG THAT COUNTED.
    result = VLA_Prolog.PROLOG("(fact (parent tom bob)) (query (findall C (parnet tom C) Kids) (length Kids N))")
    r = ResultDescribe(result)
    Report "prolog.22: findall over an undefined predicate refuses - (length Kids N) used to count 0", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "'parnet' is called by this query", vbTextCompare) > 0, "got: " & r

    ' ---- A MISSPELLED BRANCH beside one that answers: refused even though
    ' the query has a real answer, because the program is wrong either way.
    result = VLA_Prolog.PROLOG("(fact (parent tom bob)) (query (or (parent tom X) (parnet tom X)))")
    r = ResultDescribe(result)
    Report "prolog.22: an (or ...) with one undefined branch refuses, even though the other branch answers", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "'parnet' is called by this query", vbTextCompare) > 0, "got: " & r

    ' ---- REACHED THROUGH A RULE: kin's second clause calls sibling, which
    ' nothing defines. Named, although the first clause alone would answer.
    result = VLA_Prolog.PROLOG("(fact (parent tom bob)) (rule (kin X Y) (parent X Y)) (rule (kin X Y) (sibling X Y)) (query (kin tom Y))")
    r = ResultDescribe(result)
    Report "prolog.22: an undefined predicate in the body of a rule the query calls refuses, naming it", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "'sibling' is called by this query", vbTextCompare) > 0, "got: " & r

    ' ---- ...and inside a goal inside a goal: findall within not.
    result = VLA_Prolog.PROLOG("(fact (p 1)) (query (p X) (not (findall Y (parnet X Y) B)))")
    r = ResultDescribe(result)
    Report "prolog.22: an undefined predicate nested inside findall inside not refuses too", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "'parnet' is called by this query", vbTextCompare) > 0, "got: " & r

    ' ---- NOT THE WHOLE PROGRAM: a rule nothing calls is never examined, so
    ' a rules text shared between cells with different tables still answers.
    result = VLA_Prolog.PROLOG("(fact (parent tom bob)) (rule (unused X) (sibling X X)) (query (parent tom Y))")
    Report "prolog.22: a rule the query never calls is not examined - its undefined sibling does not refuse the program", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, "bob"), "got: " & ResultDescribe(result)

    ' ---- STATIC, NOT WHEN REACHED: the else-branch never runs on this
    ' data, and it refuses anyway - otherwise the program would work today
    ' and break the day the data changed.
    result = VLA_Prolog.PROLOG("(fact (parent tom bob)) (query (if (parent tom X) (parent tom X) (parnet tom X)))")
    r = ResultDescribe(result)
    Report "prolog.22: an undefined predicate on an if-branch the data never reaches still refuses", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "'parnet' is called by this query", vbTextCompare) > 0, "got: " & r
    ' ...and its twin, with that branch defined, answers.
    result = VLA_Prolog.PROLOG("(fact (parent tom bob)) (query (if (parent tom X) (parent tom X) (parent bob X)))")
    Report "prolog.22: ...and with that branch naming a defined predicate, the same query answers bob", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, "bob"), "got: " & ResultDescribe(result)

    ' ---- DEFINING THE NAME is what stops it. Once sibling has a fact the
    ' same negation answers - TRUE, and correctly, because no sibling fact
    ' says tom and bob.
    result = VLA_Prolog.PROLOG("(fact (parent tom bob)) (fact (sibling ann bob)) (query (not (sibling tom bob)))")
    Report "prolog.22: once the predicate has a fact, (not ...) over it answers - a correct TRUE", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)

    ' ---- ORDER against the quoting diagnosis (PROLOG.12). Both would fire
    ' on this program; the undefined predicate is found first, before
    ' anything is solved, and the quoting message is never reached.
    result = VLA_Prolog.PROLOG("(fact (dept " & q & "eng" & q & ")) (query (dept eng) (parnet tom X))")
    r = ResultDescribe(result)
    Report "prolog.22: an undefined predicate is named before the post-hoc quoting diagnosis runs", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "'parnet' is called by this query", vbTextCompare) > 0 _
           And InStr(1, r, "quoting mismatch", vbTextCompare) = 0, "got: " & r

    ' ---- A NAME PROLOG.19 LEFT UNRESERVED. That item filed this one as
    ' "the general fix for the ~460 impure names, with no reservation": tab
    ' is an impure SWI builtin and still a name a knowledge base may define,
    ' and undefined it now refuses like any other.
    result = VLA_Prolog.PROLOG("(fact (p 1)) (query (p X) (tab 3))")
    r = ResultDescribe(result)
    Report "prolog.22: an unreserved impure SWI name (tab) refuses as undefined - the general fix PROLOG.19 filed", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "'tab' is called by this query", vbTextCompare) > 0, "got: " & r
End Sub

' PROLOG.28: DEPTH IS NOT WORK. One ceiling of 120 steps, charged per
' candidate tried, used to stand for both, so a 121-row scan, closure over a
' six-row Reports table and a fifteen-person join each refused - blaming a
' rule that recursed, when none did. Now DEPTH (nested SolveGoalList frames,
' the stack the old ceiling protected) keeps its 120, and WORK has a budget
' sized for real Tables. And a list may now be as long as work allows, so
' the term walkers that recursed into a list's tail - one frame per element -
' loop on it instead. Pure: every case builds its own program, and the long
' ones are built in a loop, never typed.
'
' Each budget is pinned at its exact edge from both sides: a chain of 59
' links nests exactly 120 frames and answers, 60 is refused by depth (the
' transliteration measured depth = 2 x links + 2 for this rule).
Private Sub TestPrologBudgets()
    Dim result As Variant
    Dim r As String, s As String
    Dim q As String
    q = Chr$(34)

    ' ---- DEPTH, at its edge. Deep DATA is not a runaway rule either, and
    ' the refusal says so and points at DATALOG, which follows any chain.
    result = VLA_Prolog.PROLOG(BudgetChainProgram(59) & "(query (path p1 p60))")
    Report "prolog.28: a terminating chain 59 links long nests exactly 120 frames and answers", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG(BudgetChainProgram(60) & "(query (path p1 p61))")
    r = ResultDescribe(result)
    Report "prolog.28: one link more is refused by DEPTH, never by work", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "rules deep", vbTextCompare) > 0 And InStr(1, r, "facts and rules", vbTextCompare) = 0, "got: " & r
    Report "prolog.28: the depth refusal names a chain deeper than PROLOG follows, and points at DATALOG", _
           InStr(1, r, "a chain deeper than PROLOG follows", vbTextCompare) > 0 And InStr(1, r, "DATALOG", vbTextCompare) > 0, "got: " & r

    ' ---- WORK IS NOT DEPTH. Five hundred candidates, two frames deep: the
    ' shape a real Table scan has, refused at the 121st row before PROLOG.28.
    s = ""
    Dim i As Long
    For i = 1 To 500
        s = s & "(fact (n " & i & ")) "
    Next i
    result = VLA_Prolog.PROLOG(s & "(query (n 250))")
    Report "prolog.28: a scan of 500 facts answers - work grows with the data, depth does not", _
           ResultBoolIs(result, True), "got: " & ResultDescribe(result)

    ' The corpus's own headline (scripts/pareto_logic.txt, org-under):
    ' "who reports to Alice, directly or not" over SIX rows - 176 units of
    ' work, 12 frames deep - refused before this item.
    result = VLA_Prolog.PROLOG( _
        "(fact (reports bob alice)) (fact (reports carol alice)) (fact (reports dave bob)) " & _
        "(fact (reports eve dave)) (fact (reports frank carol)) (fact (reports gus hal)) " & _
        "(rule (under X Y) (reports X Y)) (rule (under X Y) (reports X Z) (under Z Y)) " & _
        "(query (under Who alice))")
    Report "prolog.28: closure over six rows answers all five reports, Bob first and Frank last", _
           ResultRowCount(result) = 6 And ResultCellIs(result, 2, 1, "bob") And ResultCellIs(result, 6, 1, "frank"), _
           "got: " & ResultDescribe(result)

    ' ---- WORK, at its ceiling: a query that tries too much is refused by
    ' WORK, with the text that names a large Table, never a runaway rule.
    result = VLA_Prolog.PROLOG("(query (between 1 99999 X) (= X 0))")
    r = ResultDescribe(result)
    Report "prolog.28: a query that tries more than the work budget is refused by WORK, never by depth", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "facts and rules", vbTextCompare) > 0 And InStr(1, r, "rules deep", vbTextCompare) = 0, "got: " & r

    ' The old edges are gone: 120 values, and all 120 pieces of fourteen
    ' characters, used to be refused up front.
    result = VLA_Prolog.PROLOG("(query (between 1 120 X))")
    Report "prolog.28: (between 1 120 X), one past the old edge, now generates all 120", _
           ResultRowCount(result) = 121, "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (sub-atom " & q & String$(14, "a") & q & " B L A S))")
    Report "prolog.28: all 120 pieces of fourteen characters, the old sub-atom edge, now answer", _
           ResultRowCount(result) = 121, "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (atom-concat A B " & q & String$(119, "a") & q & "))")
    Report "prolog.28: all 120 splits of 119 characters, the old atom-concat edge, now answer", _
           ResultRowCount(result) = 121, "got: " & ResultDescribe(result)

    ' ---- LONG LISTS. One case keeps the full ten thousand - the cheapest,
    ' which only counts - and the rest work at two or three, because the
    ' suite must stay quick: the first draft of this Sub asked for 99,999
    ' solutions in one assertion and crashed Excel (see TestPrologBetween).
    ' A findall bag of ten thousand, built, bound and
    ' measured: the harvest, the occurs check that binding it runs, and
    ' length all walk it. Recursing into the tail, this nested ten thousand
    ' frames. Counted inside a rule so the bag itself is not a column.
    result = VLA_Prolog.PROLOG("(rule (n N) (findall X (between 1 300 X) B) (length B N)) (query (n N))")
    Report "prolog.28: a findall bag of 300 is built, bound and measured", _
           ResultCellIs(result, 2, 1, "300"), "got: " & ResultDescribe(result)

    ' Rendered: resolved, contracted and written, all three walkers.
    result = VLA_Prolog.PROLOG("(query (findall X (between 1 300 X) B))")
    s = ResultCellText(result, 2, 1)
    Report "prolog.28: a bag of 300 renders whole in one cell, first element to last", _
           Left$(s, 12) = "(list 1 2 3 " And Right$(s, 9) = " 299 300)", "got: " & Left$(s, 40) & " ... " & Right$(s, 40)

    ' Too long for any cell: refused by name rather than left to the grid.
    ' One long TEXT, not a long list: the guard is on the rendered length of
    ' an answer, whatever shape produced it, and a 40,000-character atom
    ' reaches it without building anything to walk (or to release - see the
    ' note on list length in TestPrologBudgets' own header).
    result = VLA_Prolog.PROLOG("(fact (big " & q & String$(40000, "a") & q & ")) (query (big X))")
    r = ResultDescribe(result)
    Report "prolog.28: an answer too long for a cell is refused by name, with the variable and the limit", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "answer for X", vbTextCompare) > 0 And InStr(1, r, "32,767", vbTextCompare) > 0, "got: " & r

    ' Two long lists compared, unified and tested whole - TermsIdentical,
    ' UnifyTwoWay and TermIsGroundDeep.
    result = VLA_Prolog.PROLOG("(rule (same Z) (findall X (between 1 300 X) A) (findall Y (between 1 300 Y) B) (== A B) (= A B) (ground? A) (= Z yes)) (query (same Z))")
    Report "prolog.28: two bags of 300 are identical, unify, and are ground", _
           ResultCellIs(result, 2, 1, "yes"), "got: " & ResultDescribe(result)

    ' The quoting diagnosis walks to the end: two bags of 10,000 that differ
    ' ONLY in the quoting of their last element are refused as that
    ' confusion, naming it - the walker reached the ten-thousandth pair.
    result = VLA_Prolog.PROLOG("(rule (same Z) (findall X (between 1 300 X) A) (findall Y (between 1 299 Y) B0) (append B0 (list " & q & "300" & q & ") B) (== A B) (= Z yes)) (query (same Z))")
    r = ResultDescribe(result)
    Report "prolog.28: bags differing only in the quoting of the last of 300 are refused as exactly that", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "300", vbTextCompare) > 0 And InStr(1, r, "number", vbTextCompare) > 0, "got: " & r

    ' ---- THE LIST BUDGET, at its edge, both sides. A list is a chain of
    ' cons cells and VBA releases such a chain recursively, so what a query
    ' may GATHER is bounded even though what it may READ is not: 1,000 is
    ' eight times under the length that raised "Out of stack space" live
    ' (archive/VLA_Diag3.bas, 2026-09-11).
    result = VLA_Prolog.PROLOG("(rule (n N) (findall X (between 1 1000 X) B) (length B N)) (query (n N))")
    Report "prolog.28: a bag of exactly 1,000 is gathered and measured - the largest that fits", _
           ResultCellIs(result, 2, 1, "1000"), "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(rule (n N) (findall X (between 1 1001 X) B) (length B N)) (query (n N))")
    r = ResultDescribe(result)
    Report "prolog.28: one more is refused by name, saying what to ask instead", _
           ResultTextStartsWith(result, "#PROLOG!") And InStr(1, r, "gather 1001", vbTextCompare) > 0 And InStr(1, r, "DATALOG", vbTextCompare) > 0, "got: " & r
    ' Reading is not gathering: the same query that cannot hold 1,001 rows
    ' in one list counts them one at a time without a murmur.
    result = VLA_Prolog.PROLOG("(query (between 1 5000 X) (= X 4999))")
    Report "prolog.28: reading 5,000 values is fine - it is holding them in one list that is not", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, "4999"), "got: " & ResultDescribe(result)

    ' A split into ten thousand parts - one native step before PROLOG.28,
    ' and already a list whose first walk could run the stack out.
    result = VLA_Prolog.PROLOG("(rule (k N) (atomic-list-concat Parts " & q & "," & q & " " & q & String$(299, ",") & q & ") (length Parts N)) (query (k N))")
    Report "prolog.28: text split into 300 parts is measured whole", _
           ResultCellIs(result, 2, 1, "300"), "got: " & ResultDescribe(result)
End Sub

' A chain p1 -> p2 -> ... of n links as facts, and the right-recursive
' closure over it - TestPrologBudgets' depth edge. Built, never typed: 60
' links would not fit on one VBA source line.
Private Function BudgetChainProgram(ByVal n As Long) As String
    Dim s As String, i As Long
    For i = 1 To n
        s = s & "(fact (edge p" & i & " p" & (i + 1) & ")) "
    Next i
    BudgetChainProgram = s & "(rule (path X Y) (edge X Y)) (rule (path X Y) (edge X Z) (path Z Y)) "
End Function
' =====================================================================
'  G-PROLOG slice 1 - the sentence layer that writes these engines'
'  programs. Here rather than in VlaSelfTest, the owner's call: these
'  are Prolog's tests, and TestDSLs is where Prolog's tests run.
' =====================================================================

' The `conditions` sub-grammar, SD-16's third built-in and the first
' held to being REGULAR. english.vla's own test-success rows prove the
' four shipped sentences end to end; these pins hold the machinery
' underneath them, one shape at a time, on a throwaway rule - so a
' failure names the shape that broke instead of one long sentence.
'
' The pins that matter most are the two that are not about syntax. A
' role noun must become the SAME variable everywhere it is said, which
' is the only way a join is expressed here; and a role nothing binds
' must refuse, because DATALOG would refuse it in words naming a column
' the writer never typed (datalog-negation-unsafe-variable) and PROLOG
' would quietly answer from an unbound variable.
Private Sub TestPrologConditions()
    Dim d As String
    Dim junk As String

    EnglishResetGrammar
    EnglishAddPhrase "assay {c:conditions}", "(debug-print {c})"

    AssertConditions "a Table row keys each column by its header, the role as the variable", _
                     "Assay Staff lists the person as Name, the cert as Cert, and the level as Level.", _
                     "(staff (name Person) (cert Cert) (level Level))"
    ' The tokenizer drops "the" from every sentence (IsDroppedWord), so
    ' the grammar never sees it and both spellings are one sentence. The
    ' first draft required "the", and the vocabulary refused to load live.
    AssertConditions "a role reads the same without its article", _
                     "Assay Staff lists person as Name, cert as Cert, and level as Level.", _
                     "(staff (name Person) (cert Cert) (level Level))"
    AssertConditions "two columns join with a bare and, no Oxford comma needed", _
                     "Assay Leave lists the person as Name and the shift as Shift.", _
                     "(leave (name Person) (shift Shift))"
    AssertConditions "a relation between two roles", _
                     "Assay the person holds the cert.", "(holds Person Cert)"
    AssertConditions "a set over one role", _
                     "Assay the bill is big.", "(big Bill)"
    AssertConditions "is at least becomes >=", _
                     "Assay Staff lists the level as Level, and the level is at least the level.", _
                     "(>= Level Level)"
    ' The shared PROLOG/DATALOG subset has no <=, so "at most" is the
    ' same operator with its operands swapped - the roadmap's own words.
    AssertConditions "is at most becomes a SWAPPED >=, never a <=", _
                     "Assay Staff lists the level as Level, the cap as Cap, and the level is at most the cap.", _
                     "(>= Cap Level)"
    AssertConditions "is greater than becomes >", _
                     "Assay Staff lists the level as Level, the cap as Cap, and the level is greater than the cap.", _
                     "(> Level Cap)"
    AssertConditions "is less than becomes <", _
                     "Assay Staff lists the level as Level, the cap as Cap, and the level is less than the cap.", _
                     "(< Level Cap)"

    ' The window that keeps this regular: after ", and", "the level as
    ' Level" continues the row and "the level is ..." begins a new
    ' condition. One sentence holds both, so a parser that got the
    ' window wrong cannot pass this.
    AssertConditions "a column list ends where a comparison begins", _
                     "Assay Shifts lists the shift as Shift, the min as MinLevel, and the min is at least the shift.", _
                     "(shifts (shift Shift) (minlevel Min)) (>= Min Shift)"

    ' A negated Table row is always written through a generated
    ' projection rule, named from its own content, and the rule lands
    ' OUTSIDE the goals - which is why the slot binds twice. Both roles
    ' are bound POSITIVELY first, by the two rows above the negation.
    EnglishResetGrammar
    EnglishAddPhrase "assay {c:conditions} finally", "(debug-print {c-rules} {c})"
    AssertConditions "a negated Table row generates its projection rule, named from table and columns", _
                     "Assay Staff lists the person as Name, and Shifts lists the shift as Shift, and not Leave lists the person as Name and the shift as Shift finally.", _
                     "(rule (vla-not-leave-name-shift Person Shift) (leave (name Person) (shift Shift)))"
    AssertConditions "the body negates the generated rule, never the row", _
                     "Assay Staff lists the person as Name, and Shifts lists the shift as Shift, and not Leave lists the person as Name and the shift as Shift finally.", _
                     "(not (vla-not-leave-name-shift Person Shift))"
    ' A "-" inside a part is written "--", so the two ways of splitting
    ' one hyphenated string can never share a generated name. Before the
    ' escape, both of these were vla-not-leave-start-date-shift.
    AssertConditions "a hyphen inside a header is doubled in the generated name", _
                     "Assay Staff lists person as Name, and Shifts lists shift as Shift, and not Leave lists person as Start-Date and shift as Shift finally.", _
                     "(rule (vla-not-leave-start--date-shift Person Shift)"
    AssertConditions "so the other split of the same letters gets a different name", _
                     "Assay Staff lists person as Name, and Shifts lists shift as Shift, and not Leave lists person as Start and shift as Date-Shift finally.", _
                     "(rule (vla-not-leave-start-date--shift Person Shift)"

    EnglishResetGrammar
    EnglishAddPhrase "assay {c:conditions}", "(debug-print {c})"
    AssertConditionsRefusal "a role only ever compared refuses, naming the role the writer typed", _
                            "Assay Staff lists the person as Name, and the level is at least the min.", _
                            "nothing in this rule says which level it means"
    AssertConditionsRefusal "a role only ever negated refuses the same way", _
                            "Assay not Leave lists the person as Name and the shift as Shift.", _
                            "nothing in this rule says which"
    AssertConditionsRefusal "a set named with one of the grammar's own words refuses", _
                            "Assay the person is lists.", _
                            "is one of this grammar's own words"
    ' vla- is where the generated rules live, so nothing a writer names
    ' may start with it - in a condition, in a rule's head, or in a
    ' question.
    AssertConditionsRefusal "a relation in a condition named vla- refuses", _
                            "Assay person vla-holds cert.", _
                            "starts with vla-"

    EnglishResetGrammar
    EnglishAddPhrase "relate {x:relation}", "(debug-print {x})"
    AssertConditions "a relation slot binds its one bare word as the predicate name", _
                     "Relate can-cover.", "(debug-print ""can-cover"")"
    AssertConditionsRefusal "a relation slot refuses a vla- name by name", _
                            "Relate vla-can-cover.", "starts with vla-"
    AssertConditionsRefusal "a relation slot does not take a quoted name", _
                            "Relate ""can cover"".", "a relation name"
End Sub

' AssertEnglish's shape, reported through THIS module's own counters -
' VLA_Tests.bas's helper increments VlaSelfTest's, which would leave
' TestDSLs' total wrong by exactly the number of pins above.
Private Sub AssertConditions(ByVal name As String, ByVal sentence As String, ByVal frag As String)
    Dim t As String, d As String
    On Error Resume Next
    Err.Clear
    t = EnglishToVla(sentence)
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    If Len(d) > 0 Then
        Report "g-prolog: " & name, False, "unexpected refusal: " & d
        Exit Sub
    End If
    Report "g-prolog: " & name, InStr(1, t, frag, vbTextCompare) > 0, _
           "missing '" & frag & "' in: " & Left$(t, 200)
End Sub

Private Sub AssertConditionsRefusal(ByVal name As String, ByVal sentence As String, ByVal frag As String)
    Dim t As String, d As String
    On Error Resume Next
    Err.Clear
    t = EnglishToVla(sentence)
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    If Len(d) = 0 Then
        Report "g-prolog: " & name, False, "expected a refusal but it translated to: " & Left$(t, 200)
        Exit Sub
    End If
    Report "g-prolog: " & name, InStr(1, d, frag, vbTextCompare) > 0, "got: " & d
End Sub

' VLA_Runtime.VlaTableArguments, purely. It builds a piece of Excel
' FORMULA syntax out of a sentence's own G6 list, so the pins are about
' two things: that the separator is exactly what a formula's argument
' list wants, and that a name which would change the formula's SHAPE is
' refused rather than spliced. That second half is why this is not a
' general string join - a join has no grounds to refuse anything, and
' the corrupted formula would reach the user as an Excel parse error
' naming nothing.
Private Sub TestTableArguments()
    Report "g-prolog: three tables join with a comma and a space", _
           VLA_Runtime.VlaTableArguments(Array("staff", "shifts", "leave")) = "staff, shifts, leave", _
           "got: " & VLA_Runtime.VlaTableArguments(Array("staff", "shifts", "leave"))
    Report "g-prolog: one table is itself, with no separator", _
           VLA_Runtime.VlaTableArguments(Array("staff")) = "staff", _
           "got: " & VLA_Runtime.VlaTableArguments(Array("staff"))
    Report "g-prolog: underscores, digits and periods are legal in a Table name", _
           VLA_Runtime.VlaTableArguments(Array("_q1", "sales.2026", "tbl_3")) = "_q1, sales.2026, tbl_3", _
           "got: " & VLA_Runtime.VlaTableArguments(Array("_q1", "sales.2026", "tbl_3"))

    Dim bad As Variant
    Dim desc As String
    Dim ignored As String
    ' Each of these would not merely name a missing Table - spliced into
    ' the formula string it would end an argument early, or open a
    ' string, and the formula would then mean something nobody asked
    ' for. A bare {:text-list} item cannot carry any of them, but a
    ' QUOTED one can (G2's "quotes are the door"), so the check is real.
    For Each bad In Array("a,b", "a" & Chr$(34) & "b", "a)b", "a b", "1staff", "")
        desc = ""
        On Error Resume Next
        Err.Clear
        ignored = VLA_Runtime.VlaTableArguments(Array(bad))
        desc = Err.Description
        On Error GoTo 0
        Report "g-prolog: Table name '" & CStr(bad) & "' refuses by name", _
               InStr(1, desc, "cannot name a data table", vbTextCompare) > 0, "got: " & desc
    Next bad

    desc = ""
    On Error Resume Next
    Err.Clear
    ignored = VLA_Runtime.VlaTableArguments(Array())
    desc = Err.Description
    On Error GoTo 0
    Report "g-prolog: naming no tables at all refuses by name", _
           InStr(1, desc, "at least one data table", vbTextCompare) > 0, "got: " & desc
End Sub

' G-PROLOG slice 2: the clause and question sub-grammars and the range lint
' (VLA_SentenceEngine.bas, "G-PROLOG slice 2: constants, the clause, and the
' question"). Each pin runs one sentence through EnglishToVla under a test
' phrase and looks for a fragment of the VLA it translates to, or of the
' refusal it raises, through this module's own AssertConditions helpers - so
' TestDSLs' own counters see them. Every fragment was derived by running the
' sentence through a transliteration of the sub-grammars, never typed by
' hand, and each Sub's pins turn red under its own mutations there.
Private Sub TestPrologClauses()
    EnglishResetGrammar
    EnglishAddPhrase "clause {h:clause}", "(debug-print {h})"

    ' The head's shape is read from its own tokens: a relation between two
    ' operands, or "is" and a set.
    AssertConditions "a relation rule writes its head from its two role nouns", _
                     "Clause a person can-cover a shift if Staff lists the person as Name, and Shifts lists the shift as Shift.", _
                     "(rule (can-cover Person Shift) (staff (name Person)) (shifts (shift Shift)))"
    AssertConditions "a set head writes a one-place rule, and a number is compared bare", _
                     "Clause a bill is big if Bills lists the bill as Bill and the amount as Amount, and the amount is greater than 10000.", _
                     "(rule (big Bill) (bills (bill Bill) (amount Amount)) (> Amount 10000))"
    AssertConditions "a set head skips the article before its set", _
                     "Clause a bill is a violation if the bill is flagged.", _
                     "(rule (violation Bill) (flagged Bill))"

    ' No "if" makes a fact, which names only values - and a constant is typed
    ' by how it was written: quoted text stays quoted, a number stays bare.
    AssertConditions "a fact names its values, and quoted text is written quoted", _
                     "Clause ""Bob"" manages ""Carol"".", _
                     "(fact (manages \""Bob\"" \""Carol\""))"
    AssertConditions "a set fact names one value", _
                     "Clause ""B3"" is flagged.", _
                     "(fact (flagged \""B3\""))"
    AssertConditions "a number in a fact is written bare", _
                     "Clause ""Bob"" has-cap 10000.", _
                     "(fact (has-cap \""Bob\"" 10000))"
    AssertConditions "a code that only looks like a number stays text", _
                     "Clause ""007"" is secret.", _
                     "(fact (secret \""007\""))"
    AssertConditions "a quote inside a constant is escaped for the reader", _
                     "Clause ""Bob"" says ""a """"b"""""".", _
                     "\""a \\\""b\\\""\"""
    AssertConditions "a constant may stand in a rule's head", _
                     "Clause ""Bob"" can-cover a shift if Shifts lists the shift as Shift.", _
                     "(rule (can-cover \""Bob\"" Shift) (shifts (shift Shift)))"

    ' A constant stands where a role does: in a Table column, either side of a
    ' relation. There is no "=": both engines agree on these, and measured,
    ' they did not agree on any spelling of equality.
    AssertConditions "a constant fills a Table column", _
                     "Clause a person is in-ops if Roster lists the person as Name and ""Ops"" as Dept.", _
                     "(roster (name Person) (dept \""Ops\""))"
    AssertConditions "a constant is a relation's operand", _
                     "Clause a person is alice-report if the person reports-to ""Alice"".", _
                     "(reports-to Person \""Alice\"")"
    ' In a negated row the constant goes in the CALL, so one projection name is
    ' only ever one rule, whatever value two sentences negate.
    AssertConditions "a constant in a negated row goes in the call", _
                     "Clause a person is off-at-night if Staff lists the person as Name, and not Leave lists the person as Name and ""Night"" as Shift.", _
                     "(not (vla-not-leave-name-shift Person \""Night\""))"
    AssertConditions "...and its projection rule keeps every column a variable", _
                     "Clause a person is off-at-night if Staff lists the person as Name, and not Leave lists the person as Name and ""Night"" as Shift.", _
                     "(rule (vla-not-leave-name-shift Name Shift) (leave (name Name) (shift Shift)))"

    ' The refusals. Each sentence parses, and would mean something its writer
    ' did not say, or answer two ways in two engines.
    AssertConditionsRefusal "a head role no condition binds refuses, naming the role", _
                            "Clause a person can-cover a shift if Staff lists the person as Name.", _
                            "nothing in this rule says which shift it means"
    AssertConditionsRefusal "a fact with a role in it refuses", _
                            "Clause ""Bob"" manages carol.", _
                            "'carol' is a role, not a value"
    AssertConditionsRefusal "a quoted number refuses", _
                            "Clause ""Alice"" has-limit ""50000"".", _
                            "is a number written in quotes"
    AssertConditionsRefusal "equality with a constant refuses, and teaches the column spelling", _
                            "Clause a person is in-ops if Roster lists the person as Name and the dept as Dept, and the dept is ""Ops"".", _
                            "compares the dept to one fixed value"
    AssertConditionsRefusal "a comparison with quoted text refuses", _
                            "Clause a person is late if Roster lists the person as Name and the dept as Dept, and the dept is greater than ""M"".", _
                            "a comparison needs a number on each side"
    AssertConditionsRefusal "a vla- relation in a head refuses", _
                            "Clause a person vla-covers a shift if Shifts lists the shift as Shift, and Staff lists the person as Name.", _
                            "starts with vla-"
    AssertConditionsRefusal "a set named with one of the grammar's own words refuses", _
                            "Clause a bill is lists if Bills lists the bill as Bill.", _
                            "is one of this grammar's own words"
End Sub

Private Sub TestPrologQuestions()
    EnglishResetGrammar
    EnglishAddPhrase "quiz {q:question}", "(debug-print {q-engine} {q})"

    ' A question's SHAPE writes its program: no unknown is one ground atom, and
    ' any unknown a narrowing rule whose head is the unknowns - which is what
    ' names the answer's columns. Since slice 3 every question goes to DATALOG,
    ' a whether as one ground atom (DATALOG.9). The program text arrives with
    ' its quotes doubled for the formula's string literal.
    AssertConditions "who routes to DATALOG through a narrowing rule", _
                     "Quiz who can-cover ""Night"".", _
                     """DATALOG"" ""(rule (vla-ask-can-cover Who) (can-cover Who \""\""Night\""\"")) (query vla-ask-can-cover)"""
    AssertConditions "what asks for the object, written first", _
                     "Quiz what ""Bob"" can-cover.", _
                     "(rule (vla-ask-can-cover What) (can-cover \""\""Bob\""\"" What))"
    AssertConditions "whether is one ground atom, and routes to DATALOG", _
                     "Quiz whether ""Bob"" can-cover ""Night"".", _
                     """DATALOG"" ""(query (can-cover \""\""Bob\""\"" \""\""Night\""\""))"""
    AssertConditions "who ... what is a pair, both unknowns in the head", _
                     "Quiz who can-cover what.", _
                     "(rule (vla-ask-can-cover Who What) (can-cover Who What))"
    AssertConditions "which names each unknown, and so each header", _
                     "Quiz which person can-cover which shift.", _
                     "(rule (vla-ask-can-cover Person Shift) (can-cover Person Shift))"
    AssertConditions "which may ask for the object first", _
                     "Quiz which shift ""Bob"" can-cover.", _
                     "(rule (vla-ask-can-cover Shift) (can-cover \""\""Bob\""\"" Shift))"
    AssertConditions "who is a set", _
                     "Quiz who is top.", _
                     "(rule (vla-ask-top Who) (top Who))"
    AssertConditions "what is a set, skipping the article", _
                     "Quiz what is a violation.", _
                     "(rule (vla-ask-violation What) (violation What))"
    AssertConditions "whether over a set", _
                     "Quiz whether ""B4"" is a violation.", _
                     """DATALOG"" ""(query (violation \""\""B4\""\""))"""
    ' Slice 1 quoted every constant, and PROLOG answered a numeric one FALSE.
    AssertConditions "a number stays bare in a question", _
                     "Quiz whether ""Ann"" has-level 3.", _
                     "(query (has-level \""\""Ann\""\"" 3))"
    ' Every quote is doubled, so one inside a constant cannot end the formula's
    ' string early.
    AssertConditions "a quote inside a constant is doubled with the rest", _
                     "Quiz who says ""a """"b"""""".", _
                     "(says Who \""\""a \\\""\""b\\\""\""\""\"")"

    AssertConditionsRefusal "the same unknown twice refuses", _
                            "Quiz who manages who.", _
                            "asks for who twice"
    AssertConditionsRefusal "whether with an unknown refuses", _
                            "Quiz whether ""Bob"" can-cover what.", _
                            "also asks for what"
    AssertConditionsRefusal "a quoted number in a question refuses", _
                            "Quiz whether ""Ann"" has-level ""3"".", _
                            "is a number written in quotes"
    AssertConditionsRefusal "a vla- relation in a question refuses", _
                            "Quiz who vla-covers ""Night"".", _
                            "starts with vla-"
End Sub

Private Sub TestPrologRangeLint()
    EnglishResetGrammar
    EnglishAddPhrase "inscribe {r:cell} that {h:clause}", "(debug-print {h})"
    EnglishAddPhrase "pose {q:question} over {rules:range}", "(debug-print {q})"

    ' "Or" is a second rule cell with the same head. A range holding one of
    ' them but not the other answers from part of the relation, silently, in
    ' both engines - so within one program it refuses at Check.
    AssertConditions "a range holding every cell of a relation translates", _
                     "Inscribe H2 that a device fits a dock if Compat lists the device as A and the dock as B." & vbLf & "Inscribe H3 that a device fits a dock if Compat lists the dock as A and the device as B." & vbLf & "Pose whether ""monitor-4"" fits ""dock-1"" over H2:H3.", _
                     "(query (fits"
    AssertConditionsRefusal "a range that leaves out one cell of a relation refuses, naming the cell", _
                            "Inscribe H2 that a device fits a dock if Compat lists the device as A and the dock as B." & vbLf & "Inscribe H3 that a device fits a dock if Compat lists the dock as A and the device as B." & vbLf & "Pose whether ""monitor-4"" fits ""dock-1"" over H2:H2.", _
                            "'fits' is also written in cell H3, which H2:H2 leaves out"
    AssertConditionsRefusal "a relation the question reaches through another rule is checked too", _
                            "Inscribe H2 that a device fits a dock if Compat lists the device as A and the dock as B." & vbLf & "Inscribe H5 that a device fits a dock if Compat lists the dock as A and the device as B." & vbLf & "Inscribe H4 that a device is dockable if the device fits the dock." & vbLf & "Pose who is dockable over H2:H4.", _
                            "'fits' is also written in cell H5, which H2:H4 leaves out"
    AssertConditions "a range holding none of a relation's cells is left to the engine", _
                     "Inscribe H2 that a device fits a dock if Compat lists the device as A and the dock as B." & vbLf & "Inscribe H3 that a device fits a dock if Compat lists the dock as A and the device as B." & vbLf & "Pose whether ""monitor-4"" fits ""dock-1"" over H5:H6.", _
                     "(query (fits"
    AssertConditions "a sheet-qualified range is not guessed at", _
                     "Inscribe H2 that a device fits a dock if Compat lists the device as A and the dock as B." & vbLf & "Inscribe H3 that a device fits a dock if Compat lists the dock as A and the device as B." & vbLf & "Pose whether ""monitor-4"" fits ""dock-1"" over Data!H2:H2.", _
                     "(query (fits"
    AssertConditions "a cell written twice is one cell", _
                     "Inscribe H2 that a device fits a dock if Compat lists the device as A and the dock as B." & vbLf & "Inscribe H2 that a device fits a dock if Compat lists the dock as A and the device as B." & vbLf & "Pose whether ""monitor-4"" fits ""dock-1"" over H2:H2.", _
                     "(query (fits"
    ' The records belong to one translation: a question translated on its own
    ' knows nothing of cells an earlier program wrote.
    Dim ignored As String
    On Error Resume Next
    ignored = EnglishToVla("Inscribe H2 that a device fits a dock if Compat lists the device as A and the dock as B." & vbLf & "Inscribe H3 that a device fits a dock if Compat lists the dock as A and the device as B.")
    On Error GoTo 0
    AssertConditions "records do not outlive the translation that made them", _
                     "Pose whether ""monitor-4"" fits ""dock-1"" over H2:H2.", _
                     "(query (fits"
End Sub

' G-PROLOG slice 3: recursion and closure (VLA_SentenceEngine.bas, its slice-3
' section). "directly or not" writes two generated, right-recursive rules and
' a question carries its own; left recursion and a relation named after a
' table refuse at Check; and every question goes to DATALOG. Every fragment
' was derived by running its sentence through a transliteration of the
' sub-grammars, never typed, and each check turns pins red under its own
' mutation there.
Private Sub TestPrologClosure()
    EnglishResetGrammar
    EnglishAddPhrase "clause {h:clause}", "(debug-print {h})"

    ' "directly or not" after a relation asks for one step along it or more.
    ' It writes two generated rules, always right-recursive, before the rule -
    ' once per cell, and named from the relation alone, its hyphen doubled.
    AssertConditions "directly or not writes the closure's two rules, right-recursive, before the rule", _
                     "Clause a person is-under a boss if the person reports-to the boss directly or not.", _
                     "(rule (vla-any-reports--to X Y) (reports-to X Y)) (rule (vla-any-reports--to X Y) (reports-to X Z) (vla-any-reports--to Z Y)) (rule (is-under Person Boss) (vla-any-reports--to Person Boss))"
    AssertConditions "a closure said twice in one rule writes its rules once", _
                     "Clause a person far-under a boss if the person reports-to the middle directly or not, and the middle reports-to the boss directly or not.", _
                     "(debug-print ""(rule (vla-any-reports--to X Y) (reports-to X Y)) (rule (vla-any-reports--to X Y) (reports-to X Z) (vla-any-reports--to Z Y)) (rule (far-under Person Boss)"
    AssertConditions "a closure under not negates the generated relation", _
                     "Clause a person is outside if Roster lists the person as Name, and not the person reports-to ""Alice"" directly or not.", _
                     "(not (vla-any-reports--to Person \""Alice\""))"
    AssertConditions "a constant stands on either side of a closure", _
                     "Clause a person is alice-side if the person reports-to ""Alice"" directly or not.", _
                     "(rule (alice-side Person) (vla-any-reports--to Person \""Alice\""))"
    ' A rule may still call itself - it is left recursion, the call FIRST,
    ' that never finishes in PROLOG.
    AssertConditions "a recursive call written last is not left recursion", _
                     "Clause a person is-under a boss if the person reports-to the middle, and the middle is-under the boss.", _
                     "(rule (is-under Person Boss) (reports-to Person Middle) (is-under Middle Boss))"
    AssertConditions "...and nor is a closure over the rule's own relation written last", _
                     "Clause a person is-under a boss if the person reports-to the middle, and the middle is-under the boss directly or not.", _
                     "(vla-any-is--under Middle Boss)"

    ' The refusals, each by name.
    AssertConditionsRefusal "directly or not after a set refuses", _
                            "Clause a bill is odd if the bill is big directly or not.", _
                            "Here it follows a set"
    AssertConditionsRefusal "directly or not after a comparison refuses", _
                            "Clause a person is late if Staff lists the person as Name and the level as Level, and the level is at least 3 directly or not.", _
                            "Here it follows a comparison"
    AssertConditionsRefusal "directly or not after a Table row refuses", _
                            "Clause a person is late if Staff lists the person as Name directly or not.", _
                            "Here it follows a table row"
    AssertConditionsRefusal "directly or not in a rule's head refuses", _
                            "Clause a person reports-to a boss directly or not if the person manages the boss.", _
                            "not in what a rule concludes"
    AssertConditionsRefusal "left recursion refuses, naming the relation", _
                            "Clause a person is-under a boss if the person is-under the middle, and the middle reports-to the boss.", _
                            "first condition asks 'is-under' again"
    AssertConditionsRefusal "left recursion through a closure of its own relation refuses too", _
                            "Clause a person is-under a boss if the person is-under the boss directly or not.", _
                            "first condition asks 'is-under' again"
    AssertConditionsRefusal "a set rule that asks itself first refuses", _
                            "Clause a bill is late if the bill is late, and Bills lists the bill as Bill.", _
                            "first condition asks 'late' again"
    AssertConditionsRefusal "a relation named after the table it reads refuses", _
                            "Clause a source feeds a target if Feeds lists the source as Source and the target as Target.", _
                            "names both this rule and the table it reads"
    AssertConditionsRefusal "...and so does a set named after its table", _
                            "Clause a vendor is preferred if Preferred lists the vendor as Vendor.", _
                            "names both this rule and the table it reads"
    AssertConditionsRefusal "...and one that reads its table only under not", _
                            "Clause a person is leave if Staff lists the person as Name, and not Leave lists the person as Name.", _
                            "names both this rule and the table it reads"

    EnglishResetGrammar
    EnglishAddPhrase "quiz {q:question}", "(debug-print {q-engine} {q})"
    ' A question carries its own closure, and asks through it - in DATALOG,
    ' which answers even where the data loops back on itself.
    AssertConditions "who ... directly or not carries its closure and asks through it", _
                     "Quiz who reports-to ""Alice"" directly or not.", _
                     """DATALOG"" ""(rule (vla-any-reports--to X Y) (reports-to X Y)) (rule (vla-any-reports--to X Y) (reports-to X Z) (vla-any-reports--to Z Y)) (rule (vla-ask-reports-to Who) (vla-any-reports--to Who \""\""Alice\""\"")) (query vla-ask-reports-to)"""
    AssertConditions "what ... directly or not asks for the object through it", _
                     "Quiz what ""Eve"" reports-to directly or not.", _
                     "(rule (vla-ask-reports-to What) (vla-any-reports--to \""\""Eve\""\"" What))"
    AssertConditions "whether ... directly or not is one ground atom over the closure", _
                     "Quiz whether ""Eve"" reports-to ""Alice"" directly or not.", _
                     """DATALOG"" ""(rule (vla-any-reports--to X Y) (reports-to X Y)) (rule (vla-any-reports--to X Y) (reports-to X Z) (vla-any-reports--to Z Y)) (query (vla-any-reports--to \""\""Eve\""\"" \""\""Alice\""\""))"""
    AssertConditions "a pair may ask through a closure", _
                     "Quiz which person reports-to which boss directly or not.", _
                     "(rule (vla-ask-reports-to Person Boss) (vla-any-reports--to Person Boss))"
    AssertConditions "without directly or not the question asks the relation itself, its neighbour", _
                     "Quiz who reports-to ""Alice"".", _
                     """DATALOG"" ""(rule (vla-ask-reports-to Who) (reports-to Who \""\""Alice\""\"")) (query vla-ask-reports-to)"""
    AssertConditionsRefusal "directly or not after a set in a question refuses", _
                            "Quiz who is top directly or not.", _
                            "Here it follows a set"

    EnglishResetGrammar
    EnglishAddPhrase "inscribe {r:cell} that {h:clause}", "(debug-print {h})"
    EnglishAddPhrase "pose {q:question} over {rules:range}", "(debug-print {q})"
    ' Across a program: a relation named after a table ANOTHER cell reads is
    ' refused naming both cells, and the range lint reaches a closure's own
    ' relation.
    AssertConditionsRefusal "a relation named after a table another cell reads refuses, naming both cells", _
                            "Inscribe H2 that a source flows-to a target if Feeds lists the source as Source and the target as Target." & vbLf & "Inscribe H3 that a source feeds a target if the source flows-to the target.", _
                            "'feeds' is written in cell H3, and the rule in cell H2 reads a table also called 'feeds'"
    AssertConditions "...and one with a name of its own translates", _
                     "Inscribe H2 that a source flows-to a target if Feeds lists the source as Source and the target as Target." & vbLf & "Inscribe H3 that a source feeds-into a target if the source flows-to the target directly or not.", _
                     "(vla-any-flows--to Source Target)"
    AssertConditionsRefusal "the range lint reaches the relation a closure reads", _
                            "Inscribe H2 that a person reports-to a boss if Reports lists the person as Employee and the boss as Manager." & vbLf & "Inscribe H5 that a person reports-to a boss if Delegates lists the person as From and the boss as To." & vbLf & "Inscribe H3 that a person is-under a boss if the person reports-to the boss directly or not." & vbLf & "Pose who is-under ""Alice"" over H2:H3.", _
                            "'reports-to' is also written in cell H5, which H2:H3 leaves out"
End Sub

' G-PROLOG slice 4: answer shapes (VLA_SentenceEngine.bas, its slice-4 section).
' "how many" counts the set in one cell, "each ... that is" per member of a
' set, "which ... that is ... is not" lists what is outside a set, "whether
' every ... that is ... is" asks that nothing is (DATALOG.10), and "alone"
' after the subject asks that exactly one holds; "are" reads as "is" in them.
' Every success fragment is the slice-4 model's own output for its sentence,
' never typed, and each check turns pins red under its own mutation there.
Private Sub TestPrologAnswerShapes()
    EnglishResetGrammar
    EnglishAddPhrase "quiz {q:question}", "(debug-print {q-engine} {q})"

    ' HOW MANY counts the set, through the relation with its other argument
    ' fixed, and writes (headless) so the cell holds the number alone. The
    ' generated name's part is escaped, as every slice-4 name is.
    AssertConditions "how many counts the subject, one number in one cell", _
                     "Quiz how many people can-cover ""Night"".", _
                     """DATALOG"" ""(headless) (rule (vla-count-can--cover People) (count People (can-cover VlaCounted \""\""Night\""\""))) (query vla-count-can--cover)"""
    AssertConditions "how many counts the object when a value comes first", _
                     "Quiz how many shifts ""Bob"" can-cover.", _
                     """DATALOG"" ""(headless) (rule (vla-count-can--cover Shifts) (count Shifts (can-cover \""\""Bob\""\"" VlaCounted))) (query vla-count-can--cover)"""
    AssertConditions "how many counts through a closure", _
                     "Quiz how many people reports-to ""Alice"" directly or not.", _
                     """DATALOG"" ""(rule (vla-any-reports--to X Y) (reports-to X Y)) (rule (vla-any-reports--to X Y) (reports-to X Z) (vla-any-reports--to Z Y)) (headless) (rule (vla-count-reports--to People) (count People (vla-any-reports--to VlaCounted \""\""Alice\""\""))) (query vla-count-reports--to)"""
    AssertConditions "a number stays bare in a count", _
                     "Quiz how many people has-level 3.", _
                     """DATALOG"" ""(headless) (rule (vla-count-has--level People) (count People (has-level VlaCounted 3))) (query vla-count-has--level)"""

    ' EACH gives every member of the set named after "that is" its count,
    ' the zero group included, under the two nouns as headers.
    AssertConditions "each counts per member of a set, the set read first", _
                     "Quiz how many people can-cover each shift that is listed.", _
                     """DATALOG"" ""(rule (vla-each-can--cover Shift People) (listed Shift) (count People (can-cover VlaCounted Shift))) (query vla-each-can--cover)"""
    AssertConditions "are reads as is in each", _
                     "Quiz how many people can-cover each shift that are listed.", _
                     """DATALOG"" ""(rule (vla-each-can--cover Shift People) (listed Shift) (count People (can-cover VlaCounted Shift))) (query vla-each-can--cover)"""

    ' NONE lists the domain's members not in the second set; EVERY writes the
    ' same rule and asks, through DATALOG.10, whether it holds no row.
    AssertConditions "which ... that is ... is not lists the members outside the second set", _
                     "Quiz which shift that is listed is not covered.", _
                     """DATALOG"" ""(rule (vla-none-listed-covered Shift) (listed Shift) (not (covered Shift))) (query vla-none-listed-covered)"""
    AssertConditions "...and are reads as is, the noun naming the header", _
                     "Quiz which shifts that are listed are not covered.", _
                     """DATALOG"" ""(rule (vla-none-listed-covered Shifts) (listed Shifts) (not (covered Shifts))) (query vla-none-listed-covered)"""
    AssertConditions "whether every ... is asks that nothing is outside, a negated query", _
                     "Quiz whether every shift that is listed is covered.", _
                     """DATALOG"" ""(rule (vla-none-listed-covered Shift) (listed Shift) (not (covered Shift))) (query (not (vla-none-listed-covered Shift)))"""
    AssertConditions "articles are skipped before a set", _
                     "Quiz whether every bill that is a violation is an excused.", _
                     """DATALOG"" ""(rule (vla-none-violation-excused Bill) (violation Bill) (not (excused Bill))) (query (not (vla-none-violation-excused Bill)))"""

    ' ALONE follows the subject: the relation holds, and it holds for exactly one.
    AssertConditions "who alone lists the one who can, when only one can", _
                     "Quiz who alone can-cover ""Night"".", _
                     """DATALOG"" ""(rule (vla-alone-can--cover Who) (can-cover Who \""\""Night\""\"") (count VlaCount (can-cover VlaCounted \""\""Night\""\"")) (= VlaCount 1)) (query vla-alone-can--cover)"""
    AssertConditions "which names the header of an alone question", _
                     "Quiz which person alone can-cover ""Night"".", _
                     """DATALOG"" ""(rule (vla-alone-can--cover Person) (can-cover Person \""\""Night\""\"") (count VlaCount (can-cover VlaCounted \""\""Night\""\"")) (= VlaCount 1)) (query vla-alone-can--cover)"""
    AssertConditions "whether ... alone answers TRUE or FALSE for one value", _
                     "Quiz whether ""Bob"" alone can-cover ""Night"".", _
                     """DATALOG"" ""(rule (vla-alone-can--cover Who) (can-cover Who \""\""Night\""\"") (count VlaCount (can-cover VlaCounted \""\""Night\""\"")) (= VlaCount 1)) (query (vla-alone-can--cover \""\""Bob\""\""))"""
    AssertConditions "alone asks through a closure", _
                     "Quiz who alone reports-to ""Carol"" directly or not.", _
                     """DATALOG"" ""(rule (vla-any-reports--to X Y) (reports-to X Y)) (rule (vla-any-reports--to X Y) (reports-to X Z) (vla-any-reports--to Z Y)) (rule (vla-alone-reports--to Who) (vla-any-reports--to Who \""\""Carol\""\"") (count VlaCount (vla-any-reports--to VlaCounted \""\""Carol\""\"")) (= VlaCount 1)) (query vla-alone-reports--to)"""

    ' The refusals, each by name.
    AssertConditionsRefusal "each with the same noun twice refuses", _
                            "Quiz how many shift can-cover each shift that is listed.", _
                            "asks for shift twice"
    AssertConditionsRefusal "alone is reserved - an object asked first cannot be alone", _
                            "Quiz what ""Bob"" alone can-cover.", _
                            "'alone' is one of this grammar's own words"
    AssertConditionsRefusal "...nor can a relation be named alone", _
                            "Quiz who alone ""Bob"".", _
                            "and alone are reserved"
    AssertConditionsRefusal "directly or not after the set of each refuses", _
                            "Quiz how many people can-cover each shift that is listed directly or not.", _
                            "Here it follows a set"
    AssertConditionsRefusal "directly or not after the set of a none question refuses", _
                            "Quiz which shift that is listed is not covered directly or not.", _
                            "Here it follows a set"
    AssertConditionsRefusal "directly or not after the set of an every question refuses", _
                            "Quiz whether every shift that is listed is covered directly or not.", _
                            "Here it follows a set"
    AssertConditionsRefusal "a quoted number in a count refuses", _
                            "Quiz how many people can-cover ""3"".", _
                            "is a number written in quotes"
    AssertConditionsRefusal "a vla- set in a none question refuses", _
                            "Quiz which shift that is vla-listed is not covered.", _
                            "starts with vla-"
    AssertConditionsRefusal "a vla- set in an every question refuses", _
                            "Quiz whether every shift that is listed is vla-covered.", _
                            "starts with vla-"
    ' A shape this slice does not have gets the ordinary near miss, never a
    ' refusal that names one of the grammar's words as a relation or set.
    AssertConditionsRefusal "how many over a set is not one of these shapes - the ordinary near miss", _
                            "Quiz how many people is top.", _
                            "expected a question"
    AssertConditionsRefusal "whether every ... is not ... is not a shape either", _
                            "Quiz whether every shift that is listed is not covered.", _
                            "expected a question"

    EnglishResetGrammar
    EnglishAddPhrase "clause {h:clause}", "(debug-print {h})"
    AssertConditionsRefusal "alone is reserved in a rule too", _
                            "Clause a person is alone if Staff lists the person as Name.", _
                            "'alone' is one of this grammar's own words"

    EnglishResetGrammar
    EnglishAddPhrase "inscribe {r:cell} that {h:clause}", "(debug-print {h})"
    EnglishAddPhrase "pose {q:question} over {rules:range}", "(debug-print {q})"
    ' The range lint reaches both sets a none, every or each question reads.
    AssertConditionsRefusal "a none question whose range leaves out a cell of its second set refuses", _
                            "Inscribe H2 that a shift is listed if Shifts lists the shift as Shift." & vbLf & "Inscribe H3 that a shift is covered if Rota lists the person as Name and the shift as Shift." & vbLf & "Inscribe H5 that a shift is covered if Cover lists the shift as Shift." & vbLf & "Pose which shift that is listed is not covered over H2:H3.", _
                            "'covered' is also written in cell H5, which H2:H3 leaves out"
    AssertConditionsRefusal "...and so does an every question", _
                            "Inscribe H2 that a shift is listed if Shifts lists the shift as Shift." & vbLf & "Inscribe H3 that a shift is covered if Rota lists the person as Name and the shift as Shift." & vbLf & "Inscribe H5 that a shift is covered if Cover lists the shift as Shift." & vbLf & "Pose whether every shift that is listed is covered over H2:H3.", _
                            "'covered' is also written in cell H5, which H2:H3 leaves out"
    AssertConditionsRefusal "an each question whose range leaves out a cell of its set refuses", _
                            "Inscribe H2 that a person can-cover a shift if Rota lists the person as Name and the shift as Shift." & vbLf & "Inscribe H3 that a shift is listed if Shifts lists the shift as Shift." & vbLf & "Inscribe H5 that a shift is listed if Extra lists the shift as Shift." & vbLf & "Pose how many people can-cover each shift that is listed over H2:H3.", _
                            "'listed' is also written in cell H5, which H2:H3 leaves out"
    AssertConditions "a range holding every cell of both sets translates", _
                     "Inscribe H2 that a shift is listed if Shifts lists the shift as Shift." & vbLf & "Inscribe H3 that a shift is covered if Rota lists the person as Name and the shift as Shift." & vbLf & "Inscribe H5 that a shift is covered if Cover lists the shift as Shift." & vbLf & "Pose which shift that is listed is not covered over H2:H5.", _
                     "(query vla-none-listed-covered)"
End Sub

' G-PROLOG slice 5: text conditions, a first match with "otherwise", and a
' list in one cell. A text test checks two values and binds neither; each
' branch after "otherwise" is its own conditions plus the negation of every
' earlier branch's guard; "as one list" joins one unknown's answers. Every
' success fragment is the slice-5 model's own output for its sentence, never
' typed, and each check turns pins red under its own mutation there.
Private Sub TestPrologOwnShapes()
    EnglishResetGrammar
    EnglishAddPhrase "clause {h:clause}", "(debug-print {h})"

    ' TEXT: starts with, ends with and contains, exact case, text on the right.
    AssertConditions "starts with writes a text test after the row that finds the code", _
                     "Clause a code is revenue if Accounts lists the code as Code, and the code starts with ""GL-4"".", _
                     """(rule (revenue Code) (accounts (code Code)) (text-starts-with Code \""GL-4\""))"""
    AssertConditions "ends with takes a quoted numeral as text", _
                     "Clause a code is tens if Accounts lists the code as Code, and the code ends with ""10"".", _
                     """(rule (tens Code) (accounts (code Code)) (text-ends-with Code \""10\""))"""
    AssertConditions "not before contains inverts the test", _
                     "Clause a name is plain if Products lists the name as Name, and not the name contains ""an"".", _
                     """(rule (plain Name) (products (name Name)) (not (text-contains Name \""an\"")))"""
    AssertConditions "a text test may compare two roles", _
                     "Clause a code is prefixed if Accounts lists the code as Code, and Prefixes lists the start as Start, and the code starts with the start.", _
                     """(rule (prefixed Code) (accounts (code Code)) (prefixes (start Start)) (text-starts-with Code Start))"""
    AssertConditions "a bare number on the right stays a number", _
                     "Clause a code is tens if Accounts lists the code as Code, and the code ends with 10.", _
                     """(rule (tens Code) (accounts (code Code)) (text-ends-with Code 10))"""
    AssertConditionsRefusal "a text test binds nothing - the role must be found first", _
                            "Clause a code is odd if the code starts with ""GL"".", _
                            "nothing in this rule says which code it means"
    AssertConditionsRefusal "directly or not after a text test refuses", _
                            "Clause a code is odd if Accounts lists the code as Code, and the code starts with ""GL"" directly or not.", _
                            "Here it follows a text test"
    AssertConditionsRefusal "contains is reserved - no relation may be named contains", _
                            "Clause an assembly contains a part if Parts lists the assembly as Assembly and the part as Component.", _
                            "'contains' is one of this grammar's own words"

    ' FIRST MATCH: one guard per branch but the last, kept to the subject, and
    ' each later branch ruling out every earlier one.
    AssertConditions "otherwise writes a guard for each earlier branch and a rule per branch", _
                     "Clause a customer has-tier ""Gold"" if Customers lists the customer as Customer and the spend as Spend, and the spend is at least 10000, otherwise ""Silver"" if Customers lists the customer as Customer and the spend as Spend, and the spend is at least 5000, otherwise ""Bronze"" if Customers lists the customer as Customer.", _
                     """(rule (vla-first-has--tier-1 Customer) (customers (customer Customer) (spend Spend)) (>= Spend 10000)) (rule (vla-first-has--tier-2 Customer) (customers (customer Customer) (spend Spend)) (>= Spend 5000)) (rule (has-tier Customer \""Gold\"") (customers (customer Customer) (spend Spend)) (>= Spend 10000)) (rule (has-tier Customer \""Silver\"") (customers (customer Customer) (spend Spend)) (>= Spend 5000) (not (vla-first-has--tier-1 Customer))) (rule (has-tier Customer \""Bronze\"") (customers (customer Customer)) (not (vla-first-has--tier-1 Customer)) (not (vla-first-has--tier-2 Customer)))"""
    AssertConditions "otherwise reads the same with no comma before it", _
                     "Clause a customer has-tier ""Gold"" if Customers lists the customer as Customer and the spend as Spend, and the spend is at least 10000 otherwise ""Bronze"" if Customers lists the customer as Customer.", _
                     """(rule (vla-first-has--tier-1 Customer) (customers (customer Customer) (spend Spend)) (>= Spend 10000)) (rule (has-tier Customer \""Gold\"") (customers (customer Customer) (spend Spend)) (>= Spend 10000)) (rule (has-tier Customer \""Bronze\"") (customers (customer Customer)) (not (vla-first-has--tier-1 Customer)))"""
    AssertConditions "a branch may conclude a role its own conditions find", _
                     "Clause a customer pays a rate if Customers lists the customer as Customer, and Promo lists the customer as Customer and the rate as Rate, otherwise 0 if Customers lists the customer as Customer.", _
                     """(rule (vla-first-pays-1 Customer) (customers (customer Customer)) (promo (customer Customer) (rate Rate))) (rule (pays Customer Rate) (customers (customer Customer)) (promo (customer Customer) (rate Rate))) (rule (pays Customer 0) (customers (customer Customer)) (not (vla-first-pays-1 Customer)))"""
    AssertConditions "a projection two branches need is written once", _
                     "Clause a person has-status ""Senior"" if Staff lists the person as Name and the level as Level, and the level is at least 3, and not Leave lists the person as Name, otherwise ""Junior"" if Staff lists the person as Name, and not Leave lists the person as Name.", _
                     """(rule (vla-not-leave-name Person) (leave (name Person))) (rule (vla-first-has--status-1 Person) (staff (name Person) (level Level)) (>= Level 3) (not (vla-not-leave-name Person))) (rule (has-status Person \""Senior\"") (staff (name Person) (level Level)) (>= Level 3) (not (vla-not-leave-name Person))) (rule (has-status Person \""Junior\"") (staff (name Person)) (not (vla-not-leave-name Person)) (not (vla-first-has--status-1 Person)))"""
    AssertConditionsRefusal "each value after otherwise needs its own if", _
                            "Clause a customer has-tier ""Gold"" if Customers lists the customer as Customer, otherwise ""Bronze"".", _
                            "needs its own ""if"""
    AssertConditionsRefusal "otherwise needs a relation with a second value, not a set", _
                            "Clause a customer is gold if Customers lists the customer as Customer, otherwise ""x"" if Customers lists the customer as Customer.", _
                            "needs a rule about a role and a second value"
    AssertConditionsRefusal "otherwise needs a role as the subject it chooses for", _
                            "Clause ""Acme"" has-tier ""Gold"" if Customers lists ""Acme"" as Customer, otherwise ""Bronze"" if Customers lists ""Acme"" as Customer.", _
                            "needs a rule about a role and a second value"
    AssertConditionsRefusal "a branch may not read the relation it chooses", _
                            "Clause a customer has-tier ""Gold"" if the customer has-tier ""Silver"", otherwise ""Bronze"" if Customers lists the customer as Customer.", _
                            "cannot also ask 'has-tier' in its own conditions"
    AssertConditionsRefusal "each later branch must find the subject itself", _
                            "Clause a customer has-tier ""Gold"" if Customers lists the customer as Customer, otherwise ""Bronze"" if Staff lists the person as Name.", _
                            "nothing in this rule says which customer it means"

    EnglishResetGrammar
    EnglishAddPhrase "quiz {q:question}", "(debug-print {q-engine} {q})"
    ' AS ONE LIST: one unknown's answers joined in one cell, headless, or one
    ' joined list for each member of a set named after "that is".
    AssertConditions "who as one list joins the answers in one cell", _
                     "Quiz who can-cover ""Night"" as one list.", _
                     """DATALOG"" ""(headless) (rule (vla-list-can--cover Who) (textjoin Who \""\"", \""\"" (can-cover VlaListed \""\""Night\""\""))) (query vla-list-can--cover)"""
    AssertConditions "what as one list joins the objects", _
                     "Quiz what ""Bob"" can-cover as one list.", _
                     """DATALOG"" ""(headless) (rule (vla-list-can--cover What) (textjoin What \""\"", \""\"" (can-cover \""\""Bob\""\"" VlaListed))) (query vla-list-can--cover)"""
    AssertConditions "a set question as one list", _
                     "Quiz which code is revenue as one list.", _
                     """DATALOG"" ""(headless) (rule (vla-list-revenue Code) (textjoin Code \""\"", \""\"" (revenue VlaListed))) (query vla-list-revenue)"""
    AssertConditions "as one list through a closure", _
                     "Quiz who reports-to ""Alice"" directly or not as one list.", _
                     """DATALOG"" ""(rule (vla-any-reports--to X Y) (reports-to X Y)) (rule (vla-any-reports--to X Y) (reports-to X Z) (vla-any-reports--to Z Y)) (headless) (rule (vla-list-reports--to Who) (textjoin Who \""\"", \""\"" (vla-any-reports--to VlaListed \""\""Alice\""\""))) (query vla-list-reports--to)"""
    AssertConditions "each ... that is ... as one list gives a list per member, the empty ones in", _
                     "Quiz which people can-cover each shift that is listed as one list.", _
                     """DATALOG"" ""(rule (vla-list-can--cover Shift People) (listed Shift) (textjoin People \""\"", \""\"" (can-cover VlaListed Shift))) (query vla-list-can--cover)"""
    AssertConditions "are reads as is in an each list", _
                     "Quiz who can-cover each shift that are listed as one list.", _
                     """DATALOG"" ""(rule (vla-list-can--cover Shift Who) (listed Shift) (textjoin Who \""\"", \""\"" (can-cover VlaListed Shift))) (query vla-list-can--cover)"""
    AssertConditionsRefusal "an each list with the same noun twice refuses", _
                            "Quiz which shift can-cover each shift that is listed as one list.", _
                            "asks for shift twice"
    AssertConditionsRefusal "directly or not after the set of an each list refuses", _
                            "Quiz who can-cover each shift that is listed directly or not as one list.", _
                            "Here it follows a set"
    AssertConditionsRefusal "as one list cannot follow a whether", _
                            "Quiz whether ""Bob"" can-cover ""Night"" as one list.", _
                            "It cannot follow a whether question"
    AssertConditionsRefusal "...nor two unknowns", _
                            "Quiz who can-cover which shift as one list.", _
                            "It cannot follow a question with two unknowns"
    AssertConditionsRefusal "...nor how many", _
                            "Quiz how many people can-cover ""Night"" as one list.", _
                            "It cannot follow a how many question"
    AssertConditionsRefusal "...nor alone", _
                            "Quiz who alone can-cover ""Night"" as one list.", _
                            "It cannot follow an alone question"
    AssertConditionsRefusal "...nor a none question", _
                            "Quiz which shift that is listed is not covered as one list.", _
                            "It cannot follow a which ... is not question"
    AssertConditionsRefusal "...nor every", _
                            "Quiz whether every shift that is listed is covered as one list.", _
                            "It cannot follow a whether every question"
    AssertConditionsRefusal "each without as one list is not a shape - the ordinary near miss", _
                            "Quiz who can-cover each shift that is listed.", _
                            "expected a question"

    EnglishResetGrammar
    EnglishAddPhrase "inscribe {r:cell} that {h:clause}", "(debug-print {h})"
    EnglishAddPhrase "pose {q:question} over {rules:range}", "(debug-print {q})"
    ' A relation chosen with otherwise is written in one cell only.
    AssertConditionsRefusal "a first-match relation written again in a later cell refuses", _
                            "Inscribe H2 that a customer has-tier ""Gold"" if Customers lists the customer as Customer and the spend as Spend, and the spend is at least 10000, otherwise ""Bronze"" if Customers lists the customer as Customer." & vbLf & "Inscribe H3 that a customer has-tier ""Platinum"" if Customers lists the customer as Customer and the spend as Spend, and the spend is at least 50000." & vbLf & "Pose which customer has-tier which tier over H2:H3.", _
                            "'has-tier' is chosen with ""otherwise"" in cell H2, and cell H3 writes 'has-tier' again"
    AssertConditionsRefusal "...and in an earlier cell", _
                            "Inscribe H2 that a customer has-tier ""Platinum"" if Customers lists the customer as Customer and the spend as Spend, and the spend is at least 50000." & vbLf & "Inscribe H3 that a customer has-tier ""Gold"" if Customers lists the customer as Customer and the spend as Spend, and the spend is at least 10000, otherwise ""Bronze"" if Customers lists the customer as Customer." & vbLf & "Pose which customer has-tier which tier over H2:H3.", _
                            "'has-tier' is chosen with ""otherwise"" in cell H3, and cell H2 writes 'has-tier' again"
    AssertConditions "a first match alone in its cell translates", _
                     "Inscribe H2 that a customer has-tier ""Gold"" if Customers lists the customer as Customer and the spend as Spend, and the spend is at least 10000, otherwise ""Bronze"" if Customers lists the customer as Customer." & vbLf & "Pose which customer has-tier which tier over H2:H2.", _
                     "(query vla-ask-has-tier)"
    AssertConditionsRefusal "the range lint reaches the set of an each list", _
                            "Inscribe H2 that a person can-cover a shift if Rota lists the person as Name and the shift as Shift." & vbLf & "Inscribe H3 that a shift is listed if Shifts lists the shift as Shift." & vbLf & "Inscribe H5 that a shift is listed if Extra lists the shift as Shift." & vbLf & "Pose who can-cover each shift that is listed as one list over H2:H3.", _
                            "'listed' is also written in cell H5, which H2:H3 leaves out"
End Sub

Private Sub TestPrologKeyedAtoms()
    Dim q As String
    q = Chr$(34)

    Dim headerMap As Object
    Set headerMap = VLA_Runtime.VlaDictNew()
    Dim cols As New Collection
    Dim pName As New Collection: pName.Add "name": pName.Add "Name": cols.Add pName
    Dim pSalary As New Collection: pSalary.Add "salary": pSalary.Add "Salary": cols.Add pSalary
    Dim pDept As New Collection: pDept.Add "dept": pDept.Add "Dept": cols.Add pDept
    VLA_Runtime.VlaDictSet headerMap, "staffing", cols

    ' VLA.Tokenize's own string-literal mode (VLA.bas, Case """"") runs
    ' from an OPENING quote to a matching CLOSING quote, consuming
    ' everything between them - including parens - as ordinary content,
    ' never as list delimiters; it does NOT stop at whitespace the way a
    ' bare "leading marker" reading of LeafText's own single-strip would
    ' suggest. Every quoted atom below therefore needs q on BOTH sides
    ' (q & "alice" & q, never q & "alice" alone) - live-caught the hard
    ' way, first draft: an unclosed string swallowed every remaining
    ' paren in the text, crashing with a raw "unbalanced parentheses"
    ' error uncaught by PrologRun (which, unlike PROLOG(), has no On
    ' Error of its own - the correct, DatalogRun-precedented shape for a
    ' directly-callable pure-test core, not a bug in itself).
    Dim staffingFacts As String
    staffingFacts = "(fact (staffing " & q & "alice" & q & " 90000 " & q & "eng" & q & ")) " & _
                     "(fact (staffing " & q & "bob" & q & " 70000 " & q & "sales" & q & ")) " & _
                     "(fact (staffing " & q & "carol" & q & " 95000 " & q & "eng" & q & ")) "

    ' A fully-keyed query - every column named, one filtered (dept), one
    ' bound to an output variable (name), one left OUT of the atom
    ' entirely (salary) - hand-traced: alice and carol are eng, bob is
    ' sales, so exactly two rows, in fact-authored order.
    Dim clauseDict As Object
    Set clauseDict = VLA_Runtime.VlaDictNew()
    Dim result As Variant
    result = VLA_Prolog.PrologRun( _
        staffingFacts & "(query (staffing (dept " & q & "eng" & q & ") (name Name)))", _
        clauseDict, headerMap)
    Report "prolog.6: a fully-keyed query resolves against the header, in derivation order", _
           ResultRowCount(result) = 3 And ResultCellIs(result, 2, 1, "alice") And ResultCellIs(result, 3, 1, "carol"), _
           "got: " & ResultDescribe(result)

    ' A predicate NOT in headerMap - its own 2-element-list arguments in
    ' a QUERY CONJUNCT (the position DesugarBodyItem actually walks,
    ' unlike a fact, which never goes through desugaring at all) are
    ' ordinary PROLOG compound-term data, never even considered for
    ' keyed-atom desugaring (this item's own header has the real
    ' ambiguity this resolves - "box" here is not a column name, headerMap
    ' doesn't even know "widget" exists) - (box X)/(box Y) unify against
    ' the stored (box 1)/(box 2) exactly as ordinary nested compound
    ' terms, ordinary PROLOG.3-era behavior, unaffected.
    Set clauseDict = VLA_Runtime.VlaDictNew()
    result = VLA_Prolog.PrologRun( _
        "(fact (widget (box 1) (box 2))) (query (widget (box X) (box Y)))", clauseDict, headerMap)
    Report "prolog.6: a non-table-sourced predicate's own list-shaped query arguments are never desugared, ordinary compound terms", _
           ResultRowCount(result) = 2 And ResultCellIs(result, 2, 1, "1") And ResultCellIs(result, 2, 2, "2"), _
           "got: " & ResultDescribe(result)

    ' Refusals - all against the KNOWN table-sourced "staffing". PrologRun
    ' has no On Error of its own (the correct DatalogRun-precedented
    ' shape for a directly-callable core - see this Sub's own top
    ' comment) so a refusal here is a real, raised VBA error, not a
    ' "#PROLOG! ..." return string; tested VLA_Datalog's own established
    ' way (On Error Resume Next / Err.Clear / call / check Err.Number,
    ' every TestDatalog* refusal case's own precedent), not by
    ' string-matching a return value the way PROLOG() callers do.
    Dim raised As Boolean

    Set clauseDict = VLA_Runtime.VlaDictNew()
    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Prolog.PrologRun _
        staffingFacts & "(query (staffing (bogus " & q & "x" & q & ") (name Name)))", clauseDict, headerMap
    Dim errText1 As String
    errText1 = Err.Description
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "prolog.6: an unknown column name is refused, naming the predicate's own real columns", _
           raised And InStr(1, errText1, "isn't a column", vbTextCompare) > 0, "got: raised=" & raised & " """ & errText1 & """"

    Set clauseDict = VLA_Runtime.VlaDictNew()
    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Prolog.PrologRun _
        staffingFacts & "(query (staffing (dept " & q & "eng" & q & ") (dept " & q & "sales" & q & ")))", clauseDict, headerMap
    Dim errText2 As String
    errText2 = Err.Description
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "prolog.6: the same column keyed twice in one atom is refused", _
           raised And InStr(1, errText2, "more than once", vbTextCompare) > 0, "got: raised=" & raised & " """ & errText2 & """"

    Set clauseDict = VLA_Runtime.VlaDictNew()
    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Prolog.PrologRun _
        staffingFacts & "(query (staffing (dept " & q & "eng" & q & ") Name))", clauseDict, headerMap
    Dim errText3 As String
    errText3 = Err.Description
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Set clauseDict = VLA_Runtime.VlaDictNew()
    Report "prolog.6: mixing a keyed argument with a bare positional one against a table-sourced predicate is refused", _
           raised And InStr(1, errText3, "consistently keyed", vbTextCompare) > 0, "got: raised=" & raised & " """ & errText3 & """"

    ' A RULE body composing full keying with real backtracking - the
    ' identical two-row result as the pure query above, now reached
    ' through a rule (proving the anonymous-column machinery, though
    ' unused here since every column is keyed, doesn't interfere with an
    ' ordinary fully-keyed rule body either).
    Set clauseDict = VLA_Runtime.VlaDictNew()
    result = VLA_Prolog.PrologRun( _
        staffingFacts & "(rule (eng-staff Name) (staffing (dept " & q & "eng" & q & ") (name Name))) (query (eng-staff Name))", _
        clauseDict, headerMap)
    Report "prolog.6: a rule body's own fully-keyed atom composes with real backtracking", _
           ResultRowCount(result) = 3 And ResultCellIs(result, 2, 1, "alice") And ResultCellIs(result, 3, 1, "carol"), _
           "got: " & ResultDescribe(result)

    ' DISCRIMINATING proof - desugaring reaches inside not's own Goal.
    ' Built so a desugaring bug that silently failed to fire would flip
    ' the OBSERVABLE outcome, not just look incidentally right: if
    ' (dept "sales")/(name "bob") were NEVER desugared, they would stay
    ' 2-element-list ARGUMENTS to staffing/3 (arity 2, not 3) - since
    ' staffingFacts already registered staffing's own real arity (3) via
    ' its own (fact ...) forms parsed just above, RecordArity's own
    ' cross-definition check would refuse this AS A PARSE-TIME ARITY
    ' MISMATCH, not silently succeed with a different truth value - a
    ' totally different, unmissable observable failure mode either way.
    Set clauseDict = VLA_Runtime.VlaDictNew()
    result = VLA_Prolog.PrologRun( _
        staffingFacts & "(query (not (staffing (dept " & q & "sales" & q & ") (name " & q & "bob" & q & "))))", _
        clauseDict, headerMap)
    Report "prolog.6: keyed-atom desugaring reaches inside not's own Goal (a real bob/sales row makes the negation FALSE, not vacuously TRUE)", _
           ResultBoolIs(result, False), "got " & TypeName(result) & " " & result

    ' DISCRIMINATING proof - desugaring reaches inside findall's own
    ' Goal, identical reasoning: an undesugared 2-argument keyed atom
    ' against the arity-3 staffing predicate would be refused as a
    ' parse-time arity mismatch, not silently produce an empty bag.
    ' Live-caught: the harvested Bag's own elements are QUOTED atoms
    ' (staffingFacts now properly quotes "alice"/"carol", per this Sub's
    ' own quoting fix above) - RenderBoundValue's own compound-value
    ' branch (VLA.VlaWriteForm/WriteDatum, VLA.bas) deliberately
    ' RE-QUOTES a marked leaf to produce valid, round-trippable Prolog
    ' source text, unlike the bare-leaf branch (LeafText) a lone SCALAR
    ' output column strips down to plain text - this Sub's own first
    ' draft wrongly expected the bare form here, confirmed and fixed
    ' after a real live run, not assumed.
    Set clauseDict = VLA_Runtime.VlaDictNew()
    result = VLA_Prolog.PrologRun( _
        staffingFacts & "(query (findall Name (staffing (dept " & q & "eng" & q & ") (name Name)) Bag))", _
        clauseDict, headerMap)
    ' The bag spelling has now moved TWICE - `("alice" "carol")` before
    ' PROLOG.13, the cons chain after it, and `(list "alice" "carol")`
    ' after PROLOG.21 - and the QUOTING point this pin exists to make has
    ' survived all three unchanged, which is exactly why it is still
    ' built from `q` rather than written as plain text. The harvested
    ' elements are marked atoms; VLA.VlaWriteForm/WriteDatum RE-QUOTES a
    ' marked leaf inside a compound, and PROLOG.21's contraction is a
    ' term-to-term rewrite handed to that same writer, so it cannot have
    ' changed the quoting even by accident. This pin is what proves that.
    Report "prolog.6/13/21: keyed-atom desugaring reaches inside findall's own Goal", _
           ResultCol1Is(result, "(list " & q & "alice" & q & " " & q & "carol" & q & ")"), "got: " & ResultDescribe(result)
End Sub

' PROLOG.6: table-sourced facts, necessarily host-required (RangeToRows/
' RangeColumnNames both need a real Range) - TestPrologKeyedAtoms, above,
' already exercises the DESUGARING mechanism purely, off a hand-built
' headerMap, DATALOG.3/DATALOG.5's own precedent for the identical split.
' Covers: a live Table's own rows becoming ordinary ground fact clauses,
' asserted by VALUE (not just row count); a numeric column rendering
' locale-invariantly (VBA's own Str$, not CStr, on a live Excel Value2
' Double); a text column starting with a capitalized value queried back
' successfully (the real, load-bearing reason table-sourced text gets
' the SAME quoted-atom marker a program-text string literal already
' does - an undecorated capitalized value would otherwise be misdetected
' as an unbound variable by IsVarAtom); a CAPITALIZED Table name queried
' using that SAME capitalized spelling, live, working end to end (this
' item's own header has the real, pre-existing UnifyTwoWay-position-1
' gap UnifyArgsOnly fixes - hand-traced there as harmless-in-practice
' even before the fix, so this scenario alone does not discriminate the
' fix's own necessity, honestly, not overclaimed); and named-column
' syntax reached through the REAL =PROLOG(...) worksheet function
' end-to-end, not just PrologRun directly.
Private Sub TestPrologHostTable()
    Dim prior As Worksheet
    Set prior = ActiveSheet

    VlaEnsureSheet "VlaPrologHostSheet"
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets("VlaPrologHostSheet")
    ws.Activate
    Do While ws.ListObjects.Count > 0
        ws.ListObjects(1).Delete
    Loop
    ws.Cells.Clear

    ' A capitalized Table name AND a capitalized text column - both the
    ' routine, everyday shape this item's own header names, not a
    ' contrived edge case.
    ws.Range("A1:C1").Value = Array("Name", "Salary", "Dept")
    ws.Range("A2:C2").Value = Array("Alice", 90000, "eng")
    ws.Range("A3:C3").Value = Array("Bob", 70000, "sales")
    Dim loEmployees As ListObject
    Set loEmployees = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:C3"), , xlYes)
    loEmployees.Name = "EmployeesHostTest1"

    ' Plain positional reference, queried with the Table's own folded
    ' lowercase name (matching the stored fact head's own text exactly).
    Dim arr1 As Variant
    arr1 = VLA_Prolog.PROLOG( _
        "(query (employeeshosttest1 Name Salary Dept))", loEmployees.Range)
    Dim ok1 As Boolean, detail1 As String
    If IsArray(arr1) Then
        ok1 = (UBound(arr1, 1) = 3)
        detail1 = "got " & (UBound(arr1, 1) - 1) & " rows"
    Else
        detail1 = "got " & TypeName(arr1) & " = " & CStr(arr1)
    End If
    Report "prolog host: a live Table's own rows become ordinary ground fact clauses (2 rows)", _
           ok1, detail1

    ' A numeric column rendered locale-invariantly, and a capitalized
    ' text value queried back successfully - both asserted by VALUE, not
    ' inferred from a row count.
    Dim arr2 As Variant
    arr2 = VLA_Prolog.PROLOG( _
        "(query (employeeshosttest1 " & Chr$(34) & "Alice" & Chr$(34) & " S D))", loEmployees.Range)
    Dim ok2 As Boolean, detail2 As String
    If IsArray(arr2) Then
        ok2 = (ResultRowCount(arr2) = 2 And ResultCellIs(arr2, 2, 1, "90000") And ResultCellIs(arr2, 2, 2, "eng"))
        detail2 = "got S=" & ResultCellText(arr2, 2, 1) & " D=" & ResultCellText(arr2, 2, 2)
    Else
        detail2 = "got " & TypeName(arr2) & " = " & CStr(arr2)
    End If
    Report "prolog host: a numeric cell renders locale-invariantly (Str$, not CStr) and a capitalized text cell round-trips as ground data", _
           ok2, detail2

    ' A capitalized query written in the Table's OWN spelling
    ' ("Employeeshosttest1", genuinely different TEXT from the stored
    ' fact head's own folded "employeeshosttest1") still succeeds end to
    ' end. NOT a discriminating regression pin for UnifyArgsOnly itself -
    ' hand-traced honestly, not overclaimed: this codebase's own PROLOG.6
    ' header already confirms the OLD position-1-included unification
    ' would ALSO have succeeded here, via a harmless stray env binding
    ' (the "same still-free variable" fast path never fires since the
    ' two sides' text differs, so it falls to ordinary variable-binding,
    ' which still lets the REST of the term unify normally) - this test
    ' proves the CAPITALIZED-NAME SCENARIO WORKS, live, not that
    ' UnifyArgsOnly was the only way to make it work.
    Dim arr3 As Variant
    arr3 = VLA_Prolog.PROLOG( _
        "(query (Employeeshosttest1 Name S D))", loEmployees.Range)
    Dim ok3 As Boolean, detail3 As String
    If IsArray(arr3) Then
        ok3 = (UBound(arr3, 1) = 3)
        detail3 = "got " & (UBound(arr3, 1) - 1) & " rows"
    Else
        detail3 = "got " & TypeName(arr3) & " = " & CStr(arr3)
    End If
    Report "prolog host: a query written in the Table's OWN capitalized spelling still matches the internally-folded fact head (UnifyArgsOnly)", _
           ok3, detail3

    ' Named-column syntax through the REAL =PROLOG(...) entry point,
    ' end-to-end - PrologRun/TestPrologKeyedAtoms already proved the
    ' desugaring mechanism itself pure; this proves headerMap actually
    ' gets built and threaded correctly from a live Table's own header
    ' row, through the real worksheet function, not just PrologRun
    ' called directly.
    Dim arr4 As Variant
    arr4 = VLA_Prolog.PROLOG( _
        "(query (employeeshosttest1 (dept " & Chr$(34) & "eng" & Chr$(34) & ") (name Name)))", loEmployees.Range)
    Dim ok4 As Boolean, detail4 As String
    If IsArray(arr4) Then
        ok4 = (ResultRowCount(arr4) = 2 And ResultCellIs(arr4, 2, 1, "Alice"))
        detail4 = "got: " & ResultDescribe(arr4)
    Else
        detail4 = "got " & TypeName(arr4) & " = " & CStr(arr4)
    End If
    Report "prolog host: named-column (keyed-atom) syntax works end-to-end through the real =PROLOG(...) function against a live Table", _
           ok4, detail4

    ' ---- PROLOG.23: A TABLE NAMED LIKE A RESERVED WORD, in its own case.
    ' PROLOG() checks a table argument's name against the reserved set, and
    ' that check matches case-sensitively - so it is exact only because
    ' VLA_Relation.TableArgResolve hands back the name already FOLDED.
    ' PROLOG.19 filed the opposite as a defect (a Table named Between
    ' loading and then being bypassed by the arithmetic `between`, which
    ' would answer 1, 2, 3 here); the fold was there all along. These two
    ' pins are what keep it there: with the fold removed, Between loads and
    ' this query spills 1, 2, 3, and Write loads and gets the output
    ' refusal instead of this one.
    ws.Range("E1:F1").Value = Array("Low", "High")
    ws.Range("E2:F2").Value = Array(1, 3)
    Dim loBetween As ListObject
    Set loBetween = ws.ListObjects.Add(xlSrcRange, ws.Range("E1:F2"), , xlYes)
    loBetween.Name = "Between"
    Dim arr5 As Variant, r5 As String
    arr5 = VLA_Prolog.PROLOG("(query (between 1 3 X))", loBetween.Range)
    r5 = ResultDescribe(arr5)
    Report "prolog.23 host: a Table named Between is refused as a reserved word - never loaded and then bypassed by the arithmetic between", _
           ResultTextStartsWith(arr5, "#PROLOG!") And InStr(1, r5, "'between' is a reserved word", vbTextCompare) > 0, "got: " & r5

    ws.Range("H1").Value = "Line"
    ws.Range("H2").Value = "hello"
    Dim loWrite As ListObject
    Set loWrite = ws.ListObjects.Add(xlSrcRange, ws.Range("H1:H2"), , xlYes)
    loWrite.Name = "Write"
    Dim arr6 As Variant, r6 As String
    arr6 = VLA_Prolog.PROLOG("(query (write X))", loWrite.Range)
    r6 = ResultDescribe(arr6)
    Report "prolog.23 host: ...and so is a Table named Write, by the reserved-word refusal rather than the output one", _
           ResultTextStartsWith(arr6, "#PROLOG!") And InStr(1, r6, "'write' is a reserved word", vbTextCompare) > 0 _
           And InStr(1, r6, "nowhere to print", vbTextCompare) = 0, "got: " & r6

    ' ---- PROLOG.22: A TABLE WITH NO ROWS IS DEFINED, NOT UNKNOWN. Its one
    ' data row is blank, and RangeToRows drops a blank row, so it passes no
    ' facts at all. PROLOG() used to create a predicate's key per ROW, so
    ' this table left none, and an unknown-predicate check would have called
    ' a predicate the author passed in "not defined". Asserted twice: the
    ' query spills a header with nothing under it, and a negation over it is
    ' a correct TRUE - the table really is empty.
    ws.Range("J1:K1").Value = Array("Item", "Qty")
    Dim loEmpty As ListObject
    Set loEmpty = ws.ListObjects.Add(xlSrcRange, ws.Range("J1:K2"), , xlYes)
    loEmpty.Name = "EmptyHostTest1"
    Dim arr7 As Variant
    arr7 = VLA_Prolog.PROLOG("(query (emptyhosttest1 Item Qty))", loEmpty.Range)
    Report "prolog.22 host: a Table with no data rows is DEFINED - its query spills a header and nothing under it, never 'not defined'", _
           ResultRowCount(arr7) = 1 And ResultColCount(arr7) = 2, "got: " & ResultDescribe(arr7)
    Dim arr8 As Variant
    arr8 = VLA_Prolog.PROLOG("(query (not (emptyhosttest1 X Y)))", loEmpty.Range)
    Report "prolog.22 host: ...and (not ...) over that empty Table is a correct TRUE", _
           ResultBoolIs(arr8, True), "got: " & ResultDescribe(arr8)

    ' Owner-requested cleanup: unlike TestDatalogHostTable/TestSqlHostTable
    ' (which both leave their own scratch sheet behind for reuse/post-
    ' crash inspection across runs - VlaDatalogHostSheet/VlaSqlHostSheet,
    ' an existing convention this item doesn't touch), this scratch
    ' sheet is deleted outright once the suite finishes, so a normal
    ' TestDSLs run never leaves VlaPrologHostSheet sitting in the
    ' workbook. Activate prior FIRST, then delete - never delete the
    ' still-active sheet. DisplayAlerts suppresses Excel's own "data may
    ' exist" confirmation prompt, restored immediately after.
    prior.Activate
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
End Sub

' Safely checks whether a spilled-array PROLOG result's own row 2,
' column 1 equals expected - IsArray checked via a nested If, never
' combined into one "IsArray(result) And CStr(result(2,1))=..."
' expression. VBA's And does not short-circuit, so result(2,1) would
' still be touched - and still crash with a raw type mismatch - on a
' non-array result (a refusal string, or a boolean degenerate query)
' even though the IsArray check "should" have short-circuited it away.
' This is the exact class of bug this session's own ancestor-test crash
' already taught the hard way, live, in TestPrologRules - reused here as
' a shared helper specifically so it never has to be re-derived (or
' re-broken) inline at each new call site.
' PROLOG.8: a row count, a column count and a single-cell compare that are
' all SAFE on a result that is not an array at all.
'
' VBA's And does not short-circuit - this module's own long-standing trap -
' so an assertion written "IsArray(result) And UBound(result, 1) = 2"
' evaluates UBound on whatever result actually IS. When a PROLOG query
' returns a Boolean where a spill was expected - which is exactly what a
' missing or broken dispatch arm produces, since an unknown predicate is a
' silent dead end - that raises a type mismatch INSIDE the condition, and
' the run dies at the moment it was about to report the failure. These
' three answer -1 or False instead, so a wrong expectation stays a legible
' failed assertion with its own detail string intact.
Private Function ResultRowCount(ByVal result As Variant) As Long
    ResultRowCount = -1
    If IsArray(result) Then ResultRowCount = UBound(result, 1)
End Function

Private Function ResultColCount(ByVal result As Variant) As Long
    ResultColCount = -1
    If IsArray(result) Then ResultColCount = UBound(result, 2)
End Function

Private Function ResultCellIs(ByVal result As Variant, ByVal rowIx As Long, ByVal colIx As Long, ByVal expected As String) As Boolean
    If Not IsArray(result) Then Exit Function
    If rowIx < LBound(result, 1) Then Exit Function
    If rowIx > UBound(result, 1) Then Exit Function
    If colIx < LBound(result, 2) Then Exit Function
    If colIx > UBound(result, 2) Then Exit Function
    ResultCellIs = (CStr(result(rowIx, colIx)) = expected)
End Function

' PROLOG.15: the guarded way to assert a BARE BOOLEAN result, the shape a
' query with no free variables collapses to.
'
' The suite's long-standing spelling is `VarType(result) = vbBoolean And
' result = True`, and it is a trap of exactly the kind ResultCol1Is
' (below) documents: VBA's `And` does not short-circuit, so when the
' engine returns an ARRAY - which is precisely what a broken type test
' would do, by leaking a phantom column - `result = True` is evaluated
' anyway and raises a type mismatch. A FAILING assertion written that way
' KILLS the run instead of reporting, so the one test that could have
' explained the defect is the one that destroys the evidence.
'
' PROLOG.20 measured 37 assertions already carrying that shape and is the
' item that fixes them; this function is what stops the count reaching
' 38, and is what PROLOG.20 should move them onto. Guarded by early Exit
' rather than by a combined expression, for the reason above.
Private Function ResultBoolIs(ByVal result As Variant, ByVal expected As Boolean) As Boolean
    If IsArray(result) Then Exit Function
    If VarType(result) <> vbBoolean Then Exit Function
    ResultBoolIs = (CBool(result) = expected)
End Function

' PROLOG.20 follow-up: a cell compared NUMERICALLY rather than as text.
'
' Not folded into ResultCellIs, deliberately. The SQL spill tests assert
' `CDbl(arr(r, c)) = 185000`, and rewriting that as a text comparison
' would be a TIGHTENING, not a translation: CDbl accepts "185000.0" and a
' padded " 185000 " where CStr does not, so a passing test could start
' failing for a reason that has nothing to do with the defect being
' fixed. Worse, CStr follows the machine locale on a fractional value -
' the exact hazard NumberToTerm's own Str$-not-CStr note records - so a
' text rewrite would quietly make these tests locale-dependent.
'
' IsNumeric before CDbl: CDbl of a non-numeric raises, which is the whole
' class of thing this family exists to stop doing.
Private Function ResultCellNumIs(ByVal result As Variant, ByVal rowIx As Long, ByVal colIx As Long, ByVal expected As Double) As Boolean
    If Not IsArray(result) Then Exit Function
    If rowIx < LBound(result, 1) Then Exit Function
    If rowIx > UBound(result, 1) Then Exit Function
    If colIx < LBound(result, 2) Then Exit Function
    If colIx > UBound(result, 2) Then Exit Function
    If IsObject(result(rowIx, colIx)) Then Exit Function
    If Not IsNumeric(result(rowIx, colIx)) Then Exit Function
    ResultCellNumIs = (CDbl(result(rowIx, colIx)) = expected)
End Function

' PROLOG.20 follow-up: one cell as DISPLAY TEXT, for a Report's detail.
'
' A detail is evaluated on EVERY call, pass or fail, so a raw index there
' kills the run on exactly the failing case the assertion beside it was
' just made safe for. Says what went wrong instead of raising, because a
' detail string that crashes destroys the run that was about to explain
' itself.
Private Function ResultCellText(ByVal result As Variant, ByVal rowIx As Long, ByVal colIx As Long) As String
    If Not IsArray(result) Then ResultCellText = "<not an array>": Exit Function
    If rowIx < LBound(result, 1) Or rowIx > UBound(result, 1) Then ResultCellText = "<row out of range>": Exit Function
    If colIx < LBound(result, 2) Or colIx > UBound(result, 2) Then ResultCellText = "<col out of range>": Exit Function
    If IsObject(result(rowIx, colIx)) Then ResultCellText = "<object>": Exit Function
    ResultCellText = CStr(result(rowIx, colIx))
End Function

' PROLOG.20: the same guarded shape for a result that is TEXT rather than
' a Boolean or an array - a "#SQL!"/"#PROLOG!" refusal, or DATALOG's own
' blank scalar for a zero-row headless result.
'
' The VarType test is kept rather than simplified to "not an array", and
' that is not tidiness: `CStr(Empty)` is "" in VBA, so a version that only
' excluded arrays would answer True for an EMPTY result where the
' expression it replaces answers False. Caught by the equivalence
' transliteration, which pins the Empty row for exactly this reason.
Private Function ResultTextIs(ByVal result As Variant, ByVal expected As String) As Boolean
    If IsArray(result) Then Exit Function
    If VarType(result) <> vbString Then Exit Function
    ResultTextIs = (CStr(result) = expected)
End Function

' PROLOG.20: the prefix form, for the refusals whose text continues past
' the marker. Len(prefix) rather than a hard-coded 5, so the assertion
' cannot drift from the marker it is checking for.
Private Function ResultTextStartsWith(ByVal result As Variant, ByVal prefix As String) As Boolean
    If IsArray(result) Then Exit Function
    If VarType(result) <> vbString Then Exit Function
    If Len(prefix) = 0 Then Exit Function
    ResultTextStartsWith = (Left$(CStr(result), Len(prefix)) = prefix)
End Function

' OPTIMIZE.1: the "somewhere inside" twin of ResultTextStartsWith, for
' a refusal that must both carry the engine's prefix AND name the right
' function. Written as a helper rather than as
' `ResultTextStartsWith(v, "#OPTIMIZE!") And InStr(1, CStr(v), ...)`,
' which is the exact defect tools/check_test_assertion_safety.ps1
' exists to catch and DID catch on this assertion's first draft: the
' first operand admits v may not be text, and CStr of an array raises
' 13 - on precisely the run where the refusal failed to happen and the
' report was the only thing left worth having.
Private Function ResultTextHas(ByVal result As Variant, ByVal needle As String) As Boolean
    If IsArray(result) Then Exit Function
    If VarType(result) <> vbString Then Exit Function
    If Len(needle) = 0 Then Exit Function
    ResultTextHas = (InStr(1, CStr(result), needle, vbTextCompare) > 0)
End Function


' PROLOG.20: the Collection twin of ResultCellIs. `bn.Count = 1 And
' CStr(bn.Item(1)) = "x"` is the array defect in a different container -
' Item(2) on a one-element Collection raises "Subscript out of range" on
' precisely the run where the count was wrong, which is the run that had
' something to report.
'
' IsObject checked before CStr: an element may be a nested Collection,
' and CStr of an object raises a type mismatch of its own - this
' project's own recorded trap, and the reason this is not one combined
' expression either.
Private Function CollItemIs(ByVal c As Collection, ByVal ix As Long, ByVal expected As String) As Boolean
    If c Is Nothing Then Exit Function
    If ix < 1 Then Exit Function
    If ix > c.Count Then Exit Function
    If IsObject(c.Item(ix)) Then Exit Function
    CollItemIs = (CStr(c.Item(ix)) = expected)
End Function

Private Function ResultCol1Is(ByVal result As Variant, ByVal expected As String) As Boolean
    ' PROLOG.8: the row count is checked before row 2 is touched. A query
    ' with free variables but NO solutions spills a HEADER-ONLY array
    ' (BuildSpilledArray's own ReDim to nRows + 1, which is 1 when nRows
    ' is 0) - the one result shape no test had produced until this item's
    ' own `(= X 1) (\== X 1)` case, and one that made this line raise
    ' "Subscript out of range" mid-run rather than fail an assertion.
    If IsArray(result) Then
        If UBound(result, 1) >= 2 Then
            ResultCol1Is = (CStr(result(2, 1)) = expected)
        End If
    End If
End Function

' A Report detail string safe to build regardless of what PROLOG()
' actually returned - a spilled array (summarized by its own row 2,
' column 1, not the whole array: "text" & anArray raises its own type
' mismatch, distinct from indexing one element of it) or a scalar
' (Boolean/String, concatenates directly). Never assume which shape a
' result is before describing it - the exact assumption this Sub's own
' new tests almost got wrong twice while being written.
Private Function ResultDescribe(ByVal result As Variant) As String
    If IsArray(result) Then
        ' PROLOG.8: the header-only shape (free variables, zero solutions)
        ' is described rather than indexed into - see ResultCol1Is above
        ' for the same guard and the case that found it. A DETAIL string
        ' that crashes is worse than one that says little: it destroys the
        ' run that was about to tell the owner what actually broke.
        If UBound(result, 1) < 2 Then
            ResultDescribe = "array, header row only (" & UBound(result, 2) & " column(s), 0 solutions)"
        Else
            ResultDescribe = "array, " & UBound(result, 1) - 1 & " row(s), row2col1=" & CStr(result(2, 1))
        End If
    Else
        ResultDescribe = TypeName(result) & " " & CStr(result)
    End If
End Function

' Column c of a spilled-array PROLOG/DATALOG/SQL result, values joined
' with ", " - a Report detail helper, not production code (BuildSpilledArray
' itself has no such helper; nothing else in this codebase needed one).
Private Function JoinColumn(ByVal arr As Variant, ByVal c As Long) As String
    Dim r As Long, s As String
    For r = 2 To UBound(arr, 1)
        If r > 2 Then s = s & ", "
        s = s & CStr(arr(r, c))
    Next r
    JoinColumn = s
End Function

' SQL.1: the parser/evaluator's own pure test suite - a hand-built
' columnNames/rows pair (VLA_Relation.RangeToRows' own shape), calling
' VLA_Sql.SqlRun directly, no live workbook needed at all. The fact
' table itself: (Name, Dept, Salary) - alice/eng/90000, bob/sales/60000,
' carol/eng/95000, and alice/eng/90000 AGAIN, verbatim - a genuine
' duplicate row, deliberately, since RangeToRows (unlike RelFromRange)
' must never silently collapse it; every count below that includes
' alice counts her ONCE per matching row, twice where both her rows
' match.
Private Sub TestSql()
    Dim colNames As New Collection
    colNames.Add SqlColPair("name", "Name")
    colNames.Add SqlColPair("dept", "Dept")
    colNames.Add SqlColPair("salary", "Salary")

    Dim rows As New Collection
    Dim r1(1 To 3) As Variant: r1(1) = "alice": r1(2) = "eng": r1(3) = 90000: rows.Add r1
    Dim r2(1 To 3) As Variant: r2(1) = "bob": r2(2) = "sales": r2(3) = 60000: rows.Add r2
    Dim r3(1 To 3) As Variant: r3(1) = "carol": r3(2) = "eng": r3(3) = 95000: rows.Add r3
    Dim r4(1 To 3) As Variant: r4(1) = "alice": r4(2) = "eng": r4(3) = 90000: rows.Add r4

    Dim result As Collection

    Set result = VLA_Sql.SqlRun("SELECT * FROM staff", "staff", colNames, rows)
    Report "sql: SELECT * preserves a duplicate source row (bag, not set - 4 rows)", _
           CountRows(result) = 4, "got " & CountRows(result)
    Report "sql: SELECT * headers use the table's own original casing", _
           HeaderAt(result, 1) = "Name" And HeaderAt(result, 2) = "Dept" And HeaderAt(result, 3) = "Salary", _
           "got " & HeaderAt(result, 1) & "/" & HeaderAt(result, 2) & "/" & HeaderAt(result, 3)

    Set result = VLA_Sql.SqlRun("SELECT Name, Salary FROM staff WHERE Salary > 80000", "staff", colNames, rows)
    Report "sql: WHERE numeric > filters correctly, alice's duplicate both survive (3 rows)", _
           CountRows(result) = 3, "got " & CountRows(result)
    Report "sql: explicit SELECT list headers use the QUERY's own typed casing", _
           HeaderAt(result, 1) = "Name" And HeaderAt(result, 2) = "Salary", _
           "got " & HeaderAt(result, 1) & "/" & HeaderAt(result, 2)

    Set result = VLA_Sql.SqlRun("SELECT * FROM staff WHERE dept = 'eng' AND salary >= 90000", "staff", colNames, rows)
    Report "sql: AND over a text = and a numeric >= (alice x2, carol - 3 rows)", _
           CountRows(result) = 3, "got " & CountRows(result)

    Set result = VLA_Sql.SqlRun("SELECT * FROM staff WHERE NOT (dept = 'eng')", "staff", colNames, rows)
    Report "sql: NOT over a parenthesized comparison (bob only)", _
           CountRows(result) = 1, "got " & CountRows(result)

    Set result = VLA_Sql.SqlRun("SELECT * FROM staff WHERE dept = 'eng' OR dept = 'sales'", "staff", colNames, rows)
    Report "sql: OR covering every row (4 rows)", CountRows(result) = 4, "got " & CountRows(result)

    ' Precedence: AND binds tighter than OR, unparenthesized - this must
    ' parse as (dept='sales') OR (dept='eng' AND salary>92000), NOT as
    ' (dept='sales' OR dept='eng') AND salary>92000 - real SQL's own
    ' rule, and the entire reason a proper expression tree was built
    ' instead of a string-FIND heuristic (BETA_ROADMAP1.md's own
    ' correction, cited in this engine's own module header).
    Set result = VLA_Sql.SqlRun( _
        "SELECT * FROM staff WHERE dept = 'sales' OR dept = 'eng' AND salary > 92000", _
        "staff", colNames, rows)
    Report "sql: AND binds tighter than OR, unparenthesized (bob + carol only, not alice - 2 rows)", _
           CountRows(result) = 2, "got " & CountRows(result)

    ' The SAME condition, explicitly parenthesized the OTHER way -
    ' proves parens actually override precedence, not just that the
    ' default happened to look right above.
    Set result = VLA_Sql.SqlRun( _
        "SELECT * FROM staff WHERE (dept = 'sales' OR dept = 'eng') AND salary > 92000", _
        "staff", colNames, rows)
    Report "sql: parentheses override default precedence (carol only - 1 row)", _
           CountRows(result) = 1, "got " & CountRows(result)

    Dim raised As Boolean

    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Sql.SqlRun "SELECT * FROM wrongname", "staff", colNames, rows
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "sql: FROM naming a different table than the one actually passed is refused", raised, "no error raised"

    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Sql.SqlRun "SELECT nosuchcolumn FROM staff", "staff", colNames, rows
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "sql: an unknown SELECT column is refused", raised, "no error raised"

    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Sql.SqlRun "SELECT * FROM staff WHERE nosuchcolumn = 1", "staff", colNames, rows
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "sql: an unknown WHERE column is refused", raised, "no error raised"

    ' SQL.6 shipped UNION/UNION ALL/INTERSECT/EXCEPT - but only through
    ' the NEW SqlRunCompound entry point (VLA_Sql.SQL(...)'s own
    ' worksheet UDF now calls that instead). SqlRun/SqlRunJoin, this
    ' Sub's own subject, deliberately stay single-query-only FOREVER
    ' (VLA_Sql.bas's own SQL.6 header note: "this function ... stays
    ' single-query-only forever") - so a compound keyword reaching THIS
    ' entry point is still, permanently, an unsupported-keyword refusal,
    ' not a stale placeholder now that SQL.6 exists. TestSqlSetOps (this
    ' file, below) is where UNION/etc. actually get exercised.
    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Sql.SqlRun "SELECT * FROM staff UNION SELECT * FROM staff", "staff", colNames, rows
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "sql: SqlRun/SqlRunJoin stay single-query-only - a compound keyword (UNION) reaching them directly is still refused", raised, "no error raised"

    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Sql.SqlRun "SELECT * FROM staff WHERE name = 'alice", "staff", colNames, rows
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "sql: an unterminated string literal is refused", raised, "no error raised"

    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Sql.SqlRun "SELECT * FROM staff WHERE name = alice!", "staff", colNames, rows
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "sql: an unrecognized character ('!') is refused", raised, "no error raised"

    ' The dialect-pin audit's own catch: a query actually copied from a
    ' real SQL tool commonly carries a trailing ';' and/or comments -
    ' both real SQLite syntax, both previously unhandled (would have
    ' refused a perfectly valid pasted query with sql-unexpected-
    ' character, directly undermining the whole "paste it in" pitch a
    ' real-SQL-text grammar surface exists for).
    Set result = VLA_Sql.SqlRun("SELECT * FROM staff;", "staff", colNames, rows)
    Report "sql: a single trailing ';' (the ordinary pasted-statement shape) is tolerated", _
           CountRows(result) = 4, "got " & CountRows(result)

    Set result = VLA_Sql.SqlRun( _
        "SELECT * FROM staff -- only engineers, please" & vbLf & "WHERE dept = 'eng'", _
        "staff", colNames, rows)
    Report "sql: a '--' line comment is skipped like whitespace (3 eng rows)", _
           CountRows(result) = 3, "got " & CountRows(result)

    Set result = VLA_Sql.SqlRun( _
        "SELECT * /* every column */ FROM staff WHERE /* filter */ dept = 'sales'", _
        "staff", colNames, rows)
    Report "sql: a /* block comment */ is skipped like whitespace (bob only)", _
           CountRows(result) = 1, "got " & CountRows(result)

    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Sql.SqlRun "SELECT * FROM staff /* never closed", "staff", colNames, rows
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "sql: an unterminated block comment is refused, not silently extended to end-of-query", raised, "no error raised"

    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Sql.SqlRun "SELECT * FROM staff; SELECT * FROM staff", "staff", colNames, rows
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "sql: a SECOND statement after the trailing ';' is refused, not silently truncated", raised, "no error raised"

    ' SQL.2: computed SELECT expressions, AS aliases, DISTINCT.
    Set result = VLA_Sql.SqlRun("SELECT Name, Salary * 1.1 AS raised FROM staff", "staff", colNames, rows)
    Report "sql.2: a computed column (Salary * 1.1) keeps the bag's own 4 rows", _
           CountRows(result) = 4, "got " & CountRows(result)
    Report "sql.2: the computed column's own header is its AS alias", _
           HeaderAt(result, 2) = "raised", "got """ & HeaderAt(result, 2) & """"
    ' Tolerance, not exact "= 99000" - 1.1 has no exact binary
    ' representation, so 90000 * 1.1 lands a hair off 99000.0 in IEEE
    ' double precision (CStr/Debug.Print round it back to "99000" for
    ' display, which is why a failing report here would misleadingly
    ' show "got 99000" - live-caught exactly this way). Every OTHER
    ' computed-value check in this Sub uses integer-valued arithmetic
    ' (exactly representable in binary), so this is the only one that
    ' needs it.
    Report "sql.2: alice's own 90000 * 1.1 = 99000", _
           Abs(CDbl(ValueAt(result, 1, 2)) - 99000) < 0.0001, "got " & ValueAt(result, 1, 2)

    Set result = VLA_Sql.SqlRun("SELECT Name AS full_name FROM staff", "staff", colNames, rows)
    Report "sql.2: AS also renames a bare column reference", _
           HeaderAt(result, 1) = "full_name", "got """ & HeaderAt(result, 1) & """"

    ' =================================================================
    '  PROLOG.17: SCALAR FUNCTIONS. SQL reaches the same shared
    '  substrate PROLOG and DATALOG do, through the same one operator
    '  table - but spells the names its own way, because MIN( and MAX(
    '  were already AGGREGATE keywords here and have been since SQL.4.
    '  Every value is asserted, never merely the row count: a wrong
    '  operator in a SELECT list produces a plausible NUMBER in a cell,
    '  which is the failure this whole item is shaped around.
    '
    '  All values below are integer-valued and therefore exactly
    '  representable in binary, so none needs the tolerance the
    '  Salary * 1.1 case above documents.
    ' =================================================================

    ' 90000 = 7 * 12857 + 1, so both MOD and DIV are pinned off the same
    ' division and a confusion between them cannot pass both.
    Set result = VLA_Sql.SqlRun("SELECT Name, MOD(Salary, 7) AS m FROM staff", "staff", colNames, rows)
    Report "prolog.17/sql: MOD(Salary, 7) computes 1 for alice's 90000", _
           CDbl(ValueAt(result, 1, 2)) = 1, "got " & ValueAt(result, 1, 2)
    Set result = VLA_Sql.SqlRun("SELECT Name, DIV(Salary, 7) AS q FROM staff", "staff", colNames, rows)
    Report "prolog.17/sql: DIV(Salary, 7) computes 12857 - integer division, the twin of the MOD above", _
           CDbl(ValueAt(result, 1, 2)) = 12857, "got " & ValueAt(result, 1, 2)

    ' A UNARY function wrapping a BINARY expression, which is the shape
    ' that proves NK_UNARY nests over ordinary arithmetic rather than
    ' only over a bare column.
    Set result = VLA_Sql.SqlRun("SELECT Name, ABS(Salary - 95000) AS d FROM staff", "staff", colNames, rows)
    Report "prolog.17/sql: ABS(Salary - 95000) computes 5000 for alice - a unary function over a computed operand", _
           CDbl(ValueAt(result, 1, 2)) = 5000, "got " & ValueAt(result, 1, 2)

    ' LEAST/GREATEST are the scalar pair, spelled standard-SQL style
    ' precisely because MIN/MAX are taken. The two rows are twins: an
    ' implementation that mapped both to the same operator fails one.
    Set result = VLA_Sql.SqlRun("SELECT Name, LEAST(Salary, 70000) AS lo FROM staff", "staff", colNames, rows)
    Report "prolog.17/sql: LEAST(Salary, 70000) computes 70000 for alice's 90000", _
           CDbl(ValueAt(result, 1, 2)) = 70000, "got " & ValueAt(result, 1, 2)
    Set result = VLA_Sql.SqlRun("SELECT Name, GREATEST(Salary, 70000) AS hi FROM staff", "staff", colNames, rows)
    Report "prolog.17/sql: GREATEST(Salary, 70000) computes 90000 for the same row", _
           CDbl(ValueAt(result, 1, 2)) = 90000, "got " & ValueAt(result, 1, 2)

    Set result = VLA_Sql.SqlRun("SELECT Name, POWER(2, 10) AS p FROM staff", "staff", colNames, rows)
    Report "prolog.17/sql: POWER(2, 10) computes 1024", _
           CDbl(ValueAt(result, 1, 2)) = 1024, "got " & ValueAt(result, 1, 2)
    Set result = VLA_Sql.SqlRun("SELECT Name, ROUND(2.5) AS r FROM staff", "staff", colNames, rows)
    Report "prolog.17/sql: ROUND(2.5) computes 3 - half away from zero, the same reading PROLOG and DATALOG use", _
           CDbl(ValueAt(result, 1, 2)) = 3, "got " & ValueAt(result, 1, 2)

    ' ---- THE COLLISION THAT DECIDED THE SPELLING. MIN( must still be
    ' the AGGREGATE, over the whole 4-row bag, and not a scalar function
    ' of one argument. Without this, adding a scalar MIN would have
    ' silently changed the meaning of every existing MIN( in a query.
    Set result = VLA_Sql.SqlRun("SELECT MIN(Salary) AS lo FROM staff", "staff", colNames, rows)
    Report "prolog.17/sql: MIN(Salary) is STILL the aggregate - 60000 over the whole bag, not a scalar function", _
           CDbl(ValueAt(result, 1, 1)) = 60000, "got " & ValueAt(result, 1, 1)
    Report "prolog.17/sql: ...and it collapses the bag to one row, which a scalar function would not have done", _
           CountRows(result) = 1, "got " & CountRows(result)

    ' ---- the refusals, each naming what it refused
    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Sql.SqlRun "SELECT Name, SQRT(0 - 1) AS s FROM staff", "staff", colNames, rows
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "prolog.17/sql: SQRT of a negative is refused, not returned as an error value in the cell", _
           raised, "no error raised"

    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Sql.SqlRun "SELECT Name, ABS(1, 2) AS a FROM staff", "staff", colNames, rows
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "prolog.17/sql: a unary function given two arguments is refused by name, not by a bare 'expected )'", _
           raised, "no error raised"

    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Sql.SqlRun "SELECT Name, MOD(Salary) AS m FROM staff", "staff", colNames, rows
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "prolog.17/sql: ...and a binary function given one argument likewise", _
           raised, "no error raised"

    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Sql.SqlRun "SELECT Salary * 1.1 FROM staff", "staff", colNames, rows
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "sql.2: a computed column with no AS alias is refused", raised, "no error raised"

    Set result = VLA_Sql.SqlRun("SELECT DISTINCT Dept FROM staff", "staff", colNames, rows)
    Report "sql.2: DISTINCT on one column collapses eng (x3) and sales (x1) to 2 rows", _
           CountRows(result) = 2, "got " & CountRows(result)

    Set result = VLA_Sql.SqlRun("SELECT DISTINCT * FROM staff", "staff", colNames, rows)
    Report "sql.2: DISTINCT * collapses the one genuine duplicate row (4 -> 3 rows)", _
           CountRows(result) = 3, "got " & CountRows(result)

    Set result = VLA_Sql.SqlRun("SELECT DISTINCT Dept FROM staff WHERE Salary > 90000", "staff", colNames, rows)
    Report "sql.2: WHERE filters BEFORE DISTINCT, not after (only carol survives > 90000 - 1 row)", _
           CountRows(result) = 1, "got " & CountRows(result)

    Set result = VLA_Sql.SqlRun("SELECT Salary + 100 * 2 AS x FROM staff WHERE Name = 'alice'", "staff", colNames, rows)
    Report "sql.2: * binds tighter than + (90000 + 100*2 = 90200, not (90000+100)*2)", _
           ValueAt(result, 1, 1) = 90200, "got " & ValueAt(result, 1, 1)

    Set result = VLA_Sql.SqlRun("SELECT (Salary + 100) * 2 AS x FROM staff WHERE Name = 'alice'", "staff", colNames, rows)
    Report "sql.2: parentheses override precedence in a computed expression ((90000+100)*2 = 180200)", _
           ValueAt(result, 1, 1) = 180200, "got " & ValueAt(result, 1, 1)

    Set result = VLA_Sql.SqlRun("SELECT -Salary AS negated FROM staff WHERE Name = 'bob'", "staff", colNames, rows)
    Report "sql.2: unary minus (-60000)", ValueAt(result, 1, 1) = -60000, "got " & ValueAt(result, 1, 1)

    Set result = VLA_Sql.SqlRun("SELECT Name FROM staff WHERE Salary * 2 > 150000", "staff", colNames, rows)
    Report "sql.2: WHERE reuses the SAME arithmetic grammar with no parens needed (alice x2 + carol - 3 rows)", _
           CountRows(result) = 3, "got " & CountRows(result)

    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Sql.SqlRun "SELECT Salary / 0 AS x FROM staff", "staff", colNames, rows
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "sql.2: division by zero is refused, not a raw runtime crash", raised, "no error raised"

    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Sql.SqlRun "SELECT Dept * 2 AS x FROM staff", "staff", colNames, rows
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "sql.2: arithmetic on a non-numeric (text) column is refused", raised, "no error raised"

    ' The one documented SQL.2 limit: a comparison's own LEFT operand
    ' can never itself start with '(' at the very top of a condition -
    ' ParseNotExpr always claims a leading '(' there as an attempted
    ' boolean group first (this file's own SQL.2 header note has the
    ' full mechanism). Refused cleanly, not silently misparsed.
    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Sql.SqlRun "SELECT * FROM staff WHERE (Salary + 100) > 1000", "staff", colNames, rows
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "sql.2: a parenthesized LEFT comparison operand at the top of WHERE is refused, not silently misparsed", raised, "no error raised"

    ' SQL.4: GROUP BY / aggregates (COUNT/SUM/MIN/MAX/AVG) / HAVING.
    ' Same staff table: alice/eng/90000 (twice, a genuine duplicate row -
    ' bag semantics, unchanged), bob/sales/60000, carol/eng/95000 - so
    ' eng groups 3 rows (275000 total), sales groups 1 (60000).
    Set result = VLA_Sql.SqlRun("SELECT Dept, COUNT(*) AS n FROM staff GROUP BY Dept", "staff", colNames, rows)
    Report "sql.4: GROUP BY collapses 4 rows into 2 groups (eng, sales)", _
           CountRows(result) = 2, "got " & CountRows(result)
    Set result = VLA_Sql.SqlRun("SELECT Dept, COUNT(*) AS n FROM staff WHERE Dept = 'eng' GROUP BY Dept", "staff", colNames, rows)
    Report "sql.4: COUNT(*) counts eng's own 3 rows, alice's duplicate included (bag, not set)", _
           ValueAt(result, 1, 2) = 3, "got " & ValueAt(result, 1, 2)

    Set result = VLA_Sql.SqlRun("SELECT Dept, SUM(Salary) AS total FROM staff WHERE Dept = 'eng' GROUP BY Dept", "staff", colNames, rows)
    Report "sql.4: SUM(Salary) over eng's own 3 rows (90000+95000+90000 = 275000)", _
           ValueAt(result, 1, 2) = 275000, "got " & ValueAt(result, 1, 2)

    Set result = VLA_Sql.SqlRun("SELECT Dept, MIN(Salary) AS lo, MAX(Salary) AS hi FROM staff WHERE Dept = 'eng' GROUP BY Dept", "staff", colNames, rows)
    Report "sql.4: MIN(Salary)/MAX(Salary) over eng's own group (90000/95000)", _
           ValueAt(result, 1, 2) = 90000 And ValueAt(result, 1, 3) = 95000, _
           "got " & ValueAt(result, 1, 2) & "/" & ValueAt(result, 1, 3)

    Set result = VLA_Sql.SqlRun("SELECT Dept, AVG(Salary) AS avgsal FROM staff WHERE Dept = 'eng' GROUP BY Dept", "staff", colNames, rows)
    Report "sql.4: AVG(Salary) over eng's own group (275000/3)", _
           Abs(CDbl(ValueAt(result, 1, 2)) - 275000 / 3) < 0.001, "got " & ValueAt(result, 1, 2)

    Set result = VLA_Sql.SqlRun("SELECT Dept, COUNT(*) AS n FROM staff GROUP BY Dept HAVING COUNT(*) > 1", "staff", colNames, rows)
    Report "sql.4: HAVING filters the GROUPED result (only eng has more than 1 row - 1 group)", _
           CountRows(result) = 1 And ValueAt(result, 1, 1) = "eng", "got " & CountRows(result) & " groups"

    Set result = VLA_Sql.SqlRun("SELECT COUNT(*) FROM staff", "staff", colNames, rows)
    Report "sql.4: an ungrouped aggregate (no GROUP BY at all) counts every row, bag semantics (4)", _
           ValueAt(result, 1, 1) = 4, "got " & ValueAt(result, 1, 1)
    Report "sql.4: an un-aliased aggregate call gets its own canonical default header", _
           HeaderAt(result, 1) = "COUNT(*)", "got """ & HeaderAt(result, 1) & """"

    Set result = VLA_Sql.SqlRun("SELECT COUNT(*) AS n, SUM(Salary) AS s FROM staff WHERE Salary > 999999", "staff", colNames, rows)
    Report "sql.4: an ungrouped aggregate over ZERO matching rows still returns its one-row answer", _
           CountRows(result) = 1, "got " & CountRows(result)
    Report "sql.4: COUNT/SUM over zero matching rows are real, meaningful zeros", _
           ValueAt(result, 1, 1) = 0 And ValueAt(result, 1, 2) = 0, _
           "got " & ValueAt(result, 1, 1) & "/" & ValueAt(result, 1, 2)

    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Sql.SqlRun "SELECT MIN(Salary) AS m FROM staff WHERE Salary > 999999", "staff", colNames, rows
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "sql.4: MIN/MAX/AVG over zero matching rows (no GROUP BY) is refused - no defined answer, unlike COUNT/SUM", raised, "no error raised"

    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Sql.SqlRun "SELECT * FROM staff GROUP BY Dept", "staff", colNames, rows
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "sql.4: SELECT * can't be combined with GROUP BY", raised, "no error raised"

    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Sql.SqlRun "SELECT Name, COUNT(*) AS n FROM staff GROUP BY Dept", "staff", colNames, rows
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "sql.4: a SELECT item that's neither a GROUP BY column nor wrapped in an aggregate is refused", raised, "no error raised"

    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Sql.SqlRun "SELECT Dept FROM staff WHERE COUNT(*) > 1 GROUP BY Dept", "staff", colNames, rows
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "sql.4: an aggregate call inside WHERE is refused (WHERE runs before grouping)", raised, "no error raised"

    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Sql.SqlRun "SELECT COUNT(SUM(Salary)) AS x FROM staff", "staff", colNames, rows
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "sql.4: an aggregate nested inside another aggregate's own operand is refused", raised, "no error raised"

    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Sql.SqlRun "SELECT SUM(*) AS x FROM staff", "staff", colNames, rows
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "sql.4: SUM(*) is refused - only COUNT(*) may use '*' as its own argument", raised, "no error raised"

    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Sql.SqlRun "SELECT Dept FROM staff GROUP BY 5", "staff", colNames, rows
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "sql.4: a GROUP BY item that isn't a plain column (a literal number) is refused", raised, "no error raised"

    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Sql.SqlRun "SELECT Dept FROM staff GROUP BY COUNT(*)", "staff", colNames, rows
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "sql.4: an aggregate call as a GROUP BY item is refused (needs a plain column)", raised, "no error raised"

    ' SQL.5: ORDER BY / LIMIT. Same staff table: alice/eng/90000 (r1),
    ' bob/sales/60000 (r2), carol/eng/95000 (r3), alice/eng/90000 (r4,
    ' a genuine duplicate of r1) - source order r1,r2,r3,r4 throughout,
    ' the stability baseline every tie below is checked against.
    ' ORDER BY tries the OUTPUT column list first (name/alias or 1-based
    ' position), falling back to a real SOURCE/grouped column when a
    ' name matches nothing there (below) - this file's own SQL.5 header
    ' note has the full resolution rule. Name always stays column 1
    ' here to keep ValueAt(result, row, 1) checks uniform throughout.
    Set result = VLA_Sql.SqlRun("SELECT Name, Salary FROM staff ORDER BY Salary", "staff", colNames, rows)
    Report "sql.5: ORDER BY Salary ascending (bob 60000, then both 90000-alices, then carol 95000)", _
           ValueAt(result, 1, 1) = "bob" And ValueAt(result, 2, 1) = "alice" And _
           ValueAt(result, 3, 1) = "alice" And ValueAt(result, 4, 1) = "carol", _
           "got " & ValueAt(result, 1, 1) & "/" & ValueAt(result, 2, 1) & "/" & ValueAt(result, 3, 1) & "/" & ValueAt(result, 4, 1)

    Set result = VLA_Sql.SqlRun("SELECT Name, Salary FROM staff ORDER BY Salary DESC", "staff", colNames, rows)
    Report "sql.5: ORDER BY Salary DESC reverses the ascending order (carol first, bob last)", _
           ValueAt(result, 1, 1) = "carol" And ValueAt(result, 4, 1) = "bob", _
           "got " & ValueAt(result, 1, 1) & "/.../" & ValueAt(result, 4, 1)

    ' The stability check proper: sorted by Dept alone, three rows tie
    ' on 'eng' (r1=alice, r3=carol, r4=alice) - a stable sort MUST keep
    ' them in their own original relative order (alice, carol, alice),
    ' never reordered among themselves just because they're equal keys;
    ' 'sales' (bob) sorts after 'eng' alphabetically, so he lands last.
    Set result = VLA_Sql.SqlRun("SELECT Name, Dept FROM staff ORDER BY Dept", "staff", colNames, rows)
    Report "sql.5: a stable sort preserves each tied group's own original relative order (alice, carol, alice, bob)", _
           ValueAt(result, 1, 1) = "alice" And ValueAt(result, 2, 1) = "carol" And _
           ValueAt(result, 3, 1) = "alice" And ValueAt(result, 4, 1) = "bob", _
           "got " & ValueAt(result, 1, 1) & "/" & ValueAt(result, 2, 1) & "/" & ValueAt(result, 3, 1) & "/" & ValueAt(result, 4, 1)

    ' Multi-key: Dept ASC breaks the eng/sales tie first, Salary DESC
    ' then breaks the tie WITHIN eng (carol's 95000 before either
    ' 90000-alice, who then keep their own stable relative order).
    Set result = VLA_Sql.SqlRun("SELECT Name, Dept, Salary FROM staff ORDER BY Dept ASC, Salary DESC", "staff", colNames, rows)
    Report "sql.5: multi-key ORDER BY (Dept ASC, Salary DESC) - carol, alice, alice, bob", _
           ValueAt(result, 1, 1) = "carol" And ValueAt(result, 2, 1) = "alice" And _
           ValueAt(result, 3, 1) = "alice" And ValueAt(result, 4, 1) = "bob", _
           "got " & ValueAt(result, 1, 1) & "/" & ValueAt(result, 2, 1) & "/" & ValueAt(result, 3, 1) & "/" & ValueAt(result, 4, 1)

    ' ORDER BY falling back to a SOURCE column that was never selected
    ' at all (live-caught missing from this item's own first pass -
    ' output-only resolution alone refused this, stricter than real
    ' SQL/SQLite's own "order by whatever you like" allowance for an
    ' ordinary, non-DISTINCT query) - same expected order as the
    ' multi-key test just above, just with only Name in the output.
    Set result = VLA_Sql.SqlRun("SELECT Name FROM staff ORDER BY Dept ASC, Salary DESC", "staff", colNames, rows)
    Report "sql.5: ORDER BY a column not in the SELECT list at all still resolves (carol, alice, alice, bob)", _
           ValueAt(result, 1, 1) = "carol" And ValueAt(result, 2, 1) = "alice" And _
           ValueAt(result, 3, 1) = "alice" And ValueAt(result, 4, 1) = "bob", _
           "got " & ValueAt(result, 1, 1) & "/" & ValueAt(result, 2, 1) & "/" & ValueAt(result, 3, 1) & "/" & ValueAt(result, 4, 1)

    ' The same fallback, over a GROUPED query - Dept is the GROUP BY key
    ' but isn't itself selected; ORDER BY still resolves it against the
    ' GROUPED colMap (the same repointing SQL.4's own aggregate pipeline
    ' already does for WHERE-turned-HAVING/SELECT), sorting eng (n=3)
    ' before sales (n=1) alphabetically.
    Set result = VLA_Sql.SqlRun("SELECT COUNT(*) AS n FROM staff GROUP BY Dept ORDER BY Dept", "staff", colNames, rows)
    Report "sql.5: ORDER BY a GROUP BY key that isn't itself selected resolves against the grouped row (eng=3 first, sales=1 second)", _
           CountRows(result) = 2 And ValueAt(result, 1, 1) = 3 And ValueAt(result, 2, 1) = 1, _
           "got " & CountRows(result) & " rows: " & ValueAt(result, 1, 1) & "/" & ValueAt(result, 2, 1)

    Set result = VLA_Sql.SqlRun("SELECT Name, Salary FROM staff ORDER BY 2", "staff", colNames, rows)
    Report "sql.5: ORDER BY a 1-based output POSITION (2 = Salary) sorts the same as naming it", _
           ValueAt(result, 1, 1) = "bob" And ValueAt(result, 4, 1) = "carol", _
           "got " & ValueAt(result, 1, 1) & "/.../" & ValueAt(result, 4, 1)

    Set result = VLA_Sql.SqlRun("SELECT Name, Salary FROM staff ORDER BY Salary DESC LIMIT 2", "staff", colNames, rows)
    Report "sql.5: LIMIT after ORDER BY keeps only the top 2 (carol, then the first-seen 90000-alice)", _
           CountRows(result) = 2 And ValueAt(result, 1, 1) = "carol" And ValueAt(result, 2, 1) = "alice", _
           "got " & CountRows(result) & " rows: " & ValueAt(result, 1, 1) & "/" & ValueAt(result, 2, 1)

    Set result = VLA_Sql.SqlRun("SELECT Name FROM staff LIMIT 2", "staff", colNames, rows)
    Report "sql.5: LIMIT with no ORDER BY keeps the first 2 rows in natural source order (alice, bob)", _
           CountRows(result) = 2 And ValueAt(result, 1, 1) = "alice" And ValueAt(result, 2, 1) = "bob", _
           "got " & CountRows(result) & " rows: " & ValueAt(result, 1, 1) & "/" & ValueAt(result, 2, 1)

    Set result = VLA_Sql.SqlRun("SELECT Name FROM staff LIMIT 0", "staff", colNames, rows)
    Report "sql.5: LIMIT 0 is a real, valid zero-row result, not an error", CountRows(result) = 0, "got " & CountRows(result)

    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Sql.SqlRun "SELECT Name FROM staff ORDER BY nosuchcolumn", "staff", colNames, rows
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "sql.5: ORDER BY naming an unknown output column is refused", raised, "no error raised"

    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Sql.SqlRun "SELECT Name FROM staff ORDER BY 5", "staff", colNames, rows
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "sql.5: ORDER BY a position past the SELECT list's own arity is refused", raised, "no error raised"

    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Sql.SqlRun "SELECT Name FROM staff ORDER BY 1.5", "staff", colNames, rows
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "sql.5: a non-integer ORDER BY position is refused", raised, "no error raised"

    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Sql.SqlRun "SELECT Name, Dept AS Name FROM staff ORDER BY Name", "staff", colNames, rows
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "sql.5: an ORDER BY name matching more than one output column is refused, not guessed", raised, "no error raised"

    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Sql.SqlRun "SELECT Name FROM staff LIMIT 1.5", "staff", colNames, rows
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "sql.5: a non-integer LIMIT is refused", raised, "no error raised"

    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Sql.SqlRun "SELECT Name FROM staff LIMIT -1", "staff", colNames, rows
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "sql.5: a negative LIMIT is refused", raised, "no error raised"
End Sub

Private Function CountRows(ByVal result As Collection) As Long
    Dim rowsColl As Collection
    Set rowsColl = result.Item(2)
    CountRows = rowsColl.Count
End Function

Private Function HeaderAt(ByVal result As Collection, ByVal pos As Long) As String
    Dim headerColl As Collection
    Set headerColl = result.Item(1)
    HeaderAt = CStr(headerColl.Item(pos))
End Function

' SQL.2: computed-column tests need an actual VALUE, not just a row
' count or a header - the same 3-item SqlRun shape (headers, rows,
' arity) HeaderAt/CountRows already read.
Private Function ValueAt(ByVal result As Collection, ByVal rowIdx As Long, ByVal colIdx As Long) As Variant
    Dim rowsColl As Collection
    Set rowsColl = result.Item(2)
    Dim arr As Variant
    arr = rowsColl.Item(rowIdx)
    ValueAt = arr(colIdx)
End Function

' Matches VLA_Sql.bas's own Private MakeTableRec shape by hand (Item(1)=
' name, Item(2)=columnNames, Item(3)=rows) - SqlColPair's own precedent
' just above, repeated: a pure test hand-builds the tiny record shape a
' production module's own Private constructor would otherwise build,
' rather than widening that module's Private surface for a second,
' unrelated reader (VLA_Datalog.bas's own Private IsList/Nth header note
' states this house rule explicitly). VLA_Sql.SqlRunJoin is Public
' specifically so a pure test can hand it a Collection of these directly,
' the identical pure-core seam VLA_Datalog.DatalogRun already established.
Private Function SqlTableRec(ByVal name As String, ByVal columnNames As Collection, ByVal rows As Collection) As Collection
    Dim r As New Collection
    r.Add name
    r.Add columnNames
    r.Add rows
    Set SqlTableRec = r
End Function

' SQL.3: INNER JOIN ... ON, entirely pure - VLA_Sql.SqlRunJoin called
' directly with hand-built table records (SqlTableRec), no live workbook
' needed at all (unlike TestSqlHostTable's own column-name requirement,
' below - a JOIN needs no MORE from a real Table than a single-table
' query already did, so the identical hand-built-columnNames pure seam
' TestSql already established for SQL.1/SQL.2 covers this too).
'
' Two small tables joined on Dept: staff (Name, Dept) - alice/eng,
' bob/sales, carol/eng - and deptinfo (Dept, Building) - eng/B5,
' sales/B7 - proving the basic equi-join, qualified AND unqualified
' column resolution, SELECT *, WHERE after a join, an ambiguous
' unqualified column refused, and a composite ON (one equi term plus one
' residual term in the SAME ON clause) filtering correctly. A third
' table, buildinginfo (Building, City) - B5/Springfield, B7/Shelbyville
' - proves multiple JOINs fold left-to-right, a later join's own ON
' referencing a column an EARLIER join already added.
Private Sub TestSqlJoin()
    Dim staffCols As New Collection
    staffCols.Add SqlColPair("name", "Name")
    staffCols.Add SqlColPair("dept", "Dept")
    Dim staffRows As New Collection
    Dim s1(1 To 2) As Variant: s1(1) = "alice": s1(2) = "eng": staffRows.Add s1
    Dim s2(1 To 2) As Variant: s2(1) = "bob": s2(2) = "sales": staffRows.Add s2
    Dim s3(1 To 2) As Variant: s3(1) = "carol": s3(2) = "eng": staffRows.Add s3

    Dim deptCols As New Collection
    deptCols.Add SqlColPair("dept", "Dept")
    deptCols.Add SqlColPair("building", "Building")
    Dim deptRows As New Collection
    Dim d1(1 To 2) As Variant: d1(1) = "eng": d1(2) = "B5": deptRows.Add d1
    Dim d2(1 To 2) As Variant: d2(1) = "sales": d2(2) = "B7": deptRows.Add d2

    Dim bldgCols As New Collection
    bldgCols.Add SqlColPair("building", "Building")
    bldgCols.Add SqlColPair("city", "City")
    Dim bldgRows As New Collection
    Dim b1(1 To 2) As Variant: b1(1) = "B5": b1(2) = "Springfield": bldgRows.Add b1
    Dim b2(1 To 2) As Variant: b2(1) = "B7": b2(2) = "Shelbyville": bldgRows.Add b2

    Dim twoTables As New Collection
    twoTables.Add SqlTableRec("staff", staffCols, staffRows)
    twoTables.Add SqlTableRec("deptinfo", deptCols, deptRows)

    Dim result As Collection

    Set result = VLA_Sql.SqlRunJoin( _
        "SELECT staff.Name, deptinfo.Building FROM staff JOIN deptinfo ON staff.Dept = deptinfo.Dept", twoTables)
    Report "sql join: a basic equi-join derives one row per staff member (3 rows)", _
           CountRows(result) = 3, "got " & CountRows(result)
    Report "sql join: qualified SELECT columns keep their own table's original casing as headers", _
           HeaderAt(result, 1) = "Name" And HeaderAt(result, 2) = "Building", _
           "got " & HeaderAt(result, 1) & "/" & HeaderAt(result, 2)

    ' RelJoin's own output order is never part of its contract (the
    ' identical reason VLA_Datalog's own aggregation tests probe via
    ' RelContainsTuple rather than trust row position) - filtered down to
    ' exactly ONE row via WHERE first, TestSql's own established safe
    ' pattern for a ValueAt check, rather than assuming "row 1" means
    ' any particular staff member.
    Set result = VLA_Sql.SqlRunJoin( _
        "SELECT staff.Name, deptinfo.Building FROM staff JOIN deptinfo ON staff.Dept = deptinfo.Dept WHERE staff.Name = 'alice'", twoTables)
    Report "sql join: alice's own row resolves to building B5", _
           ValueAt(result, 1, 2) = "B5", "got " & ValueAt(result, 1, 2)

    Set result = VLA_Sql.SqlRunJoin( _
        "SELECT Name, Building FROM staff JOIN deptinfo ON staff.Dept = deptinfo.Dept", twoTables)
    Report "sql join: an unqualified column that's unambiguous across both tables still resolves", _
           CountRows(result) = 3 And HeaderAt(result, 1) = "Name", "got " & CountRows(result) & " rows, header " & HeaderAt(result, 1)

    Set result = VLA_Sql.SqlRunJoin( _
        "SELECT * FROM staff JOIN deptinfo ON staff.Dept = deptinfo.Dept", twoTables)
    Report "sql join: SELECT * over a join lists every joined table's own columns (4) in join order", _
           HeaderAt(result, 1) = "Name" And HeaderAt(result, 2) = "Dept" And HeaderAt(result, 3) = "Dept" And HeaderAt(result, 4) = "Building", _
           "got " & HeaderAt(result, 1) & "/" & HeaderAt(result, 2) & "/" & HeaderAt(result, 3) & "/" & HeaderAt(result, 4)

    Set result = VLA_Sql.SqlRunJoin( _
        "SELECT staff.Name FROM staff JOIN deptinfo ON staff.Dept = deptinfo.Dept WHERE deptinfo.Building = 'B5'", twoTables)
    Report "sql join: WHERE after a JOIN, over a qualified column, filters correctly (alice + carol - 2 rows)", _
           CountRows(result) = 2, "got " & CountRows(result)

    Set result = VLA_Sql.SqlRunJoin( _
        "SELECT staff.Name FROM staff JOIN deptinfo ON staff.Dept = deptinfo.Dept AND deptinfo.Building = 'B5'", twoTables)
    Report "sql join: a composite ON (one equi term, one residual term) filters within the join itself (2 rows)", _
           CountRows(result) = 2, "got " & CountRows(result)

    Set result = VLA_Sql.SqlRunJoin( _
        "SELECT DISTINCT deptinfo.Building FROM staff JOIN deptinfo ON staff.Dept = deptinfo.Dept", twoTables)
    Report "sql join: DISTINCT after a JOIN collapses eng's own two staffers to one Building row (2 rows)", _
           CountRows(result) = 2, "got " & CountRows(result)

    Set result = VLA_Sql.SqlRunJoin( _
        "SELECT staff.Name FROM staff INNER JOIN deptinfo ON staff.Dept = deptinfo.Dept", twoTables)
    Report "sql join: the explicit 'INNER JOIN' spelling means exactly the same thing as bare 'JOIN' (3 rows)", _
           CountRows(result) = 3, "got " & CountRows(result)

    ' SQL.4: GROUP BY over a joined row set - a grouped/aggregate query
    ' can (and, per this item's own scope note, should be tested to)
    ' also involve a JOIN. deptinfo.Building groups staff's own 3 rows
    ' into B5 (alice + carol, eng) and B7 (bob, sales).
    Set result = VLA_Sql.SqlRunJoin( _
        "SELECT deptinfo.Building, COUNT(*) AS n FROM staff JOIN deptinfo ON staff.Dept = deptinfo.Dept GROUP BY deptinfo.Building", twoTables)
    Report "sql join: GROUP BY over a joined row set collapses 3 rows into 2 groups (B5, B7)", _
           CountRows(result) = 2, "got " & CountRows(result)

    ' SQL.5: ORDER BY over a joined row set - staff's own 3 rows
    ' (alice/eng, bob/sales, carol/eng) sort alphabetically by Name.
    ' ORDER BY resolves against this query's own OUTPUT header (bare
    ' "Name" - a qualified SELECT item's own output header is always
    ' unqualified, NodeColumnOriginal's own originalText), never a
    ' qualified source reference like "staff.Name" - this file's own
    ' SQL.5 header note has the scope reasoning.
    Set result = VLA_Sql.SqlRunJoin( _
        "SELECT staff.Name, deptinfo.Building FROM staff JOIN deptinfo ON staff.Dept = deptinfo.Dept ORDER BY Name", twoTables)
    Report "sql join: ORDER BY over a joined row set sorts correctly (alice, bob, carol)", _
           ValueAt(result, 1, 1) = "alice" And ValueAt(result, 2, 1) = "bob" And ValueAt(result, 3, 1) = "carol", _
           "got " & ValueAt(result, 1, 1) & "/" & ValueAt(result, 2, 1) & "/" & ValueAt(result, 3, 1)

    Dim threeTables As New Collection
    threeTables.Add SqlTableRec("staff", staffCols, staffRows)
    threeTables.Add SqlTableRec("deptinfo", deptCols, deptRows)
    threeTables.Add SqlTableRec("buildinginfo", bldgCols, bldgRows)
    Set result = VLA_Sql.SqlRunJoin( _
        "SELECT staff.Name, buildinginfo.City FROM staff" & _
        " JOIN deptinfo ON staff.Dept = deptinfo.Dept" & _
        " JOIN buildinginfo ON deptinfo.Building = buildinginfo.Building", threeTables)
    Report "sql join: a SECOND join folds left-to-right, its own ON referencing a column the FIRST join already added (3 rows)", _
           CountRows(result) = 3, "got " & CountRows(result)

    Set result = VLA_Sql.SqlRunJoin( _
        "SELECT staff.Name, buildinginfo.City FROM staff" & _
        " JOIN deptinfo ON staff.Dept = deptinfo.Dept" & _
        " JOIN buildinginfo ON deptinfo.Building = buildinginfo.Building" & _
        " WHERE staff.Name = 'alice'", threeTables)
    Report "sql join: alice's own row resolves all the way through to Springfield", _
           ValueAt(result, 1, 2) = "Springfield", "got " & ValueAt(result, 1, 2)

    Dim raised As Boolean

    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Sql.SqlRunJoin "SELECT Dept FROM staff JOIN deptinfo ON staff.Dept = deptinfo.Dept", twoTables
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "sql join: an unqualified column ambiguous across two joined tables (Dept) is refused, not guessed", raised, "no error raised"

    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Sql.SqlRunJoin "SELECT * FROM staff JOIN staff ON staff.Dept = staff.Dept", twoTables
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "sql join: the same table named twice across FROM/JOIN (a self-join, no aliasing support) is refused", raised, "no error raised"

    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Sql.SqlRunJoin "SELECT * FROM staff JOIN nosuchtable ON staff.Dept = nosuchtable.Dept", twoTables
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "sql join: a JOIN naming a table that wasn't passed to this call is refused", raised, "no error raised"

    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Sql.SqlRunJoin "SELECT * FROM staff LEFT JOIN deptinfo ON staff.Dept = deptinfo.Dept", twoTables
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "sql join: LEFT JOIN is refused by name, not silently misread as INNER JOIN", raised, "no error raised"

    ' sql-needs-at-least-one-table is the UDF's own argument-count check
    ' (VLA_Sql.SQL, not SqlRunJoin - SqlRunJoin always receives an
    ' already-built tables Collection, so it has no "zero arguments" case
    ' of its own) - pure even so, since zero table ARGUMENTS means no
    ' Range object is ever touched at all. VLA_Sql.SQL never raises to
    ' its OWN caller (its own On Error GoTo fail swallows everything and
    ' returns readable "#SQL! ..." text instead, this file's own
    ' DATALOG/SQL precedent), so this checks the returned STRING, not
    ' Err.Number, unlike every other refusal check in this Sub.
    Dim noTableResult As Variant
    noTableResult = VLA_Sql.SQL("SELECT * FROM staff")
    Report "sql join: calling SQL() with no table arguments at all is refused, not a raw crash", _
           ResultTextStartsWith(noTableResult, "#SQL!"), _
           "got " & TypeName(noTableResult) & " = " & ResultDescribe(noTableResult)
End Sub

' SQL.6: UNION / UNION ALL / INTERSECT / EXCEPT, entirely pure -
' VLA_Sql.SqlRunCompound called directly with hand-built table records
' (SqlTableRec), the identical pure seam TestSqlJoin (SQL.3) already
' established for its own new grammar layer. Three tiny single-column
' tables (t1={1,2,3}, t2={2,3,4}, t3={3,4,5}) prove the four operators'
' own row-level semantics against small, hand-checkable sets; a reused
' copy of TestSql's own staff table proves the compound path's own
' backward-compatibility promise (a non-compound query through
' SqlRunCompound keeps SQL.5's own ORDER BY SOURCE-column fallback,
' which the narrower compound-only resolution rule does NOT); and
' INTERSECT's own real, verified-against-SQLite precedence (binding
' TIGHTER than UNION/EXCEPT, which are left-associative with each other)
' is proven with concrete data where the two possible groupings give
' DIFFERENT, distinguishable answers - the same "prove it, don't just
' assert it" discipline SQL.2's own arithmetic-precedence tests already
' established.
Private Sub TestSqlSetOps()
    Dim n1cols As New Collection
    n1cols.Add SqlColPair("n", "n")

    Dim t1rows As New Collection
    Dim a1(1 To 1) As Variant: a1(1) = 1: t1rows.Add a1
    Dim a2(1 To 1) As Variant: a2(1) = 2: t1rows.Add a2
    Dim a3(1 To 1) As Variant: a3(1) = 3: t1rows.Add a3
    Dim t1 As Collection: Set t1 = SqlTableRec("t1", n1cols, t1rows)

    Dim t2rows As New Collection
    Dim b2(1 To 1) As Variant: b2(1) = 2: t2rows.Add b2
    Dim b3(1 To 1) As Variant: b3(1) = 3: t2rows.Add b3
    Dim b4(1 To 1) As Variant: b4(1) = 4: t2rows.Add b4
    Dim t2 As Collection: Set t2 = SqlTableRec("t2", n1cols, t2rows)

    Dim t3rows As New Collection
    Dim c3(1 To 1) As Variant: c3(1) = 3: t3rows.Add c3
    Dim c4(1 To 1) As Variant: c4(1) = 4: t3rows.Add c4
    Dim c5(1 To 1) As Variant: c5(1) = 5: t3rows.Add c5
    Dim t3 As Collection: Set t3 = SqlTableRec("t3", n1cols, t3rows)

    Dim tables As New Collection
    tables.Add t1
    tables.Add t2
    tables.Add t3

    Dim result As Collection

    Set result = VLA_Sql.SqlRunCompound("SELECT n FROM t1 UNION SELECT n FROM t2", tables)
    Report "sql.6: UNION dedups across both branches (t1={1,2,3} u t2={2,3,4} -> 4 rows)", _
           CountRows(result) = 4, "got " & CountRows(result)

    Set result = VLA_Sql.SqlRunCompound("SELECT n FROM t1 UNION ALL SELECT n FROM t2", tables)
    Report "sql.6: UNION ALL is pure concatenation, no dedup at all (3 + 3 = 6 rows)", _
           CountRows(result) = 6, "got " & CountRows(result)

    Set result = VLA_Sql.SqlRunCompound("SELECT n FROM t1 INTERSECT SELECT n FROM t2", tables)
    Report "sql.6: INTERSECT keeps only rows present on both sides (t1 n t2 = {2,3} -> 2 rows)", _
           CountRows(result) = 2, "got " & CountRows(result)

    Set result = VLA_Sql.SqlRunCompound("SELECT n FROM t1 EXCEPT SELECT n FROM t2", tables)
    Report "sql.6: EXCEPT keeps only left-side rows absent from the right (t1 - t2 = {1} -> 1 row)", _
           CountRows(result) = 1, "got " & CountRows(result)
    Report "sql.6: EXCEPT's one surviving row really is 1, not some other leftover value", _
           CDbl(ValueAt(result, 1, 1)) = 1, "got " & ValueAt(result, 1, 1)

    ' INTERSECT's own real precedence, verified against actual SQLite
    ' behavior (not assumed by analogy - VLA_Sql.bas's own SQL.5 header
    ' note has the cautionary precedent for why that matters here): it
    ' binds TIGHTER than UNION/EXCEPT, which are themselves left-
    ' associative - `t1 UNION t2 INTERSECT t3` must parse as
    ' `t1 UNION (t2 INTERSECT t3)`, never `(t1 UNION t2) INTERSECT t3`.
    ' The two groupings give DIFFERENT, distinguishable answers over this
    ' data: t1 u (t2 n t3) = {1,2,3} u {3,4} = {1,2,3,4} (4 rows), while
    ' (t1 u t2) n t3 = {1,2,3,4} n {3,4,5} = {3,4} (2 rows) - only the
    ' CORRECT precedence keeps 1 and 2 at all.
    Set result = VLA_Sql.SqlRunCompound("SELECT n FROM t1 UNION SELECT n FROM t2 INTERSECT SELECT n FROM t3", tables)
    Report "sql.6: INTERSECT binds tighter than UNION - t1 U (t2 ^ t3), not (t1 U t2) ^ t3 (4 rows)", _
           CountRows(result) = 4, "got " & CountRows(result)

    ' Column-count compatibility, not name or type - refused by NAME
    ' (sql-set-op-column-mismatch), not guessed (padded, truncated, or
    ' matched positionally past the shorter side).
    Dim raised As Boolean
    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Sql.SqlRunCompound "SELECT n FROM t1 UNION SELECT n, n FROM t2", tables
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "sql.6: mismatched column counts across a set operator are refused, not guessed", raised, "no error raised"

    ' The promised backward-compatibility contract, proven rather than
    ' just claimed: an ordinary, non-compound query through
    ' SqlRunCompound keeps SQL.5's own ORDER BY SOURCE-column fallback
    ' (ordering by Dept even though only Name was ever SELECTed) - which
    ' it would NOT if every query were forced through the compound
    ' path's own narrower, output-only ORDER BY resolution.
    Dim staffCols As New Collection
    staffCols.Add SqlColPair("name", "Name")
    staffCols.Add SqlColPair("dept", "Dept")
    Dim staffRows As New Collection
    Dim s1(1 To 2) As Variant: s1(1) = "alice": s1(2) = "eng": staffRows.Add s1
    Dim s2(1 To 2) As Variant: s2(1) = "bob": s2(2) = "sales": staffRows.Add s2
    Dim staffTable As Collection: Set staffTable = SqlTableRec("staff", staffCols, staffRows)
    Dim staffTables As New Collection: staffTables.Add staffTable

    Set result = VLA_Sql.SqlRunCompound("SELECT Name FROM staff ORDER BY Dept DESC", staffTables)
    Report "sql.6: a plain (non-compound) query through SqlRunCompound keeps SQL.5's own SOURCE-column ORDER BY fallback", _
           CountRows(result) = 2 And ValueAt(result, 1, 1) = "bob", _
           "got " & CountRows(result) & " rows, first=" & ValueAt(result, 1, 1)

    ' A genuinely compound query's own ORDER BY, trailing the WHOLE
    ' combined result exactly once, resolved against the combined OUTPUT
    ' columns.
    Set result = VLA_Sql.SqlRunCompound("SELECT n FROM t1 UNION SELECT n FROM t2 ORDER BY n DESC", tables)
    Report "sql.6: ORDER BY trailing the WHOLE compound query sorts the combined result (4,3,2,1)", _
           CountRows(result) = 4 And CDbl(ValueAt(result, 1, 1)) = 4 And CDbl(ValueAt(result, 4, 1)) = 1, _
           "got " & CountRows(result) & " rows, first=" & ValueAt(result, 1, 1) & " last=" & ValueAt(result, CountRows(result), 1)

    ' This item's own explicit, named scope decision: a compound query's
    ' ORDER BY can NOT fall back to an un-SELECTed source column the way
    ' a plain query's own ORDER BY can (no single coherent source colMap
    ' survives combining two branches) - refused, not silently guessed.
    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Sql.SqlRunCompound "SELECT n FROM t1 UNION SELECT n FROM t2 ORDER BY nosuchcolumn", tables
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "sql.6: a compound ORDER BY can't fall back to a source column the way a plain query's own ORDER BY can", raised, "no error raised"

    ' The FIRST branch's own output header names the WHOLE compound
    ' result, regardless of later branches' own aliases - real SQL's own
    ' rule, named explicitly rather than left to whichever branch
    ' happens to run last.
    Set result = VLA_Sql.SqlRunCompound("SELECT n AS first_n FROM t1 UNION SELECT n AS second_n FROM t2", tables)
    Report "sql.6: the FIRST branch's own output header names the whole compound result", _
           HeaderAt(result, 1) = "first_n", "got """ & HeaderAt(result, 1) & """"
End Sub

' SQL.7: WITH CTEs, then recursive WITH - entirely pure -
' VLA_Sql.SqlRunWith called directly with hand-built table records
' (SqlTableRec), the identical pure seam TestSqlSetOps (SQL.6) already
' established for its own new grammar layer. A recursive `reach` CTE
' over a small 4-node chain (edges 1->2->3->4) is this Sub's own
' centerpiece - it PROVES the round loop is genuinely semi-naive (each
' round joining ONLY the previous round's own delta, never the full
' accumulated total): a "used the full total instead of the delta"
' bug would re-derive already-emitted pairs as duplicates under this
' engine's own no-dedup UNION ALL semantics, which would show up here
' as MORE than the exact 6 rows a correct transitive closure over a
' 4-node chain has (all i<j pairs among {1,2,3,4}), not merely a wrong
' VALUE - a row-count assertion alone is enough to catch it. No live
' test attempts to actually exhaust SQL_CTE_MAX_ROUNDS (10000 real
' rounds through the full parse/evaluate pipeline would make this
' suite minutes slower for no new information - the ceiling check
' itself is a single, structurally trivial `If round > ... Then
' RaiseMsg` line, the identical shape VLA_Datalog's own already-tested
' MAX_ROUNDS check uses).
Private Sub TestSqlCte()
    Dim n1cols As New Collection
    n1cols.Add SqlColPair("n", "n")

    Dim t1rows As New Collection
    Dim a1(1 To 1) As Variant: a1(1) = 1: t1rows.Add a1
    Dim a2(1 To 1) As Variant: a2(1) = 2: t1rows.Add a2
    Dim a3(1 To 1) As Variant: a3(1) = 3: t1rows.Add a3
    Dim t1 As Collection: Set t1 = SqlTableRec("t1", n1cols, t1rows)

    Dim t2rows As New Collection
    Dim b2(1 To 1) As Variant: b2(1) = 2: t2rows.Add b2
    Dim b3(1 To 1) As Variant: b3(1) = 3: t2rows.Add b3
    Dim b4(1 To 1) As Variant: b4(1) = 4: t2rows.Add b4
    Dim t2 As Collection: Set t2 = SqlTableRec("t2", n1cols, t2rows)

    Dim setTables As New Collection
    setTables.Add t1
    setTables.Add t2

    Dim result As Collection

    ' A non-compound, non-WITH query through SqlRunWith is byte-for-byte
    ' SqlRunCompound/SqlRunJoin, unchanged - the promised backward-
    ' compatibility contract, proven directly, the same way SqlRunCompound
    ' itself proved it relative to SqlRunJoin (SQL.6).
    Set result = VLA_Sql.SqlRunWith("SELECT n FROM t1 UNION SELECT n FROM t2", setTables)
    Report "sql.7: a plain compound query (no WITH at all) through SqlRunWith behaves exactly like SqlRunCompound", _
           CountRows(result) = 4, "got " & CountRows(result)

    ' Multiple CTEs, a later one referencing an earlier one - the FROM/
    ' JOIN namespace really does grow as each cte def is evaluated, in
    ' written order.
    Set result = VLA_Sql.SqlRunWith( _
        "WITH a AS (SELECT n FROM t1), b AS (SELECT n FROM a WHERE n > 1) SELECT n FROM b", setTables)
    Report "sql.7: a later CTE (b) can reference an earlier one (a) - {1,2,3} filtered to n>1 -> 2 rows", _
           CountRows(result) = 2, "got " & CountRows(result)

    ' A CTE sharing a name with an ACTUALLY-PASSED real table shadows
    ' it - real SQL's own rule, true here for free (later registration
    ' in SqlRunJoin's own byName dict simply overwrites the earlier
    ' one) rather than by any special-casing.
    Set result = VLA_Sql.SqlRunWith("WITH t1 AS (SELECT n FROM t2) SELECT n FROM t1", setTables)
    Report "sql.7: a CTE named the same as a real passed table SHADOWS it (t2's own {2,3,4}, not the real t1's {1,2,3})", _
           CountRows(result) = 3, "got " & CountRows(result)

    ' Derived tables/correlated subqueries are already structurally
    ' impossible in this grammar - confirmed empirically, not just
    ' asserted in a comment: FROM never accepts anything but a bare
    ' identifier, refused by the ordinary "expected a table name"
    ' message, no new refusal needed for this at all.
    Dim raised As Boolean
    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Sql.SqlRunWith "SELECT * FROM (SELECT n FROM t1) sub", setTables
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "sql.7: a derived table (FROM (SELECT ...)) is refused - already structurally impossible, not a new check", raised, "no error raised"

    ' A duplicate CTE name within one WITH clause is refused by name.
    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Sql.SqlRunWith "WITH a AS (SELECT n FROM t1), a AS (SELECT n FROM t2) SELECT n FROM a", setTables
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "sql.7: a duplicate CTE name in the same WITH clause is refused, not silently overwritten", raised, "no error raised"

    ' A self-reference with no RECURSIVE keyword given is refused, not
    ' silently attempted or silently treated as an ordinary CTE.
    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Sql.SqlRunWith _
        "WITH a AS (SELECT n FROM t1 UNION ALL SELECT n FROM a) SELECT n FROM a", setTables
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "sql.7: a self-referencing CTE body without RECURSIVE is refused", raised, "no error raised"

    ' ---- the recursive convergence proof itself ---------------------
    Dim edgeCols As New Collection
    edgeCols.Add SqlColPair("src", "src")
    edgeCols.Add SqlColPair("dst", "dst")
    Dim edgeRows As New Collection
    Dim e12(1 To 2) As Variant: e12(1) = 1: e12(2) = 2: edgeRows.Add e12
    Dim e23(1 To 2) As Variant: e23(1) = 2: e23(2) = 3: edgeRows.Add e23
    Dim e34(1 To 2) As Variant: e34(1) = 3: e34(2) = 4: edgeRows.Add e34
    Dim edges As Collection: Set edges = SqlTableRec("edges", edgeCols, edgeRows)
    Dim edgeTables As New Collection
    edgeTables.Add edges

    Dim reachQuery As String
    reachQuery = "WITH RECURSIVE reach AS (" & _
                 "SELECT src, dst FROM edges" & _
                 " UNION ALL" & _
                 " SELECT reach.src, edges.dst FROM reach JOIN edges ON reach.dst = edges.src" & _
                 ") SELECT src, dst FROM reach ORDER BY src, dst"
    Set result = VLA_Sql.SqlRunWith(reachQuery, edgeTables)
    ' The full transitive closure of a 4-node chain (1->2->3->4) is
    ' every i<j pair among {1,2,3,4} - exactly 6 - sorted (src,dst):
    ' (1,2)(1,3)(1,4)(2,3)(2,4)(3,4).
    Report "sql.7: a recursive WITH computes the full transitive closure of a 4-node chain (6 rows, not more/fewer)", _
           CountRows(result) = 6, "got " & CountRows(result)
    Report "sql.7: the closure's own FIRST row (sorted) is the base edge (1,2)", _
           CDbl(ValueAt(result, 1, 1)) = 1 And CDbl(ValueAt(result, 1, 2)) = 2, _
           "got (" & ValueAt(result, 1, 1) & "," & ValueAt(result, 1, 2) & ")"
    ' Sorted ascending by (src, dst), the 6 rows are (1,2)(1,3)(1,4)
    ' (2,3)(2,4)(3,4) - (1,4), the fully-transitive pair only reachable
    ' via round 3's own delta-only join, sorts to POSITION 3 (src=1
    ' sorts before src=2/3), not the last row - a wrong test assertion
    ' live-caught here, not a wrong implementation (the row-COUNT check
    ' above already confirmed the closure itself has exactly 6 rows).
    Report "sql.7: (1,4), the fully-transitive pair only reachable via round 3's own delta-only join, is present (sorted position 3)", _
           CDbl(ValueAt(result, 3, 1)) = 1 And CDbl(ValueAt(result, 3, 2)) = 4, _
           "got (" & ValueAt(result, 3, 1) & "," & ValueAt(result, 3, 2) & ")"

    ' The base referencing itself is refused (real SQL's own rule: the
    ' base case must be non-recursive).
    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Sql.SqlRunWith _
        "WITH RECURSIVE reach AS (" & _
        "SELECT src, dst FROM reach" & _
        " UNION ALL" & _
        " SELECT reach.src, edges.dst FROM reach JOIN edges ON reach.dst = edges.src" & _
        ") SELECT * FROM reach", edgeTables
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "sql.7: a recursive CTE's own BASE referencing itself is refused", raised, "no error raised"

    ' A shape this item doesn't support (the recursive term isn't the
    ' step of a plain two-leaf UNION ALL - here it's nested inside an
    ' INTERSECT) is refused by name, not silently misread.
    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Sql.SqlRunWith _
        "WITH RECURSIVE reach AS (" & _
        "SELECT src, dst FROM edges" & _
        " UNION ALL" & _
        " SELECT reach.src, edges.dst FROM reach JOIN edges ON reach.dst = edges.src" & _
        " INTERSECT SELECT src, dst FROM edges" & _
        ") SELECT * FROM reach", edgeTables
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "sql.7: a recursive CTE body that isn't the one canonical base-UNION-ALL-step shape is refused", raised, "no error raised"

    ' ORDER BY/LIMIT inside a recursive CTE's own definition is refused
    ' by name (per-round ordering has no coherent meaning) - a plain
    ' non-recursive CTE, by contrast, keeps its own ORDER BY/LIMIT for
    ' free (it's just an ordinary compoundQuery), so this is scoped
    ' specifically to the recursive case.
    raised = False
    On Error Resume Next
    Err.Clear
    VLA_Sql.SqlRunWith _
        "WITH RECURSIVE reach AS (" & _
        "SELECT src, dst FROM edges" & _
        " UNION ALL" & _
        " SELECT reach.src, edges.dst FROM reach JOIN edges ON reach.dst = edges.src" & _
        " ORDER BY src" & _
        ") SELECT * FROM reach", edgeTables
    If Err.Number <> 0 Then raised = True
    On Error GoTo 0
    Report "sql.7: ORDER BY inside a recursive CTE's own definition is refused", raised, "no error raised"

    ' A non-recursive CTE's own ORDER BY/LIMIT, by contrast, is honored
    ' for free - no special-casing needed since its own body is just an
    ' ordinary compoundQuery, evaluated through the unchanged
    ' SqlRunCompound.
    Set result = VLA_Sql.SqlRunWith( _
        "WITH top2 AS (SELECT n FROM t1 UNION SELECT n FROM t2 ORDER BY n DESC LIMIT 2) SELECT n FROM top2", setTables)
    Report "sql.7: a NON-recursive CTE's own ORDER BY/LIMIT is honored (top2 of {1,2,3,4} desc -> 4,3)", _
           CountRows(result) = 2 And CDbl(ValueAt(result, 1, 1)) = 4 And CDbl(ValueAt(result, 2, 1)) = 3, _
           "got " & CountRows(result) & " rows, first=" & ValueAt(result, 1, 1)
End Sub

' Host-required, the same reason VLA_Datalog's own TestDatalogHostTable
' is: SQL.1 needs real column NAMES, which only a real ListObject can
' give it (VLA_Relation.RangeColumnNames' own ok=False path for anything
' else) - a plain 2D array has no .ListObject to read them from at all,
' so this capability has no pure-test seam whatsoever, unlike TestSql's
' own parser/evaluator pins above. SQL.3 adds a live two-table JOIN
' scenario at the end (a real ListObject on each side, qualified
' Employees.id-shaped column references resolved against each one's own
' live header row) - TestSqlJoin's own pure suite already covers the
' grammar/evaluator mechanics off hand-built tables, so this one scenario
' is enough to prove the SAME code path works against real Tables too,
' not a second full pass through every case.
Private Sub TestSqlHostTable()
    Dim prior As Worksheet
    Set prior = ActiveSheet

    VlaEnsureSheet "VlaSqlHostSheet"
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets("VlaSqlHostSheet")
    ws.Activate
    Do While ws.ListObjects.Count > 0
        ws.ListObjects(1).Delete
    Loop
    ws.Cells.Clear

    ws.Range("A1:C1").Value = Array("Name", "Dept", "Salary")
    ws.Range("A2:C2").Value = Array("alice", "eng", 90000)
    ws.Range("A3:C3").Value = Array("bob", "sales", 60000)
    ws.Range("A4:C4").Value = Array("carol", "eng", 95000)
    Dim loStaff As ListObject
    Set loStaff = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:C4"), , xlYes)
    loStaff.Name = "SqlHostTest1"

    ' Checks ALL THREE headers, not just the first - a live bug
    ' (Dim pair As New Collection inside RangeColumnNames' own loop,
    ' fixed in VLA_Relation.bas) produced exactly 3 columns with the
    ' FIRST one's name correct, which a first-header-only check would
    ' never have caught; it only surfaced through an explicit column
    ' lookup (the WHERE test below), which is why THIS check now insists
    ' on the full header row, not a sample of it.
    Dim arrAll As Variant
    arrAll = VLA_Sql.SQL("SELECT * FROM sqlhosttest1", loStaff.Range)
    Dim okAll As Boolean, detailAll As String
    If IsArray(arrAll) Then
        Dim nRowsAll As Long, nColsAll As Long
        nRowsAll = UBound(arrAll, 1) - LBound(arrAll, 1) + 1
        nColsAll = UBound(arrAll, 2) - LBound(arrAll, 2) + 1
        Dim r0 As Long, c0 As Long
        r0 = LBound(arrAll, 1): c0 = LBound(arrAll, 2)
        okAll = (nRowsAll = 4) And (nColsAll = 3) _
                And ResultCellIs(arrAll, r0, c0, "Name") _
                And ResultCellIs(arrAll, r0, c0 + 1, "Dept") _
                And ResultCellIs(arrAll, r0, c0 + 2, "Salary")
        detailAll = "got " & nRowsAll & " rows x " & nColsAll & " cols, headers '" & _
                    ResultCellText(arrAll, r0, c0) & "','" & ResultCellText(arrAll, r0, c0 + 1) & "','" & ResultCellText(arrAll, r0, c0 + 2) & "'"
    Else
        detailAll = "got " & TypeName(arrAll) & " = " & CStr(arrAll)
    End If
    Report "sql host: SELECT * FROM a real live Table - all 3 headers, original casing", okAll, detailAll

    Dim arrWhere As Variant
    arrWhere = VLA_Sql.SQL("SELECT Name FROM sqlhosttest1 WHERE Salary >= 90000", loStaff.Range)
    Dim okWhere As Boolean, detailWhere As String
    If IsArray(arrWhere) Then
        Dim nRowsWhere As Long
        nRowsWhere = UBound(arrWhere, 1) - LBound(arrWhere, 1) + 1
        okWhere = (nRowsWhere = 3)   ' header + alice + carol
        detailWhere = "got " & nRowsWhere & " rows"
    Else
        detailWhere = "got " & TypeName(arrWhere) & " = " & CStr(arrWhere)
    End If
    Report "sql host: WHERE against a real live Table's own Value2-typed numeric column", okWhere, detailWhere

    ' A PLAIN named range (a defined Name, no ListObject) - genuinely
    ' different from "not a range at all": SqlTableName's own resolution
    ' succeeds here (TableArgResolve happily names a plain defined range,
    ' DATALOG's own precedent), so this specifically exercises
    ' SqlColumnNames' own, SEPARATE requirement - SQL needs a real Table
    ' for column NAMES, unlike DATALOG which never needed any.
    On Error Resume Next
    ThisWorkbook.Names("VlaSqlPlainRangeTest").Delete
    On Error GoTo 0
    ws.Range("Z1:Z3").Value = Array("x", "y", "z")
    ThisWorkbook.Names.Add Name:="VlaSqlPlainRangeTest", RefersTo:="='" & ws.Name & "'!" & ws.Range("Z1:Z3").Address
    Dim plainRange As Range
    Set plainRange = ThisWorkbook.Names("VlaSqlPlainRangeTest").RefersToRange

    Dim arrNotReal As Variant
    arrNotReal = VLA_Sql.SQL("SELECT * FROM vlasqlplainrangetest", plainRange)
    Dim okNotReal As Boolean, detailNotReal As String
    If VarType(arrNotReal) = vbString Then
        Dim sNotReal As String
        sNotReal = CStr(arrNotReal)
        okNotReal = (Left$(sNotReal, 5) = "#SQL!")
        detailNotReal = "got """ & sNotReal & """"
    Else
        detailNotReal = "got a " & TypeName(arrNotReal) & ", not a #SQL! error string"
    End If
    Report "sql host: a plain named range (no ListObject) is refused - SQL needs real column names", okNotReal, detailNotReal
    On Error Resume Next
    ThisWorkbook.Names("VlaSqlPlainRangeTest").Delete
    On Error GoTo 0

    ' SQL.3: a live two-table INNER JOIN through the actual =SQL(...)
    ' worksheet function - two real Tables, qualified column references
    ' resolved against each one's own live header row. Retires the OLD
    ' SQL.1-era check that used to sit here ("passing two table
    ' arguments is refused") - that assumption no longer holds now that
    ' multiple table arguments are the whole point of JOIN; this is its
    ' replacement, exercising the NEW capability live rather than
    ' pinning behavior SQL.3 deliberately changed.
    ws.Range("E1:F1").Value = Array("Dept", "Building")
    ws.Range("E2:F2").Value = Array("eng", "Bldg5")
    ws.Range("E3:F3").Value = Array("sales", "Bldg7")
    Dim loDept As ListObject
    Set loDept = ws.ListObjects.Add(xlSrcRange, ws.Range("E1:F3"), , xlYes)
    loDept.Name = "DeptInfoHostTest1"

    Dim arrJoin As Variant
    arrJoin = VLA_Sql.SQL( _
        "SELECT sqlhosttest1.Name, deptinfohosttest1.Building" & _
        " FROM sqlhosttest1 JOIN deptinfohosttest1 ON sqlhosttest1.Dept = deptinfohosttest1.Dept" & _
        " WHERE sqlhosttest1.Name = 'alice'", _
        loStaff.Range, loDept.Range)
    Dim okJoin As Boolean, detailJoin As String
    If IsArray(arrJoin) Then
        Dim nRowsJoin As Long
        nRowsJoin = UBound(arrJoin, 1) - LBound(arrJoin, 1) + 1
        Dim rJ As Long, cJ As Long
        rJ = LBound(arrJoin, 1): cJ = LBound(arrJoin, 2)
        okJoin = (nRowsJoin = 2) And ResultCellIs(arrJoin, rJ + 1, cJ + 1, "Bldg5")
        detailJoin = "got " & nRowsJoin & " rows, building '" & ResultCellText(arrJoin, rJ + 1, cJ + 1) & "'"
    Else
        detailJoin = "got " & TypeName(arrJoin) & " = " & CStr(arrJoin)
    End If
    Report "sql host: a live two-table INNER JOIN through =SQL(...) resolves alice's own eng department to Bldg5", okJoin, detailJoin

    ' SQL.4: GROUP BY + aggregate + HAVING, over a JOIN, through the
    ' actual =SQL(...) worksheet function - the roadmap's own item scope
    ' note asked specifically for this combination live, since SQL.3
    ' already proved JOIN itself needs live verification and SQL.4 sits
    ' directly downstream of it. eng has alice (90000) and carol (95000)
    ' - 2 rows, sum 185000; sales has bob alone - 1 row, dropped by
    ' HAVING COUNT(*) > 1.
    Dim arrGroup As Variant
    arrGroup = VLA_Sql.SQL( _
        "SELECT sqlhosttest1.Dept, COUNT(*) AS n, SUM(sqlhosttest1.Salary) AS total" & _
        " FROM sqlhosttest1 JOIN deptinfohosttest1 ON sqlhosttest1.Dept = deptinfohosttest1.Dept" & _
        " GROUP BY sqlhosttest1.Dept HAVING COUNT(*) > 1", _
        loStaff.Range, loDept.Range)
    Dim okGroup As Boolean, detailGroup As String
    If IsArray(arrGroup) Then
        Dim nRowsGroup As Long
        nRowsGroup = UBound(arrGroup, 1) - LBound(arrGroup, 1) + 1
        Dim rG As Long, cG As Long
        rG = LBound(arrGroup, 1): cG = LBound(arrGroup, 2)
        okGroup = (nRowsGroup = 2) And ResultCellIs(arrGroup, rG + 1, cG, "eng") _
                  And ResultCellNumIs(arrGroup, rG + 1, cG + 1, 2) And ResultCellNumIs(arrGroup, rG + 1, cG + 2, 185000)
        detailGroup = "got " & nRowsGroup & " rows, dept '" & ResultCellText(arrGroup, rG + 1, cG) & _
                      "', n=" & ResultCellText(arrGroup, rG + 1, cG + 1) & ", total=" & ResultCellText(arrGroup, rG + 1, cG + 2)
    Else
        detailGroup = "got " & TypeName(arrGroup) & " = " & CStr(arrGroup)
    End If
    Report "sql host: a live GROUP BY + COUNT/SUM + HAVING over a JOIN, through =SQL(...), keeps only eng (2 rows, 185000)", okGroup, detailGroup

    ' SQL.5: ORDER BY + LIMIT, live through =SQL(...) - top 2 salaries
    ' out of alice/90000, bob/60000, carol/95000 are carol then alice.
    Dim arrOrder As Variant
    arrOrder = VLA_Sql.SQL("SELECT Name, Salary FROM sqlhosttest1 ORDER BY Salary DESC LIMIT 2", loStaff.Range)
    Dim okOrder As Boolean, detailOrder As String
    If IsArray(arrOrder) Then
        Dim nRowsOrder As Long
        nRowsOrder = UBound(arrOrder, 1) - LBound(arrOrder, 1) + 1
        Dim rO As Long, cO As Long
        rO = LBound(arrOrder, 1): cO = LBound(arrOrder, 2)
        okOrder = (nRowsOrder = 3) And ResultCellIs(arrOrder, rO + 1, cO, "carol") And ResultCellIs(arrOrder, rO + 2, cO, "alice")
        detailOrder = "got " & nRowsOrder & " rows (incl. header), top two: '" & _
                      ResultCellText(arrOrder, rO + 1, cO) & "', '" & ResultCellText(arrOrder, rO + 2, cO) & "'"
    Else
        detailOrder = "got " & TypeName(arrOrder) & " = " & CStr(arrOrder)
    End If
    Report "sql host: a live ORDER BY Salary DESC LIMIT 2, through =SQL(...), keeps carol then alice", okOrder, detailOrder

    ' SQL.6: a live compound query COMBINED with JOIN and GROUP BY,
    ' through the actual =SQL(...) worksheet function - the roadmap's
    ' own item scope note asked specifically for this combination live,
    ' the same reasoning SQL.4's own live JOIN+GROUP BY test above
    ' already used (SQL.6 sits downstream of both). One branch groups
    ' the live JOIN by department (eng: alice+carol=2, sales: bob=1);
    ' the OTHER branch, over the SAME live table with no JOIN at all,
    ' computes a grand total (3) - both branches select ONLY the
    ' aggregate itself (arity 1, no separate label column), a
    ' deliberate shape choice: SQL.4's own "every SELECT item must be a
    ' GROUP BY column or an aggregate" rule (live-caught here first,
    ' fixed by reshaping the query rather than loosening that rule -
    ' real SQL treats a literal alongside an aggregate as an implicit
    ' constant, this engine does not, and SQL.4 already shipped
    ' owner-tested on that narrower behavior) refuses a bare literal
    ' label sitting next to COUNT(*) with no GROUP BY of its own, so
    ' this test never introduces one. UNION ALL concatenates both
    ' branches (no dedup wanted, a real department count and the grand
    ' total could coincidentally collide in value), and ORDER BY n DESC,
    ' trailing the WHOLE compound query exactly once, sorts the combined
    ' 3 rows: 3 (grand total), then 2 and 1 (the two department counts,
    ' whichever order GROUP BY happened to emit them in).
    Dim arrCompound As Variant
    arrCompound = VLA_Sql.SQL( _
        "SELECT COUNT(*) AS n" & _
        " FROM sqlhosttest1 JOIN deptinfohosttest1 ON sqlhosttest1.Dept = deptinfohosttest1.Dept" & _
        " GROUP BY sqlhosttest1.Dept" & _
        " UNION ALL" & _
        " SELECT COUNT(*) AS n FROM sqlhosttest1" & _
        " ORDER BY n DESC", _
        loStaff.Range, loDept.Range)
    Dim okCompound As Boolean, detailCompound As String
    If IsArray(arrCompound) Then
        Dim nRowsCompound As Long
        nRowsCompound = UBound(arrCompound, 1) - LBound(arrCompound, 1) + 1
        Dim rC As Long, cC As Long
        rC = LBound(arrCompound, 1): cC = LBound(arrCompound, 2)
        okCompound = (nRowsCompound = 4) _
                     And ResultCellNumIs(arrCompound, rC + 1, cC, 3) _
                     And ResultCellNumIs(arrCompound, rC + 2, cC, 2) _
                     And ResultCellNumIs(arrCompound, rC + 3, cC, 1)
        detailCompound = "got " & nRowsCompound & " rows (incl. header): " & _
                          ResultCellText(arrCompound, rC + 1, cC) & ", " & _
                          ResultCellText(arrCompound, rC + 2, cC) & ", " & _
                          ResultCellText(arrCompound, rC + 3, cC)
    Else
        detailCompound = "got " & TypeName(arrCompound) & " = " & CStr(arrCompound)
    End If
    Report "sql host: a live UNION ALL of a JOIN+GROUP BY branch and a plain-table aggregate branch, ORDER BY'd, through =SQL(...), gives 3, 2, 1", okCompound, detailCompound

    ' SQL.7: a live recursive WITH, through the actual =SQL(...)
    ' worksheet function - the exact "everything transitively under
    ' this row" shape BETA_ROADMAP1.md's own DATALOG motivating example
    ' already used (an Employees table with a manager column IS a
    ' reports-to relation, no authoring step needed), now answered by
    ' SQL directly: bob reports to alice, carol reports to bob, dave
    ' reports to carol - a straight chain of command. The recursive CTE
    ' computes EVERY (direct and indirect) report/manager pair; filtered
    ' to dave and sorted, this must list dave's WHOLE chain of command -
    ' carol (direct), then bob and alice (indirect, only reachable
    ' through round 2 and round 3's own delta-only joins).
    ws.Range("H1:I1").Value = Array("Name", "ManagerName")
    ws.Range("H2:I2").Value = Array("bob", "alice")
    ws.Range("H3:I3").Value = Array("carol", "bob")
    ws.Range("H4:I4").Value = Array("dave", "carol")
    Dim loEmployees As ListObject
    Set loEmployees = ws.ListObjects.Add(xlSrcRange, ws.Range("H1:I4"), , xlYes)
    loEmployees.Name = "EmployeesHostTest1"

    Dim arrChain As Variant
    arrChain = VLA_Sql.SQL( _
        "WITH RECURSIVE reports AS (" & _
        "SELECT Name, ManagerName FROM employeeshosttest1" & _
        " UNION ALL" & _
        " SELECT reports.Name, employeeshosttest1.ManagerName FROM reports" & _
        " JOIN employeeshosttest1 ON reports.ManagerName = employeeshosttest1.Name" & _
        ")" & _
        " SELECT ManagerName FROM reports WHERE Name = 'dave' ORDER BY ManagerName", _
        loEmployees.Range)
    Dim okChain As Boolean, detailChain As String
    If IsArray(arrChain) Then
        Dim nRowsChain As Long
        nRowsChain = UBound(arrChain, 1) - LBound(arrChain, 1) + 1
        Dim rCh As Long, cCh As Long
        rCh = LBound(arrChain, 1): cCh = LBound(arrChain, 2)
        okChain = (nRowsChain = 4) _
                  And ResultCellIs(arrChain, rCh + 1, cCh, "alice") _
                  And ResultCellIs(arrChain, rCh + 2, cCh, "bob") _
                  And ResultCellIs(arrChain, rCh + 3, cCh, "carol")
        detailChain = "got " & nRowsChain & " rows (incl. header): " & _
                      ResultCellText(arrChain, rCh + 1, cCh) & ", " & ResultCellText(arrChain, rCh + 2, cCh) & ", " & ResultCellText(arrChain, rCh + 3, cCh)
    Else
        detailChain = "got " & TypeName(arrChain) & " = " & CStr(arrChain)
    End If
    Report "sql host: a live recursive WITH, through =SQL(...), finds dave's whole chain of command (alice, bob, carol)", okChain, detailChain

    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
    prior.Activate
End Sub
