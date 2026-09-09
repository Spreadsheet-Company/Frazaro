Attribute VB_Name = "VLA_Tests_Query"
Option Explicit
Public Const VLA_TESTS_QUERY_VERSION As String = "PROLOG.6"
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
'  crossing never has to happen a second time). SQL/PROLOG/SOLVE's own
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
' SOLVE). Needs a live workbook open (TestDatalogHostTable, below,
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
    TestDatalogHostTable
    TestUnify
    TestGRenderUnify
    TestUnifyTwoWay
    TestProlog
    TestPrologRules
    TestPrologArithmetic
    TestPrologComparison
    TestPrologUnification
    TestPrologTypeTests
    TestPrologBetween
    TestPrologQuotedVersusBare
    TestPrologNegation
    TestPrologFindall
    TestPrologCut
    TestPrologKeyedAtoms
    TestPrologHostTable
    TestSql
    TestSqlJoin
    TestSqlSetOps
    TestSqlCte
    TestSqlHostTable

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
           UBound(arrHeadless, 1) = VLA_Relation.RelCount(rel) And (CStr(arrHeadless(1, 1)) = "tom" Or CStr(arrHeadless(1, 1)) = "bob"), _
           "shape mismatch"

    Set result = VLA_Datalog.DatalogRun("(headless) (rule (nothing_here X) (never_true X)) (query nothing_here)")
    Set rel = VLA_Runtime.VlaDictGet(result.Item(2), result.Item(1))
    Dim arrEmpty As Variant
    arrEmpty = VLA_Relation.RelToSpilledArray(rel, , CBool(result.Item(4)))
    Report "datalog: a zero-row headless result is a safe blank scalar, not a crash", _
           VLA_Relation.RelCount(rel) = 0 And VarType(arrEmpty) = vbString And CStr(arrEmpty) = "", _
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
           (UBound(arr2, 1) = VLA_Relation.RelCount(rel) + 1) And CStr(arr2(1, 1)) = "A" And CStr(arr2(1, 2)) = "B", _
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
                   (CStr(arrNarrow(LBound(arrNarrow, 1), LBound(arrNarrow, 2))) = "alice" Or _
                    CStr(arrNarrow(LBound(arrNarrow, 1), LBound(arrNarrow, 2))) = "bob")
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

    On Error Resume Next
    ThisWorkbook.Names("VlaDatalogAliasTest").Delete
    On Error GoTo 0
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
    prior.Activate
End Sub

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
           ok And bn.Count = 1 And CStr(bn.Item(1)) = "x" And CStr(bv.Item(1)) = "hello", "got ok=" & ok

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
           ok And envN.Count = 1 And CStr(envN.Item(1)) = "X" And CStr(envT.Item(1)) = "a", "got ok=" & ok

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
           UBound(result, 1) = 2 And UBound(result, 2) = 1 And CStr(result(1, 1)) = "X" And CStr(result(2, 1)) = "bob", _
           "got shape/contents mismatch"

    ' Two conjuncts sharing a variable - the whole point of a
    ' conjunctive query - must intersect bindings across both.
    result = VLA_Prolog.PROLOG("(fact (parent tom bob)) (fact (parent bob liz)) (query (parent tom Y) (parent Y Z))")
    Report "prolog: two conjuncts sharing a variable intersect bindings correctly", _
           UBound(result, 1) = 2 And UBound(result, 2) = 2 And CStr(result(2, 1)) = "bob" And CStr(result(2, 2)) = "liz", _
           "got: Y=" & result(2, 1) & " Z=" & result(2, 2)

    ' Multiple solutions, in fact-authored order (a bag, not a set).
    result = VLA_Prolog.PROLOG("(fact (parent tom bob)) (fact (parent tom ann)) (query (parent tom X))")
    Report "prolog: multiple matching facts produce multiple solution rows, in authored order", _
           UBound(result, 1) = 3 And CStr(result(2, 1)) = "bob" And CStr(result(3, 1)) = "ann", _
           "got " & UBound(result, 1) - 1 & " rows"

    ' Bag semantics: the SAME fact authored twice produces two identical
    ' rows - real Prolog's own behavior, not deduped the way DATALOG's
    ' fixpoint-derived relations are.
    result = VLA_Prolog.PROLOG("(fact (parent tom bob)) (fact (parent tom bob)) (query (parent tom X))")
    Report "prolog: a fact authored twice produces two identical solution rows (bag, not set)", _
           UBound(result, 1) = 3 And CStr(result(2, 1)) = "bob" And CStr(result(3, 1)) = "bob", _
           "got " & UBound(result, 1) - 1 & " rows"

    ' A query with no free variables at all collapses to a boolean
    ' scalar - TRUE when at least one solution exists, FALSE otherwise.
    result = VLA_Prolog.PROLOG("(fact (parent tom bob)) (query (parent tom bob))")
    Report "prolog: a fully-ground query with no free variables returns TRUE, not an array", _
           VarType(result) = vbBoolean And result = True, "got " & TypeName(result) & " " & result

    result = VLA_Prolog.PROLOG("(fact (parent tom bob)) (query (parent tom ann))")
    Report "prolog: a fully-ground query that matches nothing returns FALSE", _
           VarType(result) = vbBoolean And result = False, "got " & TypeName(result) & " " & result

    ' A compound-term argument (the boundary DATALOG stays clear of) -
    ' bound and rendered back through VLA.VlaWriteForm.
    result = VLA_Prolog.PROLOG("(fact (likes tom (color red))) (query (likes tom X))")
    Report "prolog: a variable bound to a compound term renders via VlaWriteForm", _
           UBound(result, 1) = 2 And CStr(result(2, 1)) = "(color red)", "got: " & result(2, 1)

    ' A predicate with zero arguments (nullary) is legal here - unlike
    ' DATALOG, this engine never sizes an array off arity.
    result = VLA_Prolog.PROLOG("(fact (raining)) (query (raining))")
    Report "prolog: a nullary (zero-argument) predicate is legal and answers TRUE", _
           VarType(result) = vbBoolean And result = True, "got " & TypeName(result) & " " & result

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
           VarType(result) = vbBoolean And result = True, "got " & TypeName(result) & " " & result

    result = VLA_Prolog.PROLOG( _
        "(fact (parent tom bob)) (fact (parent bob liz)) " & _
        "(rule (grandparent X Z) (parent X Y) (parent Y Z)) " & _
        "(query (grandparent tom Z))")
    Report "prolog.4: a non-recursive rule's own free variable resolves through its body", _
           UBound(result, 1) = 2 And CStr(result(2, 1)) = "liz", "got: " & result(2, 1)

    ' A predicate defined by a FACT and a RULE together - both fire, in
    ' written order, with no special-casing needed anywhere in the
    ' solver (this module's own header has why).
    result = VLA_Prolog.PROLOG( _
        "(fact (likes tom pizza)) (fact (foodie ann)) " & _
        "(rule (likes X sushi) (foodie X)) " & _
        "(query (likes X Y))")
    Report "prolog.4: a predicate defined by both a fact and a rule tries both, in written order", _
           UBound(result, 1) = 3 And CStr(result(2, 1)) = "tom" And CStr(result(2, 2)) = "pizza" _
           And CStr(result(3, 1)) = "ann" And CStr(result(3, 2)) = "sushi", _
           "got " & (UBound(result, 1) - 1) & " rows"

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
           UBound(result, 1) = 4 And CStr(result(2, 1)) = "bob" And CStr(result(3, 1)) = "liz" And CStr(result(4, 1)) = "ann", _
           "got " & (UBound(result, 1) - 1) & " rows: " & JoinColumn(result, 1)

    ' A negative case over the SAME recursive program - tom is not his
    ' own ancestor, and the search must terminate (not loop forever)
    ' finding out.
    result = VLA_Prolog.PROLOG( _
        "(fact (parent tom bob)) (fact (parent bob liz)) (fact (parent liz ann)) " & _
        "(rule (ancestor X Y) (parent X Y)) " & _
        "(rule (ancestor X Y) (parent X Z) (ancestor Z Y)) " & _
        "(query (ancestor tom tom))")
    Report "prolog.4: a fully-ground recursive query that doesn't hold terminates and returns FALSE", _
           VarType(result) = vbBoolean And result = False, "got " & TypeName(result) & " " & result

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
    Report "prolog.4: a genuinely non-terminating rule is refused by the resolution-step ceiling, not left to hang or crash", _
           InStr(1, ceilingResult, "resolution steps", vbTextCompare) > 0, "got: " & ceilingResult

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
    Report "prolog.4: a cyclic graph's own free-variable reachability query is refused by the ceiling, not left to hang - a real, not just single-clause, non-termination shape", _
           InStr(1, cycleResult, "resolution steps", vbTextCompare) > 0, "got: " & cycleResult

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
           UBound(result, 1) = 2 And CStr(result(2, 1)) = "7" And CStr(result(2, 2)) = "(box 7)", _
           "got: Y=" & result(2, 1) & " Z=" & result(2, 2)
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
           UBound(result, 1) = 2 And CStr(result(2, 1)) = "10" And CStr(result(2, 2)) = "20", _
           "got: X=" & result(2, 1) & " Y=" & result(2, 2)

    ' Nested arithmetic, evaluated bottom-up in a pure query - the shape
    ' DATALOG's own flat (let Z (op X Y)) structurally can't express at
    ' all, since DATALOG forbids compound terms entirely.
    result = VLA_Prolog.PROLOG("(query (is X (+ (* 2 3) 4)))")
    Report "prolog.5.1: nested arithmetic evaluates bottom-up", _
           UBound(result, 1) = 2 And CStr(result(2, 1)) = "10", "got: " & result(2, 1)

    ' is checking an ALREADY-bound target via ordinary unification - real
    ' Prolog's own is/2 semantics, free from reusing UnifyTwoWay rather
    ' than DATALOG's own bind-only (let ...), which requires a brand-new
    ' variable and cannot express this at all.
    result = VLA_Prolog.PROLOG("(query (is 10 (+ 4 6)))")
    Report "prolog.5.1: is against an already-ground target succeeds when the value matches", _
           VarType(result) = vbBoolean And result = True, "got " & TypeName(result) & " " & result

    result = VLA_Prolog.PROLOG("(query (is 10 (+ 4 5)))")
    Report "prolog.5.1: is against an already-ground target fails when the value doesn't match", _
           VarType(result) = vbBoolean And result = False, "got " & TypeName(result) & " " & result

    ' Two chained is calls in ONE rule body - the first call's own result
    ' must be visible to the second through the same threaded env.
    result = VLA_Prolog.PROLOG("(rule (twice-plus-one X Y) (is Temp (* X 2)) (is Y (+ Temp 1))) (query (twice-plus-one 5 Y))")
    Report "prolog.5.1: two chained is calls in one rule body share the same threaded bindings", _
           UBound(result, 1) = 2 And CStr(result(2, 1)) = "11", "got: " & result(2, 1)

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

    r = CStr(VLA_Prolog.PROLOG("(query (is X (mod 5 2)))"))
    Report "prolog.5.1: an unrecognized arithmetic operator is refused at parse time", _
           InStr(1, r, "isn't an arithmetic operator", vbTextCompare) > 0, "got: " & r

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
    If IsArray(result) Then doubledOk = (UBound(result, 1) = 2 And CStr(result(2, 1)) = "alice")
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

    r = CStr(VLA_Prolog.PROLOG("(fact (thing (mod 5 2))) (rule (compute X) (thing Y) (is X (+ Y 1))) (query (compute X))"))
    Report "prolog.5.1 REGRESSION: a variable bound to a 3-item term with an unrecognized operator symbol inside is is refused, not crashed", _
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
    r = CStr(VLA_Prolog.PROLOG("(query (is X (+ () 3)))"))
    Report "prolog.5.1: an empty () arithmetic operand is refused, not crashed", _
           InStr(1, r, "exactly two operands", vbTextCompare) > 0, "got: " & r

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
    Report "prolog.7: (< 1 2) succeeds", VarType(result) = vbBoolean And result = True, "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (< 2 1))")
    Report "prolog.7: (< 2 1) fails", VarType(result) = vbBoolean And result = False, "got: " & ResultDescribe(result)

    result = VLA_Prolog.PROLOG("(query (> 2 1))")
    Report "prolog.7: (> 2 1) succeeds", VarType(result) = vbBoolean And result = True, "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (> 1 2))")
    Report "prolog.7: (> 1 2) fails", VarType(result) = vbBoolean And result = False, "got: " & ResultDescribe(result)

    result = VLA_Prolog.PROLOG("(query (=< 1 2))")
    Report "prolog.7: (=< 1 2) succeeds", VarType(result) = vbBoolean And result = True, "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (=< 2 1))")
    Report "prolog.7: (=< 2 1) fails", VarType(result) = vbBoolean And result = False, "got: " & ResultDescribe(result)

    result = VLA_Prolog.PROLOG("(query (>= 2 1))")
    Report "prolog.7: (>= 2 1) succeeds", VarType(result) = vbBoolean And result = True, "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (>= 1 2))")
    Report "prolog.7: (>= 1 2) fails", VarType(result) = vbBoolean And result = False, "got: " & ResultDescribe(result)

    result = VLA_Prolog.PROLOG("(query (=:= 2 2))")
    Report "prolog.7: (=:= 2 2) succeeds", VarType(result) = vbBoolean And result = True, "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (=:= 2 3))")
    Report "prolog.7: (=:= 2 3) fails", VarType(result) = vbBoolean And result = False, "got: " & ResultDescribe(result)

    result = VLA_Prolog.PROLOG("(query (=\= 2 3))")
    Report "prolog.7: (=\= 2 3) succeeds", VarType(result) = vbBoolean And result = True, "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (=\= 2 2))")
    Report "prolog.7: (=\= 2 2) fails", VarType(result) = vbBoolean And result = False, "got: " & ResultDescribe(result)

    ' ---- the boundary cases a wrong spelling-translation would survive.
    ' =< must be <= and not <; >= must be >= and not >. Without these,
    ' mapping =< to "<" would pass every test above.
    result = VLA_Prolog.PROLOG("(query (=< 2 2))")
    Report "prolog.7: =< is inclusive at equality (it is <=, not <)", _
           VarType(result) = vbBoolean And result = True, "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (>= 2 2))")
    Report "prolog.7: >= is inclusive at equality (it is >=, not >)", _
           VarType(result) = vbBoolean And result = True, "got: " & ResultDescribe(result)

    ' =:= is NUMERIC equality, so two different spellings of one number
    ' are equal - which text comparison would get wrong.
    result = VLA_Prolog.PROLOG("(query (=:= 2.0 2))")
    Report "prolog.7: =:= compares numerically, so 2.0 =:= 2", _
           VarType(result) = vbBoolean And result = True, "got: " & ResultDescribe(result)

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
           VarType(result) = vbBoolean And result = True, "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (= 2.0 2))")
    Report "prolog.7/8: `=` is NOT =:= - unification is structural, so 2.0 does not unify with 2", _
           VarType(result) = vbBoolean And result = False, "got: " & ResultDescribe(result)

    ' `<=` is not a Prolog spelling at all (real Prolog writes =<), so it
    ' must stay an ordinary unknown predicate rather than be accepted.
    result = VLA_Prolog.PROLOG("(query (<= 1 2))")
    Report "prolog.7: `<=` is not a Prolog operator and is not accepted as one", _
           VarType(result) = vbBoolean And result = False, "got: " & ResultDescribe(result)

    ' ---- the motivating case: a THRESHOLD filter, which PROLOG.6's own
    ' README example had to route around. Strict non-empty subset, so
    ' deleting the dispatch arm empties it rather than merely reordering.
    result = VLA_Prolog.PROLOG("(fact (emp alice 90000)) (fact (emp bob 70000)) (query (emp N S) (> S 80000))")
    Report "prolog.7: a comparison filters a real backtracking search to a strict subset", _
           IsArray(result) And UBound(result, 1) = 2 And ResultCol1Is(result, "alice"), _
           "got: " & ResultDescribe(result)

    ' A comparison NEVER binds, so it contributes no output column: the
    ' query above has exactly two, N and S, not a third for the goal.
    Report "prolog.7: a comparison goal contributes no output column (it never binds)", _
           IsArray(result) And UBound(result, 2) = 2, _
           "got columns: " & UBound(result, 2)

    ' The complement, proving the filter is a real test and not a
    ' constant: the same program with the comparison inverted selects the
    ' OTHER employee, so neither row is being dropped for another reason.
    result = VLA_Prolog.PROLOG("(fact (emp alice 90000)) (fact (emp bob 70000)) (query (emp N S) (< S 80000))")
    Report "prolog.7: inverting the comparison selects the complementary row", _
           IsArray(result) And UBound(result, 1) = 2 And ResultCol1Is(result, "bob"), _
           "got: " & ResultDescribe(result)

    ' ---- nested arithmetic on either side, proving EvalArithTerm's own
    ' recursion is reached through a comparison and not only through is.
    result = VLA_Prolog.PROLOG("(query (> (+ 2 3) 4))")
    Report "prolog.7: a nested arithmetic expression evaluates on the left side", _
           VarType(result) = vbBoolean And result = True, "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (> 10 (* 2 3)))")
    Report "prolog.7: a nested arithmetic expression evaluates on the right side", _
           VarType(result) = vbBoolean And result = True, "got: " & ResultDescribe(result)

    ' ---- a comparison inside a RULE body, so the goal is freshened per
    ' invocation before it is dispatched (FreshenTerm keeps position 1 -
    ' the operator - verbatim and rewrites only the argument variables).
    result = VLA_Prolog.PROLOG("(fact (emp alice 90000)) (fact (emp bob 70000)) (rule (rich N) (emp N S) (> S 80000)) (query (rich N))")
    Report "prolog.7: a comparison in a rule body survives freshening and filters correctly", _
           IsArray(result) And UBound(result, 1) = 2 And ResultCol1Is(result, "alice"), _
           "got: " & ResultDescribe(result)

    ' ---- a comparison nested inside `not`, proving the dispatch is
    ' reached through SolveIsolated's own bounded sub-call too.
    result = VLA_Prolog.PROLOG("(fact (emp alice 90000)) (fact (emp bob 70000)) (query (emp N S) (not (> S 80000)))")
    Report "prolog.7: a comparison composes with `not`", _
           IsArray(result) And UBound(result, 1) = 2 And ResultCol1Is(result, "bob"), _
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

    r = CStr(VLA_Prolog.PROLOG("(query (< 1 (mod 5 2)))"))
    Report "prolog.7: an unrecognized operator inside a comparison operand is refused at parse time", _
           InStr(1, r, "isn't an arithmetic operator", vbTextCompare) > 0, "got: " & r
    Report "prolog.7: that refusal names (< ...), not (is ...)", _
           InStr(1, r, "(< ...)", vbTextCompare) > 0 And InStr(1, r, "(is ...)", vbTextCompare) = 0, "got: " & r

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
           VarType(result) = vbBoolean And result = True, "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (= bob ann))")
    Report "prolog.8: (= bob ann) fails", _
           VarType(result) = vbBoolean And result = False, "got: " & ResultDescribe(result)

    result = VLA_Prolog.PROLOG("(query (\= bob ann))")
    Report "prolog.8: (\= bob ann) succeeds - the two do not unify", _
           VarType(result) = vbBoolean And result = True, "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (\= bob bob))")
    Report "prolog.8: (\= bob bob) fails - they do unify", _
           VarType(result) = vbBoolean And result = False, "got: " & ResultDescribe(result)

    result = VLA_Prolog.PROLOG("(query (== bob bob))")
    Report "prolog.8: (== bob bob) succeeds", _
           VarType(result) = vbBoolean And result = True, "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (== bob ann))")
    Report "prolog.8: (== bob ann) fails", _
           VarType(result) = vbBoolean And result = False, "got: " & ResultDescribe(result)

    result = VLA_Prolog.PROLOG("(query (\== bob ann))")
    Report "prolog.8: (\== bob ann) succeeds", _
           VarType(result) = vbBoolean And result = True, "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (\== bob bob))")
    Report "prolog.8: (\== bob bob) fails", _
           VarType(result) = vbBoolean And result = False, "got: " & ResultDescribe(result)

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
           VarType(result) = vbBoolean And result = False, "got: " & ResultDescribe(result)

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
           VarType(result) = vbBoolean And result = True, "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (== X Y))")
    Report "prolog.8: (== X Y) fails - two DISTINCT free variables are not the same variable", _
           VarType(result) = vbBoolean And result = False, "got: " & ResultDescribe(result)

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
           VarType(result) = vbBoolean And result = False, "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (== (f 1 2) (f 1 2)))")
    Report "prolog.8: (== (f 1 2) (f 1 2)) succeeds - structural identity recurses", _
           VarType(result) = vbBoolean And result = True, "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (== (f 1 2) (f 1)))")
    Report "prolog.8: ...and differing arity is not identical", _
           VarType(result) = vbBoolean And result = False, "got: " & ResultDescribe(result)

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
           VarType(result) = vbBoolean And result = False, "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (=:= 2.0 2))")
    Report "prolog.8: ...while (=:= 2.0 2) still succeeds - the two are not aliases", _
           VarType(result) = vbBoolean And result = True, "got: " & ResultDescribe(result)

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
           VarType(result) = vbBoolean And result = True, "got: " & ResultDescribe(result)

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
           VarType(result) = vbBoolean And result = True, "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (not (= bob bob)))")
    Report "prolog.8: ...and its twin correctly fails", _
           VarType(result) = vbBoolean And result = False, "got: " & ResultDescribe(result)

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
           VarType(result) = vbBoolean And result = False, "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (\== X (f X)))")
    Report "prolog.8: ...and (\== X (f X)) an ordinary True - the asymmetry with `\=` is deliberate", _
           VarType(result) = vbBoolean And result = True, "got: " & ResultDescribe(result)

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
           VarType(result) = vbBoolean And result = True, "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (var? bob))")
    Report "prolog.9: (var? bob) fails - a ground atom is not a variable", _
           VarType(result) = vbBoolean And result = False, "got: " & ResultDescribe(result)

    result = VLA_Prolog.PROLOG("(query (nonvar? bob))")
    Report "prolog.9: (nonvar? bob) succeeds", _
           VarType(result) = vbBoolean And result = True, "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (nonvar? X))")
    Report "prolog.9: (nonvar? X) fails - X is free", _
           VarType(result) = vbBoolean And result = False, "got: " & ResultDescribe(result)

    result = VLA_Prolog.PROLOG("(query (atom? bob))")
    Report "prolog.9: (atom? bob) succeeds", _
           VarType(result) = vbBoolean And result = True, "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (atom? 42))")
    Report "prolog.9: (atom? 42) fails - 42 is a number, not an atom", _
           VarType(result) = vbBoolean And result = False, "got: " & ResultDescribe(result)

    result = VLA_Prolog.PROLOG("(query (number? 42))")
    Report "prolog.9: (number? 42) succeeds", _
           VarType(result) = vbBoolean And result = True, "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (number? bob))")
    Report "prolog.9: (number? bob) fails", _
           VarType(result) = vbBoolean And result = False, "got: " & ResultDescribe(result)

    result = VLA_Prolog.PROLOG("(query (atomic? bob))")
    Report "prolog.9: (atomic? bob) succeeds - an atom is atomic", _
           VarType(result) = vbBoolean And result = True, "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (atomic? 42))")
    Report "prolog.9: (atomic? 42) succeeds - a number is atomic too", _
           VarType(result) = vbBoolean And result = True, "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (atomic? (f a)))")
    Report "prolog.9: (atomic? (f a)) fails - a compound term is not atomic", _
           VarType(result) = vbBoolean And result = False, "got: " & ResultDescribe(result)

    result = VLA_Prolog.PROLOG("(query (compound? (f a)))")
    Report "prolog.9: (compound? (f a)) succeeds", _
           VarType(result) = vbBoolean And result = True, "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (compound? bob))")
    Report "prolog.9: (compound? bob) fails", _
           VarType(result) = vbBoolean And result = False, "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (compound? 42))")
    Report "prolog.9: (compound? 42) fails", _
           VarType(result) = vbBoolean And result = False, "got: " & ResultDescribe(result)

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
           VarType(result) = vbBoolean And result = False, "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (atom? ""42""))")
    Report "prolog.9/10: ...and (atom? ""42"") SUCCEEDS - it is an atom, which is the same judgement seen from the other side", _
           VarType(result) = vbBoolean And result = True, "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (atomic? ""42""))")
    Report "prolog.9/10: (atomic? ""42"") succeeds - both readings agree it is atomic; only the atom/number split is at stake", _
           VarType(result) = vbBoolean And result = True, "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (atom? ""eng""))")
    Report "prolog.9/10: (atom? ""eng"") succeeds - a marked string is an atom, capitalisation or not", _
           VarType(result) = vbBoolean And result = True, "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (var? ""Hello""))")
    Report "prolog.9/10: (var? ""Hello"") FAILS - a capitalised STRING is not an unbound variable (the marker is read on the RAW text)", _
           VarType(result) = vbBoolean And result = False, "got: " & ResultDescribe(result)

    ' ---- the coherence law the transliteration asserted 100 times, run
    ' once here against the real engine: every atomic term is an atom or
    ' a number, never both and never neither.
    result = VLA_Prolog.PROLOG("(query (atomic? bob) (atom? bob) (nonvar? bob))")
    Report "prolog.9: bob is atomic AND an atom AND nonvar, all three together", _
           VarType(result) = vbBoolean And result = True, "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (atomic? 42) (number? 42) (nonvar? 42))")
    Report "prolog.9: 42 is atomic AND a number AND nonvar, all three together", _
           VarType(result) = vbBoolean And result = True, "got: " & ResultDescribe(result)

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
           VarType(result) = vbBoolean And result = False, "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (= eng sales))")
    Report "prolog.10: two BARE symbols that genuinely differ still answer False, no refusal", _
           VarType(result) = vbBoolean And result = False, "got: " & ResultDescribe(result)
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
           VarType(result) = vbBoolean And result = False, "got: " & ResultDescribe(result)
    r = CStr(VLA_Prolog.PROLOG("(query (= (f " & Chr$(34) & "eng" & Chr$(34) & ") (f eng)))"))
    Report "prolog.10: ...but the SAME functor with a marker-only argument does refuse, nested", _
           InStr(1, r, "different things", vbTextCompare) > 0, "got: " & r

    ' ---- THE SCOPE DECISION, pinned. Clause matching must stay SILENT,
    ' and this is the half that would break if the refusal were moved
    ' into UnifyTwoWay: adding a non-matching fact would turn a working
    ' query into an error, which no definite-clause program may do.
    result = VLA_Prolog.PROLOG("(fact (color " & Chr$(34) & "red" & Chr$(34) & ")) (fact (color red)) (query (color red))")
    Report "prolog.10: MONOTONICITY - a near-miss against another clause never aborts a query that has a real match", _
           VarType(result) = vbBoolean And result = True, "got: " & ResultDescribe(result)
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
           VarType(result) = vbBoolean And result = True, "got: " & ResultDescribe(result)
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
           VarType(result) = vbBoolean And result = False, "got: " & ResultDescribe(result)
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
           VarType(result) = vbBoolean And result = True, "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (between 1 10 50))")
    Report "prolog.9: (between 1 10 50) fails - out of range", _
           VarType(result) = vbBoolean And result = False, "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (between 1 10 1))")
    Report "prolog.9: (between 1 10 1) succeeds - the low bound is INCLUSIVE", _
           VarType(result) = vbBoolean And result = True, "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (between 1 10 10))")
    Report "prolog.9: (between 1 10 10) succeeds - the high bound is INCLUSIVE", _
           VarType(result) = vbBoolean And result = True, "got: " & ResultDescribe(result)
    result = VLA_Prolog.PROLOG("(query (between 1 10 11))")
    Report "prolog.9: (between 1 10 11) fails - one past the high bound", _
           VarType(result) = vbBoolean And result = False, "got: " & ResultDescribe(result)

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
           VarType(result) = vbBoolean And result = True, "got: " & ResultDescribe(result)
    r = CStr(VLA_Prolog.PROLOG("(query (between 1 1000000 X))"))
    Report "prolog.9: ...but (between 1 1000000 X) is REFUSED rather than hanging or grinding to the step ceiling", _
           InStr(1, r, "would generate", vbTextCompare) > 0, "got: " & r
    Report "prolog.9: the range refusal names the range and the ceiling, and does NOT blame a runaway rule", _
           InStr(1, r, "1000000", vbTextCompare) > 0 And InStr(1, r, "base case", vbTextCompare) = 0, "got: " & r

    ' ---- the ceiling boundary, both sides. The dispatch charges one
    ' step for the goal itself before a single value is generated, so the
    ' largest range that fits is PROLOG_MAX_STEPS - 1 = 119. One more
    ' must produce the RANGE refusal, never the step-ceiling one - that
    ' off-by-one is the whole reason the up-front check exists.
    result = VLA_Prolog.PROLOG("(query (between 1 119 X))")
    Report "prolog.9: (between 1 119 X) is the largest range that fits - 119 rows plus a header", _
           ResultRowCount(result) = 120, "got: " & ResultDescribe(result)
    r = CStr(VLA_Prolog.PROLOG("(query (between 1 120 X))"))
    Report "prolog.9: (between 1 120 X) is one too many, and says so as a RANGE problem not a runaway-rule one", _
           InStr(1, r, "would generate", vbTextCompare) > 0 And InStr(1, r, "base case", vbTextCompare) = 0, "got: " & r

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
           VarType(result) = vbBoolean And result = True, "got " & TypeName(result) & " " & result

    ' A directly-ground goal, negated: fails (a matching fact exists).
    result = VLA_Prolog.PROLOG("(fact (parent tom bob)) (query (not (parent tom bob)))")
    Report "prolog.5.2: not fails when the negated ground goal has a solution", _
           VarType(result) = vbBoolean And result = False, "got " & TypeName(result) & " " & result

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
           UBound(result, 1) = 3 And CStr(result(2, 1)) = "tom" And CStr(result(3, 1)) = "liz", _
           "got " & (UBound(result, 1) - 1) & " rows: " & JoinColumn(result, 1)

    ' The real design fork, hand-traced rather than assumed: a variable
    ' appearing ONLY inside a negated goal (foo/1 doesn't even exist, so
    ' the negation trivially succeeds) must never surface as a phantom
    ' one-column "X" output - it must collapse to the same boolean
    ' scalar shape as any other query with no real free variables.
    result = VLA_Prolog.PROLOG("(query (not (foo X)))")
    Report "prolog.5.2: a variable appearing only inside a negated goal is never collected as a free output column", _
           VarType(result) = vbBoolean And result = True, "got " & TypeName(result) & " " & result

    ' The shared step ceiling, run to real exhaustion inside a negated
    ' goal - proving PROLOG_MAX_STEPS is one running total the whole
    ' query shares, not a fresh allowance handed to every `not`.
    r = CStr(VLA_Prolog.PROLOG("(rule (loop X) (loop X)) (query (not (loop a)))"))
    Report "prolog.5.2: a genuinely non-terminating negated goal is refused by the shared resolution-step ceiling, not left to hang or crash", _
           InStr(1, r, "resolution steps", vbTextCompare) > 0, "got: " & r

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
           VarType(result) = vbBoolean And result = True, "got " & TypeName(result) & " " & result

    ' Nested (double) negation, both ways - proving the recursive
    ' dispatch survives nesting with no special-casing needed anywhere.
    result = VLA_Prolog.PROLOG("(fact (p a)) (query (not (not (p a))))")
    Report "prolog.5.2: double negation of a goal that holds is TRUE", _
           VarType(result) = vbBoolean And result = True, "got " & TypeName(result) & " " & result

    result = VLA_Prolog.PROLOG("(fact (p a)) (query (not (not (p b))))")
    Report "prolog.5.2: double negation of a goal that doesn't hold is FALSE", _
           VarType(result) = vbBoolean And result = False, "got " & TypeName(result) & " " & result

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
           VarType(result) = vbBoolean And result = False, "got " & TypeName(result) & " " & result
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
    result = VLA_Prolog.PROLOG("(fact (color red)) (fact (color green)) (fact (color blue)) (query (findall X (color X) Bag))")
    Report "prolog.5.3: findall harvests every solution into a list, in derivation order", _
           ResultCol1Is(result, "(red green blue)"), "got: " & ResultDescribe(result)

    ' Zero solutions collapses to an empty list, never an error - real
    ' findall's own signature behavior (unlike bagof/setof).
    result = VLA_Prolog.PROLOG("(query (findall X (nonexistent X) Bag))")
    Report "prolog.5.3: findall over a goal with no solutions binds Bag to an empty list, not an error", _
           ResultCol1Is(result, "()"), "got: " & ResultDescribe(result)

    ' A COMPOUND Template - proves SubstituteTemplate's own
    ' reconstruction of Template's structure per solution, not just a
    ' flat list of Template's own variable values.
    result = VLA_Prolog.PROLOG( _
        "(fact (person alice 30)) (fact (person bob 25)) " & _
        "(query (findall (pair Name Age) (person Name Age) Bag))")
    Report "prolog.5.3: a compound Template is reconstructed per solution, not flattened", _
           ResultCol1Is(result, "((pair alice 30) (pair bob 25))"), "got: " & ResultDescribe(result)

    ' findall composing with an OUTER already-bound variable across real
    ' backtracking - the classic "group by" idiom: for each department
    ' (bound by an earlier, ordinary conjunct), harvest its own
    ' employees into a separate list.
    result = VLA_Prolog.PROLOG( _
        "(fact (dept eng)) (fact (dept sales)) " & _
        "(fact (employee alice eng)) (fact (employee bob eng)) (fact (employee carol sales)) " & _
        "(query (dept D) (findall E (employee E D) Bag))")
    Dim groupByOk As Boolean
    If IsArray(result) Then
        groupByOk = (UBound(result, 1) = 3 And CStr(result(2, 1)) = "eng" And CStr(result(2, 2)) = "(alice bob)" _
                     And CStr(result(3, 1)) = "sales" And CStr(result(3, 2)) = "(carol)")
    End If
    Report "prolog.5.3: findall composes with an outer already-bound variable across real backtracking (group-by)", _
           groupByOk, "got: " & ResultDescribe(result)

    ' An already-bound Bag argument is genuinely CHECKED via unification,
    ' not just always bound fresh - both the matching and mismatching
    ' case.
    result = VLA_Prolog.PROLOG("(fact (color red)) (fact (color green)) (query (findall X (color X) (red green)))")
    Report "prolog.5.3: an already-bound Bag matching the harvested list succeeds", _
           VarType(result) = vbBoolean And result = True, "got " & TypeName(result) & " " & result

    result = VLA_Prolog.PROLOG("(fact (color red)) (fact (color green)) (query (findall X (color X) (green red)))")
    Report "prolog.5.3: an already-bound Bag NOT matching the harvested list (wrong order) fails", _
           VarType(result) = vbBoolean And result = False, "got " & TypeName(result) & " " & result

    ' The shared step ceiling, run to real exhaustion inside findall's
    ' own Goal - the identical sub-call microscope `not` already proves
    ' this for.
    r = CStr(VLA_Prolog.PROLOG("(rule (loop X) (loop X)) (query (findall X (loop a) Bag))"))
    Report "prolog.5.3: a genuinely non-terminating goal inside findall is refused by the shared resolution-step ceiling, not left to hang or crash", _
           InStr(1, r, "resolution steps", vbTextCompare) > 0, "got: " & r

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
    Report "prolog.5.3: a Template containing a literal (not X)-shaped sub-term still has its own variable collected correctly", _
           ResultCol1Is(result, "((not a))"), "got: " & ResultDescribe(result)

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
    Report "prolog.5.3: a bare ! as findall's own goal is legal (PROLOG.5.4) and harvests exactly one Template instantiation", _
           ResultCol1Is(result, "(done)"), "got: " & ResultDescribe(result)
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
    If IsArray(result) Then commitOk = (UBound(result, 1) = 2 And CStr(result(2, 1)) = "a")
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
    If IsArray(result) Then scopeOk = (UBound(result, 1) = 2 And CStr(result(2, 1)) = "a" And CStr(result(2, 2)) = "1")
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
        boundaryOk = (UBound(result, 1) = 3 And CStr(result(2, 1)) = "a" And CStr(result(2, 2)) = "1" _
                      And CStr(result(3, 1)) = "a" And CStr(result(3, 2)) = "2")
    End If
    Report "prolog.5.4: cut does not escape its own clause to prune a caller's later choice point, and goals to cut's own right keep full backtracking", _
           boundaryOk, "got: " & ResultDescribe(result)

    ' Cut's own opacity across `not`'s isolated sub-search (real Prolog's
    ' own rule): blocked/0 reaches a cut (pruning thing's own second
    ' fact) and then FAILS outright (nonexistent/1 has no facts), so
    ' blocked/0 has zero solutions and (not (blocked)) is always TRUE,
    ' regardless of Y. If the internal cut leaked out of the isolated
    ' sub-search, the outer p(Y) loop would wrongly stop after Y=a; a
    ' correctly-opaque cut leaves it fully backtracking over both facts.
    result = VLA_Prolog.PROLOG( _
        "(fact (p a)) (fact (p b)) (fact (thing 1)) (fact (thing 2)) " & _
        "(rule (blocked) (thing W) ! (nonexistent Z)) " & _
        "(query (p Y) (not (blocked)))")
    Dim notOpaqueOk As Boolean
    If IsArray(result) Then notOpaqueOk = (UBound(result, 1) = 3 And CStr(result(2, 1)) = "a" And CStr(result(3, 1)) = "b")
    Report "prolog.5.4: a cut fired inside not's own isolated goal never leaks out to prune the outer query's own unrelated choice point", _
           notOpaqueOk, "got: " & ResultDescribe(result)

    ' The identical opacity proof for findall's own isolated sub-search:
    ' grab/1 reaches a cut after its first thing-fact, so grab(Z) always
    ' harvests exactly (1), regardless of Y - if the internal cut leaked,
    ' the outer p(Y) loop would wrongly stop after Y=a.
    result = VLA_Prolog.PROLOG( _
        "(fact (p a)) (fact (p b)) (fact (thing 1)) (fact (thing 2)) " & _
        "(rule (grab X) (thing X) !) " & _
        "(query (p Y) (findall Z (grab Z) Bag))")
    Dim findallOpaqueOk As Boolean
    If IsArray(result) Then
        findallOpaqueOk = (UBound(result, 1) = 3 And CStr(result(2, 1)) = "a" And CStr(result(2, 2)) = "(1)" _
                           And CStr(result(3, 1)) = "b" And CStr(result(3, 2)) = "(1)")
    End If
    Report "prolog.5.4: a cut fired inside findall's own isolated goal never leaks out to prune the outer query's own unrelated choice point", _
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
    If IsArray(result) Then multiCutOk = (UBound(result, 1) = 2 And CStr(result(2, 1)) = "1" And CStr(result(2, 2)) = "10")
    Report "prolog.5.4: two cuts in one clause body compose correctly, sharing one barrier, without double-firing", _
           multiCutOk, "got: " & ResultDescribe(result)

    ' The shared step ceiling, re-proven with the two new cut-signal
    ' parameters now threaded through every recursive frame - the
    ' identical genuinely non-terminating rule PROLOG.4/5.2/5.3 each
    ' already run to real exhaustion, with no cut anywhere in it, so this
    ' isolates whether cut's own new machinery alone changed the ceiling.
    r = CStr(VLA_Prolog.PROLOG("(rule (loop X) (loop X)) (query (loop a))"))
    Report "prolog.5.4: a genuinely non-terminating rule (no cut involved) is still refused cleanly by the shared step ceiling with cut's new parameters threaded through every frame", _
           InStr(1, r, "resolution steps", vbTextCompare) > 0, "got: " & r
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
           UBound(result, 1) = 3 And CStr(result(2, 1)) = "alice" And CStr(result(3, 1)) = "carol", _
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
           UBound(result, 1) = 2 And CStr(result(2, 1)) = "1" And CStr(result(2, 2)) = "2", _
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
           UBound(result, 1) = 3 And CStr(result(2, 1)) = "alice" And CStr(result(3, 1)) = "carol", _
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
           VarType(result) = vbBoolean And result = False, "got " & TypeName(result) & " " & result

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
    Report "prolog.6: keyed-atom desugaring reaches inside findall's own Goal", _
           ResultCol1Is(result, "(" & q & "alice" & q & " " & q & "carol" & q & ")"), "got: " & ResultDescribe(result)
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
        ok2 = (UBound(arr2, 1) = 2 And CStr(arr2(2, 1)) = "90000" And CStr(arr2(2, 2)) = "eng")
        detail2 = "got S=" & arr2(2, 1) & " D=" & arr2(2, 2)
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
        ok4 = (UBound(arr4, 1) = 2 And CStr(arr4(2, 1)) = "Alice")
        detail4 = "got: " & ResultDescribe(arr4)
    Else
        detail4 = "got " & TypeName(arr4) & " = " & CStr(arr4)
    End If
    Report "prolog host: named-column (keyed-atom) syntax works end-to-end through the real =PROLOG(...) function against a live Table", _
           ok4, detail4

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
           VarType(noTableResult) = vbString And Left$(CStr(noTableResult), 5) = "#SQL!", _
           "got " & TypeName(noTableResult) & " = " & CStr(noTableResult)
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
                And (CStr(arrAll(r0, c0)) = "Name") _
                And (CStr(arrAll(r0, c0 + 1)) = "Dept") _
                And (CStr(arrAll(r0, c0 + 2)) = "Salary")
        detailAll = "got " & nRowsAll & " rows x " & nColsAll & " cols, headers '" & _
                    CStr(arrAll(r0, c0)) & "','" & CStr(arrAll(r0, c0 + 1)) & "','" & CStr(arrAll(r0, c0 + 2)) & "'"
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
        okJoin = (nRowsJoin = 2) And (CStr(arrJoin(rJ + 1, cJ + 1)) = "Bldg5")
        detailJoin = "got " & nRowsJoin & " rows, building '" & CStr(arrJoin(rJ + 1, cJ + 1)) & "'"
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
        okGroup = (nRowsGroup = 2) And (CStr(arrGroup(rG + 1, cG)) = "eng") _
                  And (CDbl(arrGroup(rG + 1, cG + 1)) = 2) And (CDbl(arrGroup(rG + 1, cG + 2)) = 185000)
        detailGroup = "got " & nRowsGroup & " rows, dept '" & CStr(arrGroup(rG + 1, cG)) & _
                      "', n=" & CStr(arrGroup(rG + 1, cG + 1)) & ", total=" & CStr(arrGroup(rG + 1, cG + 2))
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
        okOrder = (nRowsOrder = 3) And (CStr(arrOrder(rO + 1, cO)) = "carol") And (CStr(arrOrder(rO + 2, cO)) = "alice")
        detailOrder = "got " & nRowsOrder & " rows (incl. header), top two: '" & _
                      CStr(arrOrder(rO + 1, cO)) & "', '" & CStr(arrOrder(rO + 2, cO)) & "'"
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
                     And (CDbl(arrCompound(rC + 1, cC)) = 3) _
                     And (CDbl(arrCompound(rC + 2, cC)) = 2) _
                     And (CDbl(arrCompound(rC + 3, cC)) = 1)
        detailCompound = "got " & nRowsCompound & " rows (incl. header): " & _
                          CStr(arrCompound(rC + 1, cC)) & ", " & _
                          CStr(arrCompound(rC + 2, cC)) & ", " & _
                          CStr(arrCompound(rC + 3, cC))
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
                  And (CStr(arrChain(rCh + 1, cCh)) = "alice") _
                  And (CStr(arrChain(rCh + 2, cCh)) = "bob") _
                  And (CStr(arrChain(rCh + 3, cCh)) = "carol")
        detailChain = "got " & nRowsChain & " rows (incl. header): " & _
                      CStr(arrChain(rCh + 1, cCh)) & ", " & CStr(arrChain(rCh + 2, cCh)) & ", " & CStr(arrChain(rCh + 3, cCh))
    Else
        detailChain = "got " & TypeName(arrChain) & " = " & CStr(arrChain)
    End If
    Report "sql host: a live recursive WITH, through =SQL(...), finds dave's whole chain of command (alice, bob, carol)", okChain, detailChain

    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
    prior.Activate
End Sub
