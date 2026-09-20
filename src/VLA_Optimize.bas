Attribute VB_Name = "VLA_Optimize"
Option Explicit
Public Const VLA_OPTIMIZE_VERSION As String = "OPTIMIZE.2"

' =====================================================================
'  VLA_Optimize - OPTIMIZE.2: constraints against the one world.
'  OPTIMIZE.1's base case, plus checks that can only eliminate.
'
'  WHAT OPTIMIZE.2 ADDS, and it is reachability rather than wording.
'  With no choice, OPTIMIZE.1's fixpoint produces exactly ONE world, so
'  (forbid BODY...) and (require CONSEQUENT BODY...) are CHECKS over
'  that world - never a reason to try another, since there is none. A
'  check that holds anywhere is a VIOLATION, and a program with a
'  violation answers VLA_OPTIMIZE_NO_SCHEDULE, whose exact words
'  OPTIMIZE.1 reserved and this item may not reword. What this item
'  owns is naming the violating CONSTRAINT and ROWS beside those words,
'  and never letting any of it become a #OPTIMIZE! refusal.
'
'  HOW A CONSTRAINT IS GROUND, and this was the item's architectural
'  fork (the owner's call, 2026-09-19). The roadmap asked for a body
'  "instantiated over the world by the same join DATALOG runs for a
'  rule body", and EvalRuleBody/RunOneRulePass/RunFixpointForRules are
'  all Private in VLA_Datalog.bas. Rather than expose the join, move it
'  to VLA_Relation.bas (blocked outright: that file is LAYER 0 and may
'  not call VLA_Messages, which the join does, a dozen times) or
'  duplicate it here (where it would drift), a constraint is REWRITTEN
'  INTO AN ORDINARY RULE and handed back to DATALOG:
'
'    (forbid B1 B2)          ->  (rule (vla-check-1 "1" V...) B1 B2)
'    (require CONS B...)     ->  (rule (vla-check-2 "2" V...) B... (not CONS))
'    (require GROUND)        ->  (rule (vla-check-3 "3") (not GROUND))
'
'  and a violation is exactly a non-empty derived relation. So the join
'  is not borrowed - the body IS a rule body - and aggregates, keyed
'  atoms, text tests, comparisons, stratification and CheckRuleSafety
'  all arrive for free and cannot drift. The one seam this needed is at
'  the PROGRAM level, not the join: VLA_Datalog.DatalogRunForms, which
'  promises only "a program is a list of forms, and here is the entry
'  that takes them". Everything about grounding stays Private.
'
'  WHY (not CONS) GOES LAST in a rewritten require: CheckRuleSafety
'  walks a body in WRITTEN ORDER, and a negated atom's variables must
'  already be bound when it is reached. Appending it after the whole
'  body is what makes DATALOG's own safety condition say exactly the
'  right thing about a require - an unbound consequent variable is an
'  EXISTENTIAL, which OPTIMIZE.1 settled needs a derived rule and no
'  new form. It is refused here first, in OPTIMIZE's own words
'  (optimize-require-unbound-consequent), so a user who never wrote
'  `not` is never shown a message about negation.
'
'  WHY A GROUND REQUIREMENT IS A CHECK AND NOT A FACT. With a choice,
'  (require (in "carbon-frame")) FORCES an atom, because the world is
'  being built. Here it cannot: the world is the data, and asserting
'  the atom would silently add a row to the user's own Table and report
'  no violation - the exact opposite of a compliance check, which
'  exists to notice that T1 is not approved. So it is checked, never
'  forced, and (not GROUND) is the whole body.
'
'  WHAT IS STILL REFUSED, and why this item did not make it reachable.
'  A program carrying a CHOICE (or a preference, objective, kept
'  schedule or effort) still refuses exactly as at OPTIMIZE.1, before
'  any check runs. The counting pre-checks OPTIMIZE.0's oracle asked
'  for are BUILT and PINNED below and reached by nothing, on the
'  owner's call (2026-09-19): every one of section 17's four
'  NONE-COUNT keys takes its DEMAND from a choice form, so reaching
'  them here would mean inferring demand without the grounder
'  OPTIMIZE.3 builds - a shadow grounder, written to be thrown away -
'  and an arithmetic slip in it would land as a WRONG ANSWER in a cell
'  ("no schedule satisfies every rule" for a roster that has one)
'  rather than as a red pin. OPTIMIZE.3 supplies demand and capacity
'  from its own grounding and calls these same functions unchanged.
'
'  WHAT THIS IS. A program with no choice, no constraint and no
'  objective IS a DATALOG program: one guaranteed, unique world, and no
'  search of any kind. So =OPTIMIZE(...) reads the same
'  (fact ...)/(rule ...)/(query ...) forms through VLA.VlaReadForms,
'  hands them to VLA_Datalog.DatalogRun WHOLESALE, and spills the
'  relation the query names, headers first, exactly as DATALOG does.
'  That is standing decision 1's shape ("one search, many views") from
'  the first day, because the decision relation of a real program is
'  just the relation a query names.
'
'  WHAT IT SETTLES RATHER THAN EXECUTES. The six forms of the three
'  ingredients are PARSED here - their shapes checked, their spellings
'  fixed - and then refused by name, so OPTIMIZE.2 and OPTIMIZE.3 add
'  machinery and not grammar, and so nothing a user writes today has to
'  change spelling later (SD-4). The five result states are reserved
'  here for the same reason: no later item may invent a sixth.
'
'  THE SPELLINGS, settled 2026-09-19 (the owner's calls), and the three
'  SHAPES settled by this item's own scoping run against
'  scripts/pareto_logic.txt section 17, whose twenty-one entries every
'  form below has to be able to say plainly:
'
'    (choose-exactly  N chosen pool [(per group...)])
'    (choose-at-least N chosen pool [(per group...)])
'    (choose-at-most  N chosen pool [(per group...)])
'    (choose-between  LO HI chosen pool [(per group...)])
'    (choose-any      chosen pool)
'    (require CONSEQUENT BODY...)          (forbid BODY...)
'    (prefer BODY... [(cost T)])           (avoid BODY... [(cost T)])
'    (minimize BODY... [(cost T)])         (maximize BODY... [(cost T)])
'    (minimise ...)                        (maximise ...)
'    (fewest-changes-from TableName)
'    (effort quick|normal|thorough)   or   (effort N)
'
'  and OPTIMIZE_STATUS(rules, tables...) as a companion FUNCTION rather
'  than a (status) form, since status is a question about the answer and
'  not a rule. A bare (choose ...) is refused by name, listing the five.
'  A derived rule stays `rule`, unchanged from DATALOG.
'
'  WHY (per ...) AND WHY choose-any - the two shapes section 17 forced,
'  and the only places this item's scoping changed anything:
'    - The roadmap's own three-slot example cannot say the KILLER CASE's
'      own choice line. "Every shift gets exactly the people it needs"
'      takes its count from the Need column of a Table, so N must be a
'      variable, and a three-slot form has nowhere to bind one.
'      (per ...) binds it, and holds as many group atoms as the
'      sentence needs without nesting - optimize-duty-rotation groups by
'      two ("every month, every person gets exactly one duty"). With no
'      (per ...) the count is over the whole pool at once.
'    - "Each extra is in the quote or not" (optimize-quote-bike) and
'      "each link is switched on, or not" (optimize-network-connect) are
'      per-POOL-ROW choices with no group and no count.
'      `choose-at-least 0` reads as a no-op per group where the sentence
'      needs "at most one per row", so it would have had to mean
'      something other than it says - the test the other four passed.
'      Hence a fifth form, and the four keep their meanings.
'
'  WHAT IS NOT HERE, on purpose: any search at all; a (status) form; a
'  lock form ("these shifts are locked" is `require` over a Table of
'  locked rows - the owner's call, 2026-09-19); and any number behind
'  the three effort levels, which OPTIMIZE.3 and OPTIMIZE.6 set once
'  they have measured. A number published before it means anything is a
'  number users tune around.
'
'  LAYER:     Engine (beside VLA_Datalog, VLA_Sql, VLA_Prolog)
'  MAY CALL:  VLA (VlaReadForms, and VlaWriteForm for the one line of
'             DISPLAY text the violations table carries - never on the
'             path between what a user wrote and what this engine
'             answers, which is what DatalogRunForms exists to keep
'             clear), VLA_Datalog (DatalogRunForms - the whole engine),
'             VLA_Relation (TableArgResolve/RelFromRange/
'             RangeColumnNames/RelToSpilledArray/RelNew/RelTryAdd/
'             RelArity/RelCount/RelTuples, and RaiseTableArgRefusal for
'             a reason no wrapper here words itself), VLA_Digest (the
'             memo key), VLA_Identity (Fold), VLA_Messages (every
'             refusal), VLA_Runtime (VlaDict*)
'  SHIPS:     add-in (VLA_Build.bas's own mods array) and the dev rig
'  PAYS INTO: OPTIMIZE.2 through OPTIMIZE.10, and G-OPTIMIZE, which may
'             only ever write the shapes named above.
'  REASON:    standing decision 4, 2026-09-18: VLA_Optimize.bas is the
'             engine's home; substrate any engine could use lives in
'             VLA_Relation.bas (DATALOG.15's spill-as-table read is the
'             first of it, built ahead of this item); every optimize-*
'             refusal id lives in VLA_Messages.bas.
'
'  WHAT CHANGED IN VLA_Datalog.bas, and it is three things and no more
'  (OPTIMIZE.1 changed nothing at all, and said so): DatalogRun split
'  into a two-line reader and DatalogRunForms; ParseProgram takes the
'  forms the reader produced; and DefinedNamesSentence never offers a
'  name under the reserved `vla-` prefix, so a user who mistyped a
'  predicate is never told to match it against `vla-check-1`. Decision
'  4's "coordinate per change" is paid here rather than dodged.
'
'  THE RESERVED PREFIX. Every name this engine generates is
'  `vla-check-N`, N being the constraint's own 1-based position in
'  WRITTEN ORDER among the constraints - deterministic, hand-derivable
'  from the program text alone, and injective, since no two
'  constraints share a position. A user who writes a predicate under
'  `vla-check-` is refused by name (optimize-reserved-predicate), so
'  the generated name can never collide with one a person meant; and an
'  Excel Table cannot be named with a hyphen at all, so a table
'  argument could not reach it even if OPTIMIZE allowed it. This is
'  DATALOG.5's own AnonymousColumnVarName discipline - derived from
'  written position, never a counter or a gensym - applied one level up.
'
'  `vla-check-` AND NOT `vla-`, which the scoping run found the hard
'  way: `vla-` is the project's umbrella for generated identifiers and
'  is ALREADY OCCUPIED. The sentence layer emits `vla-ask-can-cover`
'  and `vla-not-leave-name-shift` into ordinary DATALOG programs, two
'  of which the parity table carries, so reserving the umbrella here
'  would have refused a program DATALOG answers - and would have made
'  this engine refuse its own generator's output the day a Frazaro
'  sentence compiles to an OPTIMIZE call. One sub-prefix per generator
'  is the project's rule, and this is this generator's.
' =====================================================================

' ---- the five result states, reserved here -------------------------
'
' Reserved, not all reachable: this version only ever answers
' PROVEN_BEST, because a program with no choice has exactly one world.
' The other four have their words pinned so a later item inherits a
' sentence rather than inventing one, and so the name's own honesty rule
' - an answer may never be CALLED best unless it is proven best - is
' written down before anything can break it.
Public Const VLA_OPTIMIZE_REFUSED As Long = 0
Public Const VLA_OPTIMIZE_NO_SCHEDULE As Long = 1
Public Const VLA_OPTIMIZE_NONE_IN_BUDGET As Long = 2
Public Const VLA_OPTIMIZE_BEST_IN_BUDGET As Long = 3
Public Const VLA_OPTIMIZE_PROVEN_BEST As Long = 4

' ---- the three effort levels ---------------------------------------
'
' Deliberately ZERO, every one of them. The level a program names is
' recorded and the form is then refused, so no default is observable
' yet; OPTIMIZE.3 and OPTIMIZE.6 set these from their own
' measurements, in units of WORK (decisions and conflicts), never
' seconds - standing decision 2, so that a faster machine proves the
' same thing a slower one does.
Public Const VLA_OPTIMIZE_WORK_QUICK As Long = 0
Public Const VLA_OPTIMIZE_WORK_NORMAL As Long = 0
Public Const VLA_OPTIMIZE_WORK_THOROUGH As Long = 0

' The level a program with no (effort ...) form gets, written down now
' so the day a number lands behind it, the default is already stated.
Public Const VLA_OPTIMIZE_EFFORT_DEFAULT As String = "normal"

' ---- the form kinds, for the refusal that names the right one ------
'
' OPT_KIND_CONSTRAINT is gone with optimize-constraint-not-yet, which
' this item made false: a constraint is no longer an unbuilt
' ingredient, it is the thing this item builds. The numbers of the
' others are left exactly where they were rather than closed up, since
' nothing outside this module reads them and a renumber would be a
' diff with no meaning in it.
Private Const OPT_KIND_NONE As Long = 0
Private Const OPT_KIND_CHOICE As Long = 1
Private Const OPT_KIND_PREFERENCE As Long = 3
Private Const OPT_KIND_OBJECTIVE As Long = 4
Private Const OPT_KIND_KEPT As Long = 5
Private Const OPT_KIND_EFFORT As Long = 6

' ---- the generated name, and the violations table ------------------
'
' The prefix every name this engine invents lives under. See the
' module header: derived from a constraint's written position, refused
' at parse when a user writes one, and unreachable by an Excel Table
' name, which may not contain a hyphen.
Private Const OPT_CHECK_PREFIX As String = "vla-check-"

' The violations table's own three columns, fixed for every program so
' that a formula reading the spill has a stable shape whatever the
' rules say. A constraint's arity varies; these do not.
Private Const OPT_VIOL_COL_CHECK As String = "Check"
Private Const OPT_VIOL_COL_RULE As String = "Rule"
Private Const OPT_VIOL_COL_WHERE As String = "Where"

' How many broken checks OPTIMIZE_STATUS names before it stops
' counting them out - DefinedNamesSentence's own cap, for the same
' reason: a sentence is read, and a hundred of them is not a sentence.
Private Const OPT_STATUS_CHECK_CAP As Long = 6

' ---- the session memo ----------------------------------------------
'
'  Standing decision 1's mechanism, and the one piece of state in this
'  module. Keyed on the CONTENT of the rules text and of every input
'  Table (VLA_Digest's step-1 key, which is injective), so a
'  recalculation whose inputs did not change costs a hash, and a second
'  cell asking for the same answer's STATUS costs nothing.
'
'  IT MUST TOLERATE BEING LOST AT ANY MOMENT, and this is measured, not
'  assumed: OPTIMIZE.0's probe found that a VBA project reset wipes
'  every module variable - editing the project does it, and so does the
'  End button on the Esc dialog - and that after a reset the next
'  structural change re-ran EVERY UDF in the workbook (C7). So the memo
'  may only make an answer FASTER and may never change what it is, which
'  is what decision 2 requires anyway. Losing it costs time and nothing
'  else, and the pins cover exactly that: clear it between two calls and
'  the answer is the same.
'
'  A refusal is never memoized. Only an answer is, so a program that
'  cannot be read pays for its refusal every time - which is right, since
'  a refusal is cheap and a stale one would be a lie.
Private Const OPT_MEMO_CAP As Long = 16
Private mMemo As Object
Private mMemoOrder As Collection
Private mMemoRuns As Long

' =====================================================================
'  THE WORKSHEET SURFACE
' =====================================================================

' =OPTIMIZE(rules, table1, table2, ...) - the same argument shape as
' =DATALOG(...), and for the same reasons (see VLA_Datalog.DATALOG's own
' header for the full account of rulesText, Excel Tables, plain named
' ranges, a spilled range's name and header row, and column-scoped
' table arguments; every one of them behaves identically here, because
' it is the identical substrate underneath).
'
' A refusal returns as readable TEXT in the cell ("#OPTIMIZE! ...")
' rather than Excel's own opaque #VALUE!, the convention all four
' engines share. Note which word the prefix carries: the refusal's
' WORDS may be DATALOG's, where the shared engine raised them, and the
' prefix still says which function the user called.
Public Function OPTIMIZE(ByVal rulesText As String, ParamArray tables() As Variant) As Variant
    On Error GoTo fail
    Dim relations As Object
    Dim headerMap As Object
    Dim i As Long
    Set relations = VLA_Runtime.VlaDictNew()
    Set headerMap = VLA_Runtime.VlaDictNew()
    For i = LBound(tables) To UBound(tables)
        Dim nm As String
        nm = OptimizeTableArgName(tables(i))
        VLA_Runtime.VlaDictSet relations, nm, VLA_Relation.RelFromRange(tables(i))
        Dim colsOk As Boolean
        Dim cols As Collection
        Set cols = VLA_Relation.RangeColumnNames(tables(i), colsOk)
        If colsOk Then VLA_Runtime.VlaDictSet headerMap, nm, cols
    Next i

    Dim result As Collection
    Set result = OptimizeRun(rulesText, relations, headerMap)
    OPTIMIZE = OptimizeAnswerOf(result)
    Exit Function
fail:
    OPTIMIZE = "#OPTIMIZE! " & Err.Description
End Function

' =OPTIMISE(...) - a one-line alias forwarding to the same engine, so a
' British user's formula works. Excel's own functions have no British
' spellings; this is a courtesy, not a convention, and it is a courtesy
' rather than a second implementation on purpose: there is exactly one
' engine, one memo and one set of refusals behind both names.
Public Function OPTIMISE(ByVal rulesText As String, ParamArray tables() As Variant) As Variant
    On Error GoTo fail
    Dim relations As Object
    Dim headerMap As Object
    Dim i As Long
    Set relations = VLA_Runtime.VlaDictNew()
    Set headerMap = VLA_Runtime.VlaDictNew()
    For i = LBound(tables) To UBound(tables)
        Dim nm As String
        nm = OptimizeTableArgName(tables(i))
        VLA_Runtime.VlaDictSet relations, nm, VLA_Relation.RelFromRange(tables(i))
        Dim colsOk As Boolean
        Dim cols As Collection
        Set cols = VLA_Relation.RangeColumnNames(tables(i), colsOk)
        If colsOk Then VLA_Runtime.VlaDictSet headerMap, nm, cols
    Next i

    Dim result As Collection
    Set result = OptimizeRun(rulesText, relations, headerMap)
    OPTIMISE = OptimizeAnswerOf(result)
    Exit Function
fail:
    OPTIMISE = "#OPTIMIZE! " & Err.Description
End Function

' =OPTIMIZE_STATUS(rules, table1, ...) - the words about the same
' answer: whether it is proven best, the best found within a budget, or
' that no schedule satisfies every rule.
'
' An underscore rather than Excel's own dot (FORECAST.ETS.STAT): a VBA
' procedure name cannot contain a dot.
'
' IT COMPUTES THE ANSWER ITSELF (the owner's call, 2026-09-19, this
' item's scoping fork 3) rather than reading a memo the sibling cell
' may never have filled. That is what makes it a PURE FUNCTION of its
' own arguments: it cannot depend on calculation order, on whether an
' OPTIMIZE cell exists anywhere, or on what ran before it in the
' session - and it answers the same thing after a project reset has
' wiped the memo. The alternatives were a "not calculated yet" answer
' and a refusal, and both would make a cell's value depend on history,
' which is decision 2 broken by the side door (OPTIMIZE.7's warm-start
' trap, arriving through a status cell). The price is stated plainly: a
' status cell asked with a cold memo pays for the answer. With a warm
' one - the ordinary case, since the OPTIMIZE cell beside it filled the
' memo - it costs a hash.
Public Function OPTIMIZE_STATUS(ByVal rulesText As String, ParamArray tables() As Variant) As Variant
    On Error GoTo fail
    Dim relations As Object
    Dim headerMap As Object
    Dim i As Long
    Set relations = VLA_Runtime.VlaDictNew()
    Set headerMap = VLA_Runtime.VlaDictNew()
    For i = LBound(tables) To UBound(tables)
        Dim nm As String
        nm = OptimizeTableArgName(tables(i))
        VLA_Runtime.VlaDictSet relations, nm, VLA_Relation.RelFromRange(tables(i))
        Dim colsOk As Boolean
        Dim cols As Collection
        Set cols = VLA_Relation.RangeColumnNames(tables(i), colsOk)
        If colsOk Then VLA_Runtime.VlaDictSet headerMap, nm, cols
    Next i

    Dim result As Collection
    Set result = OptimizeRun(rulesText, relations, headerMap)
    OPTIMIZE_STATUS = CStr(result.Item(7))
    Exit Function
fail:
    OPTIMIZE_STATUS = "#OPTIMIZE! " & Err.Description
End Function

' =OPTIMIZE_VIOLATIONS(rules, table1, ...) - which rules are broken,
' and where. The compliance pack's own evidence, and the reason this
' is a third FUNCTION rather than something in the OPTIMIZE cell (the
' owner's call, 2026-09-19).
'
' THE TENSION IT RESOLVES, stated because it is real. Standing
' decision 1 says the OPTIMIZE cell spills the DECISION. OPTIMIZE.1
' reserved the "no schedule" answer's shape as the header row with
' nothing under it, measured: a word in the cell would not spill at
' all, so a reader's N2# becomes #REF!, and a word in a ROW would be
' counted as data. But Contemplation 8 and section 17's
' optimize-sod-check both say the valuable answer to a zero-choice
' check IS the violating rows, named. One cell cannot be both without
' changing SHAPE between runs - schedule columns when satisfiable,
' violation columns when not - which breaks every formula reading the
' spill on precisely the run that matters, and makes "no schedule"
' indistinguishable by shape from an answer. So the evidence gets its
' own cell, and the OPTIMIZE cell keeps the shape OPTIMIZE.1 reserved.
'
' THREE COLUMNS, FIXED FOR EVERY PROGRAM: Check (the constraint's own
' 1-based written position, which is also the N in its vla-check-N),
' Rule (the constraint exactly as the user wrote it, on one line), and
' Where (the variable bindings that made it hold). A constraint's own
' arity varies from program to program and from rule to rule; these
' three do not, so a FILTER or a COUNTIFS over this spill is written
' once and keeps working.
'
' With nothing broken it spills its own header row and nothing under
' it - the same shape, for the same measured reason.
'
' It COMPUTES THE ANSWER ITSELF, exactly as OPTIMIZE_STATUS does and
' for the identical reason: a pure function of its own arguments,
' answering the same after a project reset wiped the memo.
Public Function OPTIMIZE_VIOLATIONS(ByVal rulesText As String, ParamArray tables() As Variant) As Variant
    On Error GoTo fail
    Dim relations As Object
    Dim headerMap As Object
    Dim i As Long
    Set relations = VLA_Runtime.VlaDictNew()
    Set headerMap = VLA_Runtime.VlaDictNew()
    For i = LBound(tables) To UBound(tables)
        Dim nm As String
        nm = OptimizeTableArgName(tables(i))
        VLA_Runtime.VlaDictSet relations, nm, VLA_Relation.RelFromRange(tables(i))
        Dim colsOk As Boolean
        Dim cols As Collection
        Set cols = VLA_Relation.RangeColumnNames(tables(i), colsOk)
        If colsOk Then VLA_Runtime.VlaDictSet headerMap, nm, cols
    Next i

    Dim result As Collection
    Set result = OptimizeRun(rulesText, relations, headerMap)
    OPTIMIZE_VIOLATIONS = VLA_Relation.RelToSpilledArray(result.Item(8), _
        Array(OPT_VIOL_COL_CHECK, OPT_VIOL_COL_RULE, OPT_VIOL_COL_WHERE))
    Exit Function
fail:
    OPTIMIZE_VIOLATIONS = "#OPTIMIZE! " & Err.Description
End Function

' =====================================================================
'  THE ENGINE SEAM
' =====================================================================

' The pure entry point every test calls, and what the three worksheet
' functions above share. Returns a Collection whose FIRST FIVE items are
' DatalogRun's own, item for item and unchanged - which is what lets the
' parity pin compare the two engines directly rather than through a
' translation - plus two of this engine's own:
'
'   1 queryName        4 headless
'   2 relations        5 the Boolean answer, or Empty
'   3 head names, or Empty
'   6 the result state (one of the five VLA_OPTIMIZE_* constants)
'   7 the status words
'   8 OPTIMIZE.2: the violations, as a three-column Relation
'
' Items 6 upward are this engine's own and the parity pin never reads
' them; items 1 to 5 stay DatalogRun's, item for item, which is what
' lets that pin compare the two engines directly. Adding item 8 could
' not disturb it, and adding the rewritten CHECK RULES to the program
' could not either: they define only vla-check-N, which no parity
' program queries and no parity program names.
Public Function OptimizeRun(ByVal rulesText As String, Optional ByVal baseRelations As Object, _
                            Optional ByVal headerMap As Object) As Collection
    Dim relations As Object
    If baseRelations Is Nothing Then
        Set relations = VLA_Runtime.VlaDictNew()
    Else
        Set relations = baseRelations
    End If
    Dim hMap As Object
    If headerMap Is Nothing Then
        Set hMap = VLA_Runtime.VlaDictNew()
    Else
        Set hMap = headerMap
    End If

    ' THE RUN GETS ITS OWN DICTIONARY, and this is a correctness fix
    ' rather than hygiene (found by this item's own pins, live).
    ' DatalogRunForms writes every relation it derives into the dict it
    ' is given, and this engine now derives `vla-check-N` relations of
    ' its own. A caller that reuses one dict across two calls - which
    ' every pure test does, and which the worksheet functions do not -
    ' would therefore hand the SECOND program the FIRST program's
    ' violations, still sitting under the same generated name, and a
    ' clean program would answer "no schedule satisfies every rule" on
    ' the strength of a rule it does not contain. Copying is the fix,
    ' and it is a shallow copy: the Relation objects are shared exactly
    ' as DATALOG shares them, so nothing else changes.
    '
    ' A reserved name in the caller's own dict is dropped rather than
    ' copied, so a dict already polluted by an older build cannot carry
    ' the same fault back in.
    '
    ' THE MEMO KEY IS TAKEN OVER THIS COPY, deliberately. Keyed over the
    ' caller's dict it would include any stale `vla-check-N` relation,
    ' which makes a cell's key depend on what ran before it - decision 2
    ' broken by the side door, and the same hazard OPTIMIZE_STATUS
    ' computing its own answer exists to avoid. For a clean dict the two
    ' are byte-identical, since a copy preserves insertion order.
    Dim working As Object
    Set working = VLA_Runtime.VlaDictNew()
    Dim baseName As Variant
    For Each baseName In VLA_Runtime.VlaDictKeys(relations)
        If Not IsGeneratedCheckName(CStr(baseName)) Then
            VLA_Runtime.VlaDictSet working, CStr(baseName), _
                VLA_Runtime.VlaDictGet(relations, baseName)
        End If
    Next baseName
    Set relations = working

    ' The memo is consulted before the program is even read: the key is
    ' a hash of the text and the Tables, so a hit means this exact
    ' question was answered in this session, and the answer cannot have
    ' changed (decision 2). A miss costs one hash.
    Dim memoKey As String
    memoKey = OptimizeMemoKey(rulesText, relations, hMap)
    If MemoHas(memoKey) Then
        Set OptimizeRun = MemoGet(memoKey)
        Exit Function
    End If

    ' The program is read ONCE, here, and the forms travel the rest of
    ' the way as objects. Every form's SHAPE is checked, the five
    ' ingredients this version still does not execute are refused by
    ' name, and the constraints are collected in written order.
    Dim forms As Collection
    Set forms = VLA.VlaReadForms(rulesText)
    Dim constraints As Collection
    ScanProgram forms, constraints

    ' Each constraint becomes an ordinary rule whose head collects the
    ' rows that violate it, and checkVars carries the head's own
    ' variable names so that ONE walk decides both what the head holds
    ' and what the violations table says it means. Two walks could
    ' disagree and mislabel a column; one cannot.
    Dim checkVars As Collection
    Dim program As Collection
    Set program = RewriteProgram(forms, constraints, checkVars)

    Dim r As Collection
    Set r = VLA_Datalog.DatalogRunForms(program, relations, hMap)
    mMemoRuns = mMemoRuns + 1

    Dim violations As Collection
    Set violations = CollectViolations(constraints, checkVars, r.Item(2))
    Dim stateId As Long
    If VLA_Relation.RelCount(violations) > 0 Then
        stateId = VLA_OPTIMIZE_NO_SCHEDULE
    Else
        stateId = VLA_OPTIMIZE_PROVEN_BEST
    End If

    Dim outp As Collection
    Set outp = New Collection
    Dim item3 As Variant
    outp.Add r.Item(1)
    outp.Add r.Item(2)
    CopyVariant item3, r.Item(3)
    outp.Add item3
    outp.Add r.Item(4)
    outp.Add r.Item(5)
    outp.Add stateId
    outp.Add StatusSentence(stateId, constraints, violations)
    outp.Add violations

    MemoPut memoKey, outp
    Set OptimizeRun = outp
End Function

' The answer a cell holds, from OptimizeRun's own result: a Boolean when
' the query was one fact written out whole (DATALOG.9), otherwise the
' queried relation spilled headers-first.
'
' A "no schedule" answer is THE HEADER ROW WITH NOTHING UNDER IT - the
' owner's call, and it needs no code of its own, because it is exactly
' what a DATALOG query with no rows already answers. The reasoning
' matters for every later item: a single word in the cell ("None") would
' not spill at all, so a reader's N2# becomes #REF! (measured live,
' DATALOG.15 step 12); a word in a ROW would become data - a person
' named None, counted by COUNTIFS. The words live in OPTIMIZE_STATUS. A
' refusal (malformed input) stays #OPTIMIZE! text, where breaking the
' readers is the right thing.
Private Function OptimizeAnswerOf(ByVal result As Collection) As Variant
    ' OPTIMIZE.2: when a check is broken there IS no schedule, so the
    ' queried relation's own rows are not the answer - they are the
    ' data of a world that satisfies no rule, and spilling them would
    ' hand a reader a roster that is illegal. The reserved shape goes
    ' out instead: the same headers, with nothing under them.
    '
    ' A query written as ONE FACT (DATALOG.9) has no header row to
    ' empty, so it answers FALSE - the same "the empty answer, in this
    ' shape" the spill gives, said the only way a Boolean can say it.
    ' Which rule is broken, and where, is OPTIMIZE_STATUS's and
    ' OPTIMIZE_VIOLATIONS's to say, and they say it either way.
    Dim broken As Boolean
    broken = (CLng(result.Item(6)) = VLA_OPTIMIZE_NO_SCHEDULE)
    If Not IsEmpty(result.Item(5)) Then
        If broken Then
            OptimizeAnswerOf = False
        Else
            OptimizeAnswerOf = result.Item(5)
        End If
        Exit Function
    End If
    Dim queried As Collection
    Set queried = VLA_Runtime.VlaDictGet(result.Item(2), CStr(result.Item(1)))
    If broken Then Set queried = VLA_Relation.RelNew(VLA_Relation.RelArity(queried))
    Dim isHeadless As Boolean
    isHeadless = result.Item(4)
    If IsEmpty(result.Item(3)) Then
        OptimizeAnswerOf = VLA_Relation.RelToSpilledArray(queried, , isHeadless)
    Else
        OptimizeAnswerOf = VLA_Relation.RelToSpilledArray(queried, result.Item(3), isHeadless)
    End If
End Function

' The exact sentence each result state says. Reserved for all five, and
' the wording is the contract: OPTIMIZE.2 must say the first, OPTIMIZE.3
' the second, OPTIMIZE.6 the last two, and none of them may reword one.
' The name's own honesty rule lives in the difference between the last
' two.
Public Function OptimizeStatusWords(ByVal stateId As Long) As String
    Select Case stateId
    Case VLA_OPTIMIZE_REFUSED
        OptimizeStatusWords = "this program could not be read; the cell says why"
    Case VLA_OPTIMIZE_NO_SCHEDULE
        OptimizeStatusWords = "no schedule satisfies every rule"
    Case VLA_OPTIMIZE_NONE_IN_BUDGET
        OptimizeStatusWords = "no schedule found within the budget; there may be one"
    Case VLA_OPTIMIZE_BEST_IN_BUDGET
        OptimizeStatusWords = "best found within the budget, not proven best"
    Case VLA_OPTIMIZE_PROVEN_BEST
        OptimizeStatusWords = "proven best: this program makes no choices, so it has exactly one answer and nothing was searched"
    Case Else
        ' Unreachable while the five above are the only states there
        ' are, which is what reserving them here is for. Named rather
        ' than silently blank, so a sixth state invented by a later item
        ' announces itself.
        VLA_Messages.RaiseMsg "optimize-unknown-result-state", "state", CStr(stateId)
    End Select
End Function

' =====================================================================
'  THE SIX FORMS: parsed, shape-checked, and refused by name
' =====================================================================

' Walks every top-level form once. A form that is DATALOG's own is left
' entirely alone; one of the six is shape-checked here and remembered;
' anything else whose head can be read at all is refused as unknown,
' with the OPTIMIZE form list in the message (which is why this cannot
' simply inherit datalog-unknown-top-form, whose own text names three
' forms). A form whose head CANNOT be read - not a list, empty, a list
' as its head - is left for DatalogRun to refuse in its own words, so
' that "the same forms as DATALOG" stays literally true.
'
' ORDER OF REFUSALS, stated because it is a choice: every form's SHAPE
' is checked first, in written order, and only then is the first of the
' six refused as not-yet-built. So a misspelled form later in the text
' wins over an earlier well-formed one - the shape error is the more
' actionable of the two, and a user fixing it will meet the not-yet
' refusal next.
Private Sub ScanProgram(ByVal forms As Collection, ByRef constraints As Collection)
    Set constraints = New Collection
    Dim firstHead As String
    Dim firstKind As Long
    firstKind = OPT_KIND_NONE
    Dim f As Variant
    For Each f In forms
        RefuseReservedNames f
        Dim head As String
        head = OptTopHead(f)
        If Len(head) > 0 Then
            Dim kind As Long
            kind = OPT_KIND_NONE
            Select Case head
            Case "query"
                ' The one predicate position that is a BARE NAME rather
                ' than a list head, so RefuseReservedNames' own walk
                ' cannot see it. Without this, (query vla-check-1)
                ' spills this engine's internal check relation straight
                ' into a cell. A query written as one fact - DATALOG.9's
                ' (query (p a)) - puts a LIST here instead, and the walk
                ' has already read its head.
                Dim qForm As Collection
                Set qForm = f
                If qForm.Count >= 2 Then
                    Dim qRaw As Variant
                    NthInto qRaw, qForm, 2
                    RefuseReservedName qRaw
                End If
            Case "fact", "rule", "headless"
                ' DATALOG's own, and untouched. `rule` stays `rule`
                ' (the owner's call): a rule whose body reads chosen
                ' rows means what it says, and making users write
                ' `recursive-rule` would freeze an implementation
                ' detail into their workbooks. What OPTIMIZE.5 refuses
                ' is narrower - see optimize-recursion-through-choices,
                ' reserved below and unreachable until a choice runs.
            Case "choose"
                VLA_Messages.RaiseMsg "optimize-choose-bare"
            Case "choose-exactly", "choose-at-least", "choose-at-most"
                CheckChoiceForm f, head, 1
                kind = OPT_KIND_CHOICE
            Case "choose-between"
                CheckChoiceForm f, head, 2
                kind = OPT_KIND_CHOICE
            Case "choose-any"
                CheckChooseAnyForm f, head
                kind = OPT_KIND_CHOICE
            Case "require"
                ' OPTIMIZE.2: a constraint is no longer an unbuilt
                ' ingredient. Its shape is checked here and it is kept,
                ' in written order, because that order is the N in its
                ' own vla-check-N and the Check column of the
                ' violations table.
                CheckRequireForm f, head
                constraints.Add f
            Case "forbid"
                CheckForbidForm f, head
                constraints.Add f
            Case "prefer", "avoid"
                CheckBodyWithCostForm f, head
                kind = OPT_KIND_PREFERENCE
            Case "minimize", "maximize", "minimise", "maximise"
                CheckBodyWithCostForm f, head
                kind = OPT_KIND_OBJECTIVE
            Case "fewest-changes-from"
                CheckKeptForm f, head
                kind = OPT_KIND_KEPT
            Case "effort"
                CheckEffortForm f, head
                kind = OPT_KIND_EFFORT
            Case Else
                VLA_Messages.RaiseMsg "optimize-unknown-top-form", "head", head
            End Select
            If kind <> OPT_KIND_NONE And firstKind = OPT_KIND_NONE Then
                firstKind = kind
                firstHead = head
            End If
        End If
    Next f
    If firstKind <> OPT_KIND_NONE Then RefuseNotYet firstHead, firstKind
End Sub

' The one refusal every still-unbuilt OPTIMIZE form reaches today. One
' id per INGREDIENT rather than one per spelling: a user who wrote
' choose-at-most and a user who wrote choose-between have the same
' problem, and the form they wrote is in the message either way.
'
' OPTIMIZE.2 removed the CONSTRAINT arm, and with it
' optimize-constraint-not-yet: a constraint is checked now, so a
' message saying "this version checks none" would be false. The other
' five are untouched, which is what the owner's call on this item's
' fork 3 bought - a program with a choice refuses exactly as it did,
' in exactly the words it did, so no pin over them had to be reread.
Private Sub RefuseNotYet(ByVal head As String, ByVal kind As Long)
    Select Case kind
    Case OPT_KIND_CHOICE
        VLA_Messages.RaiseMsg "optimize-choice-not-yet", "form", head
    Case OPT_KIND_PREFERENCE
        VLA_Messages.RaiseMsg "optimize-preference-not-yet", "form", head
    Case OPT_KIND_OBJECTIVE
        VLA_Messages.RaiseMsg "optimize-objective-not-yet", "form", head
    Case OPT_KIND_KEPT
        VLA_Messages.RaiseMsg "optimize-kept-not-yet", "form", head
    Case OPT_KIND_EFFORT
        VLA_Messages.RaiseMsg "optimize-effort-not-yet", "form", head
    End Select
End Sub

' =====================================================================
'  THE RESERVED PREFIX
' =====================================================================

' Every name this engine invents lives under `vla-`, so a user may not
' write one. Walks a whole form to any depth rather than checking a
' top-level head alone, because a predicate name appears in a rule
' body, in a keyed atom and inside a constraint just as readily as in a
' fact - and one collision anywhere would let a user's own rows arrive
' in a check's relation, or a check's rows in a user's.
'
' A QUOTED token is never a predicate name, so it is never tested: a
' cell that happens to contain the text vla-check-1 is data, and data
' is not refused. Nor is a bare DATA value - (fact (p vla-thing)) names
' no predicate vla-thing, and refusing a constant would be a different
' and much wider rule than the one this reserves.
'
' A PREDICATE POSITION is therefore exactly two things: the first
' element of any list, at any depth, and the bare NAME a (query ...)
' carries - which is not a list, so the walk below would never reach
' it, and (query vla-check-1) would otherwise have spilled this
' engine's own internal relation straight into a cell. ScanProgram
' passes that one separately.
Private Sub RefuseReservedNames(ByVal form As Variant)
    If Not IsObject(form) Then Exit Sub
    Dim lst As Collection
    Set lst = form
    If lst.Count < 1 Then Exit Sub
    Dim h As Variant
    NthInto h, lst, 1
    If Not IsObject(h) Then RefuseReservedName h
    Dim i As Long
    For i = 1 To lst.Count
        Dim part As Variant
        NthInto part, lst, i
        If IsObject(part) Then RefuseReservedNames part
    Next i
End Sub

' One bare token in a predicate position.
'
' THE SUB-PREFIX, NOT `vla-` ITSELF, and this is not a nicety. The
' project's own rule for a generated identifier is that it lives under
' `vla-` with a PER-GENERATOR sub-prefix, and `vla-` is already in use:
' the sentence layer emits `vla-ask-can-cover` and
' `vla-not-leave-name-shift` into perfectly ordinary DATALOG programs,
' and two of them are in the parity table. Reserving the whole of
' `vla-` here would refuse a program DATALOG answers - a parity
' divergence, and worse, the day a Frazaro sentence compiles to an
' OPTIMIZE call it would refuse its own generator's output. So this
' reserves exactly what this generator writes, `vla-check-`, and every
' other generator keeps its own.
Private Sub RefuseReservedName(ByVal raw As Variant)
    If IsObject(raw) Then Exit Sub
    Dim s As String
    s = CStr(raw)
    If Left$(s, 1) = Chr$(34) Then Exit Sub
    If Not IsGeneratedCheckName(s) Then Exit Sub
    VLA_Messages.RaiseMsg "optimize-reserved-predicate", "name", s, "prefix", OPT_CHECK_PREFIX
End Sub

' The one place the sub-prefix is compared, so the refusal above and
' OptimizeRun's own dictionary copy can never come to disagree about
' what this engine reserves.
Private Function IsGeneratedCheckName(ByVal s As String) As Boolean
    IsGeneratedCheckName = _
        (Left$(VLA_Identity.Fold(s), Len(OPT_CHECK_PREFIX)) = OPT_CHECK_PREFIX)
End Function

' =====================================================================
'  CONSTRAINTS, REWRITTEN INTO RULES
' =====================================================================

' The program DATALOG is actually given: every form the user wrote
' except the constraints, then one generated rule per constraint, in
' written order. The user's own sub-forms are REUSED by reference and
' never copied or re-serialised, so what DATALOG grounds is
' structurally the same objects the reader produced from the text.
'
' Order within the program does not matter to DATALOG - ParseProgram
' gathers facts and rules and then runs the fixpoint - so the generated
' rules go at the end, where they read as what they are.
Private Function RewriteProgram(ByVal forms As Collection, ByVal constraints As Collection, _
                                ByRef checkVars As Collection) As Collection
    Dim outp As Collection
    Set outp = New Collection
    Dim f As Variant
    For Each f In forms
        Select Case OptTopHead(f)
        Case "require", "forbid"
            ' Replaced below, by exactly one generated rule each.
        Case Else
            outp.Add f
        End Select
    Next f

    Set checkVars = New Collection
    Dim i As Long
    For i = 1 To constraints.Count
        Dim vars As Collection
        Dim rewritten As Collection
        Set rewritten = RewriteConstraint(constraints.Item(i), i, vars)
        checkVars.Add vars
        outp.Add rewritten
    Next i
    Set RewriteProgram = outp
End Function

' One constraint, as the rule that collects the rows breaking it.
'
'   (forbid B1 B2)        ->  (rule (vla-check-N "N" V...) B1 B2)
'   (require CONS B...)   ->  (rule (vla-check-N "N" V...) B... (not CONS))
'   (require GROUND)      ->  (rule (vla-check-N "N") (not GROUND))
'
' THE LEADING "N" IS NOT DECORATION. DATALOG refuses a zero-argument
' predicate outright (datalog-predicate-needs-argument, its own MVP
' scope boundary: an arity-0 Relation runs into the inverted-bounds
' ReDim hazard VLA_Relation documents), and a constraint whose body
' binds no variable at all - every one of the three shapes above can be
' written that way - would need exactly that. Carrying the check's own
' number as a constant first slot gives every generated head an arity
' of at least one, uniformly, with no second shape to reason about, and
' it is the number the violations table's Check column then reads back.
Private Function RewriteConstraint(ByVal form As Variant, ByVal n As Long, _
                                   ByRef vars As Collection) As Collection
    Dim lst As Collection
    Set lst = form
    Dim isRequire As Boolean
    isRequire = (OptTopHead(form) = "require")
    Dim bodyFrom As Long
    If isRequire Then bodyFrom = 3 Else bodyFrom = 2

    Set vars = New Collection
    Dim seen As Object
    Set seen = VLA_Runtime.VlaDictNew()
    Dim i As Long
    For i = bodyFrom To lst.Count
        Dim part As Variant
        NthInto part, lst, i
        CollectBoundVars part, vars, seen
    Next i

    Dim consequent As Variant
    If isRequire Then
        NthInto consequent, lst, 2
        CheckConsequent consequent, seen
    End If

    Dim headAtom As Collection
    Set headAtom = New Collection
    headAtom.Add OPT_CHECK_PREFIX & n
    headAtom.Add Chr$(34) & CStr(n)
    Dim v As Variant
    For Each v In vars
        headAtom.Add CStr(v)
    Next v

    Dim ruleForm As Collection
    Set ruleForm = New Collection
    ruleForm.Add "rule"
    ruleForm.Add headAtom
    For i = bodyFrom To lst.Count
        Dim bodyPart As Variant
        NthInto bodyPart, lst, i
        ruleForm.Add bodyPart
    Next i
    If isRequire Then
        ' LAST, deliberately: CheckRuleSafety walks a body in written
        ' order, so every positive atom has bound its variables by the
        ' time the negated consequent is reached. See the module header.
        Dim notForm As Collection
        Set notForm = New Collection
        notForm.Add "not"
        notForm.Add consequent
        ruleForm.Add notForm
    End If
    Set RewriteConstraint = ruleForm
End Function

' Which variables a body item BINDS - the head of the generated rule is
' exactly these, so that CheckRuleSafety can never refuse a rule this
' module built. The rule is DATALOG's own, read off CheckRuleSafety's
' Select Case rather than guessed:
'
'   not, the six comparisons, the three text tests   bind nothing
'   count, sum, let, textjoin                        bind their result
'                                                    variable, and NOT
'                                                    the source atom's
'                                                    own new variables
'   anything else (a positive atom)                  binds every
'                                                    variable among its
'                                                    arguments
'
' The wrapper list is the one place this module has to know something
' about DATALOG's body grammar, and it is held to DATALOG's own by
' tools/check_optimize_body_wrappers.ps1 - because the failure if the
' two drift is a new DATALOG wrapper read here as a positive atom, an
' unbound variable in a generated head, and a refusal naming
' vla-check-N at a user who never wrote it.
Private Sub CollectBoundVars(ByVal item As Variant, ByVal vars As Collection, ByVal seen As Object)
    If Not IsObject(item) Then Exit Sub
    Dim lst As Collection
    Set lst = item
    If lst.Count < 1 Then Exit Sub
    Dim h As Variant
    NthInto h, lst, 1
    If IsObject(h) Then Exit Sub
    ' CStr, NOT OptAtomText, and the difference is the leading quote of
    ' a string literal. DATALOG classifies a body item by
    ' Fold(CStr(wrapperRaw)) with the quote left ON, so a quoted
    ' ("not" X) is NOT its `not` wrapper but an ordinary positive atom
    ' whose predicate happens to be spelled not. Stripping the quote
    ' here would classify it the other way and quietly drop X from the
    ' evidence. Matching DATALOG exactly is the whole point of holding
    ' the two lists together.
    Select Case VLA_Identity.Fold(CStr(h))
    Case "not", ">", "<", "<=", ">=", "=", "<>", _
         "text-starts-with", "text-ends-with", "text-contains"
        ' Binds nothing at all.
    Case "count", "sum", "let", "textjoin"
        If lst.Count >= 2 Then AddVarAt lst, 2, vars, seen
    Case Else
        Dim i As Long
        For i = 2 To lst.Count
            Dim arg As Variant
            NthInto arg, lst, i
            If IsObject(arg) Then
                ' DATALOG.5's keyed atom: every argument is a
                ' (header value) pair, so only the VALUE can be a
                ' variable. Reading the whole pair would make a column
                ' named with a capital letter - Name, Task, Shift -
                ' look exactly like one, and put a name nothing binds
                ' into the head.
                Dim pair As Collection
                Set pair = arg
                If pair.Count = 2 Then AddVarAt pair, 2, vars, seen
            Else
                AddVarToken arg, vars, seen
            End If
        Next i
    End Select
End Sub

Private Sub AddVarAt(ByVal lst As Collection, ByVal at As Long, ByVal vars As Collection, ByVal seen As Object)
    Dim raw As Variant
    NthInto raw, lst, at
    If IsObject(raw) Then Exit Sub
    AddVarToken raw, vars, seen
End Sub

' First appearance wins, and each variable appears once. The ORDER is
' the order the user wrote them in, which is what makes the violations
' table's Where column hand-derivable from the constraint alone.
Private Sub AddVarToken(ByVal raw As Variant, ByVal vars As Collection, ByVal seen As Object)
    If Not OptIsVariable(raw) Then Exit Sub
    Dim s As String
    s = CStr(raw)
    If VLA_Runtime.VlaDictHas(seen, s) Then Exit Sub
    VLA_Runtime.VlaDictSet seen, s, True
    vars.Add s
End Sub

' A require's consequent, checked in OPTIMIZE's own words before
' DATALOG ever sees the (not ...) this module wrapped around it. Both
' refusals exist so that a user who wrote `require` is never shown a
' message about negation, which they did not write.
Private Sub CheckConsequent(ByVal consequent As Variant, ByVal seen As Object)
    If Not IsObject(consequent) Then Exit Sub
    Dim lst As Collection
    Set lst = consequent
    If lst.Count < 1 Then Exit Sub
    Dim h As Variant
    NthInto h, lst, 1
    If IsObject(h) Then Exit Sub
    ' Classified by CStr, for CollectBoundVars' own reason; NAMED by
    ' OptAtomText, because a message showing a user their own word
    ' should not show them the reader's quote character with it.
    Dim w As String
    w = VLA_Identity.Fold(OptAtomText(h))
    Select Case VLA_Identity.Fold(CStr(h))
    Case "not", "count", "sum", "let", "textjoin", ">", "<", "<=", ">=", "=", "<>", _
         "text-starts-with", "text-ends-with", "text-contains"
        ' What must hold has to be a ROW some relation can hold, since
        ' that is what "wherever the body holds, this holds" means. A
        ' comparison or an aggregate is a test, not a row, and writing
        ' one here is almost always a (forbid ...) turned inside out.
        VLA_Messages.RaiseMsg "optimize-require-consequent-not-a-row", "word", w
    End Select
    Dim i As Long
    For i = 2 To lst.Count
        Dim arg As Variant
        NthInto arg, lst, i
        If IsObject(arg) Then
            Dim pair As Collection
            Set pair = arg
            If pair.Count = 2 Then CheckConsequentVarAt pair, 2, seen, w
        Else
            CheckConsequentVar arg, seen, w
        End If
    Next i
End Sub

Private Sub CheckConsequentVarAt(ByVal lst As Collection, ByVal at As Long, ByVal seen As Object, ByVal predicate As String)
    Dim raw As Variant
    NthInto raw, lst, at
    If IsObject(raw) Then Exit Sub
    CheckConsequentVar raw, seen, predicate
End Sub

Private Sub CheckConsequentVar(ByVal raw As Variant, ByVal seen As Object, ByVal predicate As String)
    If Not OptIsVariable(raw) Then Exit Sub
    If VLA_Runtime.VlaDictHas(seen, CStr(raw)) Then Exit Sub
    VLA_Messages.RaiseMsg "optimize-require-unbound-consequent", _
        "var", CStr(raw), "predicate", predicate
End Sub

' =====================================================================
'  THE VIOLATIONS TABLE
' =====================================================================

' Every check's own derived relation, read back out of the relations
' the run produced, as one three-column table in check order and then
' in the order the rows were derived (which is deterministic, standing
' decision 2, so the same program always spills the same table).
'
' A check with no rows contributes none, and is not named: a rule that
' holds is not evidence of anything.
Private Function CollectViolations(ByVal constraints As Collection, ByVal checkVars As Collection, _
                                   ByVal relations As Object) As Collection
    Dim rel As Collection
    Set rel = VLA_Relation.RelNew(3)
    Dim row(1 To 3) As Variant
    Dim i As Long
    For i = 1 To constraints.Count
        Dim nm As String
        nm = OPT_CHECK_PREFIX & i
        If VLA_Runtime.VlaDictHas(relations, nm) Then
            Dim derived As Collection
            Set derived = VLA_Runtime.VlaDictGet(relations, nm)
            If VLA_Relation.RelCount(derived) > 0 Then
                Dim ruleText As String
                ruleText = VLA.VlaWriteForm(constraints.Item(i))
                Dim names As Collection
                Set names = checkVars.Item(i)
                Dim t As Variant, arr() As Variant
                For Each t In VLA_Relation.RelTuples(derived)
                    arr = t
                    row(1) = i
                    row(2) = ruleText
                    row(3) = BindingWords(names, arr)
                    VLA_Relation.RelTryAdd rel, row
                Next t
            End If
        End If
    Next i
    Set CollectViolations = rel
End Function

' The bindings that made a check hold, read straight off the generated
' head: slot 1 is the check's own number and slots 2 upward are the
' variables, in the order the user wrote them. A check whose body binds
' nothing says so rather than leaving a blank cell, since a blank in an
' evidence table reads as a missing value.
Private Function BindingWords(ByVal names As Collection, ByRef arr() As Variant) As String
    If names.Count = 0 Then
        BindingWords = "(no names - this rule is about particular rows)"
        Exit Function
    End If
    Dim s As String
    Dim i As Long
    For i = 1 To names.Count
        ' The head was built from these very names, so the tuple is
        ' always long enough - and this is the one line where being
        ' wrong about that would raise rather than answer, inside the
        ' cell that exists to explain a failure. It costs one
        ' comparison to make that impossible.
        If i + 1 > UBound(arr) Then Exit For
        If i > 1 Then s = s & ", "
        s = s & CStr(names.Item(i)) & " = " & CStr(arr(i + 1))
    Next i
    BindingWords = s
End Function

' =====================================================================
'  WHAT THE STATUS CELL SAYS
' =====================================================================

' The reserved words, unchanged, and then - only when a check is broken
' - which ones and how many rows each. OPTIMIZE.1 fixed the five
' sentences and this item may not reword one, so the reserved sentence
' stays the PREFIX and everything this item owns is appended after it.
' That is also why every pin OPTIMIZE.1 wrote over these words still
' passes: each matches a prefix fragment.
Private Function StatusSentence(ByVal stateId As Long, ByVal constraints As Collection, _
                                ByVal violations As Collection) As String
    Dim words As String
    words = OptimizeStatusWords(stateId)
    If stateId <> VLA_OPTIMIZE_NO_SCHEDULE Then
        StatusSentence = words
        Exit Function
    End If
    Dim s As String
    Dim named As Long
    Dim more As Long
    Dim i As Long
    For i = 1 To constraints.Count
        Dim n As Long
        n = ViolationRowsFor(violations, i)
        If n > 0 Then
            If named >= OPT_STATUS_CHECK_CAP Then
                more = more + 1
            Else
                If named > 0 Then s = s & "; "
                s = s & "check " & i & ", " & VLA.VlaWriteForm(constraints.Item(i)) & _
                    ", is broken by " & n & " " & RowsWord(n)
                named = named + 1
            End If
        End If
    Next i
    If more > 0 Then s = s & "; and " & more & " more check" & PluralS(more) & " broken"
    StatusSentence = words & ": " & s & ". OPTIMIZE_VIOLATIONS lists them."
End Function

Private Function ViolationRowsFor(ByVal violations As Collection, ByVal checkNumber As Long) As Long
    Dim n As Long
    Dim t As Variant, arr() As Variant
    For Each t In VLA_Relation.RelTuples(violations)
        arr = t
        If CLng(arr(1)) = checkNumber Then n = n + 1
    Next t
    ViolationRowsFor = n
End Function

Private Function RowsWord(ByVal n As Long) As String
    If n = 1 Then RowsWord = "row" Else RowsWord = "rows"
End Function

Private Function PluralS(ByVal n As Long) As String
    If n <> 1 Then PluralS = "s"
End Function

' =====================================================================
'  THE COUNTING PRE-CHECKS - built and pinned, reached by nothing yet
' =====================================================================
'
'  OPTIMIZE.0's oracle could not prove a pure counting impossibility in
'  three minutes under two encodings (clingo 5.8.2, the owner, on
'  optimize-roster-loose: 63 slots a week against a capacity of 50).
'  The arithmetic does it at once, because a pigeonhole proof is
'  exponential for clause learning and one multiplication by hand. So
'  these exist before any search does.
'
'  THEY ARE REACHED BY NOTHING IN THIS VERSION, on the owner's call
'  (2026-09-19, this item's fork 3), and the reason is where a mistake
'  would land. Every one of section 17's four NONE-COUNT keys takes its
'  DEMAND from a choice form, and a choice is refused before any of
'  this runs. Reaching them here would mean inferring demand from the
'  choice forms and the rule graph WITHOUT the grounder OPTIMIZE.3
'  builds - a shadow grounder, written to be replaced - and a slip in
'  it would answer "no schedule satisfies every rule" for a roster that
'  has one. That is the worst thing this engine can do, and the name's
'  own honesty rule exists to forbid its mirror. Reserved and pinned,
'  a wrong number is a red pin instead.
'
'  THE ORDER, and it is the corpus's own. optimize-roster-senior has
'  BOTH a local reason and a counting one, and clingo found the local
'  one at once (0.003 s) where it could not find the counting one at
'  all. OptimizeFirstShortfall reports the pool check first for the
'  same reason a person would: "the only senior is on leave that day"
'  names a row the manager can change, where "seven nights need a
'  senior and one person may work five shifts" names a policy.
'
'  THE SENTENCES ARE BUILT FROM THE PROGRAM'S OWN NAMES, never from
'  English verbs. Section 17's `why` lines read "may work", "may hold",
'  "seat" - words no s-expression supplies - so what a pin can require,
'  and does, is that the sentence carries the corpus's own NUMBERS, in
'  the corpus's own order. That is exactly what
'  tools/optimize01_corpus.ps1 already checks of the corpus itself.

' A group the choice must fill, against the rows that could fill it.
' optimize-roster-senior's local reason: the night of day 3 needs one
' senior, and no senior is free that day.
Public Function OptimizePoolShortWords(ByVal groupName As String, ByVal groupKey As String, _
                                       ByVal needed As Long, ByVal poolName As String, _
                                       ByVal available As Long) As String
    If needed <= available Then Exit Function
    OptimizePoolShortWords = OptimizeStatusWords(VLA_OPTIMIZE_NO_SCHEDULE) & _
        ": " & groupName & " " & groupKey & " needs " & needed & " " & RowsWord(needed) & _
        " from '" & poolName & "', and only " & available & " " & RowsWord(available) & _
        " can fill it."
End Function

' What the rules demand of one relation, against what a per-member cap
' can supply. optimize-sod-short (8 roles, 3 people at most 2 each) and
' optimize-roster-loose (63 slots, 10 people at most 5 each).
Public Function OptimizeCapShortWords(ByVal demandName As String, ByVal demand As Long, _
                                      ByVal memberName As String, ByVal members As Long, _
                                      ByVal cap As Long) As String
    Dim capacity As Long
    capacity = members * cap
    If demand <= capacity Then Exit Function
    OptimizeCapShortWords = OptimizeStatusWords(VLA_OPTIMIZE_NO_SCHEDULE) & _
        ": the rules need " & demand & " " & RowsWord(demand) & " of '" & demandName & _
        "', and '" & memberName & "' has " & members & " " & RowsWord(members) & _
        " at most " & cap & " each, which is " & capacity & "."
End Function

' Things that must each be placed, against the groups that can hold
' them. optimize-seat-overflow (7 guests, 2 tables of 3) and
' optimize-exam-rooms (5 exams, 2 sittings of 2).
'
' Kept apart from OptimizeCapShortWords although the arithmetic is the
' same: one says what the RULES need and the other what the DATA
' brings, and a single sentence would be wrong about one of them.
Public Function OptimizeCapacityShortWords(ByVal thingName As String, ByVal things As Long, _
                                           ByVal groupName As String, ByVal groups As Long, _
                                           ByVal cap As Long) As String
    Dim capacity As Long
    capacity = groups * cap
    If things <= capacity Then Exit Function
    OptimizeCapacityShortWords = OptimizeStatusWords(VLA_OPTIMIZE_NO_SCHEDULE) & _
        ": there are " & things & " " & RowsWord(things) & " of '" & thingName & _
        "' to place, and '" & groupName & "' has " & groups & " " & RowsWord(groups) & _
        " holding at most " & cap & " each, which is " & capacity & "."
End Function

' The three in order, first non-empty wins. "" when nothing is proven
' impossible - which is never "there is a schedule", only "counting
' alone does not say there is not one".
Public Function OptimizeFirstShortfall(ByVal poolWords As String, ByVal capWords As String, _
                                       ByVal capacityWords As String) As String
    If Len(poolWords) > 0 Then
        OptimizeFirstShortfall = poolWords
    ElseIf Len(capWords) > 0 Then
        OptimizeFirstShortfall = capWords
    Else
        OptimizeFirstShortfall = capacityWords
    End If
End Function

' (choose-exactly N chosen pool [(per ...)]), and choose-between with
' two counts. countSlots is 1 or 2.
Private Sub CheckChoiceForm(ByVal form As Variant, ByVal head As String, ByVal countSlots As Long)
    Dim lst As Collection
    Set lst = form
    Dim least As Long
    least = 1 + countSlots + 2
    If lst.Count < least Or lst.Count > least + 1 Then
        VLA_Messages.RaiseMsg "optimize-form-bad-shape", "form", head, _
            "expected", ChoiceShapeWords(head, countSlots), "example", ChoiceExampleFor(head)
    End If
    Dim i As Long
    For i = 2 To 1 + countSlots
        CheckChoiceCount lst, i, head
    Next i
    ' The chosen rows, then the pool they come from. Both are atoms; a
    ' bare word or a number in either slot is the commonest way to
    ' write this wrong.
    CheckIsAtom lst, 2 + countSlots, head, "the rows being chosen"
    CheckIsAtom lst, 3 + countSlots, head, "the pool they are chosen from"
    ' choose-between's two counts, when both are plain numbers, must be
    ' in order. Left alone when either is a variable: which is larger is
    ' then a question about the data, and OPTIMIZE.3's own grounding
    ' answers it.
    If countSlots = 2 Then CheckBetweenOrder lst, head
    If lst.Count = least + 1 Then CheckPerForm lst, least + 1, head
End Sub

' (choose-any chosen pool) - each pool row in or out, no count, no
' group. The one form of the five with a fixed arity of three.
Private Sub CheckChooseAnyForm(ByVal form As Variant, ByVal head As String)
    Dim lst As Collection
    Set lst = form
    If lst.Count <> 3 Then
        VLA_Messages.RaiseMsg "optimize-form-bad-shape", "form", head, _
            "expected", "the rows being chosen, then the pool they are chosen from", _
            "example", "(choose-any (on From To) (Links From To Cost))"
    End If
    CheckIsAtom lst, 2, head, "the rows being chosen"
    CheckIsAtom lst, 3, head, "the pool they are chosen from"
End Sub

' (require CONSEQUENT BODY...) - in every world, wherever the body
' holds, the consequent holds. DATALOG's own `rule` order (the owner's
' call, this item's scoping fork 4): a reader who knows `rule` knows
' this, and `forbid` stays its exact mirror. With no body it is a plain
' ground requirement ("Ann sits at the head table"). An EXISTENTIAL
' requirement needs no new form and is not one: "every night has a
' senior on it" is a derived `rule` doing the existential, then a
' `require` over that rule's own head - which is why nothing here
' spells `exists`, and what G-OPTIMIZE will generate.
Private Sub CheckRequireForm(ByVal form As Variant, ByVal head As String)
    Dim lst As Collection
    Set lst = form
    If lst.Count < 2 Then
        VLA_Messages.RaiseMsg "optimize-form-bad-shape", "form", head, _
            "expected", "what must hold, then the rows it must hold for (none, for a plain requirement)", _
            "example", "(require (senior P) (reviews T P))"
    End If
    Dim i As Long
    For i = 2 To lst.Count
        CheckIsAtom lst, i, head, "a required row"
    Next i
End Sub

' (forbid BODY...) - no world in which the whole body holds.
Private Sub CheckForbidForm(ByVal form As Variant, ByVal head As String)
    Dim lst As Collection
    Set lst = form
    If lst.Count < 2 Then
        VLA_Messages.RaiseMsg "optimize-form-bad-shape", "form", head, _
            "expected", "the rows that may never all hold at once", _
            "example", "(forbid (prepares T P) (reviews T P))"
    End If
    Dim i As Long
    For i = 2 To lst.Count
        CheckIsAtom lst, i, head, "a forbidden row"
    Next i
End Sub

' (prefer BODY... [(cost T)]) and (avoid ...), and the same shape for
' (minimize ...)/(maximize ...) - a body, and optionally what each
' instance of it is worth.
'
' WRITTEN ORDER IS PRIORITY among the objectives (standing decision 2's
' stated order, put on the page). The value of an objective is the sum
' of its cost over every ground instance of its body that holds in the
' world, so one shape says all three kinds section 17 asks for: a COUNT
' of matching rows ("as few preparing roles as possible go to
' seniors"), a WEIGHTED sum from a Table column ("the least total
' weight of wishes broken"), and a sum over DERIVED atoms ("least
' overtime"). No separate count or sum wrapper, which would give
' `count` two meanings - DATALOG's aggregate over certain rows, and an
' objective term over chosen ones.
Private Sub CheckBodyWithCostForm(ByVal form As Variant, ByVal head As String)
    Dim lst As Collection
    Set lst = form
    If lst.Count < 2 Then
        VLA_Messages.RaiseMsg "optimize-form-bad-shape", "form", head, _
            "expected", "the rows this is about, and optionally what each one is worth", _
            "example", "(" & head & " (assign S P) (requests-off P S) (cost 1))"
    End If
    Dim last As Long
    last = lst.Count
    If IsCostForm(lst, last) Then
        CheckCostForm lst, last, head
        last = last - 1
        If last < 2 Then
            VLA_Messages.RaiseMsg "optimize-form-bad-shape", "form", head, _
                "expected", "the rows this is about, before its (cost ...)", _
                "example", "(" & head & " (assign S P) (requests-off P S) (cost 1))"
        End If
    End If
    Dim i As Long
    For i = 2 To last
        CheckIsAtom lst, i, head, "a row this is about"
    Next i
End Sub

' (fewest-changes-from LastMonth) - the kept schedule, named as an
' EXPLICIT input. Not `keep`, which in this codebase's own vocabulary
' reads as "filter to these rows" and sounds hard where this is soft;
' and never the cell's own previous value, which is OPTIMIZE.7's
' warm-start trap.
Private Sub CheckKeptForm(ByVal form As Variant, ByVal head As String)
    Dim lst As Collection
    Set lst = form
    If lst.Count <> 2 Then
        VLA_Messages.RaiseMsg "optimize-form-bad-shape", "form", head, _
            "expected", "one name: the Table holding the schedule to stay close to", _
            "example", "(fewest-changes-from LastMonth)"
    End If
    Dim raw As Variant
    NthInto raw, lst, 2
    If IsObject(raw) Then
        VLA_Messages.RaiseMsg "optimize-form-bad-shape", "form", head, _
            "expected", "one name, not a row", "example", "(fewest-changes-from LastMonth)"
    End If
    If Len(OptAtomText(raw)) = 0 Then
        VLA_Messages.RaiseMsg "optimize-form-bad-shape", "form", head, _
            "expected", "one name, and it is empty", "example", "(fewest-changes-from LastMonth)"
    End If
End Sub

' (effort quick|normal|thorough) or (effort N) for experts. "Effort"
' says WORK, never time, so determinism stays honest where
' "(budget 50000)" left "50,000 what?" on the page. The three levels
' map to a fixed work count, set once OPTIMIZE.3 and OPTIMIZE.6 have
' measured; a number is a work count written out.
Private Sub CheckEffortForm(ByVal form As Variant, ByVal head As String)
    Dim lst As Collection
    Set lst = form
    If lst.Count <> 2 Then
        VLA_Messages.RaiseMsg "optimize-form-bad-shape", "form", head, _
            "expected", "one of quick, normal or thorough, or a number of units of work", _
            "example", "(effort thorough)"
    End If
    Dim raw As Variant
    NthInto raw, lst, 2
    If IsObject(raw) Then
        VLA_Messages.RaiseMsg "optimize-form-bad-shape", "form", head, _
            "expected", "one of quick, normal or thorough, or a number", "example", "(effort thorough)"
    End If
    Dim w As String
    w = VLA_Identity.Fold(OptAtomText(raw))
    Select Case w
    Case "quick", "normal", "thorough"
        ' settled
    Case Else
        If Not IsWholeNumberText(w) Then
            VLA_Messages.RaiseMsg "optimize-effort-unknown-level", "level", OptAtomText(raw)
        End If
        If CDbl(w) <= 0 Then
            VLA_Messages.RaiseMsg "optimize-effort-unknown-level", "level", OptAtomText(raw)
        End If
    End Select
End Sub

' (per group-atom ...) - which rows the count is counted for, one atom
' per thing grouped by. Several atoms rather than one, because section
' 17 needs them: "every month, every person gets exactly one duty".
Private Sub CheckPerForm(ByVal lst As Collection, ByVal at As Long, ByVal head As String)
    Dim raw As Variant
    NthInto raw, lst, at
    If Not IsObject(raw) Then
        VLA_Messages.RaiseMsg "optimize-form-bad-shape", "form", head, _
            "expected", "a (per ...) after the pool, or nothing at all", _
            "example", "(per (Shifts S Need))"
    End If
    Dim per As Collection
    Set per = raw
    If per.Count < 2 Then
        VLA_Messages.RaiseMsg "optimize-form-bad-shape", "form", "per", _
            "expected", "at least one row to group by", "example", "(per (Shifts S Need))"
    End If
    Dim w As Variant
    NthInto w, per, 1
    If IsObject(w) Then
        VLA_Messages.RaiseMsg "optimize-form-bad-shape", "form", "per", _
            "expected", "the word per, then the rows to group by", "example", "(per (Shifts S Need))"
    End If
    If VLA_Identity.Fold(OptAtomText(w)) <> "per" Then
        VLA_Messages.RaiseMsg "optimize-form-bad-shape", "form", head, _
            "expected", "a (per ...) after the pool, or nothing at all", _
            "example", "(per (Shifts S Need))"
    End If
    Dim i As Long
    For i = 2 To per.Count
        CheckIsAtom per, i, "per", "a row to group by"
    Next i
End Sub

' A count slot: a whole number of zero or more, or a VARIABLE - and the
' variable is the whole reason (per ...) exists. "Every shift gets
' exactly the people it needs" takes its count from a Table column, so
' a form that accepted only a literal could not say the killer case's
' own choice line.
Private Sub CheckChoiceCount(ByVal lst As Collection, ByVal at As Long, ByVal head As String)
    Dim raw As Variant
    NthInto raw, lst, at
    If IsObject(raw) Then
        VLA_Messages.RaiseMsg "optimize-count-not-a-number", "form", head, "count", "a row"
    End If
    Dim s As String
    s = OptAtomText(raw)
    If OptIsVariable(raw) Then Exit Sub
    If Not IsWholeNumberText(s) Then
        VLA_Messages.RaiseMsg "optimize-count-not-a-number", "form", head, "count", s
    End If
    If CDbl(s) < 0 Then
        VLA_Messages.RaiseMsg "optimize-count-not-a-number", "form", head, "count", s
    End If
End Sub

Private Sub CheckBetweenOrder(ByVal lst As Collection, ByVal head As String)
    Dim rawLo As Variant, rawHi As Variant
    NthInto rawLo, lst, 2
    NthInto rawHi, lst, 3
    If IsObject(rawLo) Or IsObject(rawHi) Then Exit Sub
    If OptIsVariable(rawLo) Or OptIsVariable(rawHi) Then Exit Sub
    Dim loText As String, hiText As String
    loText = OptAtomText(rawLo)
    hiText = OptAtomText(rawHi)
    If Not IsWholeNumberText(loText) Then Exit Sub
    If Not IsWholeNumberText(hiText) Then Exit Sub
    If CDbl(loText) > CDbl(hiText) Then
        VLA_Messages.RaiseMsg "optimize-choose-range-inverted", "low", loText, "high", hiText
    End If
End Sub

' Whether the item at `at` is a (cost ...) form - checked before it is
' validated, so a body atom that happens to be about costs is not
' mistaken for the modifier.
Private Function IsCostForm(ByVal lst As Collection, ByVal at As Long) As Boolean
    If at < 2 Then Exit Function
    Dim raw As Variant
    NthInto raw, lst, at
    If Not IsObject(raw) Then Exit Function
    Dim c As Collection
    Set c = raw
    If c.Count < 1 Then Exit Function
    Dim w As Variant
    NthInto w, c, 1
    If IsObject(w) Then Exit Function
    IsCostForm = (VLA_Identity.Fold(OptAtomText(w)) = "cost")
End Function

' (cost T), default 1 when absent. T may be a number OR A VARIABLE, and
' the variable is not a nicety: optimize-seat-wedding carries
' Wishes(A, B, Weight) and asks for "the least total weight of wishes
' broken", so a cost restricted to a literal could not say that
' sentence. A variable must be bound by a positive row of the same
' body - DATALOG's own safety condition, which OPTIMIZE.6 will check
' against the grounding the way CheckRuleSafety already does for a rule.
Private Sub CheckCostForm(ByVal lst As Collection, ByVal at As Long, ByVal head As String)
    Dim raw As Variant
    NthInto raw, lst, at
    Dim c As Collection
    Set c = raw
    If c.Count <> 2 Then
        VLA_Messages.RaiseMsg "optimize-form-bad-shape", "form", "cost", _
            "expected", "one number, or one name bound by a row of the same rule", _
            "example", "(cost 2) or (cost Weight)"
    End If
    Dim v As Variant
    NthInto v, c, 2
    If IsObject(v) Then
        VLA_Messages.RaiseMsg "optimize-form-bad-shape", "form", "cost", _
            "expected", "one number or one name, not a row", "example", "(cost 2) or (cost Weight)"
    End If
    If OptIsVariable(v) Then Exit Sub
    If Not IsNumberText(OptAtomText(v)) Then
        VLA_Messages.RaiseMsg "optimize-form-bad-shape", "form", "cost", _
            "expected", "one number, or one name bound by a row of the same rule", _
            "example", "(cost 2) or (cost Weight)"
    End If
End Sub

' Every slot that must hold a row: a list whose own first element is a
' plain word. A bare word, a number or a quoted string in such a slot is
' the commonest way to write any of these forms wrong.
Private Sub CheckIsAtom(ByVal lst As Collection, ByVal at As Long, ByVal head As String, ByVal role As String)
    Dim raw As Variant
    NthInto raw, lst, at
    If Not IsObject(raw) Then
        VLA_Messages.RaiseMsg "optimize-not-a-row", "form", head, "role", role, "text", OptAtomText(raw)
    End If
    Dim c As Collection
    Set c = raw
    If c.Count < 1 Then
        VLA_Messages.RaiseMsg "optimize-not-a-row", "form", head, "role", role, "text", "()"
    End If
    Dim w As Variant
    NthInto w, c, 1
    If IsObject(w) Then
        VLA_Messages.RaiseMsg "optimize-not-a-row", "form", head, "role", role, "text", "a row of rows"
    End If
    If Len(OptAtomText(w)) = 0 Then
        VLA_Messages.RaiseMsg "optimize-not-a-row", "form", head, "role", role, "text", "an empty name"
    End If
End Sub

Private Function ChoiceShapeWords(ByVal head As String, ByVal countSlots As Long) As String
    If countSlots = 2 Then
        ChoiceShapeWords = "a least and a most, then the rows being chosen, then the pool, then an optional (per ...)"
    Else
        ChoiceShapeWords = "a count, then the rows being chosen, then the pool, then an optional (per ...)"
    End If
End Function

Private Function ChoiceExampleFor(ByVal head As String) As String
    If head = "choose-between" Then
        ChoiceExampleFor = "(choose-between 2 4 (assign S P) (Eligible S P) (per (Shifts S)))"
    Else
        ChoiceExampleFor = "(" & head & " N (assign S P) (Eligible S P) (per (Shifts S N)))"
    End If
End Function

' =====================================================================
'  THE MEMO
' =====================================================================

' The key: SEC.11's SHA-256 over VLA_Digest's injective framing of
' everything that can change the answer. Public so a test can read one,
' and so the framing is inspectable rather than implied.
'
' WHAT GOES IN, and why each part has to:
'   - a version tag, so that changing this layout later cannot make a
'     new key collide with an old one;
'   - the number of Tables, then for each one IN ARGUMENT ORDER its
'     name, its column count, its row count, its column NAMES (a
'     renamed header changes what a keyed rule resolves to, so it
'     changes the answer), and every tuple cell in order;
'   - the rules text, last.
'
' THE TUPLES RATHER THAN THE RANGE, deliberately. Hashing the Relation
' the engine will actually read makes the key a function of precisely
' what the engine saw - after RelFromRange dropped all-blank rows and
' RelTryAdd absorbed duplicates - and it avoids inheriting
' VLA_Relation.TupleKey's own Chr$(31) join, which is not injective.
'
' ARGUMENT ORDER MATTERS TO THE KEY and not to the answer, so the same
' Tables passed in a different order are a MISS: one needless re-solve,
' never a wrong answer. Sorting to close that would cost every call and
' buy a rare case.
Public Function OptimizeMemoKey(ByVal rulesText As String, ByVal relations As Object, _
                                ByVal headerMap As Object) As String
    Dim buf() As Byte
    Dim n As Long
    VLA_Digest.VlaKeyBegin buf, n
    VLA_Digest.VlaKeyAddText buf, n, "vla-optimize-memo-1"

    Dim names As Collection
    Set names = VLA_Runtime.VlaDictKeys(relations)
    VLA_Digest.VlaKeyAddCount buf, n, names.Count

    Dim k As Variant
    For Each k In names
        Dim nm As String
        nm = CStr(k)
        Dim rel As Collection
        Set rel = VLA_Runtime.VlaDictGet(relations, nm)
        VLA_Digest.VlaKeyAddText buf, n, nm
        VLA_Digest.VlaKeyAddCount buf, n, VLA_Relation.RelArity(rel)
        VLA_Digest.VlaKeyAddCount buf, n, VLA_Relation.RelCount(rel)
        Dim cols As Collection
        Set cols = Nothing
        If VLA_Runtime.VlaDictHas(headerMap, nm) Then
            Set cols = VLA_Runtime.VlaDictGet(headerMap, nm)
        End If
        If cols Is Nothing Then
            VLA_Digest.VlaKeyAddCount buf, n, 0
        Else
            VLA_Digest.VlaKeyAddCount buf, n, cols.Count
            Dim cp As Variant
            For Each cp In cols
                Dim pair As Collection
                Set pair = cp
                VLA_Digest.VlaKeyAddText buf, n, CStr(pair.Item(1))
                VLA_Digest.VlaKeyAddText buf, n, CStr(pair.Item(2))
            Next cp
        End If
        Dim tup As Variant
        For Each tup In VLA_Relation.RelTuples(rel)
            Dim arr() As Variant
            arr = tup
            Dim ci As Long
            For ci = LBound(arr) To UBound(arr)
                VLA_Digest.VlaKeyAddCell buf, n, arr(ci)
            Next ci
        Next tup
    Next k

    VLA_Digest.VlaKeyAddText buf, n, rulesText
    OptimizeMemoKey = VLA_Digest.VlaKeyHex(buf, n)
End Function

' How many searches have actually run in this session. The memo's own
' pin: two cells asking the same question must move this by one.
Public Function OptimizeMemoRuns() As Long
    OptimizeMemoRuns = mMemoRuns
End Function

' How many answers the memo currently holds - never more than the cap.
Public Function OptimizeMemoCount() As Long
    If mMemo Is Nothing Then Exit Function
    OptimizeMemoCount = VLA_Runtime.VlaDictKeys(mMemo).Count
End Function

' Empties the memo, exactly as a VBA project reset does, and resets the
' run counter. Public because the pins need to SIMULATE that reset: the
' memo must tolerate being lost at any moment, so the proof is to lose
' it between two calls and get the same answer.
Public Sub OptimizeMemoClear()
    Set mMemo = Nothing
    Set mMemoOrder = Nothing
    mMemoRuns = 0
End Sub

Private Sub MemoEnsure()
    If mMemo Is Nothing Then Set mMemo = VLA_Runtime.VlaDictNew()
    If mMemoOrder Is Nothing Then Set mMemoOrder = New Collection
End Sub

Private Function MemoHas(ByVal memoKey As String) As Boolean
    If mMemo Is Nothing Then Exit Function
    MemoHas = VLA_Runtime.VlaDictHas(mMemo, memoKey)
End Function

Private Function MemoGet(ByVal memoKey As String) As Collection
    Set MemoGet = VLA_Runtime.VlaDictGet(mMemo, memoKey)
End Function

' Capped at OPT_MEMO_CAP, oldest first. A cap rather than no cap
' because each entry holds a whole relation set alive, and a workbook
' of many OPTIMIZE cells would otherwise grow the session's memory
' without bound. Oldest-first rather than least-used: it needs no
' bookkeeping per hit, and a recalculation touches every cell anyway.
Private Sub MemoPut(ByVal memoKey As String, ByVal result As Collection)
    MemoEnsure
    If VLA_Runtime.VlaDictHas(mMemo, memoKey) Then Exit Sub
    VLA_Runtime.VlaDictSet mMemo, memoKey, result
    mMemoOrder.Add memoKey
    If mMemoOrder.Count <= OPT_MEMO_CAP Then Exit Sub
    ' Over the cap: the oldest entries go by REBUILDING the dictionary
    ' rather than by removing from it. The VlaDict family has no
    ' remove, and it could not simply gain one: on a host with no
    ' Scripting runtime, VlaDictNew falls back to a plain Collection of
    ' pairs, which cannot be removed from by key at all. Rebuilding
    ' works on both representations, runs at most once per put, and
    ' walks at most OPT_MEMO_CAP entries.
    Do While mMemoOrder.Count > OPT_MEMO_CAP
        mMemoOrder.Remove 1
    Loop
    Dim rebuilt As Object
    Set rebuilt = VLA_Runtime.VlaDictNew()
    Dim k As Variant
    For Each k In mMemoOrder
        Dim ks As String
        ks = CStr(k)
        If VLA_Runtime.VlaDictHas(mMemo, ks) Then
            VLA_Runtime.VlaDictSet rebuilt, ks, VLA_Runtime.VlaDictGet(mMemo, ks)
        End If
    Next k
    Set mMemo = rebuilt
End Sub

' =====================================================================
'  TABLE ARGUMENTS, and the S-expression accessors
' =====================================================================

' OPTIMIZE's own wording for a table argument that cannot be named -
' VLA_Datalog.TableArgName's twin, and its own ids for the same reason
' PROLOG and SQL have theirs: the shared resolver in VLA_Relation.bas
' returns a REASON CODE and never raises (its own LAYER 0.5 contract),
' because refusal wording is each engine's own job, and a message
' reading "every DATALOG table argument..." inside an #OPTIMIZE! cell
' would name the wrong function.
'
' The Case Else is DATALOG.15's own lesson, applied on the way in
' rather than after the fact: before that item, none of the three
' wrappers had one, so a NEW reason from TableArgResolve fell through
' every Select Case and handed back an EMPTY predicate name, silently.
' Every reason this wrapper does not word itself - an error value, and a
' spill's name or headers - is worded once, in
' VLA_Relation.RaiseTableArgRefusal.
Private Function OptimizeTableArgName(ByVal v As Variant) As String
    Dim ok As Boolean, reason As String, detail As String
    OptimizeTableArgName = VLA_Relation.TableArgResolve(v, ok, reason, detail)
    If ok Then Exit Function
    Select Case reason
    Case "not-a-range"
        VLA_Messages.RaiseMsg "optimize-table-not-a-range"
    Case "noncontiguous"
        VLA_Messages.RaiseMsg "optimize-table-noncontiguous-columns"
    Case "needs-a-name"
        VLA_Messages.RaiseMsg "optimize-table-needs-a-name"
    Case Else
        VLA_Relation.RaiseTableArgRefusal reason, detail
    End Select
End Function

' ---- minimal S-expression accessors, the same three VLA_Datalog.bas
'      keeps for the same reason (a cross-module Private call does not
'      exist in VBA). NthInto and CopyVariant are not conveniences:
'      a bare `dest = lst.Item(i)` invokes an object's DEFAULT MEMBER
'      when the item is itself a list, which for a Collection is Item,
'      which needs an index - so it raises 450 rather than storing the
'      list. VLA_Datalog.bas's own header has the full account of the
'      live incident.
Private Sub NthInto(ByRef dest As Variant, ByVal lst As Collection, ByVal i As Long)
    If IsObject(lst.Item(i)) Then
        Set dest = lst.Item(i)
    Else
        dest = lst.Item(i)
    End If
End Sub

Private Sub CopyVariant(ByRef dest As Variant, ByVal src As Variant)
    If IsObject(src) Then
        Set dest = src
    Else
        dest = src
    End If
End Sub

' A top-level form's head symbol, folded - or "" when the form has no
' readable head, which is left for DatalogRun to refuse in its own
' words rather than reworded here.
Private Function OptTopHead(ByVal form As Variant) As String
    If Not IsObject(form) Then Exit Function
    Dim lst As Collection
    Set lst = form
    If lst.Count < 1 Then Exit Function
    Dim h As Variant
    NthInto h, lst, 1
    If IsObject(h) Then Exit Function
    OptTopHead = VLA_Identity.Fold(OptAtomText(h))
End Function

' A raw reader token's own text. A leading Chr$(34) marks a quoted
' string literal (VLA.bas's reader convention) and is stripped, exactly
' as VLA_Datalog.AtomText does - so "cost" quoted is still the word
' cost, and a quoted count is still refused as not a number.
Private Function OptAtomText(ByVal raw As Variant) As String
    If IsObject(raw) Then Exit Function
    Dim s As String
    s = CStr(raw)
    If Left$(s, 1) = Chr$(34) Then
        OptAtomText = Mid$(s, 2)
    Else
        OptAtomText = s
    End If
End Function

' A variable, by DATALOG's own rule: a BARE token whose first character
' is A-Z. A quoted string is always a constant, even if its content
' starts with a capital.
Private Function OptIsVariable(ByVal raw As Variant) As Boolean
    If IsObject(raw) Then Exit Function
    Dim s As String
    s = CStr(raw)
    If Len(s) = 0 Then Exit Function
    If Left$(s, 1) = Chr$(34) Then Exit Function
    Dim c As Long
    c = AscW(Left$(s, 1))
    OptIsVariable = (c >= 65 And c <= 90)
End Function

' Digits only, at least one - never Val() or IsNumeric, either of which
' would accept "1E3", " 1", "1,5" or "&H10" and each of which reads
' differently in a different locale. A count is a count.
Private Function IsWholeNumberText(ByVal s As String) As Boolean
    If Len(s) = 0 Then Exit Function
    Dim i As Long
    For i = 1 To Len(s)
        Dim c As Long
        c = AscW(Mid$(s, i, 1))
        If c < 48 Or c > 57 Then Exit Function
    Next i
    IsWholeNumberText = True
End Function

' A cost may carry a decimal point, so this is the wider of the two -
' still digits, with at most one dot, and an optional leading minus, and
' still never Val() or IsNumeric for the same locale reason.
Private Function IsNumberText(ByVal s As String) As Boolean
    Dim body As String
    body = s
    If Left$(body, 1) = "-" Then body = Mid$(body, 2)
    If Len(body) = 0 Then Exit Function
    Dim dots As Long
    Dim i As Long
    For i = 1 To Len(body)
        Dim c As Long
        c = AscW(Mid$(body, i, 1))
        If c = 46 Then
            dots = dots + 1
            If dots > 1 Then Exit Function
        ElseIf c < 48 Or c > 57 Then
            Exit Function
        End If
    Next i
    If body = "." Then Exit Function
    IsNumberText = True
End Function
