Attribute VB_Name = "VLA_Optimize"
Option Explicit
Public Const VLA_OPTIMIZE_VERSION As String = "OPTIMIZE.3"

' =====================================================================
'  VLA_Optimize - OPTIMIZE.3 (slice 2): choice, grounding over the pool,
'  and a search that propagates. OPTIMIZE.2's constraints against the
'  one world, OPTIMIZE.1's base case beneath them.
'
'  WHAT OPTIMIZE.3 SLICE 2 ADDS. The five choice forms run: a program
'  that chooses is grounded in three passes - the certain part through
'  DATALOG's own fixpoint, then the choices and then the constraints over
'  chosen rows through slice 1's integer grounder - into atoms, counters
'  and clauses, and VLA_OptimizeSearch finds the first schedule that
'  breaks no rule, deciding the rows in the Tables' own order (the
'  owner's fork 6). The section OPTIMIZE.3: PROGRAMS THAT CHOOSE, below,
'  has the whole account. (effort ...) is the search's budget, in work;
'  OPTIMIZE.2's counting pre-checks are reached at last; a program with
'  no choice takes OPTIMIZE.2's path exactly as before. Still refused by
'  name: preferences, objectives and the kept schedule (OPTIMIZE.6 and
'  .7), rules over chosen rows (.5) and a count over them (.4).
'
'  SLICE 3 holds a formula to ceilings, and SLICE 5 adds the command a
'  formula's size refusal points to: a cell's own =OPTIMIZE(...) call run
'  again with a command's ceiling and no seconds guard, its answer written
'  to a new sheet (the section OPTIMIZE.3 slice 5: THE COMMAND).
'
'  The rest of this header is OPTIMIZE.1's and OPTIMIZE.2's, kept as
'  written; where it says a choice is refused, OPTIMIZE.3 is what changed
'  that.
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
'             clear), VLA_Datalog (DatalogRunForms - the whole engine -
'             and, since OPTIMIZE.3, DatalogGroundRules, slice 1's
'             integer grounder), VLA_Relation (TableArgResolve/
'             RelFromRange/RangeColumnNames/RelToSpilledArray/RelNew/
'             RelTryAdd/RelArity/RelCount/RelTuples, RaiseTableArgRefusal
'             for a reason no wrapper here words itself, and the
'             VlaSymbols table with VlaSymInit/VlaSymFind and
'             InvariantNumberText), VLA_OptimizeSearch (OPTIMIZE.3's
'             search, which calls nothing back), VLA_Digest (the memo
'             key), VLA_Identity (Fold), VLA_Messages (every refusal),
'             VLA_Runtime (VlaDict*), and - OPTIMIZE.3 slice 5's command
'             alone - Excel's own objects: the cell's formula and
'             Worksheet.Evaluate to read it, Worksheets.Add to write its
'             answer, the status bar, and EnableCancelKey for Esc
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
' In units of WORK - one decision or one dead end each - never seconds:
' standing decision 2, so that a faster machine proves the same thing a
' slower one does. Zero until OPTIMIZE.3 slice 2, the first search that
' consumes them, and provisional there, ten times apart, until slice 4's
' ladder measured them on OPTIMIZE.0's reference roster, 50 people over
' 4 weeks: a unit of search work at 3.7 to 4.35 us, and everything else
' a run does at 0.75 s. THOROUGH is the most search a formula's 2 s
' leaves room for on that roster - 250,000 units, about 1.1 s, and 1.8 s
' in all - and each level is a tenth of the next (the owner's call,
' 2026-09-25). OPTIMIZE.6 revisits them once searches optimise.
Public Const VLA_OPTIMIZE_WORK_QUICK As Long = 2500
Public Const VLA_OPTIMIZE_WORK_NORMAL As Long = 25000
Public Const VLA_OPTIMIZE_WORK_THOROUGH As Long = 250000

' The level a program with no (effort ...) form gets, written down now
' so the day a number lands behind it, the default is already stated.
Public Const VLA_OPTIMIZE_EFFORT_DEFAULT As String = "normal"

' ---- the form kinds, for the refusal that names the right one ------
'
' OPT_KIND_CONSTRAINT is gone with optimize-constraint-not-yet, which
' OPTIMIZE.2 made false: a constraint is no longer an unbuilt
' ingredient, it is the thing that item built. OPT_KIND_CHOICE (1) and
' OPT_KIND_EFFORT (6) went the same way at OPTIMIZE.3 slice 2, with
' optimize-choice-not-yet and optimize-effort-not-yet. The numbers of
' the others are left exactly where they were rather than closed up,
' since nothing outside this module reads them and a renumber would be
' a diff with no meaning in it.
Private Const OPT_KIND_NONE As Long = 0
Private Const OPT_KIND_PREFERENCE As Long = 3
Private Const OPT_KIND_OBJECTIVE As Long = 4
Private Const OPT_KIND_KEPT As Long = 5

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

' ---- OPTIMIZE.3: the names a choice program's grounding writes ------
'
' Every one under OPT_CHECK_PREFIX, so a user can never write one (the
' prefix is refused at parse) and DefinedNamesSentence never offers one
' back. Each is derived from written position alone - a choice form's
' own 1-based number among the choice forms, a constraint's among the
' constraints, a stub's in the order the stubs are written - and they
' are injective, since what follows the prefix never overlaps: never,
' stub-, group-, member-, rank-, clause-, or OPTIMIZE.2's bare number.
Private Const OPT_NEVER_NAME As String = "vla-check-never"
Private Const OPT_STUB_PREFIX As String = "vla-check-stub-"
Private Const OPT_GROUP_PREFIX As String = "vla-check-group-"
Private Const OPT_MEMBER_PREFIX As String = "vla-check-member-"
Private Const OPT_RANK_PREFIX As String = "vla-check-rank-"
Private Const OPT_CLAUSE_PREFIX As String = "vla-check-clause-"

' ---- OPTIMIZE.3: the seconds guard -----------------------------------
'
' OPTIMIZE.0.C: a guard exists only so that a runaway formula cannot
' hold Excel indefinitely, and when it fires the answer says so. Ten
' seconds: five times the 2 s a formula is allowed to be projected at,
' and about nine times the search thorough allows on the reference
' roster (slice 4), so a search the effort already bounds never meets it
' on a sound machine. The effort is the stop that means something; this
' is the one that depends on the machine, and the status says which one
' fired.
Private Const OPT_GUARD_SECONDS As Double = 10

' ---- OPTIMIZE.3 slice 3: the ceilings a formula's grounding is held to --
'
' Fork 1, the owner's call 2026-09-24: counted in ROWS and never in
' seconds, so a workbook refuses or answers the same on every machine. A
' formula lays out at most 100,000 rows of OPTIMIZE's own grounding - its
' choices' groups and members, and its constraints' clauses - and at
' most 50,000 in any one step: about 0.9 s by the re-fitted model of
' DATALOG's own evaluator (tools/optimize3_model.ps1), and far less on
' the integer grounder, which slice 4's ladder measures before either is
' raised. DATALOG's certain part is not counted, since DATALOG has no
' ceiling either. What counts as laid out, and when a step is counted,
' is VLA_Datalog's to say (its declarations section). A command's own
' ceiling arrives with the command, slice 5.
Private Const OPT_FORMULA_STEP_ROWS As Double = 50000
Private Const OPT_FORMULA_TOTAL_ROWS As Double = 100000

' OPTIMIZE.3 slice 5: a COMMAND's ceiling, fork 1's own number - at most
' 500,000 rows in any one step, for memory, and none in all; nor any
' seconds guard, since Esc stops a command. The effort is the program's
' own, exactly as in a formula.
Private Const OPT_COMMAND_STEP_ROWS As Double = 500000

' The ceilings and the guard the run in progress is held to, set by the
' entry point before anything is grounded: OptimizeRun a formula's,
' OptimizeRunCommand a command's (SetRunMode). A size refusal records
' which, so its words say "a formula" or "a command", and only a
' formula's point to the command.
Private mForCommand As Boolean
Private mStepCeiling As Double
Private mTotalCeiling As Double
Private mGuardSeconds As Double

' A size refusal is memoized like an answer, unlike every other refusal
' (see the memo's own header), as a two-item result: this marker, under
' the reserved prefix so no query name can be it, and the refusal's own
' words. OptimizeRun raises it again on every ask.
Private Const OPT_REFUSAL_MARK As String = "vla-check-refusal"
Private Const OPT_REFUSE_CHOICE_STEP As Long = 1
Private Const OPT_REFUSE_RULE_STEP As Long = 2
Private Const OPT_REFUSE_TOTAL As Long = 3

' ---- OPTIMIZE.3: integer tuples, and a program's ground form ---------
'
' A tuple index over Long ids: open addressing over a flat key array,
' tuple numbers in the order they were first added. The atom table is
' one of these, keyed (chosen predicate's number, argument ids..., 0
' padding), so an atom's NUMBER is its tuple number - and the numbers
' are handed out in the order the atoms are decided.
Private Type OptTupleIndex
    wid As Long
    n As Long
    keys() As Long
    slot() As Long
    mask As Long
End Type

' Everything the grounding of a choice program produces, in integers,
' plus what the words need to say which rule or group is meant. See the
' section header, OPTIMIZE.3: PROGRAMS THAT CHOOSE.
Private Type OptGround
    keyWidth As Long
    atoms As OptTupleIndex
    nForms As Long
    fHasPer() As Boolean
    fGVars() As Collection
    fGRows() As Variant
    fGWidth() As Long
    fGCount() As Long
    fCtrFirst() As Long
    fCtrN() As Long
    nCtr As Long
    ctrForm() As Long
    ctrGroup() As Long
    ctrLo() As Long
    ctrHi() As Long
    ctrStart() As Long
    nCtrMem As Long
    ctrMem() As Long
    nCl As Long
    clCheck() As Long
    clRule() As Long
    clRow() As Long
    clStart() As Long
    nClLit As Long
    clLit() As Long
    rowsLaid As Double
    peakStep As Double
End Type

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
'  A refusal is never memoized, with ONE exception. Only an answer is, so
'  a program that cannot be read pays for its refusal every time - which
'  is right, since such a refusal is cheap and a stale one would be a
'  lie. The exception is OPTIMIZE.3 slice 3's size refusal, which is not
'  cheap: it is found by laying a program out up to a ceiling, after
'  DATALOG has answered its certain part. It is memoized exactly as an
'  answer is, under the same key, so it can no more be stale than an
'  answer can - and the Function Wizard's second run of the same
'  arguments, a status cell beside it, and every recalculation cost a
'  hash, which is what fork 3 settled the wizard on.
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
'  OPTIMIZE.3 slice 5: THE COMMAND
' =====================================================================
'
'  WHAT IT IS. A formula refuses a program too large to lay out inside a
'  cell (slice 3's ceilings) and stops a search after ten seconds; the
'  command runs the same program where neither holds - a step may lay out
'  up to 500,000 rows, there is no limit in all and no seconds guard, and
'  Esc stops it. The minimal one, on purpose: it runs a cell's own
'  =OPTIMIZE(...) call again, so nothing is retyped, and writes the
'  answer as values on a new sheet, so nothing of the user's is
'  overwritten. The command form OPTIMIZE.7 plans - a Table kept from run
'  to run, progress, Continue - is built on this one.
'
'  TWO WAYS IN, the owner's call (2026-09-26), one procedure: the button
'  Frazaro > Logic Engines > Optimize Selected Cell
'  (VLA_IDE.VlaOptimizeSelectedCell, which shows the new sheet), and the
'  sentence "Optimize cell C1." (english.vla's optimize-cell, the VLA form
'  (vlaoptimizecell (range "c1")), which the interpreter reaches through
'  its own Case). Only the interpreter: a compiled program runs in the
'  user's workbook with no engine beside it, so the emitter refuses the
'  call by name (VLA.bas's RefuseInterpreterOnlyCall), and the runtime's
'  helper manifest names it so Check does not refuse every program for
'  the phrasebook macro that calls it. A program's sentence leaves the
'  active sheet as it was, so the program's next sentence acts where it
'  would have.
'
'  READING THE CELL. Its formula must be exactly one call to OPTIMIZE,
'  OPTIMISE, OPTIMIZE_STATUS or OPTIMIZE_VIOLATIONS - the four take the
'  same arguments - split at its top-level commas by OptimizeCallArgs,
'  the pure half. A cell inside a spilled answer is read as the spill's
'  first cell, which holds the formula. Each argument is evaluated on the
'  cell's own sheet (Worksheet.Evaluate), so a reference means what it
'  means in the cell: the first gives the rules' text, every other a
'  range or a Table, read into relations exactly as the worksheet
'  functions read theirs. A quoted text argument is unquoted here rather
'  than evaluated, since Evaluate reads at most 255 characters.
'
'  WHILE IT RUNS, Esc is caught (EnableCancelKey = xlErrorHandler, error
'  18 - OPTIMIZE.0's host probe, C4) and the status bar says what is
'  running; both are put back before the command returns, however it
'  ends, so an interpreted program's later statements are not left
'  armed. A refusal met on the way - the program's own, or a size refusal
'  at the command's ceiling - is raised again as optimize-command-failed,
'  naming the cell, and Esc as optimize-command-stopped. Nothing is
'  written unless the run succeeds.
'
'  THE SHEET: "Optimize C1" (" (2)" and on if the name is taken), after
'  the cell's own sheet: the status in A1, where the answer came from and
'  when in A2, the answer from A4 in OPTIMIZE's own shape, and under it
'  the rules broken, if any, in OPTIMIZE_VIOLATIONS's. Every value is
'  written as a value, and text as text - an apostrophe before it, so no
'  answer can become a formula (the guard every value Frazaro writes
'  has). Returns the sheet's name.
Public Function VlaOptimizeCell(ByVal target As Range, Optional ByVal showSheet As Boolean = False) As String
    If target.Cells.CountLarge <> 1 Then
        VLA_Messages.RaiseMsg "optimize-command-one-cell", "range", CommandCellName(target), _
            "count", OptCountText(CDbl(target.Cells.CountLarge))
    End If
    Dim cell As Range
    Set cell = CommandFormulaCell(target)
    Dim cellName As String
    cellName = CommandCellName(cell)
    Dim formulaText As String
    formulaText = CommandFormulaText(cell)
    Dim fnName As String, reason As String
    Dim args As Collection
    Set args = OptimizeCallArgs(formulaText, fnName, reason)
    If args Is Nothing Then
        VLA_Messages.RaiseMsg "optimize-command-no-formula", "cell", cellName, "why", reason
    End If

    Dim rulesText As String
    rulesText = CommandRulesText(cell, CStr(args.Item(1)), cellName)
    Dim relations As Object, headerMap As Object
    Set relations = VLA_Runtime.VlaDictNew()
    Set headerMap = VLA_Runtime.VlaDictNew()
    Dim i As Long
    For i = 2 To args.Count
        AddTableArg CommandTableArg(cell, CStr(args.Item(i)), i, cellName), relations, headerMap
    Next i

    Dim prevCancel As Long
    prevCancel = Application.EnableCancelKey
    Dim prevStatus As String
    prevStatus = CStr(Application.StatusBar)
    Dim result As Collection
    Dim failNum As Long, failText As String
    On Error GoTo runFailed
    Application.EnableCancelKey = xlErrorHandler
    Application.StatusBar = "Frazaro: running " & cellName & " as a command - Esc stops it"
    Set result = OptimizeRunCommand(rulesText, relations, headerMap)
    PutStatusBarBack prevStatus
    Application.EnableCancelKey = prevCancel
    On Error GoTo 0
    VlaOptimizeCell = WriteCommandSheet(cell, cellName, formulaText, result, showSheet)
    Exit Function
runFailed:
    failNum = Err.Number
    failText = Err.Description
    PutStatusBarBack prevStatus
    Application.EnableCancelKey = prevCancel
    If failNum = 18 Then VLA_Messages.RaiseMsg "optimize-command-stopped", "cell", cellName
    VLA_Messages.RaiseMsg "optimize-command-failed", "cell", cellName, "detail", failText
End Function

' THE PURE HALF of reading the cell: a formula's text split into its one
' OPTIMIZE call's arguments, each as written, trimmed; the call's own name
' in fnName. Nothing, and the reason in words, when the formula is not
' exactly one such call - no formula, another function, anything around
' the call, or an argument left empty. Commas split only at the call's own
' level: never inside a nested call's parentheses, an array constant's
' braces, a structured reference's brackets, a quoted text or a quoted
' sheet name. A leading @ (implicit intersection, as older Excel shows a
' formula) and an add-in's own prefix before the name
' (Frazaro.xlam!OPTIMIZE) are allowed.
Public Function OptimizeCallArgs(ByVal formulaText As String, ByRef fnName As String, _
                                 ByRef reason As String) As Collection
    fnName = ""
    reason = ""
    Dim s As String
    s = Trim$(formulaText)
    If Left$(s, 1) <> "=" Then
        reason = "it holds no formula"
        Exit Function
    End If
    s = Trim$(Mid$(s, 2))
    If Left$(s, 1) = "@" Then s = Trim$(Mid$(s, 2))
    Dim openAt As Long
    openAt = InStr(1, s, "(")
    Dim head As String
    If openAt > 0 Then head = Trim$(Left$(s, openAt - 1))
    If InStrRev(head, "!") > 0 Then head = Mid$(head, InStrRev(head, "!") + 1)
    Select Case UCase$(head)
    Case "OPTIMIZE", "OPTIMISE", "OPTIMIZE_STATUS", "OPTIMIZE_VIOLATIONS"
    Case Else
        reason = "its formula is " & formulaText
        Exit Function
    End Select

    Dim args As Collection
    Set args = New Collection
    Dim i As Long, depth As Long, startAt As Long, closeAt As Long
    Dim ch As String
    Dim inText As Boolean, inName As Boolean
    startAt = openAt + 1
    i = openAt + 1
    Do While i <= Len(s)
        ch = Mid$(s, i, 1)
        If inText Then
            If ch = """" Then
                If Mid$(s, i + 1, 1) = """" Then
                    i = i + 1
                Else
                    inText = False
                End If
            End If
        ElseIf inName Then
            If ch = "'" Then
                If Mid$(s, i + 1, 1) = "'" Then
                    i = i + 1
                Else
                    inName = False
                End If
            End If
        Else
            Select Case ch
            Case """"
                inText = True
            Case "'"
                inName = True
            Case "(", "{", "["
                depth = depth + 1
            Case ")", "}", "]"
                If depth = 0 Then
                    If ch = ")" Then closeAt = i
                    Exit Do
                End If
                depth = depth - 1
            Case ","
                If depth = 0 Then
                    args.Add Trim$(Mid$(s, startAt, i - startAt))
                    startAt = i + 1
                End If
            End Select
        End If
        i = i + 1
    Loop
    If closeAt = 0 Then
        reason = "its formula, " & formulaText & ", does not close its call"
        Exit Function
    End If
    If Len(Trim$(Mid$(s, closeAt + 1))) > 0 Then
        reason = "its formula, " & formulaText & ", has more in it than the one call"
        Exit Function
    End If
    args.Add Trim$(Mid$(s, startAt, closeAt - startAt))
    Dim k As Long
    For k = 1 To args.Count
        If Len(CStr(args.Item(k))) = 0 Then
            reason = "its formula, " & formulaText & ", leaves argument " & k & " empty"
            Exit Function
        End If
    Next k
    fnName = UCase$(head)
    Set OptimizeCallArgs = args
End Function

' The status bar as it was before the command: Excel's own when Excel had
' it - the property then reads back as the text FALSE, not VBA's False
' (live-caught, slice 5's first suite run) - or the text a program had put
' there, which a command must not take away.
Private Sub PutStatusBarBack(ByVal prevStatus As String)
    If UCase$(prevStatus) = "FALSE" Then
        Application.StatusBar = False
    Else
        Application.StatusBar = prevStatus
    End If
End Sub

' The cell whose formula is run: the cell itself, or - when it lies in a
' spilled answer with no formula of its own - the spill's first cell.
' SpillParent is Excel 365's, reached late-bound so the module still
' compiles on an Excel without it.
Private Function CommandFormulaCell(ByVal target As Range) As Range
    Set CommandFormulaCell = target
    If target.HasFormula Then Exit Function
    Dim parentCell As Object
    On Error Resume Next
    Set parentCell = CallByName(target, "SpillParent", VbGet)
    On Error GoTo 0
    If Not parentCell Is Nothing Then Set CommandFormulaCell = parentCell
End Function

Private Function CommandCellName(ByVal r As Range) As String
    CommandCellName = r.Worksheet.Name & "!" & r.Address(False, False)
End Function

' The cell's formula in Excel's own English spelling, commas between the
' arguments on every machine: Formula2 where Excel has it (the formula as
' typed, with no implicit-intersection @ added), Formula where it does not.
Private Function CommandFormulaText(ByVal cell As Range) As String
    If Not cell.HasFormula Then Exit Function
    Dim f As Variant
    On Error Resume Next
    f = CallByName(cell, "Formula2", VbGet)
    On Error GoTo 0
    If IsEmpty(f) Then f = cell.Formula
    CommandFormulaText = CStr(f)
End Function

' The rules' text, as the cell's first argument gives it: a quoted text
' unquoted as written, anything else evaluated on the cell's sheet - one
' cell's value, or a value.
Private Function CommandRulesText(ByVal cell As Range, ByVal argText As String, ByVal cellName As String) As String
    If IsOneTextLiteral(argText) Then
        CommandRulesText = Replace(Mid$(argText, 2, Len(argText) - 2), """""", """")
        Exit Function
    End If
    Dim obj As Object
    Dim v As Variant
    EvaluateArg cell, argText, cellName, "first argument", obj, v
    If Not obj Is Nothing Then
        If TypeName(obj) <> "Range" Then
            CommandArgRefused cellName, "first argument", argText, "it is not text or a cell"
        End If
        Dim r As Range
        Set r = obj
        If r.Cells.CountLarge <> 1 Then
            CommandArgRefused cellName, "first argument", argText, "it refers to " & _
                OptCountText(CDbl(r.Cells.CountLarge)) & " cells, and the rules are one cell's text"
        End If
        v = r.Value
    End If
    If IsError(v) Then
        CommandArgRefused cellName, "first argument", argText, "it is an error value, not the rules' text"
    End If
    CommandRulesText = CStr(v)
End Function

' Argument i after the first: a range or a Table, as the worksheet
' functions take it.
Private Function CommandTableArg(ByVal cell As Range, ByVal argText As String, ByVal i As Long, _
                                 ByVal cellName As String) As Range
    Dim obj As Object
    Dim v As Variant
    EvaluateArg cell, argText, cellName, "argument " & i, obj, v
    If TypeName(obj) <> "Range" Then
        CommandArgRefused cellName, "argument " & i, argText, "it is not a range or a Table"
    End If
    Set CommandTableArg = obj
End Function

' One argument evaluated on the cell's own sheet: a range in obj, or else
' a value in v. Excel's Evaluate reads at most 255 characters and says
' nothing clear past them, so a longer argument is refused with the way
' round it.
Private Sub EvaluateArg(ByVal cell As Range, ByVal argText As String, ByVal cellName As String, _
                        ByVal which As String, ByRef obj As Object, ByRef v As Variant)
    If Len(argText) > 255 Then
        CommandArgRefused cellName, which, argText, "Excel evaluates at most 255 characters of an " & _
            "argument - put it in a cell of its own and refer to that cell"
    End If
    Set obj = Nothing
    On Error Resume Next
    Set obj = cell.Worksheet.Evaluate(argText)
    If Err.Number <> 0 Then
        Err.Clear
        Set obj = Nothing
        v = cell.Worksheet.Evaluate(argText)
    End If
    On Error GoTo 0
End Sub

Private Sub CommandArgRefused(ByVal cellName As String, ByVal which As String, ByVal argText As String, _
                              ByVal why As String)
    Dim shown As String
    shown = argText
    If Len(shown) > 60 Then shown = Left$(shown, 57) & "..."
    VLA_Messages.RaiseMsg "optimize-command-argument", "cell", cellName, "which", which, "arg", shown, "why", why
End Sub

' Whether an argument is one quoted text and nothing else: quotes at both
' ends, and every quote between them doubled.
Private Function IsOneTextLiteral(ByVal t As String) As Boolean
    If Len(t) < 2 Then Exit Function
    If Left$(t, 1) <> """" Or Right$(t, 1) <> """" Then Exit Function
    Dim i As Long
    i = 2
    Do While i <= Len(t) - 1
        If Mid$(t, i, 1) = """" Then
            If i + 1 > Len(t) - 1 Then Exit Function
            If Mid$(t, i + 1, 1) <> """" Then Exit Function
            i = i + 2
        Else
            i = i + 1
        End If
    Loop
    IsOneTextLiteral = True
End Function

' One table argument into the run's relations and header map, exactly as
' the four worksheet functions each read theirs.
Private Sub AddTableArg(ByVal tbl As Range, ByVal relations As Object, ByVal headerMap As Object)
    Dim nm As String
    nm = OptimizeTableArgName(tbl)
    VLA_Runtime.VlaDictSet relations, nm, VLA_Relation.RelFromRange(tbl)
    Dim colsOk As Boolean
    Dim cols As Collection
    Set cols = VLA_Relation.RangeColumnNames(tbl, colsOk)
    If colsOk Then VLA_Runtime.VlaDictSet headerMap, nm, cols
End Sub

' The answer as values on a new sheet after the cell's own, and its name.
' A sentence in a program puts the active sheet back as it was; the button
' leaves the new sheet showing. A write that fails takes its sheet away.
Private Function WriteCommandSheet(ByVal cell As Range, ByVal cellName As String, ByVal formulaText As String, _
                                   ByVal result As Collection, ByVal showSheet As Boolean) As String
    Dim wb As Workbook
    Set wb = cell.Worksheet.Parent
    Dim wasActive As Object
    Set wasActive = ActiveSheet
    Dim prevUpdating As Boolean
    prevUpdating = Application.ScreenUpdating
    Dim prevAlerts As Boolean
    prevAlerts = Application.DisplayAlerts
    Dim ws As Worksheet
    Dim failText As String
    On Error GoTo writeFailed
    Application.ScreenUpdating = False
    Set ws = wb.Worksheets.Add(After:=cell.Worksheet)
    ws.Name = CommandSheetName(wb, "Optimize " & cell.Address(False, False))
    ws.Range("A1").Value = SafeCellValue(CStr(result.Item(7)))
    ws.Range("A2").Value = SafeCellValue("From " & cellName & ", " & formulaText & " - run as a command on " & _
        Format$(Now, "yyyy-mm-dd") & " at " & Format$(Now, "hh:nn") & ".")
    Dim lastRow As Long
    lastRow = WriteValuesAt(ws.Range("A4"), OptimizeAnswerOf(result))
    If VLA_Relation.RelCount(result.Item(8)) > 0 Then
        ws.Cells(lastRow + 2, 1).Value = "Rules broken:"
        WriteValuesAt ws.Cells(lastRow + 3, 1), VLA_Relation.RelToSpilledArray(result.Item(8), _
            Array(OPT_VIOL_COL_CHECK, OPT_VIOL_COL_RULE, OPT_VIOL_COL_WHERE))
    End If
    If Not showSheet Then
        If Not wasActive Is Nothing Then wasActive.Activate
    End If
    Application.ScreenUpdating = prevUpdating
    WriteCommandSheet = ws.Name
    Exit Function
writeFailed:
    failText = Err.Description
    Application.DisplayAlerts = False
    If Not ws Is Nothing Then ws.Delete
    Application.DisplayAlerts = prevAlerts
    Application.ScreenUpdating = prevUpdating
    VLA_Messages.RaiseMsg "optimize-command-failed", "cell", cellName, "detail", failText
End Function

' A value, or a two-dimensional answer, written at topLeft as values; the
' last row it took.
Private Function WriteValuesAt(ByVal topLeft As Range, ByVal answer As Variant) As Long
    If Not IsArray(answer) Then
        topLeft.Value = SafeCellValue(answer)
        WriteValuesAt = topLeft.Row
        Exit Function
    End If
    Dim r0 As Long, r1 As Long, c0 As Long, c1 As Long
    r0 = LBound(answer, 1)
    r1 = UBound(answer, 1)
    c0 = LBound(answer, 2)
    c1 = UBound(answer, 2)
    Dim vals() As Variant
    ReDim vals(1 To r1 - r0 + 1, 1 To c1 - c0 + 1)
    Dim r As Long, c As Long
    For r = r0 To r1
        For c = c0 To c1
            vals(r - r0 + 1, c - c0 + 1) = SafeCellValue(answer(r, c))
        Next c
    Next r
    topLeft.Resize(r1 - r0 + 1, c1 - c0 + 1).Value = vals
    WriteValuesAt = topLeft.Row + (r1 - r0)
End Function

' Text as text: an apostrophe before it, so no answer can become a
' formula or a number. Anything else as it is.
Private Function SafeCellValue(ByVal v As Variant) As Variant
    If VarType(v) = vbString Then
        If Len(v) > 0 Then
            SafeCellValue = "'" & v
            Exit Function
        End If
    End If
    SafeCellValue = v
End Function

' A sheet name the workbook does not have yet: base, or base (2), (3)...
Private Function CommandSheetName(ByVal wb As Workbook, ByVal base As String) As String
    Dim nm As String
    nm = base
    Dim n As Long
    n = 1
    Do While SheetNameTaken(wb, nm)
        n = n + 1
        nm = base & " (" & n & ")"
    Loop
    CommandSheetName = nm
End Function

Private Function SheetNameTaken(ByVal wb As Workbook, ByVal nm As String) As Boolean
    Dim sh As Object
    For Each sh In wb.Sheets
        If StrComp(sh.Name, nm, vbTextCompare) = 0 Then
            SheetNameTaken = True
            Exit Function
        End If
    Next sh
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
'   9 OPTIMIZE.3: the search's own numbers, one Variant array -
'     atoms, clauses, counters, decisions, dead ends, work, the
'     search's outcome (a VLA_OptimizeSearch OPT_SEARCH_* constant, or
'     0 when nothing was searched), atoms pruned at grounding, and - from
'     slice 3 - the rows the grounding laid out and its largest step,
'     the two numbers the ceilings hold. All zero for a program with no
'     choice.
'
' A program too large for a formula to lay out raises its size refusal
' here, whether the refusal was just found or is the memo's (see the
' memo's header): it is never returned as an answer.
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
    Dim hMap As Object
    PrepareRun baseRelations, headerMap, relations, hMap

    ' The memo is consulted before the program is even read: the key is
    ' a hash of the text and the Tables, so a hit means this exact
    ' question was answered in this session, and the answer cannot have
    ' changed (decision 2). A miss costs one hash.
    Dim memoKey As String
    memoKey = OptimizeMemoKey(rulesText, relations, hMap)
    Dim outp As Collection
    If MemoHas(memoKey) Then
        Set outp = MemoGet(memoKey)
        If IsSizeRefusal(outp) Then RaiseSizeRefusal outp
        Set OptimizeRun = outp
        Exit Function
    End If

    SetRunMode False
    Set outp = RunOnce(rulesText, relations, hMap)
    MemoPut memoKey, outp
    If IsSizeRefusal(outp) Then RaiseSizeRefusal outp
    Set OptimizeRun = outp
End Function

' OPTIMIZE.3 slice 5: the same run as a COMMAND - OptimizeRun's answer
' in OptimizeRun's shape, held to a command's ceiling instead of a
' formula's: at most 500,000 rows in any one step and none in all, and no
' seconds guard, since Esc stops a command. It neither reads nor fills the
' memo: the memo is a formula's, keyed as a formula's question, and a
' command is asked on purpose and answers again. Its size refusals say "a
' command" and point nowhere further.
Public Function OptimizeRunCommand(ByVal rulesText As String, Optional ByVal baseRelations As Object, _
                                   Optional ByVal headerMap As Object) As Collection
    Dim relations As Object
    Dim hMap As Object
    PrepareRun baseRelations, headerMap, relations, hMap
    SetRunMode True
    Dim outp As Collection
    Set outp = RunOnce(rulesText, relations, hMap)
    SetRunMode False
    If IsSizeRefusal(outp) Then RaiseSizeRefusal outp
    Set OptimizeRunCommand = outp
End Function

Private Sub SetRunMode(ByVal forCommand As Boolean)
    mForCommand = forCommand
    If forCommand Then
        mStepCeiling = OPT_COMMAND_STEP_ROWS
        mTotalCeiling = 0
        mGuardSeconds = 0
    Else
        mStepCeiling = OPT_FORMULA_STEP_ROWS
        mTotalCeiling = OPT_FORMULA_TOTAL_ROWS
        mGuardSeconds = OPT_GUARD_SECONDS
    End If
End Sub

' The caller's Tables as the run's own dictionary, and an empty header
' map where none was given.
Private Sub PrepareRun(ByVal baseRelations As Object, ByVal headerMap As Object, _
                       ByRef relations As Object, ByRef hMap As Object)
    If baseRelations Is Nothing Then
        Set relations = VLA_Runtime.VlaDictNew()
    Else
        Set relations = baseRelations
    End If
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
End Sub

' One run of the program, held to the ceilings SetRunMode set.
Private Function RunOnce(ByVal rulesText As String, ByVal relations As Object, _
                         ByVal hMap As Object) As Collection
    ' The program is read ONCE, here, and the forms travel the rest of
    ' the way as objects. Every form's SHAPE is checked, the three
    ' ingredients this version still does not execute are refused by
    ' name, and the choices and constraints are collected in written
    ' order.
    Dim forms As Collection
    Set forms = VLA.VlaReadForms(rulesText)
    Dim constraints As Collection
    Dim choices As Collection
    Dim effortWork As Long
    Dim effortWords As String
    ScanProgram forms, constraints, choices, effortWork, effortWords

    ' OPTIMIZE.3: a program that chooses nothing takes OPTIMIZE.2's path
    ' unchanged, line for line - which is what keeps the parity pin
    ' honest: every DATALOG program is a program with no choice.
    If choices.Count = 0 Then
        Set RunOnce = RunZeroChoice(forms, constraints, relations, hMap)
    Else
        Set RunOnce = RunWithChoices(forms, constraints, choices, effortWork, effortWords, relations, hMap)
    End If
End Function

' OPTIMIZE.2's whole run, moved here unchanged when OPTIMIZE.3 gave
' OptimizeRun a second path: one world, the constraints checked over it.
Private Function RunZeroChoice(ByVal forms As Collection, ByVal constraints As Collection, _
                               ByVal relations As Object, ByVal hMap As Object) As Collection
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
    ' OPTIMIZE.3's item 9, the search's own numbers - all zero here,
    ' because nothing was searched, and nothing laid out beyond what
    ' DATALOG answers, which no ceiling counts.
    outp.Add Array(0&, 0&, 0&, 0&, 0&, 0&, 0&, 0&, 0&, 0&)
    Set RunZeroChoice = outp
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
    '
    ' OPTIMIZE.3: a search that spent its budget without finding a
    ' schedule answers the same empty shape. It HAS no schedule to show,
    ' and anything else in the cell - a partial one, the certain rows -
    ' would be read as an answer; OPTIMIZE_STATUS says there may be one.
    Dim broken As Boolean
    broken = (CLng(result.Item(6)) = VLA_OPTIMIZE_NO_SCHEDULE) Or _
             (CLng(result.Item(6)) = VLA_OPTIMIZE_NONE_IN_BUDGET)
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
'
' OPTIMIZE.3: the PROVEN_BEST sentence says the program makes no choices,
' so it stays exactly as written for the programs it describes, and a
' program that DOES choose keeps its "proven best" prefix and says its
' own reason after it (ChoiceBestWords). The reserved words are never
' reworded; a choice program's are a different sentence with the same
' prefix, the pattern OPTIMIZE.2 set for "no schedule".
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
'
' OPTIMIZE.3 slice 2: the choice forms and (effort ...) are no longer
' refused - the choices are collected in written order, which is the
' order their atoms are decided in, and the effort becomes the search's
' budget. A program with no (effort ...) gets VLA_OPTIMIZE_EFFORT_DEFAULT,
' and one with two is refused: which one was meant is not ours to guess.
' A program with no choice may still carry an (effort ...) - it is simply
' never spent, which is true of a search that has nothing to decide, and
' saves a user editing the choices out from having to edit it out too.
Private Sub ScanProgram(ByVal forms As Collection, ByRef constraints As Collection, _
                        ByRef choices As Collection, ByRef effortWork As Long, _
                        ByRef effortWords As String)
    Set constraints = New Collection
    Set choices = New Collection
    Dim effortForm As Collection
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
                choices.Add f
            Case "choose-between"
                CheckChoiceForm f, head, 2
                choices.Add f
            Case "choose-any"
                CheckChooseAnyForm f, head
                choices.Add f
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
                If Not effortForm Is Nothing Then VLA_Messages.RaiseMsg "optimize-effort-twice"
                Set effortForm = f
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

    If effortForm Is Nothing Then
        effortWork = EffortWorkOf(VLA_OPTIMIZE_EFFORT_DEFAULT)
        effortWords = "the default effort, " & VLA_OPTIMIZE_EFFORT_DEFAULT & ","
    Else
        Dim raw As Variant
        NthInto raw, effortForm, 2
        effortWork = EffortWorkOf(VLA_Identity.Fold(OptAtomText(raw)))
        effortWords = VLA.VlaWriteForm(effortForm)
    End If
End Sub

' A level's work count, or a number's own value - CheckEffortForm has
' already refused anything else, a number over the Long range included.
Private Function EffortWorkOf(ByVal w As String) As Long
    Select Case w
    Case "quick"
        EffortWorkOf = VLA_OPTIMIZE_WORK_QUICK
    Case "normal"
        EffortWorkOf = VLA_OPTIMIZE_WORK_NORMAL
    Case "thorough"
        EffortWorkOf = VLA_OPTIMIZE_WORK_THOROUGH
    Case Else
        EffortWorkOf = CLng(w)
    End Select
End Function

' The one refusal every still-unbuilt OPTIMIZE form reaches today. One
' id per INGREDIENT rather than one per spelling: a user who wrote
' prefer and a user who wrote avoid have the same problem, and the form
' they wrote is in the message either way.
'
' OPTIMIZE.2 removed the CONSTRAINT arm, and with it
' optimize-constraint-not-yet: a constraint is checked now, so a
' message saying "this version checks none" would be false. OPTIMIZE.3
' slice 2 removed the CHOICE and EFFORT arms the same way, and REWORDED
' the three that remain: each justified itself with "with no choice
' there is exactly one answer", which a program that chooses makes
' false. Their new words are true of every program - this version gives
' the first answer that breaks no rule - and each keeps the fragment its
' pins match, so no pin had to be reread for the rewording.
Private Sub RefuseNotYet(ByVal head As String, ByVal kind As Long)
    Select Case kind
    Case OPT_KIND_PREFERENCE
        VLA_Messages.RaiseMsg "optimize-preference-not-yet", "form", head
    Case OPT_KIND_OBJECTIVE
        VLA_Messages.RaiseMsg "optimize-objective-not-yet", "form", head
    Case OPT_KIND_KEPT
        VLA_Messages.RaiseMsg "optimize-kept-not-yet", "form", head
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
        Case "effort"
            ' OPTIMIZE.3: accepted in a program with no choice, and never
            ' spent - there is nothing to search - so never DATALOG's.
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
        ' OPTIMIZE.3: the budget is a Long, counted without wrapping, so
        ' a number past its range is refused rather than silently cut.
        If CDbl(w) > 2147483647# Then
            VLA_Messages.RaiseMsg "optimize-effort-too-large", "level", OptAtomText(raw)
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
'  OPTIMIZE.3: PROGRAMS THAT CHOOSE
' =====================================================================
'
'  WHAT HAPPENS, in order, to a program with at least one choice form.
'
'   1. WHAT IS CHOSEN. Each choice form's chosen row names a predicate the
'      search will fill - its CHOSEN predicate. Nothing else may give it
'      rows (a table argument, a fact, a rule): a row cannot be both
'      given and chosen. Every form choosing it must agree on its arity,
'      and its row is written by position, since it has no Table and so
'      no column names.
'   2. WHAT LATER ITEMS BUILD, refused in OPTIMIZE's own words: a rule
'      that reads chosen rows (OPTIMIZE.5), a pool or (per ...) row that
'      is itself chosen, and a count, sum or textjoin over chosen rows
'      inside a constraint (OPTIMIZE.4 - a cap is a second choice form
'      over the same rows, which this version propagates natively).
'   3. PASS 1, THE CERTAIN PART: DATALOG's own fixpoint over the user's
'      facts and rules, with every constraint that reads no chosen row
'      rewritten and checked exactly as OPTIMIZE.2 checks it, and the
'      chosen predicates registered as EMPTY relations so a query or a
'      constraint may name them. Plus one never-firing STUB rule per atom
'      a later pass reads, (rule (vla-check-stub-K "K") (vla-check-never
'      "x") ATOM), because DATALOG.14's PushBoundArguments narrows a
'      recursive predicate to the one constant every reader it can SEE
'      pins - and a later pass is a reader it cannot see, which would
'      read the narrowed relation as the whole. The stub makes it
'      visible. Its first atom is the empty vla-check-never, where
'      EvalRuleBody stops at once, so it costs one empty lookup.
'   4. A constraint broken by certain rows alone is "no schedule", named
'      exactly as OPTIMIZE.2 names it, and nothing is grounded further.
'   5. PASS 2, THE CHOICES, through slice 1's integer grounder
'      (VLA_Datalog.DatalogGroundRules): per form, its GROUPS (the rows
'      of its (per ...) atoms), its MEMBERS (each pool row, with the
'      group it falls in), and one counter per group. The ATOMS - every
'      row that may be chosen - are the union of every form's members
'      (clingo's own reading of several choice rules over one
'      predicate), numbered in the order they are DECIDED: forms in
'      written order; within a form, its groups in the order of their
'      (per ...) rows, the first (per ...) atom outermost; within a
'      group, its members in the order of the pool's rows. That is the
'      owner's fork 6, the Tables' own order, made precise for a form
'      with a (per ...), and it is why "the first shift is filled first"
'      is the one-sentence account of a roster's answer.
'   6. PASS 3, THE CONSTRAINTS THAT READ CHOSEN ROWS, as clauses: each is
'      grounded as a rule over the certain relations and, for each
'      chosen predicate, the relation of its POSSIBLE rows; each row of
'      that rule is one clause. A positive chosen atom in the body adds
'      "not this atom"; a negated one adds "this atom", or nothing when
'      the atom can never be chosen, since its negation then always
'      holds; a chosen consequent of a require adds "this atom". A
'      clause left with nothing in it is broken whatever is chosen - a
'      violation, in the violations table like OPTIMIZE.2's. Its rule's
'      body is PLANNED (slice 3): walked smallest atom first, then always
'      the smallest sharing a name with what is joined, so "(forbid
'      (assign S P) (assign T P) (next S T))" joins next first - the same
'      clauses, in the planned order rather than the written one.
'   5a. THE CEILINGS (slice 3), over passes 2 and 3 together: at most
'      OPT_FORMULA_TOTAL_ROWS rows laid out, at most OPT_FORMULA_STEP_ROWS
'      in any one step. Every step's rows are known before one is made,
'      so a program past a ceiling is refused having laid out no more
'      than the ceiling - naming the choice form or the rule, and the
'      number - and the refusal is memoized, since finding it was not
'      cheap.
'   7. SINGLE-ATOM PRUNING (a clause of one literal fixes its atom at
'      grounding, and costs the search nothing) and THE COUNTING
'      PRE-CHECKS, OPTIMIZE.2's three comparisons reached at last: a
'      group needing more rows than can still fill it, or a demand no
'      capacity can meet, is "no schedule" by arithmetic, before any
'      search - the pigeonhole proof clause learning cannot make.
'   8. THE SEARCH, VLA_OptimizeSearch, within the effort.
'   9. THE ANSWER: DATALOG answers the user's own (query ...) over the
'      certain relations and the chosen rows, so a query by name, a
'      query of one fact and (headless) all mean what they mean there. A
'      chosen predicate's columns are named by the first choice form's
'      own chosen row - (assign S P) gives S and P - as a rule's head
'      names its relation's.
'
'  NOT YET, and each has its item: an objective or preference (.6), the
'  kept schedule (.7), a count over chosen rows (.4), rules over chosen
'  rows (.5), and a command for the programs a formula's ceilings refuse
'  (this item's slice 5).

Private Function RunWithChoices(ByVal forms As Collection, ByVal constraints As Collection, _
                                ByVal choices As Collection, ByVal effortWork As Long, _
                                ByVal effortWords As String, ByVal relations As Object, _
                                ByVal hMap As Object) As Collection
    Dim i As Long, c As Long
    Dim f As Variant

    ' --- 1. what is chosen ----------------------------------------------
    Dim chosen As Object
    Set chosen = VLA_Runtime.VlaDictNew()
    Dim chosenNames As Collection
    Set chosenNames = New Collection
    Dim chosenAtoms As Object
    Set chosenAtoms = VLA_Runtime.VlaDictNew()
    For i = 1 To choices.Count
        RegisterChosen choices.Item(i), chosen, chosenNames, chosenAtoms, relations
    Next i
    For Each f In forms
        RefuseDefinesChosen f, chosen
    Next f

    ' --- 2. what later items build ----------------------------------------
    For i = 1 To choices.Count
        CheckChoicePools choices.Item(i), chosen
        CheckChoiceVariables choices.Item(i)
    Next i
    Dim readsChoice() As Boolean
    ReDim readsChoice(0 To constraints.Count)
    For c = 1 To constraints.Count
        readsChoice(c) = ConstraintReadsChoice(constraints.Item(c), chosen)
    Next c

    ' --- 3. pass 1, the certain part --------------------------------------
    Dim checkVars As Collection
    Dim pass1 As Collection
    Set pass1 = BuildCertainPass(forms, constraints, choices, readsChoice, chosen, checkVars)
    Dim rel1 As Object
    Set rel1 = CopyRelations(relations)
    Dim nm As Variant
    For Each nm In chosenNames
        VLA_Runtime.VlaDictSet rel1, CStr(nm), _
            VLA_Relation.RelNew(CLng(VLA_Runtime.VlaDictGet(chosen, CStr(nm))))
    Next nm
    VLA_Runtime.VlaDictSet rel1, OPT_NEVER_NAME, VLA_Relation.RelNew(1)
    Dim r1 As Collection
    Set r1 = VLA_Datalog.DatalogRunForms(pass1, rel1, hMap)
    mMemoRuns = mMemoRuns + 1

    Dim gr As OptGround
    Dim syms As VlaSymbols
    VLA_Relation.VlaSymInit syms
    Dim chosenVal() As Long
    ReDim chosenVal(0 To 0)
    Dim stats(1 To 10) As Long

    ' --- 4. a constraint broken by certain rows alone ---------------------
    Dim violations As Collection
    Set violations = CollectViolations(constraints, checkVars, r1.Item(2))
    If VLA_Relation.RelCount(violations) > 0 Then
        Set RunWithChoices = FinishChoiceRun(forms, r1, chosen, chosenNames, chosenAtoms, gr, chosenVal, _
            False, VLA_OPTIMIZE_NO_SCHEDULE, StatusSentence(VLA_OPTIMIZE_NO_SCHEDULE, constraints, violations), _
            violations, stats, hMap, syms)
        Exit Function
    End If

    ' --- 5. pass 2, the choices -------------------------------------------
    ' A ceiling passed in either pass ends the run with the size refusal,
    ' which OptimizeRun memoizes and raises.
    Dim refusal As Collection
    GroundChoices choices, r1.Item(2), hMap, syms, chosen, chosenNames, gr, refusal
    If Not refusal Is Nothing Then
        Set RunWithChoices = refusal
        Exit Function
    End If

    ' --- 6. pass 3, the constraints over chosen rows, as clauses ----------
    GroundClauses constraints, readsChoice, r1.Item(2), hMap, syms, chosen, chosenNames, gr, violations, refusal
    If Not refusal Is Nothing Then
        Set RunWithChoices = refusal
        Exit Function
    End If
    stats(1) = gr.atoms.n
    stats(2) = gr.nCl
    stats(3) = gr.nCtr
    stats(9) = CLng(gr.rowsLaid)
    stats(10) = CLng(gr.peakStep)
    If VLA_Relation.RelCount(violations) > 0 Then
        Set RunWithChoices = FinishChoiceRun(forms, r1, chosen, chosenNames, chosenAtoms, gr, chosenVal, _
            False, VLA_OPTIMIZE_NO_SCHEDULE, StatusSentence(VLA_OPTIMIZE_NO_SCHEDULE, constraints, violations), _
            violations, stats, hMap, syms)
        Exit Function
    End If

    ' --- 7. single-atom pruning, then the counting pre-checks -------------
    Dim pruned() As Long
    Dim nPruned As Long
    PruneSingleAtoms gr, pruned, nPruned
    stats(8) = nPruned
    Dim reason As String
    reason = PoolShortReason(gr, pruned, choices, syms)
    If Len(reason) = 0 Then reason = CapShortReason(gr, pruned, choices)
    If Len(reason) > 0 Then
        Set RunWithChoices = FinishChoiceRun(forms, r1, chosen, chosenNames, chosenAtoms, gr, chosenVal, _
            False, VLA_OPTIMIZE_NO_SCHEDULE, reason, violations, stats, hMap, syms)
        Exit Function
    End If

    ' --- 8. the search ----------------------------------------------------
    Dim prob As OptSearchProblem
    BuildSearchProblem gr, prob
    Dim res As OptSearchResult
    VLA_OptimizeSearch.OptSearchRun prob, effortWork, mGuardSeconds, res
    stats(4) = res.decisions
    stats(5) = res.conflicts
    stats(6) = res.work
    stats(7) = res.outcome
    Dim stateId As Long
    Dim words As String
    Dim found As Boolean
    Select Case res.outcome
    Case VLA_OptimizeSearch.OPT_SEARCH_FOUND
        stateId = VLA_OPTIMIZE_PROVEN_BEST
        words = ChoiceBestWords(res.decisions, res.conflicts)
        chosenVal = res.value
        found = True
    Case VLA_OptimizeSearch.OPT_SEARCH_NONE
        stateId = VLA_OPTIMIZE_NO_SCHEDULE
        If res.rootConflict Then
            words = ChoiceRootWords(res, gr, constraints, choices, syms)
        Else
            words = ChoiceNoneWords(res.decisions, res.conflicts)
        End If
    Case VLA_OptimizeSearch.OPT_SEARCH_BUDGET
        stateId = VLA_OPTIMIZE_NONE_IN_BUDGET
        words = ChoiceBudgetWords(effortWords, effortWork, res.decisions, res.conflicts)
    Case VLA_OptimizeSearch.OPT_SEARCH_GUARD
        stateId = VLA_OPTIMIZE_NONE_IN_BUDGET
        words = ChoiceGuardWords(res.seconds, res.work, effortWork, effortWords)
    Case Else
        VLA_Messages.RaiseMsg "optimize-internal", "detail", "handed its search a problem the search could not read"
    End Select

    ' --- 9. the answer ----------------------------------------------------
    Set RunWithChoices = FinishChoiceRun(forms, r1, chosen, chosenNames, chosenAtoms, gr, chosenVal, _
        found, stateId, words, violations, stats, hMap, syms)
End Function

' The chosen rows - one relation per chosen predicate, in the order the
' atoms were decided, empty unless a schedule was found - and DATALOG's
' own answer to the user's (query ...) over them and every certain
' relation. Items 1 to 5 are that answer's, exactly as the zero-choice
' path's are DatalogRun's.
Private Function FinishChoiceRun(ByVal forms As Collection, ByVal r1 As Collection, ByVal chosen As Object, _
                                 ByVal chosenNames As Collection, ByVal chosenAtoms As Object, _
                                 ByRef gr As OptGround, ByRef chosenVal() As Long, ByVal found As Boolean, _
                                 ByVal stateId As Long, ByVal words As String, ByVal violations As Collection, _
                                 ByRef stats() As Long, ByVal hMap As Object, ByRef syms As VlaSymbols) As Collection
    Dim rel4 As Object
    Set rel4 = CopyRelations(r1.Item(2))
    Dim pi As Long
    Dim nm As Variant
    For Each nm In chosenNames
        pi = pi + 1
        Dim arity As Long
        arity = CLng(VLA_Runtime.VlaDictGet(chosen, CStr(nm)))
        Dim rel As Collection
        Set rel = VLA_Relation.RelNew(arity)
        If found Then
            Dim a As Long
            For a = 1 To gr.atoms.n
                If chosenVal(a) = 1 Then
                    If gr.atoms.keys((a - 1) * gr.keyWidth + 1) = pi Then
                        Dim vals() As Variant
                        vals = AtomValues(gr, a, arity, syms)
                        VLA_Relation.RelTryAdd rel, vals
                    End If
                End If
            Next a
        End If
        VLA_Runtime.VlaDictSet rel4, CStr(nm), rel
    Next nm

    Dim finalForms As Collection
    Set finalForms = New Collection
    Dim f As Variant
    For Each f In forms
        Select Case OptTopHead(f)
        Case "query", "headless"
            finalForms.Add f
        End Select
    Next f
    Dim r4 As Collection
    Set r4 = VLA_Datalog.DatalogRunForms(finalForms, rel4, hMap)

    Dim outp As Collection
    Set outp = New Collection
    outp.Add r4.Item(1)
    outp.Add r4.Item(2)
    Dim item3 As Variant
    If VLA_Runtime.VlaDictHas(chosen, CStr(r4.Item(1))) Then
        item3 = ChosenHeadNames(VLA_Runtime.VlaDictGet(chosenAtoms, CStr(r4.Item(1))))
    Else
        CopyVariant item3, r1.Item(3)
    End If
    outp.Add item3
    outp.Add r4.Item(4)
    outp.Add r4.Item(5)
    outp.Add stateId
    outp.Add words
    outp.Add violations
    outp.Add Array(stats(1), stats(2), stats(3), stats(4), stats(5), stats(6), stats(7), stats(8), _
                   stats(9), stats(10))
    Set FinishChoiceRun = outp
End Function

' ---- 1 and 2: what is chosen, and what this version refuses ----------

' A choice form's chosen row, registered; the same predicate chosen again
' must have the same arity. A chosen row is written by position - it has
' no Table, so a keyed (column value) pair has nothing to name.
Private Sub RegisterChosen(ByVal cf As Collection, ByVal chosen As Object, ByVal chosenNames As Collection, _
                           ByVal chosenAtoms As Object, ByVal relations As Object)
    Dim head As String
    head = OptTopHead(cf)
    Dim atom As Collection
    Set atom = ChoiceChosenAtom(cf)
    Dim nm As String
    nm = AtomNameOf(atom)
    If atom.Count < 2 Then
        VLA_Messages.RaiseMsg "optimize-form-bad-shape", "form", head, _
            "expected", "the rows being chosen to name at least one thing after the predicate", _
            "example", ChoiceExampleFor(head)
    End If
    Dim j As Long
    For j = 2 To atom.Count
        Dim a As Variant
        NthInto a, atom, j
        If IsObject(a) Then
            VLA_Messages.RaiseMsg "optimize-form-bad-shape", "form", head, _
                "expected", "the rows being chosen written by position, like (assign S P) - a chosen row has no Table, so it has no column names", _
                "example", ChoiceExampleFor(head)
        End If
    Next j
    Dim arity As Long
    arity = atom.Count - 1
    If VLA_Runtime.VlaDictHas(chosen, nm) Then
        If CLng(VLA_Runtime.VlaDictGet(chosen, nm)) <> arity Then
            VLA_Messages.RaiseMsg "optimize-chosen-arity", "name", nm, "form", head, _
                "arity", CStr(arity), "other", CStr(VLA_Runtime.VlaDictGet(chosen, nm))
        End If
        Exit Sub
    End If
    If VLA_Runtime.VlaDictHas(relations, nm) Then
        VLA_Messages.RaiseMsg "optimize-chosen-is-defined", "name", nm, "how", "a table argument"
    End If
    VLA_Runtime.VlaDictSet chosen, nm, arity
    chosenNames.Add nm
    VLA_Runtime.VlaDictSet chosenAtoms, nm, atom
End Sub

' A fact or a rule giving a chosen predicate rows, and a rule READING
' chosen rows - OPTIMIZE.5's, whose derived atoms the search would have
' to decide as it goes.
Private Sub RefuseDefinesChosen(ByVal f As Variant, ByVal chosen As Object)
    Dim head As String
    head = OptTopHead(f)
    If head <> "fact" And head <> "rule" Then Exit Sub
    Dim lst As Collection
    Set lst = f
    If lst.Count < 2 Then Exit Sub
    Dim a As Variant
    NthInto a, lst, 2
    If Not IsObject(a) Then Exit Sub
    Dim nm As String
    nm = AtomNameOf(a)
    If VLA_Runtime.VlaDictHas(chosen, nm) Then
        VLA_Messages.RaiseMsg "optimize-chosen-is-defined", "name", nm, "how", "a (" & head & " ...)"
    End If
    If head <> "rule" Then Exit Sub
    Dim j As Long
    For j = 3 To lst.Count
        Dim item As Variant, src As Variant
        NthInto item, lst, j
        If BodyReadKind(item, src) > 0 Then
            If VLA_Runtime.VlaDictHas(chosen, AtomNameOf(src)) Then
                VLA_Messages.RaiseMsg "optimize-rule-over-choice-not-yet", "predicate", nm, "chosen", AtomNameOf(src)
            End If
        End If
    Next j
End Sub

' A pool, and every (per ...) row, must be a row the program already
' has - not a test, not a wrapper, and not a row being chosen.
Private Sub CheckChoicePools(ByVal cf As Collection, ByVal chosen As Object)
    Dim head As String
    head = OptTopHead(cf)
    CheckPoolAtom head, ChoicePoolAtom(cf), "the pool they are chosen from", chosen
    Dim per As Collection
    Set per = ChoicePer(cf)
    If per Is Nothing Then Exit Sub
    Dim j As Long
    For j = 2 To per.Count
        Dim pa As Variant
        NthInto pa, per, j
        CheckPoolAtom head, pa, "a row to group by", chosen
    Next j
End Sub

Private Sub CheckPoolAtom(ByVal head As String, ByVal atom As Variant, ByVal role As String, _
                          ByVal chosen As Object)
    Dim lst As Collection
    Set lst = atom
    Dim h As Variant
    NthInto h, lst, 1
    Select Case VLA_Identity.Fold(CStr(h))
    Case "not", "count", "sum", "let", "textjoin", ">", "<", "<=", ">=", "=", "<>", _
         "text-starts-with", "text-ends-with", "text-contains"
        VLA_Messages.RaiseMsg "optimize-not-a-row", "form", head, "role", role, "text", OptAtomText(h)
    End Select
    If VLA_Runtime.VlaDictHas(chosen, AtomNameOf(atom)) Then
        VLA_Messages.RaiseMsg "optimize-pool-reads-choice", "form", head, "name", AtomNameOf(atom)
    End If
End Sub

' Every name in the chosen row must come from the pool or a (per ...)
' row, or the row ranges over nothing in particular; a count that is a
' name must come from a (per ...) row, since it is the count FOR a group.
Private Sub CheckChoiceVariables(ByVal cf As Collection)
    Dim head As String
    head = OptTopHead(cf)
    Dim poolVars As Collection, poolSeen As Object
    Set poolVars = New Collection
    Set poolSeen = VLA_Runtime.VlaDictNew()
    CollectBoundVars ChoicePoolAtom(cf), poolVars, poolSeen
    Dim perVars As Collection, perSeen As Object
    Set perVars = New Collection
    Set perSeen = VLA_Runtime.VlaDictNew()
    Dim per As Collection
    Set per = ChoicePer(cf)
    Dim j As Long
    If Not per Is Nothing Then
        For j = 2 To per.Count
            Dim pa As Variant
            NthInto pa, per, j
            CollectBoundVars pa, perVars, perSeen
        Next j
    End If
    Dim atom As Collection
    Set atom = ChoiceChosenAtom(cf)
    For j = 2 To atom.Count
        Dim a As Variant
        NthInto a, atom, j
        If OptIsVariable(a) Then
            If Not VLA_Runtime.VlaDictHas(poolSeen, CStr(a)) And Not VLA_Runtime.VlaDictHas(perSeen, CStr(a)) Then
                VLA_Messages.RaiseMsg "optimize-choice-unbound", "form", head, "var", CStr(a)
            End If
        End If
    Next j
    For j = 2 To 1 + ChoiceCountSlots(head)
        Dim cr As Variant
        NthInto cr, cf, j
        If OptIsVariable(cr) Then
            If Not VLA_Runtime.VlaDictHas(perSeen, CStr(cr)) Then
                VLA_Messages.RaiseMsg "optimize-count-unbound", "form", head, "var", CStr(cr)
            End If
        End If
    Next j
End Sub

' Whether a constraint reads a chosen predicate anywhere - its body, or
' a require's consequent - refusing a count, sum or textjoin over chosen
' rows, which is OPTIMIZE.4's.
Private Function ConstraintReadsChoice(ByVal form As Variant, ByVal chosen As Object) As Boolean
    Dim lst As Collection
    Set lst = form
    Dim isRequire As Boolean
    isRequire = (OptTopHead(form) = "require")
    Dim bodyFrom As Long
    If isRequire Then bodyFrom = 3 Else bodyFrom = 2
    Dim j As Long
    For j = bodyFrom To lst.Count
        Dim item As Variant, src As Variant
        NthInto item, lst, j
        Dim kind As Long
        kind = BodyReadKind(item, src)
        If kind > 0 Then
            If VLA_Runtime.VlaDictHas(chosen, AtomNameOf(src)) Then
                If kind = 3 Then
                    Dim w As Variant
                    NthInto w, item, 1
                    VLA_Messages.RaiseMsg "optimize-count-over-choice-not-yet", _
                        "word", VLA_Identity.Fold(OptAtomText(w)), "chosen", AtomNameOf(src)
                End If
                ConstraintReadsChoice = True
            End If
        End If
    Next j
    If isRequire Then
        Dim cons As Variant
        NthInto cons, lst, 2
        If VLA_Runtime.VlaDictHas(chosen, AtomNameOf(cons)) Then ConstraintReadsChoice = True
    End If
End Function

' What a rule-body item reads, by DATALOG's own shapes: 1, a positive
' atom (the item itself); 2, a negated atom - (not ATOM); 3, an
' aggregate's source - (count V ATOM), (sum V ATOM), (textjoin V S
' ATOM); 0, nothing - a comparison, a let, a text test, or a shape
' DATALOG will refuse in its own words. Classified by CStr with any
' quote left on, CollectBoundVars' own reason: a quoted "not" is an
' atom whose predicate is spelled not.
Private Function BodyReadKind(ByVal item As Variant, ByRef src As Variant) As Long
    If Not IsObject(item) Then Exit Function
    Dim lst As Collection
    Set lst = item
    If lst.Count < 1 Then Exit Function
    Dim h As Variant
    NthInto h, lst, 1
    If IsObject(h) Then Exit Function
    Select Case VLA_Identity.Fold(CStr(h))
    Case ">", "<", "<=", ">=", "=", "<>", "text-starts-with", "text-ends-with", "text-contains", "let"
        Exit Function
    Case "not"
        ' A text test under not - DATALOG.11's inverted test - comes back
        ' as 2 with the test as its source, and is harmless there: every
        ' caller asks whether the source's name is CHOSEN, or a rule's
        ' head, and a test word is neither.
        If lst.Count <> 2 Then Exit Function
        NthInto src, lst, 2
        If Not IsObject(src) Then Exit Function
        BodyReadKind = 2
    Case "count", "sum"
        If lst.Count <> 3 Then Exit Function
        NthInto src, lst, 3
        If Not IsObject(src) Then Exit Function
        BodyReadKind = 3
    Case "textjoin"
        If lst.Count <> 4 Then Exit Function
        NthInto src, lst, 4
        If Not IsObject(src) Then Exit Function
        BodyReadKind = 3
    Case Else
        CopyVariant src, item
        BodyReadKind = 1
    End Select
End Function

' An atom's predicate, folded, as DATALOG names it - "" for anything
' that is not a list with a word first.
Private Function AtomNameOf(ByVal atom As Variant) As String
    If Not IsObject(atom) Then Exit Function
    Dim lst As Collection
    Set lst = atom
    If lst.Count < 1 Then Exit Function
    Dim h As Variant
    NthInto h, lst, 1
    If IsObject(h) Then Exit Function
    AtomNameOf = VLA_Identity.Fold(OptAtomText(h))
End Function

' A negated chosen row says a row is NOT chosen, and asks it about every
' value of a name nothing else in the rule binds - a "for all" nobody
' wrote. Refused before DATALOG sees it, in words about the row the user
' did write.
Private Sub CheckNegatedChosen(ByVal atom As Variant, ByVal seen As Object)
    Dim lst As Collection
    Set lst = atom
    Dim j As Long
    For j = 2 To lst.Count
        Dim a As Variant
        NthInto a, lst, j
        If OptIsVariable(a) Then
            If Not VLA_Runtime.VlaDictHas(seen, CStr(a)) Then
                VLA_Messages.RaiseMsg "optimize-negated-choice-unbound", "var", CStr(a), "predicate", AtomNameOf(atom)
            End If
        End If
    Next j
End Sub

' ---- the choice form's parts ------------------------------------------

Private Function ChoiceCountSlots(ByVal head As String) As Long
    Select Case head
    Case "choose-exactly", "choose-at-least", "choose-at-most"
        ChoiceCountSlots = 1
    Case "choose-between"
        ChoiceCountSlots = 2
    End Select
End Function

Private Function ChoiceChosenAtom(ByVal cf As Collection) As Collection
    Set ChoiceChosenAtom = cf.Item(2 + ChoiceCountSlots(OptTopHead(cf)))
End Function

Private Function ChoicePoolAtom(ByVal cf As Collection) As Collection
    Set ChoicePoolAtom = cf.Item(3 + ChoiceCountSlots(OptTopHead(cf)))
End Function

' The (per ...) list, or Nothing when the form has none.
Private Function ChoicePer(ByVal cf As Collection) As Collection
    Dim at As Long
    at = 4 + ChoiceCountSlots(OptTopHead(cf))
    If cf.Count >= at Then Set ChoicePer = cf.Item(at)
End Function

' The rows a (per ...) groups by, as a list of atoms.
Private Function PerAtomsOf(ByVal per As Collection) As Collection
    Dim outp As Collection
    Set outp = New Collection
    Dim j As Long
    For j = 2 To per.Count
        outp.Add per.Item(j)
    Next j
    Set PerAtomsOf = outp
End Function

' A group family's name, for the counting sentences: its (per ...) rows'
' predicates, or the form itself when it counts over its whole pool.
Private Function GroupFamilyName(ByVal cf As Collection) As String
    Dim per As Collection
    Set per = ChoicePer(cf)
    If per Is Nothing Then
        GroupFamilyName = "(" & OptTopHead(cf) & " ...)"
    Else
        GroupFamilyName = PerNamesOf(per)
    End If
End Function

Private Function PerNamesOf(ByVal per As Collection) As String
    Dim s As String
    Dim j As Long
    For j = 2 To per.Count
        If j > 2 Then s = s & " and "
        s = s & AtomNameOf(per.Item(j))
    Next j
    PerNamesOf = s
End Function

' Whether a form caps every one of its groups at the same literal number
' - what the pigeonhole comparison needs of the capping side.
Private Function UniformLiteralCap(ByVal cf As Collection, ByRef cap As Long) As Boolean
    Dim head As String
    head = OptTopHead(cf)
    Dim at As Long
    Select Case head
    Case "choose-exactly", "choose-at-most"
        at = 2
    Case "choose-between"
        at = 3
    Case Else
        Exit Function
    End Select
    Dim raw As Variant
    NthInto raw, cf, at
    If OptIsVariable(raw) Then Exit Function
    cap = CLng(OptAtomText(raw))
    UniformLiteralCap = True
End Function

' A demand family that places each of its groups exactly once - the
' sentence then counts THINGS to place rather than rows the rules need.
Private Function ExactlyOneEach(ByVal cf As Collection) As Boolean
    If OptTopHead(cf) <> "choose-exactly" Then Exit Function
    If ChoicePer(cf) Is Nothing Then Exit Function
    Dim raw As Variant
    NthInto raw, cf, 2
    If OptIsVariable(raw) Then Exit Function
    ExactlyOneEach = (OptAtomText(raw) = "1")
End Function

' ---- 3: pass 1 --------------------------------------------------------

' Pass 1's program: the user's own facts, rules, query and headless
' directive; each constraint that reads only certain rows, rewritten as
' OPTIMIZE.2 checks it; and one never-firing stub per atom a later pass
' reads (the section header has why). checkVars gets one entry per
' constraint, an empty one for a constraint pass 3 grounds instead.
'
' A stub is written only for an atom over a predicate a RULE derives:
' PushBoundArguments narrows nothing else (its condition 2 - no fact and
' no Table may put rows into what it narrows), so a stub over a Table,
' a fact or a test word would be a rule with nothing to protect.
Private Function BuildCertainPass(ByVal forms As Collection, ByVal constraints As Collection, _
                                  ByVal choices As Collection, ByRef readsChoice() As Boolean, _
                                  ByVal chosen As Object, ByRef checkVars As Collection) As Collection
    Dim outp As Collection
    Set outp = New Collection
    Dim ruleHeads As Object
    Set ruleHeads = VLA_Runtime.VlaDictNew()
    Dim f As Variant
    For Each f In forms
        Select Case OptTopHead(f)
        Case "choose-exactly", "choose-at-least", "choose-at-most", "choose-between", "choose-any", _
             "require", "forbid", "effort"
        Case Else
            outp.Add f
            If OptTopHead(f) = "rule" Then
                Dim rl As Collection
                Set rl = f
                If rl.Count >= 2 Then
                    Dim rh As Variant
                    NthInto rh, rl, 2
                    If Len(AtomNameOf(rh)) > 0 Then VLA_Runtime.VlaDictSet ruleHeads, AtomNameOf(rh), True
                End If
            End If
        End Select
    Next f
    Set checkVars = New Collection
    Dim c As Long
    For c = 1 To constraints.Count
        Dim vars As Collection
        If readsChoice(c) Then
            checkVars.Add New Collection
        Else
            outp.Add RewriteConstraint(constraints.Item(c), c, vars)
            checkVars.Add vars
        End If
    Next c
    Dim stubN As Long
    Dim i As Long, j As Long
    For i = 1 To choices.Count
        Dim cf As Collection
        Set cf = choices.Item(i)
        AddStub outp, stubN, ChoicePoolAtom(cf), ruleHeads
        Dim per As Collection
        Set per = ChoicePer(cf)
        If Not per Is Nothing Then
            For j = 2 To per.Count
                AddStub outp, stubN, per.Item(j), ruleHeads
            Next j
        End If
    Next i
    For c = 1 To constraints.Count
        If readsChoice(c) Then AddConstraintStubs outp, stubN, constraints.Item(c), chosen, ruleHeads
    Next c
    Set BuildCertainPass = outp
End Function

' (rule (vla-check-stub-K "K") (vla-check-never "x") ATOM), for an atom
' over a predicate a rule derives.
Private Sub AddStub(ByVal outp As Collection, ByRef stubN As Long, ByVal atom As Variant, _
                    ByVal ruleHeads As Object)
    If Not IsObject(atom) Then Exit Sub
    If Not VLA_Runtime.VlaDictHas(ruleHeads, AtomNameOf(atom)) Then Exit Sub
    stubN = stubN + 1
    Dim head As Collection
    Set head = New Collection
    head.Add OPT_STUB_PREFIX & stubN
    head.Add Chr$(34) & CStr(stubN)
    Dim never As Collection
    Set never = New Collection
    never.Add OPT_NEVER_NAME
    never.Add Chr$(34) & "x"
    Dim r As Collection
    Set r = New Collection
    r.Add "rule"
    r.Add head
    r.Add never
    r.Add atom
    outp.Add r
End Sub

' Every certain row a constraint over chosen rows reads - positive,
' negated or aggregated, each written positively in its stub, which is
' all PushBoundArguments needs to see, and a require's certain
' consequent too.
Private Sub AddConstraintStubs(ByVal outp As Collection, ByRef stubN As Long, ByVal form As Variant, _
                               ByVal chosen As Object, ByVal ruleHeads As Object)
    Dim lst As Collection
    Set lst = form
    Dim isRequire As Boolean
    isRequire = (OptTopHead(form) = "require")
    Dim bodyFrom As Long
    If isRequire Then bodyFrom = 3 Else bodyFrom = 2
    Dim j As Long
    For j = bodyFrom To lst.Count
        Dim item As Variant, src As Variant
        NthInto item, lst, j
        If BodyReadKind(item, src) > 0 Then
            If Not VLA_Runtime.VlaDictHas(chosen, AtomNameOf(src)) Then AddStub outp, stubN, src, ruleHeads
        End If
    Next j
    If isRequire Then
        Dim cons As Variant
        NthInto cons, lst, 2
        If IsObject(cons) Then
            If Not VLA_Runtime.VlaDictHas(chosen, AtomNameOf(cons)) Then AddStub outp, stubN, cons, ruleHeads
        End If
    End If
End Sub

' The same dictionary's entries in a new dictionary - the objects are
' shared, as DATALOG shares them; only the names are the run's own.
Private Function CopyRelations(ByVal d As Object) As Object
    Dim outp As Object
    Set outp = VLA_Runtime.VlaDictNew()
    Dim k As Variant
    For Each k In VLA_Runtime.VlaDictKeys(d)
        VLA_Runtime.VlaDictSet outp, CStr(k), VLA_Runtime.VlaDictGet(d, k)
    Next k
    Set CopyRelations = outp
End Function

' (rule (NAME "N" V1 ... Vk TERM...) BODY...) - the one shape every rule
' this section generates has: its own name, its own number as a constant
' first slot (DATALOG refuses a head with no argument, and a body that
' binds nothing would need one - OPTIMIZE.2's vla-check-N reason), the
' variables it carries, any further terms exactly as written, then the
' body, whose items are the user's own forms, reused by reference.
Private Function BuildRule(ByVal headName As String, ByVal n As Long, ByVal vars As Collection, _
                           ByVal terms As Collection, ByVal body As Collection) As Collection
    Dim head As Collection
    Set head = New Collection
    head.Add headName
    head.Add Chr$(34) & CStr(n)
    Dim v As Variant
    For Each v In vars
        head.Add CStr(v)
    Next v
    If Not terms Is Nothing Then
        Dim t As Variant
        For Each t In terms
            head.Add t
        Next t
    End If
    Dim r As Collection
    Set r = New Collection
    r.Add "rule"
    r.Add head
    Dim b As Variant
    For Each b In body
        r.Add b
    Next b
    Set BuildRule = r
End Function

' ---- 5: pass 2, the choices --------------------------------------------

Private Sub GroundChoices(ByVal choices As Collection, ByVal certain As Object, ByVal hMap As Object, _
                          ByRef syms As VlaSymbols, ByVal chosen As Object, ByVal chosenNames As Collection, _
                          ByRef gr As OptGround, ByRef refusal As Collection)
    Dim nF As Long
    nF = choices.Count
    gr.nForms = nF
    ReDim gr.fHasPer(1 To nF)
    ReDim gr.fGVars(1 To nF)
    ReDim gr.fGRows(1 To nF)
    ReDim gr.fGWidth(1 To nF)
    ReDim gr.fGCount(1 To nF)
    ReDim gr.fCtrFirst(1 To nF)
    ReDim gr.fCtrN(1 To nF)

    ' The atom key: the chosen predicate's number, then its arguments'
    ' ids, padded with 0 - which no symbol is - to the widest arity.
    Dim maxArity As Long
    Dim nm As Variant
    For Each nm In chosenNames
        If CLng(VLA_Runtime.VlaDictGet(chosen, CStr(nm))) > maxArity Then
            maxArity = CLng(VLA_Runtime.VlaDictGet(chosen, CStr(nm)))
        End If
    Next nm
    gr.keyWidth = 1 + maxArity
    TupInit gr.atoms, gr.keyWidth
    gr.nCtr = 0
    gr.nCtrMem = 0
    ReDim gr.ctrForm(1 To 16)
    ReDim gr.ctrGroup(1 To 16)
    ReDim gr.ctrLo(1 To 16)
    ReDim gr.ctrHi(1 To 16)
    ReDim gr.ctrStart(1 To 17)
    gr.ctrStart(1) = 1
    ReDim gr.ctrMem(1 To 16)

    ' One batch: per form, its groups (when it has a (per ...)), its
    ' members, and - with two or more (per ...) atoms - one ranking rule
    ' per atom that names anything, for the order of the groups.
    Dim batch As Collection
    Set batch = New Collection
    ' formOf(k): the choice form batch rule k was written for, so a ceiling
    ' passed in rule k is refused in that form's name.
    Dim formOf As Collection
    Set formOf = New Collection
    Dim groupRule() As Long, memberRule() As Long
    ReDim groupRule(1 To nF)
    ReDim memberRule(1 To nF)
    Dim rankInfo() As Collection
    ReDim rankInfo(1 To nF)
    Dim i As Long, j As Long
    For i = 1 To nF
        Dim cf As Collection
        Set cf = choices.Item(i)
        Dim per As Collection
        Set per = ChoicePer(cf)
        gr.fHasPer(i) = Not per Is Nothing
        Dim gv As Collection, gSeen As Object
        Set gv = New Collection
        Set gSeen = VLA_Runtime.VlaDictNew()
        Dim body As Collection
        Set body = New Collection
        body.Add ChoicePoolAtom(cf)
        If gr.fHasPer(i) Then
            For j = 2 To per.Count
                Dim pa As Variant
                NthInto pa, per, j
                CollectBoundVars pa, gv, gSeen
                body.Add pa
            Next j
            batch.Add BuildRule(OPT_GROUP_PREFIX & i, i, gv, Nothing, PerAtomsOf(per))
            formOf.Add i
            groupRule(i) = batch.Count
        End If
        Set gr.fGVars(i) = gv
        Dim terms As Collection
        Set terms = New Collection
        Dim cAtom As Collection
        Set cAtom = ChoiceChosenAtom(cf)
        For j = 2 To cAtom.Count
            terms.Add cAtom.Item(j)
        Next j
        batch.Add BuildRule(OPT_MEMBER_PREFIX & i, i, gv, terms, body)
        formOf.Add i
        memberRule(i) = batch.Count
        Set rankInfo(i) = New Collection
        If gr.fHasPer(i) Then
            If per.Count > 2 Then
                For j = 2 To per.Count
                    NthInto pa, per, j
                    Dim av As Collection, aSeen As Object
                    Set av = New Collection
                    Set aSeen = VLA_Runtime.VlaDictNew()
                    CollectBoundVars pa, av, aSeen
                    If av.Count > 0 Then
                        Dim one As Collection
                        Set one = New Collection
                        one.Add pa
                        batch.Add BuildRule(OPT_RANK_PREFIX & i & "-" & j, i, av, Nothing, one)
                        formOf.Add i
                        rankInfo(i).Add Array(batch.Count, av)
                    End If
                Next j
            End If
        End If
    Next i

    ' In written order, as slice 2 grounds them: a member rule's first
    ' atom is its pool, and its rows must come out in the pool's order.
    ' Only the ceilings are new here - pass 3 alone is planned.
    Dim out As Collection
    Set out = VLA_Datalog.DatalogGroundRules(batch, certain, hMap, syms, stepCeiling:=mStepCeiling, _
                                             totalCeiling:=mTotalCeiling)
    Dim over As Variant
    over = VLA_Datalog.DatalogGroundOverflow()
    If Not IsEmpty(over) Then
        Set refusal = SizeRefusal(over, choices.Item(CLng(formOf.Item(CLng(over(1))))), True)
        Exit Sub
    End If
    gr.rowsLaid = VLA_Datalog.DatalogGroundRowsLaid()
    gr.peakStep = VLA_Datalog.DatalogGroundPeakStep()
    For i = 1 To nF
        GroundOneChoice choices.Item(i), i, out, groupRule(i), memberRule(i), rankInfo(i), _
            chosenNames, syms, gr
    Next i
End Sub

' One choice form's groups, atoms and counters, from its grounded rules.
Private Sub GroundOneChoice(ByVal cf As Collection, ByVal i As Long, ByVal out As Collection, _
                            ByVal gRuleIx As Long, ByVal mRuleIx As Long, ByVal rankInfo As Collection, _
                            ByVal chosenNames As Collection, ByRef syms As VlaSymbols, _
                            ByRef gr As OptGround)
    Dim head As String
    head = OptTopHead(cf)
    Dim k As Long
    k = gr.fGVars(i).Count
    Dim cAtom As Collection
    Set cAtom = ChoiceChosenAtom(cf)
    Dim m As Long
    m = cAtom.Count - 1
    Dim predIx As Long
    predIx = NameIndex(chosenNames, AtomNameOf(cAtom))

    ' --- the groups, and the order they are decided in -----------------
    Dim nG As Long, wG As Long
    Dim gRows() As Long
    Dim gIndex As OptTupleIndex
    TupInit gIndex, k
    Dim r As Long
    If gr.fHasPer(i) Then
        Dim gOut As Variant
        gOut = out.Item(gRuleIx)
        wG = CLng(gOut(1))
        nG = CLng(gOut(2))
        gRows = gOut(3)
        For r = 0 To nG - 1
            TupFind gIndex, gRows, r * wG + 1, True
        Next r
    Else
        nG = 1
        wG = 1
        ReDim gRows(0 To 1)
    End If
    gr.fGRows(i) = gRows
    gr.fGWidth(i) = wG
    gr.fGCount(i) = nG
    Dim order() As Long
    ReDim order(0 To nG)
    For r = 1 To nG
        order(r) = r
    Next r
    If rankInfo.Count > 0 Then SortGroupsByRanks order, nG, gRows, wG, gr.fGVars(i), rankInfo, out

    ' --- the members, each in its group, in the pool's order ----------
    Dim mOut As Variant
    mOut = out.Item(mRuleIx)
    Dim wM As Long, nM As Long
    wM = CLng(mOut(1))
    nM = CLng(mOut(2))
    Dim mRows() As Long
    mRows = mOut(3)
    Dim grpOf() As Long, cnt() As Long, startOf() As Long, fillAt() As Long, bucket() As Long
    ReDim grpOf(0 To nM)
    ReDim cnt(0 To nG + 1)
    Dim g As Long
    For r = 0 To nM - 1
        If gr.fHasPer(i) Then
            g = TupFind(gIndex, mRows, r * wM + 1, False)
            If g = 0 Then VLA_Messages.RaiseMsg "optimize-internal", "detail", "grounded a choice's member into no group"
        Else
            g = 1
        End If
        grpOf(r) = g
        cnt(g) = cnt(g) + 1
    Next r
    ReDim startOf(0 To nG + 1)
    ReDim fillAt(0 To nG + 1)
    For g = 1 To nG
        startOf(g + 1) = startOf(g) + cnt(g)
        fillAt(g) = startOf(g)
    Next g
    ReDim bucket(0 To nM)
    For r = 0 To nM - 1
        g = grpOf(r)
        bucket(fillAt(g)) = r
        fillAt(g) = fillAt(g) + 1
    Next r

    ' --- the counts: a literal, or the group's own value of a name ----
    Dim slots As Long
    slots = ChoiceCountSlots(head)
    Dim litCount(1 To 2) As Long, varPos(1 To 2) As Long, varName(1 To 2) As String
    Dim j As Long
    For j = 1 To slots
        Dim cr As Variant
        NthInto cr, cf, 1 + j
        If OptIsVariable(cr) Then
            varName(j) = CStr(cr)
            varPos(j) = VarPosition(gr.fGVars(i), varName(j))
        Else
            litCount(j) = CLng(OptAtomText(cr))
        End If
    Next j

    ' --- the atoms and the counters, group by group ------------------
    Dim kbuf() As Long
    ReDim kbuf(0 To gr.keyWidth)
    kbuf(1) = predIx
    gr.fCtrFirst(i) = gr.nCtr + 1
    Dim gi As Long, q As Long, ai As Long, nMem As Long, id As Long
    Dim mem() As Long
    Dim cnts(1 To 2) As Long
    Dim lo As Long, hi As Long
    For gi = 1 To nG
        g = order(gi)
        ReDim mem(1 To startOf(g + 1) - startOf(g) + 1)
        nMem = 0
        For q = startOf(g) To startOf(g + 1) - 1
            r = bucket(q)
            For ai = 1 To m
                kbuf(1 + ai) = mRows(r * wM + 1 + k + ai)
            Next ai
            nMem = nMem + 1
            mem(nMem) = TupFind(gr.atoms, kbuf, 0, True)
        Next q
        If slots > 0 Then
            For j = 1 To slots
                If varPos(j) > 0 Then
                    id = gRows((g - 1) * wG + 1 + varPos(j))
                    If Not WholeCountOf(syms, id, cnts(j)) Then
                        VLA_Messages.RaiseMsg "optimize-count-bad-value", "form", head, "var", varName(j), _
                            "group", GroupKeyWords(gr, i, g, syms), "value", SafeValueText(syms.rep(id))
                    End If
                Else
                    cnts(j) = litCount(j)
                End If
            Next j
            Select Case head
            Case "choose-exactly"
                lo = cnts(1)
                hi = cnts(1)
            Case "choose-at-least"
                lo = cnts(1)
                hi = VLA_OptimizeSearch.OPT_SEARCH_NO_MOST
            Case "choose-at-most"
                lo = 0
                hi = cnts(1)
            Case "choose-between"
                lo = cnts(1)
                hi = cnts(2)
            End Select
            AddCounter gr, i, g, lo, hi, mem, nMem
        End If
    Next gi
    gr.fCtrN(i) = gr.nCtr - gr.fCtrFirst(i) + 1
End Sub

Private Sub AddCounter(ByRef gr As OptGround, ByVal formIx As Long, ByVal groupIx As Long, _
                       ByVal lo As Long, ByVal hi As Long, ByRef mem() As Long, ByVal nMem As Long)
    Dim k As Long
    k = gr.nCtr + 1
    Do While k + 1 > UBound(gr.ctrStart)
        ReDim Preserve gr.ctrStart(1 To 2 * UBound(gr.ctrStart))
    Loop
    Do While k > UBound(gr.ctrForm)
        ReDim Preserve gr.ctrForm(1 To 2 * UBound(gr.ctrForm))
        ReDim Preserve gr.ctrGroup(1 To UBound(gr.ctrForm))
        ReDim Preserve gr.ctrLo(1 To UBound(gr.ctrForm))
        ReDim Preserve gr.ctrHi(1 To UBound(gr.ctrForm))
    Loop
    Do While gr.nCtrMem + nMem > UBound(gr.ctrMem)
        ReDim Preserve gr.ctrMem(1 To 2 * UBound(gr.ctrMem))
    Loop
    Dim j As Long
    For j = 1 To nMem
        gr.ctrMem(gr.nCtrMem + j) = mem(j)
    Next j
    gr.nCtrMem = gr.nCtrMem + nMem
    gr.ctrForm(k) = formIx
    gr.ctrGroup(k) = groupIx
    gr.ctrLo(k) = lo
    gr.ctrHi(k) = hi
    gr.nCtr = k
    gr.ctrStart(k + 1) = gr.nCtrMem + 1
End Sub

' Sorts order(1..n) so the groups run in the order of their (per ...)
' rows, the first (per ...) atom outermost: each group's own values of
' each atom's names are looked up among that atom's rows, whose positions
' are its ranks, and the groups' rank vectors compared left to right. A
' merge sort, so equal vectors keep the grounder's own order.
Private Sub SortGroupsByRanks(ByRef order() As Long, ByVal n As Long, ByRef gRows() As Long, _
                              ByVal wG As Long, ByVal gv As Collection, ByVal rankInfo As Collection, _
                              ByVal out As Collection)
    Dim nR As Long
    nR = rankInfo.Count
    Dim rk() As Long
    ReDim rk(0 To n * nR + 1)
    Dim j As Long, g As Long, p As Long, r As Long
    For j = 1 To nR
        Dim info As Variant
        info = rankInfo.Item(j)
        Dim av As Collection
        Set av = info(1)
        Dim entry As Variant
        entry = out.Item(CLng(info(0)))
        Dim wR As Long, nRows As Long
        wR = CLng(entry(1))
        nRows = CLng(entry(2))
        Dim rRows() As Long
        rRows = entry(3)
        Dim idx As OptTupleIndex
        TupInit idx, av.Count
        For r = 0 To nRows - 1
            TupFind idx, rRows, r * wR + 1, True
        Next r
        Dim posIn() As Long
        ReDim posIn(0 To av.Count)
        For p = 1 To av.Count
            posIn(p) = VarPosition(gv, CStr(av.Item(p)))
        Next p
        Dim kbuf() As Long
        ReDim kbuf(0 To av.Count)
        For g = 1 To n
            For p = 1 To av.Count
                kbuf(p) = gRows((g - 1) * wG + 1 + posIn(p))
            Next p
            rk((g - 1) * nR + j) = TupFind(idx, kbuf, 0, False)
        Next g
    Next j
    MergeSortByVectors order, n, rk, nR
End Sub

Private Sub MergeSortByVectors(ByRef order() As Long, ByVal n As Long, ByRef rk() As Long, ByVal w As Long)
    If n < 2 Then Exit Sub
    Dim tmp() As Long
    ReDim tmp(0 To n)
    Dim runLen As Long, loPos As Long, midPos As Long, hiPos As Long
    Dim a As Long, b As Long, k As Long
    runLen = 1
    Do While runLen < n
        loPos = 1
        Do While loPos <= n
            midPos = loPos + runLen - 1
            If midPos > n Then midPos = n
            hiPos = loPos + 2 * runLen - 1
            If hiPos > n Then hiPos = n
            a = loPos
            b = midPos + 1
            k = loPos
            Do While a <= midPos And b <= hiPos
                If VectorLess(rk, order(b), order(a), w) Then
                    tmp(k) = order(b)
                    b = b + 1
                Else
                    tmp(k) = order(a)
                    a = a + 1
                End If
                k = k + 1
            Loop
            Do While a <= midPos
                tmp(k) = order(a)
                a = a + 1
                k = k + 1
            Loop
            Do While b <= hiPos
                tmp(k) = order(b)
                b = b + 1
                k = k + 1
            Loop
            loPos = loPos + 2 * runLen
        Loop
        For k = 1 To n
            order(k) = tmp(k)
        Next k
        runLen = runLen * 2
    Loop
End Sub

' True when group x's rank vector comes strictly before group y's.
Private Function VectorLess(ByRef rk() As Long, ByVal x As Long, ByVal y As Long, ByVal w As Long) As Boolean
    Dim j As Long
    For j = 1 To w
        If rk((x - 1) * w + j) < rk((y - 1) * w + j) Then
            VectorLess = True
            Exit Function
        End If
        If rk((x - 1) * w + j) > rk((y - 1) * w + j) Then Exit Function
    Next j
End Function

' ---- 6: pass 3, the constraints over chosen rows ----------------------

Private Sub GroundClauses(ByVal constraints As Collection, ByRef readsChoice() As Boolean, _
                          ByVal certain As Object, ByVal hMap As Object, ByRef syms As VlaSymbols, _
                          ByVal chosen As Object, ByVal chosenNames As Collection, ByRef gr As OptGround, _
                          ByVal violations As Collection, ByRef refusal As Collection)
    gr.nCl = 0
    gr.nClLit = 0
    ReDim gr.clCheck(1 To 16)
    ReDim gr.clRule(1 To 16)
    ReDim gr.clRow(1 To 16)
    ReDim gr.clStart(1 To 17)
    gr.clStart(1) = 1
    ReDim gr.clLit(1 To 16)
    Dim c As Long
    Dim hasAny As Boolean
    For c = 1 To constraints.Count
        If readsChoice(c) Then hasAny = True
    Next c
    If Not hasAny Then Exit Sub

    ' Every chosen predicate's POSSIBLE rows, as the relation pass 3 reads.
    Dim rel3 As Object
    Set rel3 = CopyRelations(certain)
    Dim pi As Long
    Dim nm As Variant
    For Each nm In chosenNames
        pi = pi + 1
        Dim arity As Long
        arity = CLng(VLA_Runtime.VlaDictGet(chosen, CStr(nm)))
        Dim rel As Collection
        Set rel = VLA_Relation.RelNew(arity)
        Dim a As Long
        For a = 1 To gr.atoms.n
            If gr.atoms.keys((a - 1) * gr.keyWidth + 1) = pi Then
                Dim vals() As Variant
                vals = AtomValues(gr, a, arity, syms)
                VLA_Relation.RelTryAdd rel, vals
            End If
        Next a
        VLA_Runtime.VlaDictSet rel3, CStr(nm), rel
    Next nm

    ' One rule per such constraint - except one with nothing left in its
    ' body, such as a plain (require (in "carbon-frame")): that is ONE
    ' clause, its chosen rows' names all constants, and DATALOG takes no
    ' rule without a body, so it is made here as a single empty row.
    Dim batch As Collection
    Set batch = New Collection
    Dim metaOf() As Collection, ruleOf() As Long
    ReDim metaOf(0 To constraints.Count)
    ReDim ruleOf(0 To constraints.Count)
    ' checkOf(k): the constraint batch rule k grounds, for a refusal's name.
    Dim checkOf As Collection
    Set checkOf = New Collection
    For c = 1 To constraints.Count
        If readsChoice(c) Then
            Dim meta As Collection
            Dim rl As Collection
            Set rl = BuildClauseRule(constraints.Item(c), c, chosen, meta)
            Set metaOf(c) = meta
            If rl.Count > 2 Then
                batch.Add rl
                checkOf.Add c
                ruleOf(c) = batch.Count
            End If
        End If
    Next c
    ' PLANNED (slice 3): a clause is a clause whichever atom is joined
    ' first, so these rules - and only these - are walked smallest atom
    ' first. The clauses come out in the planned order; the search's
    ' answer, and its decisions and dead ends, do not depend on it, since
    ' propagation reaches the same fixpoint in any order. The count goes
    ' on from pass 2's.
    Dim out As Collection
    Set out = VLA_Datalog.DatalogGroundRules(batch, rel3, hMap, syms, planned:=True, _
                                             stepCeiling:=mStepCeiling, _
                                             totalCeiling:=mTotalCeiling, rowsBefore:=gr.rowsLaid)
    Dim over As Variant
    over = VLA_Datalog.DatalogGroundOverflow()
    If Not IsEmpty(over) Then
        Set refusal = SizeRefusal(over, constraints.Item(CLng(checkOf.Item(CLng(over(1))))), False)
        Exit Sub
    End If
    gr.rowsLaid = VLA_Datalog.DatalogGroundRowsLaid()
    If VLA_Datalog.DatalogGroundPeakStep() > gr.peakStep Then gr.peakStep = VLA_Datalog.DatalogGroundPeakStep()

    Dim lits() As Long
    ReDim lits(1 To 16)
    Dim r As Long, nl As Long, rowBase As Long, aId As Long
    Dim taut As Boolean
    For c = 1 To constraints.Count
        If Not readsChoice(c) Then GoTo nextCheck
        Dim mq As Collection
        Set mq = metaOf(c)
        Dim qVars As Collection, qPos As Collection, qNeg As Collection, qCol As Object
        Set qVars = mq.Item(2)
        Set qPos = mq.Item(3)
        Set qNeg = mq.Item(4)
        Dim qHasCons As Boolean
        qHasCons = mq.Item(5)
        Dim qCons As Variant
        NthInto qCons, mq, 6
        Set qCol = mq.Item(7)
        Dim w As Long, nRows As Long
        Dim rows() As Long
        If ruleOf(c) > 0 Then
            Dim entry As Variant
            entry = out.Item(ruleOf(c))
            w = CLng(entry(1))
            nRows = CLng(entry(2))
            rows = entry(3)
        Else
            w = 1
            nRows = 1
            ReDim rows(0 To 1)
        End If
        For r = 0 To nRows - 1
            nl = 0
            rowBase = r * w
            Dim at As Variant
            For Each at In qPos
                aId = InstantiateAtom(at, rows, rowBase, qCol, syms, gr, chosenNames)
                If aId = 0 Then
                    VLA_Messages.RaiseMsg "optimize-internal", "detail", _
                        "grounded a constraint over a row it could not find among the rows that may be chosen"
                End If
                AppendLit lits, nl, -aId
            Next at
            For Each at In qNeg
                aId = InstantiateAtom(at, rows, rowBase, qCol, syms, gr, chosenNames)
                If aId > 0 Then AppendLit lits, nl, aId
            Next at
            If qHasCons Then
                aId = InstantiateAtom(qCons, rows, rowBase, qCol, syms, gr, chosenNames)
                If aId > 0 Then AppendLit lits, nl, aId
            End If
            NormalizeClause lits, nl, taut
            If Not taut Then
                If nl = 0 Then
                    AddViolationRow violations, c, constraints.Item(c), qVars, rows, rowBase, syms
                Else
                    AddClause gr, c, ruleOf(c), r, lits, nl
                End If
            End If
        Next r
nextCheck:
    Next c
End Sub

' One constraint over chosen rows, as the rule pass 3 grounds: its body
' with every NEGATED chosen row taken out (collected, since each becomes
' a positive literal), and a require's certain consequent negated at the
' end exactly as OPTIMIZE.2 writes it; its head the check's number and
' every name the kept body binds. meta holds (check number, those names,
' the positive chosen rows, the negated ones, whether the consequent is
' chosen, the consequent, and each name's column in a grounded row).
Private Function BuildClauseRule(ByVal form As Variant, ByVal c As Long, ByVal chosen As Object, _
                                 ByRef meta As Collection) As Collection
    Dim lst As Collection
    Set lst = form
    Dim isRequire As Boolean
    isRequire = (OptTopHead(form) = "require")
    Dim bodyFrom As Long
    If isRequire Then bodyFrom = 3 Else bodyFrom = 2
    Dim kept As Collection, posAtoms As Collection, negAtoms As Collection
    Set kept = New Collection
    Set posAtoms = New Collection
    Set negAtoms = New Collection
    Dim j As Long
    For j = bodyFrom To lst.Count
        Dim item As Variant, src As Variant
        NthInto item, lst, j
        Dim kind As Long
        kind = BodyReadKind(item, src)
        If kind = 2 And VLA_Runtime.VlaDictHas(chosen, AtomNameOf(src)) Then
            negAtoms.Add src
        Else
            kept.Add item
            If kind = 1 Then
                If VLA_Runtime.VlaDictHas(chosen, AtomNameOf(src)) Then posAtoms.Add src
            End If
        End If
    Next j
    Dim cons As Variant
    Dim consChosen As Boolean
    If isRequire Then
        NthInto cons, lst, 2
        consChosen = VLA_Runtime.VlaDictHas(chosen, AtomNameOf(cons))
        If Not consChosen Then
            ' LAST, for OPTIMIZE.2's reason: CheckRuleSafety walks a body
            ' in written order, so the names are bound by the time it is
            ' reached.
            Dim notForm As Collection
            Set notForm = New Collection
            notForm.Add "not"
            notForm.Add cons
            kept.Add notForm
        End If
    End If

    Dim vars As Collection, seen As Object
    Set vars = New Collection
    Set seen = VLA_Runtime.VlaDictNew()
    Dim b As Variant
    For Each b In kept
        CollectBoundVars b, vars, seen
    Next b
    If isRequire Then CheckConsequent cons, seen
    Dim na As Variant
    For Each na In negAtoms
        CheckChosenUse na, chosen, OptTopHead(form)
        CheckNegatedChosen na, seen
    Next na
    If consChosen Then CheckChosenUse cons, chosen, OptTopHead(form)

    Dim colOf As Object
    Set colOf = VLA_Runtime.VlaDictNew()
    Dim p As Long
    For p = 1 To vars.Count
        VLA_Runtime.VlaDictSet colOf, CStr(vars.Item(p)), 1 + p
    Next p
    Set meta = New Collection
    meta.Add c
    meta.Add vars
    meta.Add posAtoms
    meta.Add negAtoms
    meta.Add consChosen
    If consChosen Then
        meta.Add cons
    Else
        meta.Add Empty
    End If
    meta.Add colOf
    Set BuildClauseRule = BuildRule(OPT_CLAUSE_PREFIX & c, c, vars, Nothing, kept)
End Function

' A negated chosen row, or a require's chosen consequent, is taken out of
' the rule pass 3 grounds - so DATALOG never sees it, and never holds it
' to the arity every other use of the predicate has. Held to it here, in
' DATALOG's own words, since a positive chosen row with the wrong arity
' meets exactly those words from the grounder; and written by position,
' since a keyed (column value) pair has no column to name.
Private Sub CheckChosenUse(ByVal atom As Variant, ByVal chosen As Object, ByVal formHead As String)
    Dim lst As Collection
    Set lst = atom
    Dim nm As String
    nm = AtomNameOf(lst)
    Dim arity As Long
    arity = CLng(VLA_Runtime.VlaDictGet(chosen, nm))
    If lst.Count - 1 <> arity Then
        VLA_Messages.RaiseMsg "datalog-arity-mismatch", "predicate", nm, "a", CStr(arity), "b", CStr(lst.Count - 1)
    End If
    Dim j As Long
    For j = 2 To lst.Count
        If IsObject(lst.Item(j)) Then
            VLA_Messages.RaiseMsg "optimize-form-bad-shape", "form", formHead, _
                "expected", "a chosen row written by position, like (assign S P) - a chosen row has no Table, so it has no column names", _
                "example", "(forbid (assign S P) (not (assign T P)) (next S T))"
        End If
    Next j
End Sub

' A chosen row as its atom number, its names read from one grounded row
' of a pass-3 rule and its constants from the text: 0 when the row is
' not among those that may be chosen at all.
Private Function InstantiateAtom(ByVal atom As Variant, ByRef rows() As Long, ByVal rowBase As Long, _
                                 ByVal colOf As Object, ByRef syms As VlaSymbols, ByRef gr As OptGround, _
                                 ByVal chosenNames As Collection) As Long
    Dim lst As Collection
    Set lst = atom
    Dim kbuf() As Long
    ReDim kbuf(0 To gr.keyWidth)
    kbuf(1) = NameIndex(chosenNames, AtomNameOf(lst))
    Dim j As Long
    For j = 2 To lst.Count
        Dim raw As Variant
        NthInto raw, lst, j
        If OptIsVariable(raw) Then
            kbuf(j) = rows(rowBase + CLng(VLA_Runtime.VlaDictGet(colOf, CStr(raw))))
        Else
            Dim id As Long
            id = VLA_Relation.VlaSymFind(syms, OptAtomText(raw))
            If id = 0 Then Exit Function
            kbuf(j) = id
        End If
    Next j
    InstantiateAtom = TupFind(gr.atoms, kbuf, 0, False)
End Function

Private Sub AppendLit(ByRef lits() As Long, ByRef nl As Long, ByVal lit As Long)
    If nl + 1 > UBound(lits) Then ReDim Preserve lits(1 To 2 * UBound(lits))
    nl = nl + 1
    lits(nl) = lit
End Sub

' Sorts a clause's literals by atom, drops a repeated one, and reports a
' clause holding an atom both ways - true in every world, so no clause.
Private Sub NormalizeClause(ByRef lits() As Long, ByRef nl As Long, ByRef taut As Boolean)
    taut = False
    Dim i As Long, j As Long, x As Long
    For i = 2 To nl
        x = lits(i)
        j = i - 1
        Do While j >= 1
            If Abs(lits(j)) < Abs(x) Or (Abs(lits(j)) = Abs(x) And lits(j) <= x) Then Exit Do
            lits(j + 1) = lits(j)
            j = j - 1
        Loop
        lits(j + 1) = x
    Next i
    Dim n As Long
    n = 0
    For i = 1 To nl
        If n > 0 Then
            If lits(n) = lits(i) Then GoTo nextLit
            If lits(n) = -lits(i) Then
                taut = True
                Exit Sub
            End If
        End If
        n = n + 1
        lits(n) = lits(i)
nextLit:
    Next i
    nl = n
End Sub

Private Sub AddClause(ByRef gr As OptGround, ByVal c As Long, ByVal ruleIx As Long, ByVal rowIx As Long, _
                      ByRef lits() As Long, ByVal nl As Long)
    Dim k As Long
    k = gr.nCl + 1
    Do While k + 1 > UBound(gr.clStart)
        ReDim Preserve gr.clStart(1 To 2 * UBound(gr.clStart))
    Loop
    Do While k > UBound(gr.clCheck)
        ReDim Preserve gr.clCheck(1 To 2 * UBound(gr.clCheck))
        ReDim Preserve gr.clRule(1 To UBound(gr.clCheck))
        ReDim Preserve gr.clRow(1 To UBound(gr.clCheck))
    Loop
    Do While gr.nClLit + nl > UBound(gr.clLit)
        ReDim Preserve gr.clLit(1 To 2 * UBound(gr.clLit))
    Loop
    Dim j As Long
    For j = 1 To nl
        gr.clLit(gr.nClLit + j) = lits(j)
    Next j
    gr.nClLit = gr.nClLit + nl
    gr.clCheck(k) = c
    gr.clRule(k) = ruleIx
    gr.clRow(k) = rowIx
    gr.nCl = k
    gr.clStart(k + 1) = gr.nClLit + 1
End Sub

' A clause left empty: the constraint is broken whatever is chosen. One
' row of the violations table, in OPTIMIZE.2's own shape - its number,
' its text and the bindings that made it hold.
Private Sub AddViolationRow(ByVal violations As Collection, ByVal c As Long, ByVal form As Variant, _
                            ByVal vars As Collection, ByRef rows() As Long, ByVal rowBase As Long, _
                            ByRef syms As VlaSymbols)
    Dim arr() As Variant
    ReDim arr(1 To 1 + vars.Count)
    arr(1) = c
    Dim p As Long
    For p = 1 To vars.Count
        arr(1 + p) = syms.rep(rows(rowBase + 1 + p))
    Next p
    Dim row(1 To 3) As Variant
    row(1) = c
    row(2) = VLA.VlaWriteForm(form)
    row(3) = BindingWords(vars, arr)
    VLA_Relation.RelTryAdd violations, row
End Sub

' ---- 7: single-atom pruning and the counting pre-checks ---------------

' A clause of one literal is not a constraint over choices at all but
' the atom's value, fixed at grounding (OPTIMIZE.0.A: "nobody works while
' on leave" takes the leave rows out of the pool). pruned(a) is -1 for an
' atom no world may choose, +1 for one every world must.
Private Sub PruneSingleAtoms(ByRef gr As OptGround, ByRef pruned() As Long, ByRef nPruned As Long)
    ReDim pruned(0 To gr.atoms.n)
    nPruned = 0
    Dim k As Long, lit As Long
    For k = 1 To gr.nCl
        If gr.clStart(k + 1) - gr.clStart(k) = 1 Then
            lit = gr.clLit(gr.clStart(k))
            If lit < 0 Then
                If pruned(-lit) <> -1 Then nPruned = nPruned + 1
                pruned(-lit) = -1
            ElseIf pruned(lit) = 0 Then
                pruned(lit) = 1
            End If
        End If
    Next k
End Sub

' The first group, in the order they are decided, that needs more rows
' than can still fill it - OptimizePoolShortWords, reached at last.
Private Function PoolShortReason(ByRef gr As OptGround, ByRef pruned() As Long, _
                                 ByVal choices As Collection, ByRef syms As VlaSymbols) As String
    Dim i As Long, k As Long, j As Long, avail As Long
    For i = 1 To gr.nForms
        For k = gr.fCtrFirst(i) To gr.fCtrFirst(i) + gr.fCtrN(i) - 1
            avail = 0
            For j = gr.ctrStart(k) To gr.ctrStart(k + 1) - 1
                If pruned(gr.ctrMem(j)) <> -1 Then avail = avail + 1
            Next j
            If gr.ctrLo(k) > avail Then
                Dim cf As Collection
                Set cf = choices.Item(i)
                Dim gName As String, gKey As String
                If gr.fHasPer(i) Then
                    gName = PerNamesOf(ChoicePer(cf))
                    gKey = GroupKeyWords(gr, i, gr.ctrGroup(k), syms)
                Else
                    gName = "(" & OptTopHead(cf) & " ...)"
                    gKey = "over its whole pool"
                End If
                PoolShortReason = OptimizePoolShortWords(gName, gKey, gr.ctrLo(k), _
                    AtomNameOf(ChoicePoolAtom(cf)), avail)
                Exit Function
            End If
        Next k
    Next i
End Function

' The pigeonhole, OPTIMIZE.2's other two comparisons: a DEMAND family
' whose groups share no atom needs the sum of its lower bounds, all
' distinct atoms; a CAPPING family over the same chosen rows, every
' group capped at one literal number, whose groups cover every atom the
' demand could still use, holds at most (its groups that touch them) x
' (the cap). The demand past the capacity is "no schedule", and it is a
' proof - the pigeonhole principle, which clause learning cannot make
' short (OPTIMIZE.0's oracle, three minutes on optimize-roster-loose).
Private Function CapShortReason(ByRef gr As OptGround, ByRef pruned() As Long, _
                                ByVal choices As Collection) As String
    Dim nA As Long
    nA = gr.atoms.n
    Dim inA() As Long, inB() As Long
    Dim fa As Long, fb As Long, k As Long, j As Long, a As Long
    Dim demand As Double, cap As Long, groupsB As Long
    Dim disjoint As Boolean, touches As Boolean, covered As Boolean
    Dim cfA As Collection, cfB As Collection
    For fa = 1 To gr.nForms
        If gr.fCtrN(fa) = 0 Then GoTo nextA
        Set cfA = choices.Item(fa)
        ReDim inA(0 To nA)
        demand = 0
        disjoint = True
        For k = gr.fCtrFirst(fa) To gr.fCtrFirst(fa) + gr.fCtrN(fa) - 1
            demand = demand + gr.ctrLo(k)
            For j = gr.ctrStart(k) To gr.ctrStart(k + 1) - 1
                a = gr.ctrMem(j)
                If inA(a) <> 0 Then disjoint = False
                inA(a) = k
            Next j
        Next k
        If Not disjoint Or demand <= 0 Or demand > 2147483647# Then GoTo nextA
        For fb = 1 To gr.nForms
            If fb = fa Then GoTo nextB
            If gr.fCtrN(fb) = 0 Then GoTo nextB
            Set cfB = choices.Item(fb)
            If AtomNameOf(ChoiceChosenAtom(cfB)) <> AtomNameOf(ChoiceChosenAtom(cfA)) Then GoTo nextB
            If Not UniformLiteralCap(cfB, cap) Then GoTo nextB
            ReDim inB(0 To nA)
            groupsB = 0
            For k = gr.fCtrFirst(fb) To gr.fCtrFirst(fb) + gr.fCtrN(fb) - 1
                touches = False
                For j = gr.ctrStart(k) To gr.ctrStart(k + 1) - 1
                    a = gr.ctrMem(j)
                    inB(a) = 1
                    If inA(a) <> 0 And pruned(a) <> -1 Then touches = True
                Next j
                If touches Then groupsB = groupsB + 1
            Next k
            covered = True
            For a = 1 To nA
                If inA(a) <> 0 And pruned(a) <> -1 And inB(a) = 0 Then
                    covered = False
                    Exit For
                End If
            Next a
            If Not covered Then GoTo nextB
            If demand > CDbl(groupsB) * cap Then
                If ExactlyOneEach(cfA) Then
                    CapShortReason = OptimizeCapacityShortWords(PerNamesOf(ChoicePer(cfA)), gr.fCtrN(fa), _
                        GroupFamilyName(cfB), groupsB, cap)
                Else
                    CapShortReason = OptimizeCapShortWords(AtomNameOf(ChoiceChosenAtom(cfA)), CLng(demand), _
                        GroupFamilyName(cfB), groupsB, cap)
                End If
                Exit Function
            End If
nextB:
        Next fb
nextA:
    Next fa
End Function

' ---- 8: the search's own problem ---------------------------------------

' Every clause, and every counter that bounds anything - an "at least
' none, at most any" counter is no constraint - tagged with its own
' number here, so a contradiction's reasons come back as rows of gr.
Private Sub BuildSearchProblem(ByRef gr As OptGround, ByRef prob As OptSearchProblem)
    VLA_OptimizeSearch.OptProblemInit prob, gr.atoms.n
    Dim buf() As Long
    ReDim buf(1 To 16)
    Dim k As Long, j As Long, n As Long
    For k = 1 To gr.nCl
        n = gr.clStart(k + 1) - gr.clStart(k)
        If n > UBound(buf) Then ReDim buf(1 To 2 * n)
        For j = 1 To n
            buf(j) = gr.clLit(gr.clStart(k) + j - 1)
        Next j
        VLA_OptimizeSearch.OptProblemAddClause prob, buf, n, k
    Next k
    For k = 1 To gr.nCtr
        If gr.ctrLo(k) > 0 Or gr.ctrHi(k) <> VLA_OptimizeSearch.OPT_SEARCH_NO_MOST Then
            n = gr.ctrStart(k + 1) - gr.ctrStart(k)
            If n > UBound(buf) Then ReDim buf(1 To 2 * n)
            For j = 1 To n
                buf(j) = gr.ctrMem(gr.ctrStart(k) + j - 1)
            Next j
            VLA_OptimizeSearch.OptProblemAddCounter prob, buf, n, gr.ctrLo(k), gr.ctrHi(k), k
        End If
    Next k
End Sub

' ---- 9: the words -------------------------------------------------------

' "proven best", and why a program that chooses may say it: with nothing
' minimized or maximized every schedule that breaks no rule is as good
' as another, and this one is the FIRST of them - OPTIMIZE.1's reserved
' sentence says the program makes no choices, which is true only of one
' that makes none, so a choice program keeps the prefix and says its own
' reason after it. "Tables and facts" since slice 3: slice 2's live pass
' showed "your Tables" on a program whose rows were all (fact ...) forms.
Private Function ChoiceBestWords(ByVal decisions As Long, ByVal conflicts As Long) As String
    ChoiceBestWords = "proven best: every rule holds, and nothing is being minimized or maximized, so no schedule is better than this one - it is the first that breaks no rule when the rows are decided in the order your Tables and facts list them (" & _
        decisions & " decision" & PluralS(decisions) & ", " & conflicts & " dead end" & PluralS(conflicts) & ")."
End Function

Private Function ChoiceNoneWords(ByVal decisions As Long, ByVal conflicts As Long) As String
    ChoiceNoneWords = OptimizeStatusWords(VLA_OPTIMIZE_NO_SCHEDULE) & _
        ": every way of making the choices was tried (" & decisions & " decision" & PluralS(decisions) & _
        ", " & conflicts & " dead end" & PluralS(conflicts) & ")."
End Function

' A contradiction reached before the first decision, and the rules it was
' walked back to: each check by number, as the violations table writes
' it, then each choice's group that took part, in the order they are
' decided - capped as the broken-check list is.
Private Function ChoiceRootWords(ByRef res As OptSearchResult, ByRef gr As OptGround, _
                                 ByVal constraints As Collection, ByVal choices As Collection, _
                                 ByRef syms As VlaSymbols) As String
    Dim seenC() As Long
    ReDim seenC(0 To constraints.Count)
    Dim j As Long, c As Long, k As Long
    For j = 1 To res.nWhyClauses
        seenC(gr.clCheck(res.whyClauses(j))) = 1
    Next j
    Dim s As String
    Dim named As Long, more As Long
    For c = 1 To constraints.Count
        If seenC(c) <> 0 Then
            If named >= OPT_STATUS_CHECK_CAP Then
                more = more + 1
            Else
                If named > 0 Then s = s & "; "
                s = s & "check " & c & ", " & VLA.VlaWriteForm(constraints.Item(c))
                named = named + 1
            End If
        End If
    Next c
    For j = 1 To res.nWhyCounters
        k = res.whyCounters(j)
        If named >= OPT_STATUS_CHECK_CAP Then
            more = more + 1
        Else
            If named > 0 Then s = s & "; "
            s = s & VLA.VlaWriteForm(choices.Item(gr.ctrForm(k)))
            If gr.fHasPer(gr.ctrForm(k)) Then
                s = s & " for " & GroupKeyWords(gr, gr.ctrForm(k), gr.ctrGroup(k), syms)
            End If
            named = named + 1
        End If
    Next j
    If more > 0 Then s = s & "; and " & more & " more"
    ChoiceRootWords = OptimizeStatusWords(VLA_OPTIMIZE_NO_SCHEDULE) & _
        ": these cannot all hold, before anything is chosen - " & s & "."
End Function

Private Function ChoiceBudgetWords(ByVal effortWords As String, ByVal budget As Long, _
                                   ByVal decisions As Long, ByVal conflicts As Long) As String
    ChoiceBudgetWords = OptimizeStatusWords(VLA_OPTIMIZE_NONE_IN_BUDGET) & ": " & effortWords & _
        " allows " & budget & " unit" & PluralS(budget) & " of work - a decision or a dead end is one each - and all of them went on " & _
        decisions & " decision" & PluralS(decisions) & " and " & conflicts & " dead end" & PluralS(conflicts) & _
        ". More effort may find one: (effort thorough), or a larger number."
End Function

Private Function ChoiceGuardWords(ByVal seconds As Double, ByVal work As Long, ByVal budget As Long, _
                                  ByVal effortWords As String) As String
    ChoiceGuardWords = OptimizeStatusWords(VLA_OPTIMIZE_NONE_IN_BUDGET) & _
        ": the search was stopped after " & VLA_Relation.InvariantNumberText(Round(seconds, 1)) & _
        " seconds by the guard that keeps a formula from holding Excel, having done " & work & _
        " of the " & budget & " units of work " & effortWords & " allows. Unlike the effort, where this stops depends on how fast the machine is."
End Function

' ---- 5a: the size refusal ------------------------------------------------

' The refusal a ceiling passed in pass 2 or 3 becomes: the two-item result
' OptimizeRun memoizes - the marker, then Array(which refusal, the form as
' the user wrote it, the reason or the running count, the ceiling, and
' whether it was a command's ceiling rather than a formula's). The words
' are made here, once, so the memo holds exactly what is raised.
Private Function SizeRefusal(ByVal over As Variant, ByVal form As Variant, ByVal isChoice As Boolean) As Collection
    Dim rec As Collection
    Set rec = New Collection
    rec.Add OPT_REFUSAL_MARK
    If CLng(over(0)) = VLA_Datalog.DATALOG_GROUND_OVER_TOTAL Then
        rec.Add Array(OPT_REFUSE_TOTAL, VLA.VlaWriteForm(form), OptCountText(CDbl(over(2))), _
                      OptCountText(mTotalCeiling), mForCommand)
    ElseIf isChoice Then
        rec.Add Array(OPT_REFUSE_CHOICE_STEP, VLA.VlaWriteForm(form), SizeWhyWords(over, False), _
                      OptCountText(mStepCeiling), mForCommand)
    Else
        rec.Add Array(OPT_REFUSE_RULE_STEP, VLA.VlaWriteForm(form), SizeWhyWords(over, True), _
                      OptCountText(mStepCeiling), mForCommand)
    End If
    Set SizeRefusal = rec
End Function

Private Function IsSizeRefusal(ByVal result As Collection) As Boolean
    If result.Count <> 2 Then Exit Function
    If IsObject(result.Item(1)) Then Exit Function
    IsSizeRefusal = (CStr(result.Item(1)) = OPT_REFUSAL_MARK)
End Function

' Slice 5: a formula's size refusal points to the command, which lays out
' more; a command's says "a command" and has nowhere further to point.
Private Sub RaiseSizeRefusal(ByVal result As Collection)
    Dim a As Variant
    a = result.Item(2)
    Dim who As String, thenWords As String
    If CBool(a(4)) Then
        who = "a command"
    Else
        who = "a formula"
        thenWords = CommandPointerWords()
    End If
    Select Case CLng(a(0))
    Case OPT_REFUSE_CHOICE_STEP
        VLA_Messages.RaiseMsg "optimize-choice-too-large", "form", a(1), "why", a(2), "who", who, _
            "ceiling", a(3), "then", thenWords
    Case OPT_REFUSE_RULE_STEP
        VLA_Messages.RaiseMsg "optimize-rule-too-large", "form", a(1), "why", a(2), "who", who, _
            "ceiling", a(3), "then", thenWords
    Case Else
        VLA_Messages.RaiseMsg "optimize-too-large", "form", a(1), "rows", a(2), "who", who, _
            "ceiling", a(3), "then", thenWords
    End Select
End Sub

' Where a formula's size refusal points, and how to get there.
Private Function CommandPointerWords() As String
    CommandPointerWords = " To run it as a command, which lays out up to " & _
        OptCountText(OPT_COMMAND_STEP_ROWS) & " rows in one step and has no limit in all, " & _
        "select this cell and choose Frazaro > Logic Engines > Optimize Selected Cell."
End Function

' The one step that passed the step ceiling, in words ending on the rows
' it would have made: an atom's own rows, rows that share no name paired
' every one with every one, a join on the names they share, or - for a
' rule DATALOG's own evaluator answered - the rule's rows. planned says
' the join was already the smallest part first, which is worth saying:
' rewriting the rule in another order would not help.
Private Function SizeWhyWords(ByVal over As Variant, ByVal planned As Boolean) As String
    Dim stepRows As String, atomRows As String, before As String, pred As String
    stepRows = OptCountText(CDbl(over(3)))
    atomRows = OptCountText(CDbl(over(5)))
    If CDbl(over(4)) = 1 Then
        before = "the one row before it"
    Else
        before = "the " & OptCountText(CDbl(over(4))) & " before it"
    End If
    pred = CStr(over(9))
    If CBool(over(8)) Then
        SizeWhyWords = "its rows come to " & stepRows
    ElseIf CBool(over(7)) Then
        SizeWhyWords = "'" & pred & "' alone holds " & stepRows & " rows"
    ElseIf CBool(over(6)) Then
        SizeWhyWords = "the " & atomRows & " rows of '" & pred & "' share no name with " & before & _
            ", so every one pairs with every one, making " & stepRows & " rows"
    Else
        SizeWhyWords = "the " & atomRows & " rows of '" & pred & "' meet " & before & _
            " on the names they share and make " & stepRows & " rows"
        If planned Then SizeWhyWords = SizeWhyWords & ", even joined smallest part first"
    End If
End Function

' A count as a user reads it: whole, with a comma between each three
' digits - the same on every machine, which Format$'s grouping is not.
Private Function OptCountText(ByVal n As Double) As String
    Dim digits As String
    digits = Format$(n, "0")
    Dim outp As String
    Dim i As Long, taken As Long
    For i = Len(digits) To 1 Step -1
        If taken > 0 And taken Mod 3 = 0 Then outp = "," & outp
        outp = Mid$(digits, i, 1) & outp
        taken = taken + 1
    Next i
    OptCountText = outp
End Function

' A group's own values of its names, "S = s3, N = 2", as the violations
' table writes bindings.
Private Function GroupKeyWords(ByRef gr As OptGround, ByVal i As Long, ByVal g As Long, _
                               ByRef syms As VlaSymbols) As String
    Dim gv As Collection
    Set gv = gr.fGVars(i)
    If gv.Count = 0 Then
        GroupKeyWords = "(no names)"
        Exit Function
    End If
    Dim rows() As Long
    rows = gr.fGRows(i)
    Dim w As Long
    w = gr.fGWidth(i)
    Dim s As String
    Dim p As Long
    For p = 1 To gv.Count
        If p > 1 Then s = s & ", "
        s = s & CStr(gv.Item(p)) & " = " & SafeValueText(syms.rep(rows((g - 1) * w + 1 + p)))
    Next p
    GroupKeyWords = s
End Function

' ---- small helpers --------------------------------------------------------

' An atom's values, as its relation holds them: each argument's symbol's
' first-seen value, which is the pool's own cell.
Private Function AtomValues(ByRef gr As OptGround, ByVal a As Long, ByVal arity As Long, _
                            ByRef syms As VlaSymbols) As Variant()
    Dim t() As Variant
    ReDim t(1 To arity)
    Dim j As Long
    For j = 1 To arity
        t(j) = syms.rep(gr.atoms.keys((a - 1) * gr.keyWidth + 1 + j))
    Next j
    AtomValues = t
End Function

' A chosen predicate's column names: its first choice form's own chosen
' row, each argument's text - (assign S P) gives S and P.
Private Function ChosenHeadNames(ByVal atom As Collection) As Variant
    Dim names() As Variant
    ReDim names(1 To atom.Count - 1)
    Dim j As Long
    For j = 2 To atom.Count
        Dim raw As Variant
        NthInto raw, atom, j
        names(j - 1) = OptAtomText(raw)
    Next j
    ChosenHeadNames = names
End Function

Private Function NameIndex(ByVal names As Collection, ByVal nm As String) As Long
    Dim i As Long
    For i = 1 To names.Count
        If CStr(names.Item(i)) = nm Then
            NameIndex = i
            Exit Function
        End If
    Next i
End Function

' A name's 1-based place among a group's names, matched as DATALOG
' matches a variable - without regard to case.
Private Function VarPosition(ByVal names As Collection, ByVal nm As String) As Long
    Dim i As Long
    For i = 1 To names.Count
        If StrComp(CStr(names.Item(i)), nm, vbTextCompare) = 0 Then
            VarPosition = i
            Exit Function
        End If
    Next i
End Function

' A count read from a Table: a whole number of zero or more, as a number
' or as the text of one. False for anything else.
Private Function WholeCountOf(ByRef syms As VlaSymbols, ByVal id As Long, ByRef n As Long) As Boolean
    Dim v As Variant
    v = syms.rep(id)
    Select Case VarType(v)
    Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, vbByte
        Dim d As Double
        d = CDbl(v)
        If d < 0 Or d > 2147483647# Then Exit Function
        If d <> Int(d) Then Exit Function
        n = CLng(d)
        WholeCountOf = True
    Case vbString
        Dim s As String
        s = CStr(v)
        If Left$(s, 1) = Chr$(34) Then s = Mid$(s, 2)
        If Not IsWholeNumberText(s) Then Exit Function
        If Len(s) > 10 Then Exit Function
        If CDbl(s) > 2147483647# Then Exit Function
        n = CLng(s)
        WholeCountOf = True
    End Select
End Function

' A value as text for a sentence, never raising on an error value.
Private Function SafeValueText(ByVal v As Variant) As String
    If IsError(v) Then
        SafeValueText = "an error value"
    ElseIf IsObject(v) Then
        SafeValueText = "an object"
    Else
        SafeValueText = CStr(v)
    End If
End Function

' ---- the tuple index ------------------------------------------------------

Private Sub TupInit(ByRef t As OptTupleIndex, ByVal wid As Long)
    t.wid = wid
    t.n = 0
    ReDim t.keys(0 To 16 * wid + 1)
    t.mask = 63
    ReDim t.slot(0 To t.mask)
End Sub

' The number of the tuple src(off + 1 .. off + wid): found, or added when
' addIt and absent; 0 when absent and not added. Hashed a step per id by
' TupHashStep (VLA_Datalog's IntHashStep says why): this table is probed in
' line, and the (h * 33) Xor id it used at slice 2 packed the atoms'
' (predicate, item, slot) keys into a run of neighbouring slots - 49,500
' atoms took 1.5 billion probes, 48.5 s, found by slice 3's timing.
Private Function TupFind(ByRef t As OptTupleIndex, ByRef src() As Long, ByVal off As Long, _
                         ByVal addIt As Boolean) As Long
    Dim h As Long, i As Long, s As Long, k As Long, kb As Long
    Dim same As Boolean
    h = 5381
    For i = 1 To t.wid
        h = TupHashStep(h, src(off + i))
    Next i
    s = h And t.mask
    Do
        k = t.slot(s)
        If k = 0 Then Exit Do
        kb = (k - 1) * t.wid
        same = True
        For i = 1 To t.wid
            If t.keys(kb + i) <> src(off + i) Then
                same = False
                Exit For
            End If
        Next i
        If same Then
            TupFind = k
            Exit Function
        End If
        s = (s + 1) And t.mask
    Loop
    If Not addIt Then Exit Function
    t.n = t.n + 1
    Do While t.n * t.wid + 1 > UBound(t.keys)
        ReDim Preserve t.keys(0 To 2 * UBound(t.keys) + 1)
    Loop
    kb = (t.n - 1) * t.wid
    For i = 1 To t.wid
        t.keys(kb + i) = src(off + i)
    Next i
    t.slot(s) = t.n
    TupFind = t.n
    If 2 * t.n > t.mask Then TupRehash t
End Function

Private Sub TupRehash(ByRef t As OptTupleIndex)
    t.mask = 2 * t.mask + 1
    ReDim t.slot(0 To t.mask)
    Dim k As Long, i As Long, h As Long, s As Long, kb As Long
    For k = 1 To t.n
        kb = (k - 1) * t.wid
        h = 5381
        For i = 1 To t.wid
            h = TupHashStep(h, t.keys(kb + i))
        Next i
        s = h And t.mask
        Do While t.slot(s) <> 0
            s = (s + 1) And t.mask
        Loop
        t.slot(s) = k
    Next k
End Sub

' One id folded into a tuple's hash - VLA_Datalog's IntHashStep, the same
' four lines, kept here rather than made Public so no worksheet function
' list gains it. See there for what it is and why.
Private Function TupHashStep(ByVal h As Long, ByVal id As Long) As Long
    h = (((h + id) And &H1FFFFF) * 1021) And &H1FFFFF
    h = h Xor (h \ 2048)
    h = (h * 1019) And &H1FFFFF
    TupHashStep = h Xor (h \ 1024)
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
