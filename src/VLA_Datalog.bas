Attribute VB_Name = "VLA_Datalog"
Option Explicit
Public Const VLA_DATALOG_VERSION As String = "DATALOG.5"
' DATALOG.5: named-column atoms - a rule-BODY atom's arguments may now
' be (header var-or-const) pairs, keyed by the target predicate's own
' column name, instead of bare positional tokens: (staffing (name X)
' (salary S)) rather than (staffing X _ S). Found by a phrasebook
' lookahead, not engine work: the first English sentence anyone would
' say about a table ("...salary is over 80000") binds a column NAME,
' and DATALOG's own purely positional engine had no way to turn that
' into "column 2" without reading the live header at translate time.
'
' DATALOG stays positional INTERNALLY - a keyed atom desugars, once, at
' parse time, to exactly the positional atom it stands for
' (DesugarBodyAtomForm, called from ParseProgram's own rule-body loop,
' BEFORE ParseAtom ever sees the form). Every column the atom does NOT
' key becomes a fresh anonymous variable (AnonymousColumnVarName) -
' deterministic, derived only from the atom's own WRITTEN position in
' its rule's body (the same `bi` index ParseProgram's own body loop
' already walks by) and the column's own 1-based header position, never
' a counter or random name (this project's own standing veto on
' gensym-style identifier invention, BETA_ROADMAP1.md's LANGUAGE +
' MACHINE section). Deliberately UNIQUE PER ATOM OCCURRENCE, not per
' column alone: two different keyed atoms both omitting the same
' column - even against the same table, even within the same rule body
' - must never be silently joined on that column's own value, so each
' gets its own fresh name. Because desugaring happens before ParseAtom
' ever runs, CheckRuleSafety and EvalRuleBody see a perfectly ordinary
' positional atom either way - neither needed a single new case for
' this (verified, not assumed): a keyed atom's own anonymous columns
' bind (or, under `not`, must already be bound) exactly the way any
' other fresh positional variable already does.
'
' Only a rule BODY position may be keyed - rule heads stay positional
' by definition (a derived predicate has no header of its own to key
' against): ParseAtom is called directly on a rule's head, never routed
' through DesugarBodyAtomForm, so a keyed-looking head argument simply
' falls through to ParseAtom's own pre-existing "no compound-term
' nesting" refusal (datalog-compound-term), unchanged. A keyed atom
' against a predicate with no header of its own - a (fact ...) block, a
' plain named range, or any rule-derived predicate, none of which
' headerMap (below) ever has an entry for - is refused the same way
' (datalog-keyed-atom-needs-header). Keyed and positional arguments may
' not mix in one atom (datalog-atom-mixed-keying); an unknown header
' name is refused listing the names that actually exist
' (datalog-unknown-column - DATALOG's own id, not a reuse of SQL's
' identically-shaped sql-unknown-column, SD-9's discipline); the same
' column keyed twice in one atom is refused (datalog-keyed-column-
' repeated). A keyed pair whose own VALUE position is itself a nested
' compound term - (salary (foo X)) - is deliberately left untouched by
' DesugarBodyAtomForm and falls through to ParseAtom's own existing
' compound-term check instead of inventing a second way to say the same
' thing - the one place this item is carving a genuinely new, narrow
' exception into that check's "no list argument, ever" rule (a keyed
' pair is a 2-element list ONLY when both its own elements are bare
' tokens), so this fallthrough is what keeps that exception exactly as
' narrow as intended.
'
' headerMap (predicate's own folded name -> VLA_Relation.RangeColumnNames'
' own (folded, original) pair Collection) is built by DATALOG()/
' DatalogRun BEFORE ParseProgram ever runs - the call ordering
' baseRelations already needed - and threaded through as ParseProgram's
' own new parameter; a predicate with no entry in headerMap is exactly
' a predicate with no header (the three cases above, uniformly).
' Consumes SQL.1's own header-name substrate directly (RangeColumnNames)
' rather than growing a twin - the earlier-simplifies-later move this
' section keeps making, one more time; VLA_Relation.bas's own DATALOG.5
' header note has the RangeColumnNames narrowing this item needed
' (keyed names resolve within a DATALOG.3 column slice only) and the
' RelFromRange has-header flag this item's own pure test needed.
'
' DATALOG.4: comparison and arithmetic built-ins in rule bodies - the
' MVP's own second and last stated limit closed (the first,
' column-scoped table arguments, was DATALOG.3 above). Two new rule-
' body shapes, both reusing ParseAtom's own atom-record shape directly
' rather than inventing a second one: a comparison filter, (> X 50000)
' with the full </<=/>/>=/=/<> set, parsed AS an atom whose own
' "predicate" is the operator symbol and whose two "args" are its
' operands (ParseAtom neither knows nor cares that the head symbol
' isn't a relation name); and an arithmetic binding form, (let Z
' (+ X Y)) - +, -, *, /, exactly one brand-new result variable, the
' same one-new-variable discipline datalog-aggregate-result-reused
' already enforces for count/sum, reused here as its own
' datalog-let-result-reused. BI_CMP/BI_LET join the BI_* family
' (MakeBodyItem/BodyItemKind); CheckRuleSafety extends the SAME
' written-order walk `not` already established - every operand variable
' must already be bound by an earlier positive body atom
' (datalog-builtin-unsafe-variable), because a comparison can never
' itself enumerate "every value greater than 50000." Unlike not/count/
' sum, neither kind touches stratification at all: BodyItemNeedsFullRelation
' stays False for both (redefined from "kind <> BI_POS" to an explicit
' BI_NOT/BI_COUNT/BI_SUM check now that a SECOND non-full-relation kind
' exists - the two concepts "doesn't need the full relation" and "is a
' valid semi-naive delta position" used to coincide when BI_POS was the
' only such kind, and no longer do), and ComputeStrata skips a BI_CMP/
' BI_LET body item entirely when building its own predicate-dependency
' graph (its own atom's "predicate" is an operator symbol, never a real
' relation - counting it would risk colliding with an actual predicate
' someone chose to name "+", however unlikely). Both evaluate INLINE,
' per row, in EvalRuleBody (ResolveOperand/BuiltinOperandIsNumeric are
' the new small helpers) - a parser case plus one evaluator branch, per
' this item's own roadmap words.
'
' The numeric-vs-text policy is DATALOG's own, not SQL's borrowed
' unchanged: DATALOG's own rule-text constants and fact-block values
' are ALWAYS plain VBA Strings (never a real Excel cell type, unlike a
' table-sourced value read via RelFromRange's own Value2 path), so
' BuiltinOperandIsNumeric treats a numeric-LOOKING string as numeric too
' (VLA_Relation.IsInvariantNumericString) - more lenient than SQL's own
' WHERE clause on purpose (VLA_Relation.bas's own DATALOG.4 header note
' has the full reasoning), or a filter like (> Salary 50000) could never
' fire on a hand-written fact at all. The actual mechanical compare/
' compute (VLA_Relation.CompareValues/ComputeArithmetic) is genuinely
' shared with VLA_Sql.bas now, not a twin grown alongside it - this
' section's own SQL.1 sequencing note, honored.
'
' DATALOG.3: column-scoped table arguments - the first of the MVP's two
' remaining stated limits to actually get built (the second, avoiding a
' full re-parse/re-fixpoint per recalc, stays deliberately unbuilt
' pending a real profiling signal - BETA_ROADMAP2.md's own scoping note).
' TableArgName gains an early refusal (datalog-table-noncontiguous-
' columns) for a multi-area (Ctrl-selected) Range, before .ListObject
' even gets a chance to resolve unpredictably against one. The actual
' narrowing lives in VLA_Relation.SourceToArray (its own DATALOG.2
' history note has the full mechanism) - TableArgName's own job, naming
' the predicate off the table's ListObject.Name, is UNCHANGED by how many
' of that table's own columns were actually passed.
'
' DATALOG.1: stratified negation (`not`), the item's own first Stretch
' sub-item, built second (after the MVP, ahead of grouped aggregation on
' purpose - the roadmap's own sequencing note: aggregation reuses this
' same stratification rather than inventing its own). A rule body
' position can now be (not (pred args...)) as well as a plain atom -
' every body item is parsed into a (kind, atom, resultVar) record
' (MakeBodyItem/BodyItemAtom/BodyItemKind) rather than a bare atom.
' CheckRuleSafety walks the body in WRITTEN order and refuses
' (datalog-negation-unsafe-variable) any (not ...) whose own variables
' aren't already bound by an earlier POSITIVE body atom in the same rule
' - negation can never itself introduce a variable, since DATALOG has no
' way to enumerate "every value not in this relation." ComputeStrata
' builds the whole program's predicate dependency graph (an edge per
' body item, tagged strict/non-strict) and solves each predicate's
' stratum via Bellman-Ford-shaped relaxation (non-strict edge:
' stratum(head) >= stratum(body); strict edge: strictly greater) - a
' predicate depending on itself through a strict edge, directly or
' through a cycle, is refused at parse time
' (datalog-negation-not-stratifiable), never detected live.
' RunStratifiedFixpoint replaces the old flat RunFixpoint (renamed
' RunFixpointForRules, unchanged internally) as DatalogRun's own entry
' point: rules are bucketed by their head predicate's stratum and run to
' a full fixpoint one stratum at a time, ascending, sharing one relations
' dict throughout - by construction, every predicate a stratum's own
' rules negate or aggregate already finished, in full, in a strictly
' earlier stratum. EvalRuleBody's negation branch (FilterOutMatching,
' VLA_Relation.RelContainsTuple) is a plain anti-join: for each row
' already accumulated, substitute its own bound values into the negated
' atom and keep the row only if that fully-resolved tuple is ABSENT from
' the target predicate's relation. A program using no `not`/`count`/
' `sum` anywhere folds every rule into stratum 0, so
' RunStratifiedFixpoint degenerates back to exactly the old single
' RunFixpointForRules call - DATALOG.0's own behavior, unchanged.
'
' DATALOG.2: grouped aggregation (`count`/`sum`), sequenced deliberately
' AFTER negation, reusing its stratification rather than inventing a
' second copy (the roadmap's own words: "an aggregate over predicate P
' is exactly as unsafe as a negated use of P if P is not fully computed
' first"). A rule body position may now also be (count Var (pred ...))
' or (sum Var (pred ...)) - BI_COUNT/BI_SUM, the same BodyItemKind
' family `not` (BI_NOT) already lives in, and BodyItemNeedsFullRelation
' is true for all three: they share one code path everywhere it matters
' (ComputeStrata's own strict-edge tagging, RunFixpointForRules' own
' delta-position skip, EvalRuleBody's own "read relations() in full,
' never deltas()" branch). Unlike `not`, an aggregate's own atom
' variables may be a MIX of already-bound (the GROUP-BY key) and not-yet-
' bound (the thing counted, or - for `sum`, exactly one such variable,
' checked at parse time, datalog-sum-needs-one-value-variable - the
' value added up); its own resultVar is a brand-new bound variable
' (refused if it collides with one already in scope,
' datalog-aggregate-result-reused), the same way a normal atom's unbound
' argument introduces one. ComputeAggregateGroups makes one pass over the
' fully-computed target relation, grouping by exactly the atom's already-
' bound positions (BoundPositionPairs/KeyFromPositions - the same
' ordering used again, against colOf's ACCUM columns instead of the
' atom's own positions, so the two sides' group keys compare equal for
' equal values); ApplyAggregate then extends every accum row with its own
' group's result (a group with no matching tuples is a real, meaningful
' zero, unlike negation's anti-join - an aggregate never drops a row).
'
' =====================================================================
'  VLA_Datalog - the QUERY AND LOGIC section's own sequencing note,
'  built: the smallest of the four engines (no compound-term
'  unification, no backtracking, no cut), proving the shared
'  VLA_Relation.bas substrate cheaply instead of discovering its
'  problems mid-way through SQL's full-relational tier or PROLOG's
'  ceiling. MVP scope: function-free Horn clauses, facts sourced from a
'  range (VLA_Relation.RelFromRange) or a (fact ...) block, semi-naive
'  bottom-up fixpoint evaluation, results as a spilled array. Stratified
'  negation and grouped aggregation (DATALOG.1/.2, above) are both built
'  - the item's own two Stretch sub-items, done.
'
'  Program text is VLA's own S-expression syntax (VLA.VlaReadForms -
'  "the one architectural gift this item gets for free from already
'  being a Lisp", the item's own words), never Prolog's dot-syntax:
'
'      (fact (predicate const1 const2 ...))
'      (rule (head-pred V1 V2 ...) (body-pred1 ...) (body-pred2 ...))
'      (rule (head-pred V1 V2 ...) (body-pred1 ...) (not (excluded V1)))
'      (rule (head-pred X N) (body-pred1 X ...) (count N (source X Y)))
'      (rule (head-pred X S) (body-pred1 X ...) (sum S (source X Amount)))
'      (rule (head-pred X) (body-pred1 X Salary) (> Salary 50000))
'      (rule (head-pred X Total) (body-pred1 X A) (body-pred2 X B) (let Total (+ A B)))
'      (rule (rich X) (staffing (name X) (salary S)) (> S 80000))
'      (query predicate-name)
'
'  A rule BODY atom's own arguments may be keyed by column name instead
'  of position, DATALOG.5 - (staffing (name X) (salary S)) rather than
'  (staffing X _ S) - resolved against the target predicate's own
'  header row (a live Table argument's own column names; never a
'  (fact ...) block, a plain named range, or a rule-derived predicate,
'  none of which have one) and desugared to the equivalent positional
'  atom before this engine ever evaluates it - see this module's own
'  header above for the full mechanism.
'
'  A rule body position may be a plain atom, (not (predicate ...)) -
'  stratified negation, DATALOG.1 - (count Var (predicate ...))/
'  (sum Var (predicate ...)) - grouped aggregation, DATALOG.2 - a
'  comparison filter, (op X Y) where op is one of </<=/>/>=/=/<>, or an
'  arithmetic binding, (let Z (op X Y)) where op is one of +/-/*// -
'  both DATALOG.4. `not`'s own variables, and a comparison/arithmetic
'  built-in's own operand variables, must ALL already be bound by an
'  earlier (written-order) POSITIVE body atom in the same rule (a
'  comparison/arithmetic built-in can no more enumerate "every value
'  greater than 50000" than negation can enumerate "every value not in
'  this relation"); `count`/`sum` may mix already-bound (group-by) and
'  not-yet-bound (aggregated-over) variables, but `sum` requires
'  EXACTLY ONE not-yet-bound variable (the value), and `let` binds
'  exactly one brand-new result variable of its own (refused if it
'  collides with one already in scope, the same discipline count/sum's
'  own result variable already follows). Either way, the predicate
'  being negated or aggregated must be fully computable before the rule
'  using it runs (checked at parse time; a cycle running back through
'  one is refused, never silently mis-evaluated) - a comparison/
'  arithmetic built-in touches no predicate at all, so this concern
'  never applies to either of DATALOG.4's own two new shapes.
'
'  A bare token whose first letter is A-Z is a VARIABLE (the classical
'  Prolog/Datalog convention); anything else - lowercase words,
'  numbers, quoted strings - is a CONSTANT, compared case-sensitively
'  (VLA_Relation.bas's own tuple-identity discipline). Exactly one
'  (query ...) is required per program - naming the relation DATALOG
'  should return - rather than guessing "the last rule's head", which
'  would silently change meaning if rules were reordered. Renaming a
'  table-derived predicate (whose name is always the Range's own
'  ListObject/defined name - see DATALOG's own header below) needs no
'  new syntax: (rule (reports_to X Y) (Employees X Y)) already
'  expresses it with the one mechanism this engine has.
'
'  Every predicate has ONE fixed arity everywhere it appears (facts,
'  rule heads, rule bodies, and any table argument sharing its name) -
'  checked and refused, never truncated or padded. Every rule's head
'  variables must all also appear in its body (the standard Datalog
'  safety/range-restriction condition) - refused otherwise, since an
'  unsafe rule's result would depend on something this engine has no
'  way to enumerate.
'
'  LAYER:     1 (VLA_Relation)
'  MAY CALL:  VLA (VlaReadForms only), VLA_Identity (Fold), VLA_Runtime
'             (VlaDictNew/Set/Get/Has/Keys - predicate-NAME lookups
'             only, never tuple data; SD-8's identifier fold is correct
'             for a predicate/variable name and wrong for a fact's own
'             values, which is why VLA_Relation.bas keeps its own,
'             separate, case-sensitive index instead), VLA_Messages
'             (RaiseMsg), VLA_Relation (the whole join/storage
'             substrate, plus DATALOG.4's own shared CompareValues/
'             ComputeArithmetic/IsInvariantNumericString).
'  Duplicates VLA.bas's own Private IsList/Nth rather than widen its
'  public surface for a second, unrelated reader - VLA_Interpreter.bas's
'  own precedent (its header note) for the identical situation.
'  SHIPS:     add-in only, as a real Excel worksheet function
'             (=DATALOG(...)) the moment the add-in loads - no
'             VBProject trust, no export, no emitted code. SEC.1 Tier-0
'             by construction: reads only the ranges/text it is given,
'             writes nothing, calls nothing outside itself. Stated
'             limit, same as deflambda's own: an exported standalone
'             workbook using =DATALOG() needs the add-in present.
'  PAYS INTO: QUERY AND LOGIC's DATALOG item; the shared-substrate bet
'             VLA_Relation.bas exists to prove, ahead of SQL/PROLOG.
' =====================================================================

' Body-item kind tags (DATALOG.1/.2/.4) - which of the six shapes a rule
' body position is: a plain atom, (not ...), (count ...), (sum ...),
' a comparison (op X Y), or (let Var (op X Y)). Declared here, at the
' top of the module, rather than next to MakeBodyItem further down, so
' every one of their many call sites below sees an ordinary top-of-
' module Const, not a declaration sandwiched between two Function
' bodies (a real, unexplained compile error this codebase hit once
' already from exactly that sandwiching, per this module's own
' defensive habit since).
Private Const BI_POS As Long = 0
Private Const BI_NOT As Long = 1
Private Const BI_COUNT As Long = 2
Private Const BI_SUM As Long = 3
Private Const BI_CMP As Long = 4
Private Const BI_LET As Long = 5

' ---- minimal S-expression accessors, duplicated from VLA.bas's own
'      Private IsList/Nth (cross-module Private calls do not exist in
'      VBA) - kept small on purpose; this module needs no more of the
'      reader's surface than "is this a list" and "fetch element i".
Private Function IsList(ByVal v As Variant) As Boolean
    IsList = IsObject(v)
End Function

Private Function Nth(ByVal lst As Collection, ByVal i As Long) As Variant
    If IsObject(lst.Item(i)) Then
        Set Nth = lst.Item(i)
    Else
        Nth = lst.Item(i)
    End If
End Function

' The one safe way to STORE Nth's result into a variable, as opposed to
' passing it straight through as another call's argument (Nth(lst, i)
' used directly as a function argument, e.g. ParseAtom(Nth(lst, 2), ...),
' is fine on its own - VBA's argument-binding preserves an object
' reference without touching it).
'
' A bare `dest = Nth(lst, i)` is a live-caught VBA trap: whenever
' lst.Item(i) is itself a list (a compound term; or - DATALOG.1/.2's own
' addition - a rule-BODY item, ALWAYS a list), Nth returns an object, and
' a plain Let assignment of an object into a Variant invokes that
' object's own DEFAULT MEMBER rather than storing the reference -
' Collection's default member is Item, which needs an index argument, so
' `dest = Nth(...)` on an object result raises runtime error 450
' ("Wrong number of arguments...") rather than storing the list.
' Live-caught exactly this way the instant DATALOG.1/.2's own body-parsing
' loop reached its first rule with a body (every body position is a list,
' so this fired unconditionally, on the very first non-fact-only
' DatalogRun call) - a bare fact-only program never reaches ParseAtom's
' own analogous spots with a compound argument outside the one existing
' test built to exercise that path, which is why DATALOG.0's own pure
' tests never surfaced this same trap already latent in Nth's OTHER
' callers (predRaw/raw/h/qRaw, fixed alongside this one, all the same
' shape, all now routed through NthInto). NthInto sidesteps this the
' same way Nth's own body already does internally (IsObject-gated
' Set-vs-Let), just one call further out, ByRef into the caller's own
' variable.
Private Sub NthInto(ByRef dest As Variant, ByVal lst As Collection, ByVal i As Long)
    If IsObject(lst.Item(i)) Then
        Set dest = lst.Item(i)
    Else
        dest = lst.Item(i)
    End If
End Sub

' Same trap, same fix, for copying one already-fetched Variant into
' another (e.g. "start atomForm off as a copy of bodyForm, then maybe
' override it") rather than fetching fresh from a list - a bare
' `dest = src` invokes src's own default member whenever src currently
' holds an object; passing src as this Sub's own ByVal Variant argument
' preserves the reference correctly (argument-binding, unlike a bare
' Let-assignment statement, does not invoke a default member), so the
' IsObject check inside sees the truth and picks Set vs plain assignment
' correctly.
Private Sub CopyVariant(ByRef dest As Variant, ByVal src As Variant)
    If IsObject(src) Then
        Set dest = src
    Else
        dest = src
    End If
End Sub

' A raw reader token -> its atom text. Refuses if v is itself a list
' (a compound term - the one restriction that makes every DATALOG()
' query provably terminate). A leading Chr$(34) marks a quoted string
' literal (VLA.bas's own reader convention); its content, already
' escape-resolved by the tokenizer, is the constant's text. A bare
' token's text is used as-is, whether it turns out to be a variable or
' a constant (IsVariableAtom decides that separately, from the SAME
' raw value, before this stripping happens).
Private Function AtomText(ByVal raw As Variant, ByVal ctx As String) As String
    If IsObject(raw) Then VLA_Messages.RaiseMsg "datalog-compound-term", "context", ctx
    Dim s As String
    s = CStr(raw)
    If Left$(s, 1) = Chr$(34) Then
        AtomText = Mid$(s, 2)
    Else
        AtomText = s
    End If
End Function

' A quoted string ("Bob") is always a constant, even if its content
' starts with a capital letter - only a BARE token starting A-Z is a
' variable.
Private Function IsVariableAtom(ByVal raw As Variant) As Boolean
    If IsObject(raw) Then Exit Function
    Dim s As String
    s = CStr(raw)
    If Left$(s, 1) = Chr$(34) Then Exit Function
    If Len(s) = 0 Then Exit Function
    Dim c As Integer
    c = AscW(Left$(s, 1))
    IsVariableAtom = (c >= 65 And c <= 90)
End Function

' ---- atom records: a plain Collection, this codebase's own ad-hoc-
'      record convention. Item(1) = folded predicate name (String).
'      Item(2) = args, a Collection of 2-item Collections, each
'      (isVariable As Boolean, text As String).
Private Function ParseAtom(ByVal form As Variant, ByVal ctx As String) As Collection
    If Not IsList(form) Then VLA_Messages.RaiseMsg "datalog-atom-not-a-list", "context", ctx
    Dim lst As Collection
    Set lst = form
    If lst.Count < 1 Then VLA_Messages.RaiseMsg "datalog-atom-empty", "context", ctx
    Dim predRaw As Variant
    NthInto predRaw, lst, 1
    If IsObject(predRaw) Then VLA_Messages.RaiseMsg "datalog-predicate-name-not-symbol", "context", ctx
    Dim predName As String
    predName = VLA_Identity.Fold(AtomText(predRaw, ctx))
    Dim args As New Collection
    Dim i As Long
    For i = 2 To lst.Count
        Dim raw As Variant
        NthInto raw, lst, i
        ' `As New` inside a loop is a real VBA trap: the auto-
        ' instantiate rule only fires when the variable is currently
        ' Nothing, which after the first iteration it never is again -
        ' every argument past the first would silently APPEND onto the
        ' SAME rec object rather than start fresh, and args.Add rec
        ' would then store that one shared, ever-growing object for
        ' every argument. Since ArgIsVar/ArgText only ever read
        ' Item(1)/Item(2), the symptom is invisible whenever an atom's
        ' arguments happen to share the same text (edge(X,X)) and
        ' silently wrong whenever they don't (grandparent(X,Z) would
        ' read "X" back for its own second argument) - live-caught via
        ' exactly that shape. Explicit Set each iteration is the fix.
        Dim rec As Collection
        Set rec = New Collection
        rec.Add IsVariableAtom(raw)
        rec.Add AtomText(raw, ctx & " (predicate '" & predName & "')")
        args.Add rec
    Next i
    ' MVP scope boundary, not an incidental restriction: a zero-
    ' argument (propositional) predicate would need an arity-0
    ' Relation, and constructing one's own arrays safely runs into the
    ' identical inverted-bounds ReDim hazard VLA_Relation.bas's own
    ' RelUnit/SourceToArray headers document - refused here, at the
    ' one chokepoint every atom (fact, rule head, rule body) already
    ' passes through, rather than partially hardened deep inside the
    ' shared substrate for a shape nothing in this engine's own corpus
    ' or tests ever needs.
    If args.Count = 0 Then VLA_Messages.RaiseMsg "datalog-predicate-needs-argument", "predicate", predName
    Dim atom As New Collection
    atom.Add predName
    atom.Add args
    Set ParseAtom = atom
End Function

Private Function AtomPred(ByVal atom As Collection) As String
    AtomPred = atom.Item(1)
End Function

Private Function AtomArity(ByVal atom As Collection) As Long
    AtomArity = atom.Item(2).Count
End Function

Private Function AtomArgAt(ByVal atom As Collection, ByVal i As Long) As Collection
    Set AtomArgAt = atom.Item(2).Item(i)
End Function

Private Function ArgIsVar(ByVal a As Collection) As Boolean
    ArgIsVar = a.Item(1)
End Function

Private Function ArgText(ByVal a As Collection) As String
    ArgText = a.Item(2)
End Function

' ---- body-item records: a rule body is a Collection of these, not of
'      bare atoms, now that a body position can be a negated atom
'      (`not`, DATALOG.1), a grouped aggregate (`count`/`sum`,
'      DATALOG.2), a comparison filter, or an arithmetic binding
'      (`let`, both DATALOG.4) as well as a plain one. Item(1) = kind
'      (Long, one of the BI_* constants declared near the top of this
'      module). Item(2) = the atom itself (Collection, ParseAtom's own
'      shape) - for BI_COUNT/BI_SUM this is the SOURCE predicate being
'      aggregated, not the head; for BI_CMP/BI_LET the operator symbol
'      (">", "+", ...) stands in as ParseAtom's own "predicate", its two
'      operands as the "args" - never a real predicate at all. Item(3) =
'      resultVar (String) - an aggregate's or a `let`'s own brand-new
'      result variable name; "" and unused for BI_POS/BI_NOT/BI_CMP.
Private Function MakeBodyItem(ByVal kind As Long, ByVal atom As Collection, ByVal resultVar As String) As Collection
    Dim item As New Collection
    item.Add kind
    item.Add atom
    item.Add resultVar
    Set MakeBodyItem = item
End Function

Private Function BodyItemKind(ByVal item As Collection) As Long
    BodyItemKind = item.Item(1)
End Function

Private Function BodyItemAtom(ByVal item As Collection) As Collection
    Set BodyItemAtom = item.Item(2)
End Function

Private Function BodyItemResultVar(ByVal item As Collection) As String
    BodyItemResultVar = item.Item(3)
End Function

' Both negation and grouped aggregation need the SAME guarantee the
' roadmap's own DATALOG.2 sequencing note names: the predicate being
' negated or aggregated must be fully, finally computed before the rule
' using it can run - never mid-fixpoint, never read through a delta.
' BI_POS is the only kind that may ever be a semi-naive delta position
' or contribute a non-strict (>=) stratum edge; BI_NOT/BI_COUNT/BI_SUM
' force a strictly higher stratum (ComputeStrata) and always read
' relations() in full (EvalRuleBody, RunFixpointForRules's own round
' loop).
'
' DATALOG.4: BI_CMP/BI_LET also return False here (a comparison or
' arithmetic built-in touches no predicate at all, so "needs the full
' relation" is meaningless for either) - but, UNLIKE BI_POS, neither is
' a valid semi-naive DELTA position either, since neither one reads any
' predicate's relations()/deltas() dict by name in the first place.
' Before DATALOG.4, "does not need the full relation" and "is a valid
' delta position" were the SAME set ({BI_POS}), so callers could use
' this one function for both questions; now that a body item can be
' BI_CMP/BI_LET too, they are two DIFFERENT sets, and a caller asking
' "is this a valid delta position" must check "kind = BI_POS" directly
' rather than "Not BodyItemNeedsFullRelation(kind)" - RunFixpointForRules'
' own round loop does exactly that now.
Private Function BodyItemNeedsFullRelation(ByVal kind As Long) As Boolean
    BodyItemNeedsFullRelation = (kind = BI_NOT Or kind = BI_COUNT Or kind = BI_SUM)
End Function

Private Function TopHead(ByVal form As Variant) As String
    If Not IsList(form) Then VLA_Messages.RaiseMsg "datalog-top-form-not-a-list"
    Dim lst As Collection
    Set lst = form
    If lst.Count < 1 Then VLA_Messages.RaiseMsg "datalog-top-form-empty"
    Dim h As Variant
    NthInto h, lst, 1
    If IsObject(h) Then VLA_Messages.RaiseMsg "datalog-top-form-bad-head"
    TopHead = VLA_Identity.Fold(CStr(h))
End Function

' Every head variable must also appear in the rule's own body - the
' standard Datalog safety/range-restriction condition. Refused rather
' than silently producing an unbound (or infinite) result. Walks
' bodyItems in WRITTEN order, deliberately: a `not`'s own variables (and,
' as of DATALOG.2, a `count`/`sum`'s own GROUP-BY variables - whichever
' of its atom's variables are already bound) must already be bound by an
' EARLIER positive atom in the same body (not merely "somewhere" in the
' body) - the same left-to-right order EvalRuleBody itself evaluates in,
' so a safety pass that accepts a program is a guarantee the evaluator
' can actually satisfy, not just a necessary condition checked against
' the wrong order.
'
' Per body-item kind:
'   BI_POS   - every variable is bound (or re-bound, harmlessly) here.
'   BI_NOT   - every variable must ALREADY be bound; contributes none.
'   BI_COUNT - any mix is fine (bound = group key, unbound = counted,
'              wildcard-style); contributes exactly its own resultVar.
'   BI_SUM   - same as BI_COUNT, but EXACTLY ONE of its atom's variables
'              may be unbound (the value being summed) - refused
'              otherwise, since summing zero or several undesignated
'              columns has no defined meaning here.
' Both aggregate kinds also refuse a resultVar name that collides with
' an already-bound variable (datalog-aggregate-result-reused) - a
' silent double-binding would be a confusing way to fail later instead.
Private Sub CheckRuleSafety(ByVal headAtom As Collection, ByVal bodyItems As Collection)
    Dim boundVars As Object
    Set boundVars = VLA_Runtime.VlaDictNew()
    Dim bi As Variant
    For Each bi In bodyItems
        Dim item As Collection
        Set item = bi
        Dim atom As Collection
        Set atom = BodyItemAtom(item)
        Dim kind As Long
        kind = BodyItemKind(item)
        Select Case kind
        Case BI_NOT
            Dim ni As Long
            For ni = 1 To AtomArity(atom)
                Dim na As Collection
                Set na = AtomArgAt(atom, ni)
                If ArgIsVar(na) Then
                    If Not VLA_Runtime.VlaDictHas(boundVars, ArgText(na)) Then
                        VLA_Messages.RaiseMsg "datalog-negation-unsafe-variable", "predicate", AtomPred(atom), "var", ArgText(na)
                    End If
                End If
            Next ni
        Case BI_COUNT, BI_SUM
            Dim newVarsSeen As Object
            Set newVarsSeen = VLA_Runtime.VlaDictNew()
            Dim gi As Long
            For gi = 1 To AtomArity(atom)
                Dim ga As Collection
                Set ga = AtomArgAt(atom, gi)
                If ArgIsVar(ga) Then
                    If Not VLA_Runtime.VlaDictHas(boundVars, ArgText(ga)) Then
                        VLA_Runtime.VlaDictSet newVarsSeen, ArgText(ga), True
                    End If
                End If
            Next gi
            If kind = BI_SUM Then
                If VLA_Runtime.VlaDictKeys(newVarsSeen).Count <> 1 Then
                    VLA_Messages.RaiseMsg "datalog-sum-needs-one-value-variable", "predicate", AtomPred(atom), "count", VLA_Runtime.VlaDictKeys(newVarsSeen).Count
                End If
            End If
            Dim resultVar As String
            resultVar = BodyItemResultVar(item)
            If VLA_Runtime.VlaDictHas(boundVars, resultVar) Then
                VLA_Messages.RaiseMsg "datalog-aggregate-result-reused", "var", resultVar
            End If
            VLA_Runtime.VlaDictSet boundVars, resultVar, True
        Case BI_CMP, BI_LET
            ' Same rule as BI_NOT, applied to a comparison/arithmetic
            ' built-in's own two operands: every variable must ALREADY
            ' be bound (a comparison/arithmetic built-in contributes no
            ' new binding of its own from its operands - only BI_LET's
            ' own separate resultVar, below, ever does). Written-order,
            ' the same walk the whole Sub already performs.
            Dim bvi As Long
            For bvi = 1 To AtomArity(atom)
                Dim bva As Collection
                Set bva = AtomArgAt(atom, bvi)
                If ArgIsVar(bva) Then
                    If Not VLA_Runtime.VlaDictHas(boundVars, ArgText(bva)) Then
                        VLA_Messages.RaiseMsg "datalog-builtin-unsafe-variable", "operator", AtomPred(atom), "var", ArgText(bva)
                    End If
                End If
            Next bvi
            If kind = BI_LET Then
                ' (let Z (op X Y)) binds its own brand-new Z, the same
                ' "refused if it collides with one already in scope"
                ' discipline BI_COUNT/BI_SUM's own resultVar already
                ' follows, above.
                Dim letResultVar As String
                letResultVar = BodyItemResultVar(item)
                If VLA_Runtime.VlaDictHas(boundVars, letResultVar) Then
                    VLA_Messages.RaiseMsg "datalog-let-result-reused", "var", letResultVar
                End If
                VLA_Runtime.VlaDictSet boundVars, letResultVar, True
            End If
        Case Else ' BI_POS
            Dim ai As Long
            For ai = 1 To AtomArity(atom)
                Dim a As Collection
                Set a = AtomArgAt(atom, ai)
                If ArgIsVar(a) Then VLA_Runtime.VlaDictSet boundVars, ArgText(a), True
            Next ai
        End Select
    Next bi
    Dim hi As Long
    For hi = 1 To AtomArity(headAtom)
        Dim ha As Collection
        Set ha = AtomArgAt(headAtom, hi)
        If ArgIsVar(ha) Then
            If Not VLA_Runtime.VlaDictHas(boundVars, ArgText(ha)) Then
                VLA_Messages.RaiseMsg "datalog-unsafe-head-variable", "predicate", AtomPred(headAtom), "var", ArgText(ha)
            End If
        End If
    Next hi
End Sub

' DATALOG.5's own deterministic, gensym-free anonymous variable name for
' a keyed atom's own UNMENTIONED column - derived ONLY from the atom's
' own written position within its rule's body (bodyItemIndex, exactly
' the `bi` loop index ParseProgram's own body-item loop already uses -
' no separate counter invented) and the column's own 1-based position
' in the target predicate's header - fully reproducible from the rule
' text alone, never a counter or random name (this project's own
' standing veto on gensym-style identifier invention). Deliberately
' UNIQUE PER ATOM OCCURRENCE, not per column alone: two different keyed
' atoms both omitting the SAME column - even against the same table,
' even within the same rule body - must never be silently joined on
' that column's own value; each occurrence gets its own fresh name, the
' same "fresh" an ordinary unbound positional variable already means.
' Starts with an uppercase letter (VLA_Datalog's own IsVariableAtom
' rule) so it re-enters ParseAtom/CheckRuleSafety/EvalRuleBody as an
' ordinary VARIABLE, not a constant, indistinguishable from one a user
' typed by hand.
Private Function AnonymousColumnVarName(ByVal bodyItemIndex As Long, ByVal colIndex As Long) As String
    AnonymousColumnVarName = "VlaAnonB" & bodyItemIndex & "C" & colIndex
End Function

' Comma-joined ORIGINAL (unfolded) column names, for datalog-unknown-
' column's own teaching-refusal wording ("its own columns are: ...") -
' mirrors sql-unknown-column's shape (BETA_ROADMAP2.md's own DATALOG.5
' words) without literally reusing its id (SD-9: ids are never reused).
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

' DATALOG.5's own desugaring: recognizes a rule-body atom whose
' arguments are ALL (header var-or-const) pairs - a 2-element list,
' first element a bare token (the column name), second element itself
' never a list (the value: a variable or a constant) - and rewrites the
' whole atom into the exact positional form it stands for, resolved
' against headerMap's own predicate-name -> VLA_Relation.RangeColumnNames'
' own (folded, original) pair Collection. Returns atomForm UNTOUCHED
' whenever it is already fully positional (no argument is a list) - the
' overwhelmingly common case, and every prior DATALOG.x item's own only
' case.
'
' Classification, once, over every argument:
'   - no argument is a list             -> unchanged (ordinary atom).
'   - EVERY argument is a well-formed
'     (header value) pair               -> a genuine keyed atom, resolved.
'   - a mix of list and non-list args   -> datalog-atom-mixed-keying.
'   - every argument is a LIST, but at least one fails the strict
'     2-element (bare, bare) pair shape - left UNTOUCHED, deliberately:
'     this function invents no second way to say "not a valid form
'     here" - ParseAtom's own pre-existing "no compound-term nesting"
'     check (datalog-compound-term) already refuses it correctly the
'     moment it walks that argument itself, including the one case this
'     item's own roadmap specifically calls out: a keyed pair whose OWN
'     value position is itself a nested compound term, e.g.
'     (salary (foo X)) - NOT silently accepted through this new opening.
'
' Called only for a body item whose atom stands for a real predicate
' reference (BI_POS/BI_NOT/BI_COUNT/BI_SUM's own source atom) - never
' BI_CMP/BI_LET's synthetic operator form (ParseProgram's own call site
' gates this), so an operand list like (> (a b) X) still reaches
' ParseAtom untouched and is refused as a compound term, not
' misdiagnosed as "mixed keying".
Private Function DesugarBodyAtomForm(ByVal bodyItemIndex As Long, ByVal atomForm As Variant, ByVal headerMap As Object) As Variant
    Dim result As Variant
    CopyVariant result, atomForm

    If IsList(atomForm) Then
        Dim lst As Collection
        Set lst = atomForm
        If lst.Count >= 2 Then
            Dim anyList As Boolean, anyBare As Boolean, allValidPairs As Boolean
            allValidPairs = True
            Dim i As Long
            For i = 2 To lst.Count
                Dim raw As Variant
                NthInto raw, lst, i
                If IsObject(raw) Then
                    anyList = True
                    Dim pairLst As Collection
                    Set pairLst = raw
                    If pairLst.Count <> 2 Then
                        allValidPairs = False
                    Else
                        Dim headerRaw As Variant, valueRaw As Variant
                        NthInto headerRaw, pairLst, 1
                        NthInto valueRaw, pairLst, 2
                        If IsObject(headerRaw) Or IsObject(valueRaw) Then allValidPairs = False
                    End If
                Else
                    anyBare = True
                End If
            Next i

            If anyList Then
                Dim predRaw As Variant
                NthInto predRaw, lst, 1
                If Not IsObject(predRaw) Then
                    Dim predFolded As String
                    predFolded = VLA_Identity.Fold(CStr(predRaw))
                    If anyBare Then
                        VLA_Messages.RaiseMsg "datalog-atom-mixed-keying", "predicate", predFolded
                    ElseIf allValidPairs Then
                        If Not VLA_Runtime.VlaDictHas(headerMap, predFolded) Then
                            VLA_Messages.RaiseMsg "datalog-keyed-atom-needs-header", "predicate", predFolded
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
                            Dim pRaw As Variant
                            NthInto pRaw, lst, i
                            Dim pLst As Collection
                            Set pLst = pRaw
                            Dim hRaw As Variant, vRaw As Variant
                            NthInto hRaw, pLst, 1
                            NthInto vRaw, pLst, 2
                            Dim hText As String
                            hText = CStr(hRaw)
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
                                VLA_Messages.RaiseMsg "datalog-unknown-column", "predicate", predFolded, "column", hText, "columns", JoinOriginalNames(headerPairs)
                            End If
                            If filled(pos) Then
                                VLA_Messages.RaiseMsg "datalog-keyed-column-repeated", "predicate", predFolded, "column", hText
                            End If
                            filled(pos) = True
                            CopyVariant slots(pos), vRaw
                        Next i

                        Dim outLst As New Collection
                        outLst.Add predRaw
                        For j = 1 To arity
                            If filled(j) Then
                                outLst.Add slots(j)
                            Else
                                outLst.Add AnonymousColumnVarName(bodyItemIndex, j)
                            End If
                        Next j
                        CopyVariant result, outLst
                    End If
                    ' Else: every argument is list-shaped, none of them
                    ' bare, but at least one fails the strict pair shape
                    ' - left as the default passthrough (result already
                    ' set to atomForm, above); ParseAtom's own existing
                    ' compound-term check refuses the malformed argument.
                End If
            End If
        End If
    End If

    ' Same Set-vs-Let care as Nth/NthInto's own bodies (this module's
    ' own top-of-file precedent) - result may hold an object (the
    ' unchanged passthrough, or a freshly-built desugared Collection)
    ' or a non-object (only if atomForm itself was never a list to
    ' begin with); a bare `DesugarBodyAtomForm = result` on an object
    ' value would invoke ITS OWN default member instead of returning
    ' the reference.
    If IsObject(result) Then
        Set DesugarBodyAtomForm = result
    Else
        DesugarBodyAtomForm = result
    End If
End Function

' Parses rulesText into facts (a Collection of ground atoms), rules (a
' Collection of 2-item Collections: headAtom, bodyAtoms), and the one
' required query predicate name.
Private Sub ParseProgram(ByVal rulesText As String, ByRef facts As Collection, ByRef rules As Collection, ByRef queryName As String, ByRef headless As Boolean, ByVal headerMap As Object)
    Set facts = New Collection
    Set rules = New Collection
    queryName = ""
    headless = False
    Dim forms As Collection
    Set forms = VLA.VlaReadForms(rulesText)
    Dim queryCount As Long
    Dim f As Variant
    For Each f In forms
        Dim lst As Collection
        Set lst = f
        Dim head As String
        head = TopHead(f)
        Select Case head
        Case "fact"
            If lst.Count <> 2 Then VLA_Messages.RaiseMsg "datalog-fact-bad-shape"
            Dim factAtom As Collection
            Set factAtom = ParseAtom(Nth(lst, 2), "a fact")
            Dim ai As Long
            For ai = 1 To AtomArity(factAtom)
                If ArgIsVar(AtomArgAt(factAtom, ai)) Then
                    VLA_Messages.RaiseMsg "datalog-fact-has-variable", "predicate", AtomPred(factAtom), "arg", ArgText(AtomArgAt(factAtom, ai))
                End If
            Next ai
            facts.Add factAtom
        Case "rule"
            If lst.Count < 3 Then VLA_Messages.RaiseMsg "datalog-rule-needs-body"
            Dim headAtom As Collection
            Set headAtom = ParseAtom(Nth(lst, 2), "a rule head")
            ' Same `As New`-in-a-loop trap as ParseAtom's own rec: this
            ' whole Case sits inside the For Each f In forms loop, so a
            ' SECOND (rule ...) form in one program would find
            ' bodyAtoms/ruleRec already non-Nothing from the first
            ' rule and silently append onto (and share) that rule's
            ' own objects instead of starting fresh - live-caught: a
            ' two-rule program derived nothing at all, because the
            ' second rule's own body had been merged onto the first's.
            Dim bodyAtoms As Collection
            Set bodyAtoms = New Collection
            Dim bi As Long
            For bi = 3 To lst.Count
                Dim bodyForm As Variant
                NthInto bodyForm, lst, bi
                Dim kind As Long
                kind = BI_POS
                Dim resultVarName As String
                resultVarName = ""
                Dim atomForm As Variant
                CopyVariant atomForm, bodyForm
                ' (not (pred ...)) / (count Var (pred ...)) / (sum Var
                ' (pred ...)) / (let Var (op X Y)) / (op X Y) - detected
                ' the same way TopHead reads a top-level form's own head
                ' symbol (VLA_Identity.Fold, so NOT/Not/not and COUNT/
                ' Count/count etc. all match - a no-op for the operator
                ' symbols below, none of which contain a letter to fold),
                ' but scoped to exactly these wrapper shapes rather than
                ' reusing TopHead itself, which would wrongly demand a
                ' (fact ...)/(rule ...)/(query ...) head here instead. A
                ' plain atom whose own predicate happens to be named
                ' "not"/"count"/"sum"/"let" or one of the six comparison
                ' operators is not reachable here - the same reserved-
                ' word tradeoff fact/rule/query/headless already make at
                ' the top level.
                If IsList(bodyForm) Then
                    Dim bLst As Collection
                    Set bLst = bodyForm
                    If bLst.Count >= 1 Then
                        Dim wrapperRaw As Variant
                        NthInto wrapperRaw, bLst, 1
                        If Not IsObject(wrapperRaw) Then
                            Dim wrapperWord As String
                            wrapperWord = VLA_Identity.Fold(CStr(wrapperRaw))
                            Select Case wrapperWord
                            Case "not"
                                If bLst.Count <> 2 Then VLA_Messages.RaiseMsg "datalog-not-bad-shape"
                                kind = BI_NOT
                                NthInto atomForm, bLst, 2
                            Case "count", "sum"
                                If bLst.Count <> 3 Then VLA_Messages.RaiseMsg "datalog-aggregate-bad-shape", "form", wrapperWord
                                Dim rvRaw As Variant
                                NthInto rvRaw, bLst, 2
                                If IsObject(rvRaw) Then
                                    VLA_Messages.RaiseMsg "datalog-aggregate-result-not-a-variable", "form", wrapperWord
                                End If
                                If Not IsVariableAtom(rvRaw) Then
                                    VLA_Messages.RaiseMsg "datalog-aggregate-result-not-a-variable", "form", wrapperWord
                                End If
                                resultVarName = AtomText(rvRaw, "an aggregate's own result variable")
                                If wrapperWord = "count" Then kind = BI_COUNT Else kind = BI_SUM
                                NthInto atomForm, bLst, 3
                            Case "let"
                                ' (let Z (+ X Y)) - DATALOG.4's own
                                ' arithmetic binding form. Same shape as
                                ' count/sum's own (Var, source-form) pair
                                ' above (a result variable, then a nested
                                ' form to parse as an atom), so the same
                                ' two checks apply; kept as its own Case
                                ' (rather than folded into "count", "sum"
                                ' above) so a malformed (let ...) gets its
                                ' own wording, not "aggregate", which
                                ' would misname the problem.
                                If bLst.Count <> 3 Then VLA_Messages.RaiseMsg "datalog-let-bad-shape"
                                Dim letRvRaw As Variant
                                NthInto letRvRaw, bLst, 2
                                If IsObject(letRvRaw) Then
                                    VLA_Messages.RaiseMsg "datalog-let-result-not-a-variable"
                                End If
                                If Not IsVariableAtom(letRvRaw) Then
                                    VLA_Messages.RaiseMsg "datalog-let-result-not-a-variable"
                                End If
                                resultVarName = AtomText(letRvRaw, "a let's own result variable")
                                kind = BI_LET
                                NthInto atomForm, bLst, 3
                            Case ">", "<", "<=", ">=", "=", "<>"
                                ' (op X Y) - DATALOG.4's own comparison
                                ' filter. atomForm already defaults to a
                                ' COPY of the whole bodyForm (set above,
                                ' before this Select Case ever runs) - a
                                ' comparison IS the whole (op left right)
                                ' form, parsed as an atom whose own
                                ' "predicate" is the operator symbol and
                                ' whose two "args" are its operands, so no
                                ' NthInto override is needed here, unlike
                                ' every other wrapper case above.
                                kind = BI_CMP
                            End Select
                        End If
                    End If
                End If
                If kind <> BI_CMP And kind <> BI_LET Then
                    ' DATALOG.5: a keyed atom is only ever a real
                    ' predicate reference - BI_POS's own atom, BI_NOT's
                    ' negated atom, or BI_COUNT/BI_SUM's own source atom
                    ' - never BI_CMP/BI_LET's synthetic operator form,
                    ' whose two "args" are plain operands, not a real
                    ' predicate's own columns. Desugars BEFORE ParseAtom
                    ' ever sees the form, so a keyed atom's own
                    ' anonymous/resolved columns reach ParseAtom (and
                    ' everything downstream of it - CheckRuleSafety,
                    ' EvalRuleBody) as a perfectly ordinary positional
                    ' atom, needing no new case anywhere else in this
                    ' module.
                    CopyVariant atomForm, DesugarBodyAtomForm(bi, atomForm, headerMap)
                End If
                Dim biAtom As Collection
                Set biAtom = ParseAtom(atomForm, "a rule body")
                If kind = BI_CMP Then
                    ' A COMPARISON is still exactly two operands - all six
                    ' of </<=/>/>=/=/<> are binary and none of PROLOG.17's
                    ' widening touches them. ParseAtom alone only enforces
                    ' "at least one", so the exact count is checked here.
                    If AtomArity(biAtom) <> 2 Then
                        VLA_Messages.RaiseMsg "datalog-builtin-needs-two-operands", "operator", AtomPred(biAtom), "count", AtomArity(biAtom), "expected", "two operands", "example", "(> X 50000) or (+ X Y)"
                    End If
                ElseIf kind = BI_LET Then
                    ' PROLOG.17: the operator set is no longer written
                    ' here. VLA_Relation.ArithOpArity is asked instead -
                    ' the SAME table PROLOG asks and the same one
                    ' ComputeArithmetic's own Case arms are checked
                    ' against - so DATALOG cannot drift into recognizing a
                    ' different set of operators from the substrate that
                    ' computes them. Refused by name, never guessed, the
                    ' identical discipline datalog-compound-term already
                    ' established for this whole engine.
                    '
                    ' ARITY COMES FROM THE OPERATOR, not from the shape:
                    ' this arm used to share the flat "exactly two" test
                    ' above with BI_CMP, which is why (let Z (abs X)) was
                    ' unwritable rather than merely unimplemented.
                    Dim wantArgsLet As Long
                    wantArgsLet = VLA_Relation.ArithOpArity(AtomPred(biAtom))
                    If wantArgsLet = 0 Then
                        VLA_Messages.RaiseMsg "datalog-unknown-arithmetic-operator", "operator", AtomPred(biAtom)
                    End If
                    If AtomArity(biAtom) <> wantArgsLet Then
                        VLA_Messages.RaiseMsg "datalog-builtin-needs-two-operands", "operator", AtomPred(biAtom), "count", AtomArity(biAtom), "expected", ArithArityWordsFor(wantArgsLet), "example", ArithArityExampleFor(wantArgsLet)
                    End If
                End If
                bodyAtoms.Add MakeBodyItem(kind, biAtom, resultVarName)
            Next bi
            CheckRuleSafety headAtom, bodyAtoms
            Dim ruleRec As Collection
            Set ruleRec = New Collection
            ruleRec.Add headAtom
            ruleRec.Add bodyAtoms
            rules.Add ruleRec
        Case "query"
            If lst.Count <> 2 Then VLA_Messages.RaiseMsg "datalog-query-bad-shape"
            Dim qRaw As Variant
            NthInto qRaw, lst, 2
            If IsObject(qRaw) Then VLA_Messages.RaiseMsg "datalog-query-not-a-symbol"
            queryCount = queryCount + 1
            queryName = VLA_Identity.Fold(AtomText(qRaw, "a query"))
        Case "headless"
            ' Opt-in, per-program, not a formula argument: ParamArray
            ' tables() must be DATALOG's own last parameter (a hard VBA
            ' rule), so a literal third argument would sit BETWEEN
            ' rulesText and the table arguments - every table-bearing
            ' formula would need a filler comma just to reach past it.
            ' A rules-text directive costs nothing for anyone not using
            ' it and needs no comma-skipping for anyone who is.
            If lst.Count <> 1 Then VLA_Messages.RaiseMsg "datalog-headless-bad-shape"
            headless = True
        Case Else
            VLA_Messages.RaiseMsg "datalog-unknown-top-form", "head", head
        End Select
    Next f
    If queryCount = 0 Then VLA_Messages.RaiseMsg "datalog-query-missing"
    If queryCount > 1 Then VLA_Messages.RaiseMsg "datalog-query-ambiguous", "count", queryCount
End Sub

Private Sub RecordArity(ByVal predArityDict As Object, ByVal predName As String, ByVal arity As Long)
    If VLA_Runtime.VlaDictHas(predArityDict, predName) Then
        Dim prev As Long
        prev = VLA_Runtime.VlaDictGet(predArityDict, predName)
        If prev <> arity Then
            VLA_Messages.RaiseMsg "datalog-arity-mismatch", "predicate", predName, "a", prev, "b", arity
        End If
    Else
        VLA_Runtime.VlaDictSet predArityDict, predName, arity
    End If
End Sub

Private Function GetOrCreateRelation(ByVal relations As Object, ByVal predName As String, ByVal arity As Long) As Collection
    If VLA_Runtime.VlaDictHas(relations, predName) Then
        Set GetOrCreateRelation = VLA_Runtime.VlaDictGet(relations, predName)
    Else
        Dim rel As Collection
        Set rel = VLA_Relation.RelNew(arity)
        VLA_Runtime.VlaDictSet relations, predName, rel
        Set GetOrCreateRelation = rel
    End If
End Function

' Filters rel down to tuples consistent with atom's own constants and
' within-atom repeated variables (e.g. (edge X X), (reports_to Y "bob")) -
' unification-lite, since no compound terms exist to unify structurally.
' RelJoin itself knows nothing about constants or repeats; this is
' where a single atom's own shape actually gets resolved, before it
' ever reaches the cross-atom join.
Private Function AtomMatches(ByVal atom As Collection, ByRef tup() As Variant) As Boolean
    Dim seen As Object
    Set seen = VLA_Runtime.VlaDictNew()
    Dim i As Long
    For i = 1 To AtomArity(atom)
        Dim a As Collection
        Set a = AtomArgAt(atom, i)
        Dim v As String
        v = CStr(tup(i))
        If ArgIsVar(a) Then
            Dim nm As String
            nm = ArgText(a)
            If VLA_Runtime.VlaDictHas(seen, nm) Then
                If StrComp(CStr(VLA_Runtime.VlaDictGet(seen, nm)), v, vbBinaryCompare) <> 0 Then Exit Function
            Else
                VLA_Runtime.VlaDictSet seen, nm, v
            End If
        Else
            If StrComp(ArgText(a), v, vbBinaryCompare) <> 0 Then Exit Function
        End If
    Next i
    AtomMatches = True
End Function

Private Function FilterAtomRelation(ByVal rel As Collection, ByVal atom As Collection) As Collection
    Dim outp As Collection
    Set outp = VLA_Relation.RelNew(VLA_Relation.RelArity(rel))
    Dim t As Variant, arr() As Variant
    For Each t In VLA_Relation.RelTuples(rel)
        arr = t
        If AtomMatches(atom, arr) Then VLA_Relation.RelTryAdd outp, arr
    Next t
    Set FilterAtomRelation = outp
End Function

' The anti-join half of negation: keeps only accum's own rows whose
' natom instantiation (every argument resolved through colOf's running
' variable->column map, or used as-is if it's a constant - CheckRuleSafety
' already guarantees every variable natom uses is already bound) is
' ABSENT from negRel. negRel Is Nothing means the negated predicate has
' no relation at all yet (never appeared as a fact/table/rule head) -
' vacuously true for every row, same as probing a genuinely empty
' relation, so every row survives unfiltered.
Private Function FilterOutMatching(ByVal accum As Collection, ByVal natom As Collection, _
                                    ByVal colOf As Object, ByVal negRel As Collection) As Collection
    Dim outp As Collection
    Set outp = VLA_Relation.RelNew(VLA_Relation.RelArity(accum))
    Dim nArity As Long
    nArity = AtomArity(natom)
    Dim t As Variant, arr() As Variant
    For Each t In VLA_Relation.RelTuples(accum)
        arr = t
        Dim keep As Boolean
        keep = True
        If Not negRel Is Nothing Then
            Dim probe() As Variant
            ReDim probe(1 To nArity)
            Dim i As Long
            For i = 1 To nArity
                Dim a As Collection
                Set a = AtomArgAt(natom, i)
                If ArgIsVar(a) Then
                    probe(i) = arr(CLng(VLA_Runtime.VlaDictGet(colOf, ArgText(a))))
                Else
                    probe(i) = ArgText(a)
                End If
            Next i
            keep = Not VLA_Relation.RelContainsTuple(negRel, probe)
        End If
        If keep Then VLA_Relation.RelTryAdd outp, arr
    Next t
    Set FilterOutMatching = outp
End Function

' A grouped aggregate's own atom position, in order, whose variable is
' already bound (i.e. a GROUP-BY key) - the shared ordering both
' ComputeAggregateGroups (reading the target relation, where a tuple's
' own values sit at these ATOM positions) and the per-accum-row lookup
' below (where the same values sit at colOf's ACCUM columns) must walk
' identically, so the two sides' own group-key strings compare equal
' for the same underlying values despite coming from different arrays.
Private Sub BoundPositionPairs(ByVal atom As Collection, ByVal colOf As Object, _
                                ByRef atomPositions As Collection, ByRef accumCols As Collection)
    Set atomPositions = New Collection
    Set accumCols = New Collection
    Dim i As Long
    For i = 1 To AtomArity(atom)
        Dim a As Collection
        Set a = AtomArgAt(atom, i)
        If ArgIsVar(a) Then
            If VLA_Runtime.VlaDictHas(colOf, ArgText(a)) Then
                atomPositions.Add i
                accumCols.Add CLng(VLA_Runtime.VlaDictGet(colOf, ArgText(a)))
            End If
        End If
    Next i
End Sub

Private Function KeyFromPositions(ByRef arr() As Variant, ByVal positions As Collection) As String
    Dim r As String
    Dim p As Variant
    For Each p In positions
        r = r & CStr(arr(CLng(p))) & Chr$(31)
    Next p
    KeyFromPositions = r
End Function

' The grouped-aggregate core: one pass over the FULLY-COMPUTED target
' relation (never a delta - BodyItemNeedsFullRelation's own guarantee),
' building groupKeyString -> running total (Long for BI_COUNT, Double
' for BI_SUM), keyed on exactly the atom's own already-bound (group-by)
' argument positions. A tuple is skipped (does not contribute to any
' group) if it is inconsistent with the atom's own UNBOUND-variable
' repeats or CONSTANT arguments - the identical consistency rule
' AtomMatches enforces for a normal atom, restricted here to the
' not-yet-globally-bound positions, since a globally-bound position is
' instead what SELECTS which group a tuple belongs to, not a filter to
' apply before grouping. For BI_SUM, CheckRuleSafety already guarantees
' exactly one such not-yet-bound variable exists (valuePos) - the value
' summed into that group's own running total. targetRel Is Nothing (the
' aggregated predicate has no relation at all, e.g. an EDB predicate no
' fact/table ever populated) returns an empty dict - every group
' defaults to zero at lookup time, below, the same as probing a
' genuinely empty relation.
Private Function ComputeAggregateGroups(ByVal kind As Long, ByVal atom As Collection, _
                                         ByVal colOf As Object, ByVal targetRel As Collection) As Object
    Dim groups As Object
    Set groups = VLA_Runtime.VlaDictNew()
    If targetRel Is Nothing Then
        Set ComputeAggregateGroups = groups
        Exit Function
    End If
    Dim atomPositions As Collection, accumCols As Collection
    BoundPositionPairs atom, colOf, atomPositions, accumCols
    Dim nArgs As Long
    nArgs = AtomArity(atom)
    Dim t As Variant, arr() As Variant
    For Each t In VLA_Relation.RelTuples(targetRel)
        arr = t
        Dim seenLocal As Object
        Set seenLocal = VLA_Runtime.VlaDictNew()
        Dim consistent As Boolean
        consistent = True
        Dim valuePos As Long
        valuePos = 0
        Dim i As Long
        For i = 1 To nArgs
            Dim a As Collection
            Set a = AtomArgAt(atom, i)
            Dim v As String
            v = CStr(arr(i))
            If ArgIsVar(a) Then
                Dim nm As String
                nm = ArgText(a)
                If Not VLA_Runtime.VlaDictHas(colOf, nm) Then
                    If VLA_Runtime.VlaDictHas(seenLocal, nm) Then
                        If StrComp(CStr(VLA_Runtime.VlaDictGet(seenLocal, nm)), v, vbBinaryCompare) <> 0 Then
                            consistent = False
                            Exit For
                        End If
                    Else
                        VLA_Runtime.VlaDictSet seenLocal, nm, v
                        If kind = BI_SUM Then valuePos = i
                    End If
                End If
            Else
                If StrComp(ArgText(a), v, vbBinaryCompare) <> 0 Then
                    consistent = False
                    Exit For
                End If
            End If
        Next i
        If consistent Then
            Dim gk As String
            gk = KeyFromPositions(arr, atomPositions)
            If kind = BI_COUNT Then
                If VLA_Runtime.VlaDictHas(groups, gk) Then
                    VLA_Runtime.VlaDictSet groups, gk, CLng(VLA_Runtime.VlaDictGet(groups, gk)) + 1
                Else
                    VLA_Runtime.VlaDictSet groups, gk, 1&
                End If
            Else
                Dim addend As Double
                addend = CDbl(arr(valuePos))
                If VLA_Runtime.VlaDictHas(groups, gk) Then
                    VLA_Runtime.VlaDictSet groups, gk, CDbl(VLA_Runtime.VlaDictGet(groups, gk)) + addend
                Else
                    VLA_Runtime.VlaDictSet groups, gk, addend
                End If
            End If
        End If
    Next t
    Set ComputeAggregateGroups = groups
End Function

' Extends accum with ONE new column (the aggregate's own resultVar): for
' every existing row, its own group-key (accumCols side of
' BoundPositionPairs) is looked up in groups; a group with no matching
' target-relation tuples at all (lookup miss) is a real, meaningful
' zero - not a row that gets dropped the way FilterOutMatching would
' drop one, so every accum row survives unconditionally here.
Private Function ApplyAggregate(ByVal kind As Long, ByVal accum As Collection, ByVal atom As Collection, _
                                 ByVal colOf As Object, ByVal groups As Object) As Collection
    Dim atomPositions As Collection, accumCols As Collection
    BoundPositionPairs atom, colOf, atomPositions, accumCols
    Dim inArity As Long
    inArity = VLA_Relation.RelArity(accum)
    Dim outp As Collection
    Set outp = VLA_Relation.RelNew(inArity + 1)
    Dim t As Variant, arr() As Variant
    For Each t In VLA_Relation.RelTuples(accum)
        arr = t
        Dim gk As String
        gk = KeyFromPositions(arr, accumCols)
        Dim aggVal As Variant
        If VLA_Runtime.VlaDictHas(groups, gk) Then
            aggVal = VLA_Runtime.VlaDictGet(groups, gk)
        ElseIf kind = BI_COUNT Then
            aggVal = 0&
        Else
            aggVal = 0#
        End If
        Dim nt() As Variant
        ReDim nt(1 To inArity + 1)
        Dim i As Long
        For i = 1 To inArity
            nt(i) = arr(i)
        Next i
        nt(inArity + 1) = aggVal
        VLA_Relation.RelTryAdd outp, nt
    Next t
    Set ApplyAggregate = outp
End Function

Private Function ColOfCount(ByVal colOf As Object) As Long
    ColOfCount = VLA_Runtime.VlaDictKeys(colOf).Count
End Function

' SQL.3: this module's own former CollToLongArray hoisted into
' VLA_Relation.bas (Public there now) the moment VLA_Sql.bas needed the
' identical Collection-of-Longs -> RelJoin's own leftCols()/rightCols()
' conversion for its own equi-condition join keys - the same second-
' consumer move as TableArgResolve/RangeColumnNames/CompareValues
' before it. Every call site below now reads VLA_Relation.CollToLongArray
' directly; behavior unchanged. CollToStringArray stays here, Private -
' DATALOG's own newVarNames are variable NAMES, not a generically
' useful shape SQL ever needs.
Private Function CollToStringArray(ByVal c As Collection) As Variant
    If c.Count = 0 Then
        CollToStringArray = Array()
    Else
        Dim r() As Variant
        ReDim r(1 To c.Count)
        Dim i As Long
        For i = 1 To c.Count: r(i) = CStr(c.Item(i)): Next i
        CollToStringArray = r
    End If
End Function

' Which of atom's own variable positions are already bound (join
' columns on both sides, via colOf's running variable->column map) vs
' brand new (kept, appended to the accumulator). A second occurrence
' of the same variable WITHIN this atom carries no new information
' (AtomMatches already enforced its equality with the first
' occurrence) and is skipped entirely, on both sides.
Private Sub BuildJoinPlan(ByVal colOf As Object, ByVal atom As Collection, _
                          ByRef leftCols As Variant, ByRef rightCols As Variant, _
                          ByRef newVarCols As Variant, ByRef newVarNames As Variant)
    Dim lc As New Collection, rc As New Collection
    Dim nc As New Collection, nn As New Collection
    Dim seenInAtom As Object
    Set seenInAtom = VLA_Runtime.VlaDictNew()
    Dim i As Long
    For i = 1 To AtomArity(atom)
        Dim a As Collection
        Set a = AtomArgAt(atom, i)
        If ArgIsVar(a) Then
            Dim nm As String
            nm = ArgText(a)
            If VLA_Runtime.VlaDictHas(seenInAtom, nm) Then
                ' repeated within this atom - nothing new, nothing to join on again
            ElseIf VLA_Runtime.VlaDictHas(colOf, nm) Then
                lc.Add CLng(VLA_Runtime.VlaDictGet(colOf, nm))
                rc.Add i
                VLA_Runtime.VlaDictSet seenInAtom, nm, True
            Else
                nc.Add i
                nn.Add nm
                VLA_Runtime.VlaDictSet seenInAtom, nm, True
            End If
        End If
    Next i
    leftCols = VLA_Relation.CollToLongArray(lc)
    rightCols = VLA_Relation.CollToLongArray(rc)
    newVarCols = VLA_Relation.CollToLongArray(nc)
    newVarNames = CollToStringArray(nn)
End Sub

' joined's tuples are accum's own accumArity columns followed by the
' filtered atom-relation's full columns (RelJoin's own left++right
' concatenation) - keep the accumulator columns as-is and append only
' the atom's brand-new variable columns (newVarCols, 1-based positions
' WITHIN the atom, offset by accumArity to land in joined).
'
' joined is RelJoin's own return value directly - a bare Collection of
' tuple arrays, per RelJoin's own header ("every matching pair,
' concatenated left-then-right"), NOT a Relation record (RelNew's own
' 3-item wrapper: arity, tuples, index). Iterated as-is; wrapping it
' through VLA_Relation.RelTuples() here would try to unwrap an already-
' unwrapped Collection, reading its SECOND TUPLE as if it were the
' outer wrapper's own tuples-sub-collection at Item(2) - live-caught:
' runtime error 424 ("Object required") on the `Set` inside RelTuples,
' since a tuple array is a Variant, not an object.
Private Function ProjectAfterJoin(ByVal joined As Collection, ByVal accumArity As Long, ByRef newVarCols As Variant) As Collection
    Dim newCount As Long
    newCount = UBound(newVarCols) - LBound(newVarCols) + 1
    Dim outArity As Long
    outArity = accumArity + newCount
    Dim outp As Collection
    Set outp = VLA_Relation.RelNew(outArity)
    Dim t As Variant, arr() As Variant
    For Each t In joined
        arr = t
        Dim nt() As Variant
        ReDim nt(1 To outArity)
        Dim i As Long
        For i = 1 To accumArity
            nt(i) = arr(i)
        Next i
        For i = 1 To newCount
            nt(accumArity + i) = arr(accumArity + newVarCols(LBound(newVarCols) + i - 1))
        Next i
        VLA_Relation.RelTryAdd outp, nt
    Next t
    Set ProjectAfterJoin = outp
End Function

' Evaluates one rule's body for one semi-naive "delta position":
' deltaPos = 0 means round 1's bootstrap (every atom reads relations()
' in full - full(IDB) is empty at that point regardless, so this is
' exactly the general formula's own round-1 case, not an
' approximation of it); deltaPos = 1..n means "body atom at that
' position reads deltas(), every other atom reads relations() in
' full" - the standard semi-naive construction. Builds the join chain
' via VLA_Relation.RelJoin starting from RelUnit() (arity 0, one empty
' tuple - the join identity), so the first body atom needs no special
' case: joining against RelUnit() with no shared columns degenerates
' to exactly that atom's own (filtered) relation.

' DATALOG.4's own comparison/arithmetic built-ins: resolves ONE operand
' (a ParseAtom-shaped arg record - already-bound VARIABLE or CONSTANT,
' CheckRuleSafety's own guarantee, never anything else) against a row
' currently being built up, arr() - the identical resolution
' FilterOutMatching's own probe-building loop already performs for a
' negated atom's operands, reused here for a comparison/arithmetic
' operand instead. A tuple cell is always a scalar (a fact's own
' constant text, or a table-sourced Value2 read) - never an object - so
' a plain assignment is safe, no CopyVariant needed.
Private Function ResolveOperand(ByVal a As Collection, ByVal colOf As Object, ByRef arr() As Variant) As Variant
    If ArgIsVar(a) Then
        ResolveOperand = arr(CLng(VLA_Runtime.VlaDictGet(colOf, ArgText(a))))
    Else
        ResolveOperand = ArgText(a)
    End If
End Function

' DATALOG's OWN numeric-ness policy for an already-resolved comparison/
' arithmetic operand - deliberately MORE LENIENT than VLA_Sql.bas's own
' WHERE-clause policy (VLA_Relation.ValueIsNumericType alone): a real
' Excel numeric type (a table-sourced value, read via RelFromRange's own
' Value2 path) counts, exactly as it does for SQL, but so does a
' numeric-LOOKING STRING (VLA_Relation.IsInvariantNumericString) - since
' DATALOG's own rule-text constants (the "50000" in (> X 50000)) and
' fact-block values are ALWAYS plain VBA Strings, never a real Excel
' cell type at all, unlike anything SQL's own WHERE clause ever
' compares. Without this, a filter like (> Salary 50000) could never
' fire on a hand-written fact - VLA_Relation.bas's own DATALOG.4 header
' note has the fuller reasoning for why this is a second, deliberate
' policy rather than a bug to reconcile with SQL's.
' PROLOG.17: the operand-count phrase and worked example a `let` refusal
' shows, given the arity VLA_Relation.ArithOpArity reported. Two small
' functions rather than one message per arity, so that
' datalog-builtin-needs-two-operands stays ONE id serving the comparison
' shape and both arithmetic shapes - and so that its BINARY rendering is
' byte-identical to DATALOG.4's own original text, which is what keeps
' this widening invisible to every test that already existed.
Private Function ArithArityWordsFor(ByVal wantArgs As Long) As String
    If wantArgs = 1 Then
        ArithArityWordsFor = "one operand"
    Else
        ArithArityWordsFor = "two operands"
    End If
End Function

Private Function ArithArityExampleFor(ByVal wantArgs As Long) As String
    If wantArgs = 1 Then
        ArithArityExampleFor = "(abs X)"
    Else
        ArithArityExampleFor = "(+ X Y)"
    End If
End Function

Private Function BuiltinOperandIsNumeric(ByVal v As Variant) As Boolean
    If VLA_Relation.ValueIsNumericType(v) Then
        BuiltinOperandIsNumeric = True
    ElseIf VarType(v) = vbString Then
        BuiltinOperandIsNumeric = VLA_Relation.IsInvariantNumericString(CStr(v))
    End If
End Function

Private Function EvalRuleBody(ByVal headAtom As Collection, ByVal bodyItems As Collection, _
                               ByVal deltaPos As Long, ByVal relations As Object, ByVal deltas As Object) As Collection
    Dim accum As Collection
    Set accum = VLA_Relation.RelUnit()
    Dim colOf As Object
    Set colOf = VLA_Runtime.VlaDictNew()

    Dim bi As Long
    For bi = 1 To bodyItems.Count
        Dim item As Collection
        Set item = bodyItems.Item(bi)
        Dim atom As Collection
        Set atom = BodyItemAtom(item)

        Dim kind As Long
        kind = BodyItemKind(item)

        If BodyItemNeedsFullRelation(kind) Then
            ' Negation and both aggregate kinds always read relations() in
            ' full, never deltas() - stratification (RunStratifiedFixpoint)
            ' guarantees the predicate they read finished, in full, before
            ' this rule's own stratum ever began, so it can never itself be
            ' mid-fixpoint here. RunFixpointForRules' own round loop never
            ' passes one of these body positions as deltaPos, for the
            ' identical reason - the "bi = deltaPos" branch below is
            ' unreachable for them, so it is not even checked.
            Dim fullRel As Collection
            Set fullRel = Nothing
            If VLA_Runtime.VlaDictHas(relations, AtomPred(atom)) Then
                Set fullRel = VLA_Runtime.VlaDictGet(relations, AtomPred(atom))
            End If
            Select Case kind
            Case BI_NOT
                Set accum = FilterOutMatching(accum, atom, colOf, fullRel)
                If VLA_Relation.RelCount(accum) = 0 Then
                    Set EvalRuleBody = New Collection
                    Exit Function
                End If
            Case BI_COUNT, BI_SUM
                Dim groups As Object
                Set groups = ComputeAggregateGroups(kind, atom, colOf, fullRel)
                Set accum = ApplyAggregate(kind, accum, atom, colOf, groups)
                VLA_Runtime.VlaDictSet colOf, BodyItemResultVar(item), VLA_Relation.RelArity(accum)
                ' An aggregate always succeeds (even a zero-tuple group
                ' yields a real, meaningful 0) - RelCount can never
                ' legitimately reach 0 here unless accum itself was
                ' already empty coming in, in which case this is simply
                ' propagating that, not a new empty-result case to guard.
                If VLA_Relation.RelCount(accum) = 0 Then
                    Set EvalRuleBody = New Collection
                    Exit Function
                End If
            End Select
        ElseIf kind = BI_CMP Then
            ' DATALOG.4's own comparison filter, (op X Y) - evaluated
            ' INLINE, per row, directly against accum's own rows so far
            ' (no predicate lookup, no join: neither operand can be
            ' anything but already-bound-variable-or-constant,
            ' CheckRuleSafety's own guarantee). ResolveOperand reads a
            ' variable's own value positionally through colOf, exactly
            ' the way FilterOutMatching's own probe-building loop already
            ' resolves a negated atom's operands.
            Dim filteredCmp As Collection
            Set filteredCmp = VLA_Relation.RelNew(VLA_Relation.RelArity(accum))
            Dim tCmp As Variant, arrCmp() As Variant
            For Each tCmp In VLA_Relation.RelTuples(accum)
                arrCmp = tCmp
                Dim lValC As Variant, rValC As Variant
                lValC = ResolveOperand(AtomArgAt(atom, 1), colOf, arrCmp)
                rValC = ResolveOperand(AtomArgAt(atom, 2), colOf, arrCmp)
                If VLA_Relation.CompareValues(AtomPred(atom), lValC, rValC, _
                        BuiltinOperandIsNumeric(lValC) And BuiltinOperandIsNumeric(rValC)) Then
                    VLA_Relation.RelTryAdd filteredCmp, arrCmp
                End If
            Next tCmp
            Set accum = filteredCmp
            If VLA_Relation.RelCount(accum) = 0 Then
                Set EvalRuleBody = New Collection
                Exit Function
            End If
        ElseIf kind = BI_LET Then
            ' DATALOG.4's own arithmetic binding, (let Z (op X Y)) -
            ' extends every one of accum's own rows with ONE new column
            ' (Z, its own resultVar), the same "extend, never drop" shape
            ' ApplyAggregate's own aggregate-column extension already
            ' uses, since a let - unlike a comparison filter - always
            ' succeeds for a row once its operands are numeric.
            Dim inArityLet As Long
            inArityLet = VLA_Relation.RelArity(accum)
            Dim extendedLet As Collection
            Set extendedLet = VLA_Relation.RelNew(inArityLet + 1)
            Dim tLet As Variant, arrLet() As Variant
            For Each tLet In VLA_Relation.RelTuples(accum)
                arrLet = tLet
                Dim lValL As Variant, rValL As Variant
                Dim okArith As Boolean, reasonArith As String
                Dim computedVal As Variant
                ' PROLOG.17: unary and binary dispatch off the SAME shared
                ' table the parse-time gate above consulted, so the two
                ' cannot disagree about an operator's arity.
                If VLA_Relation.ArithOpArity(AtomPred(atom)) = 1 Then
                    lValL = ResolveOperand(AtomArgAt(atom, 1), colOf, arrLet)
                    computedVal = VLA_Relation.ComputeArithmeticUnary(AtomPred(atom), lValL, _
                            BuiltinOperandIsNumeric(lValL), okArith, reasonArith)
                Else
                    lValL = ResolveOperand(AtomArgAt(atom, 1), colOf, arrLet)
                    rValL = ResolveOperand(AtomArgAt(atom, 2), colOf, arrLet)
                    computedVal = VLA_Relation.ComputeArithmetic(AtomPred(atom), lValL, rValL, _
                            BuiltinOperandIsNumeric(lValL) And BuiltinOperandIsNumeric(rValL), okArith, reasonArith)
                End If
                ' PROLOG.17: EVERY reason is mapped, not only the two this
                ' site used to be able to receive. The substrate now
                ' returns "domain-error", "overflow" and
                ' "unknown-operator" as well, and the old two-branch shape
                ' would have reported a (sqrt -1) as a NON-NUMERIC OPERAND
                ' - a confidently wrong explanation of a real refusal. The
                ' final Else keeps its original meaning as the
                ' not-numeric case AND catches any reason added later,
                ' which is refused rather than computed.
                If Not okArith Then
                    Select Case reasonArith
                    Case "divide-by-zero"
                        VLA_Messages.RaiseMsg "datalog-division-by-zero"
                    Case "domain-error"
                        VLA_Messages.RaiseMsg "datalog-arithmetic-domain-error", "operator", AtomPred(atom)
                    Case "overflow"
                        VLA_Messages.RaiseMsg "datalog-arithmetic-overflow", "operator", AtomPred(atom)
                    Case "unknown-operator"
                        VLA_Messages.RaiseMsg "datalog-unknown-arithmetic-operator", "operator", AtomPred(atom)
                    Case Else
                        VLA_Messages.RaiseMsg "datalog-arithmetic-non-numeric-operand", "operator", AtomPred(atom)
                    End Select
                End If
                Dim ntLet() As Variant
                ReDim ntLet(1 To inArityLet + 1)
                Dim iLet As Long
                For iLet = 1 To inArityLet
                    ntLet(iLet) = arrLet(iLet)
                Next iLet
                ntLet(inArityLet + 1) = computedVal
                VLA_Relation.RelTryAdd extendedLet, ntLet
            Next tLet
            Set accum = extendedLet
            VLA_Runtime.VlaDictSet colOf, BodyItemResultVar(item), VLA_Relation.RelArity(accum)
            If VLA_Relation.RelCount(accum) = 0 Then
                Set EvalRuleBody = New Collection
                Exit Function
            End If
        Else
            Dim srcRel As Collection
            Dim havePred As Boolean
            If bi = deltaPos Then
                havePred = VLA_Runtime.VlaDictHas(deltas, AtomPred(atom))
                If havePred Then Set srcRel = VLA_Runtime.VlaDictGet(deltas, AtomPred(atom))
            Else
                havePred = VLA_Runtime.VlaDictHas(relations, AtomPred(atom))
                If havePred Then Set srcRel = VLA_Runtime.VlaDictGet(relations, AtomPred(atom))
            End If
            If Not havePred Then
                Set EvalRuleBody = New Collection
                Exit Function
            End If

            Dim filtered As Collection
            Set filtered = FilterAtomRelation(srcRel, atom)
            If VLA_Relation.RelCount(filtered) = 0 Then
                Set EvalRuleBody = New Collection
                Exit Function
            End If

            Dim leftCols As Variant, rightCols As Variant
            Dim newVarCols As Variant, newVarNames As Variant
            BuildJoinPlan colOf, atom, leftCols, rightCols, newVarCols, newVarNames

            Dim accumArity As Long
            accumArity = ColOfCount(colOf)
            Dim joined As Collection
            Set joined = VLA_Relation.RelJoin(accum, leftCols, filtered, rightCols)
            Dim projected As Collection
            Set projected = ProjectAfterJoin(joined, accumArity, newVarCols)

            Dim ni As Long
            For ni = LBound(newVarNames) To UBound(newVarNames)
                accumArity = accumArity + 1
                VLA_Runtime.VlaDictSet colOf, newVarNames(ni), accumArity
            Next ni
            Set accum = projected

            If VLA_Relation.RelCount(accum) = 0 Then
                Set EvalRuleBody = New Collection
                Exit Function
            End If
        End If
    Next bi

    Dim outp As New Collection
    Dim rowV As Variant, rowArr() As Variant
    For Each rowV In VLA_Relation.RelTuples(accum)
        rowArr = rowV
        Dim head() As Variant
        ReDim head(1 To AtomArity(headAtom))
        Dim hi2 As Long
        For hi2 = 1 To AtomArity(headAtom)
            Dim harg As Collection
            Set harg = AtomArgAt(headAtom, hi2)
            If ArgIsVar(harg) Then
                head(hi2) = rowArr(CLng(VLA_Runtime.VlaDictGet(colOf, ArgText(harg))))
            Else
                head(hi2) = ArgText(harg)
            End If
        Next hi2
        outp.Add head
    Next rowV
    Set EvalRuleBody = outp
End Function

' Runs one rule at one delta position, merging any newly-derived
' tuples into relations(head) (the running "full" set) and, for
' whichever of them are genuinely new, into newDeltas(head) (next
' round's delta). Returns True iff at least one tuple was new.
Private Function RunOneRulePass(ByVal ruleRec As Collection, ByVal deltaPos As Long, _
                                 ByVal relations As Object, ByVal deltas As Object, ByVal newDeltas As Object) As Boolean
    Dim headAtom As Collection, bodyAtoms As Collection
    Set headAtom = ruleRec.Item(1)
    Set bodyAtoms = ruleRec.Item(2)
    Dim derived As Collection
    Set derived = EvalRuleBody(headAtom, bodyAtoms, deltaPos, relations, deltas)
    If derived.Count = 0 Then Exit Function
    Dim headPred As String
    headPred = AtomPred(headAtom)
    Dim full As Collection
    Set full = GetOrCreateRelation(relations, headPred, AtomArity(headAtom))
    Dim newRel As Collection
    If VLA_Runtime.VlaDictHas(newDeltas, headPred) Then
        Set newRel = VLA_Runtime.VlaDictGet(newDeltas, headPred)
    Else
        Set newRel = VLA_Relation.RelNew(AtomArity(headAtom))
        VLA_Runtime.VlaDictSet newDeltas, headPred, newRel
    End If
    Dim t As Variant, arr() As Variant
    Dim foundNew As Boolean
    For Each t In derived
        arr = t
        If VLA_Relation.RelTryAdd(full, arr) Then
            VLA_Relation.RelTryAdd newRel, arr
            foundNew = True
        End If
    Next t
    RunOneRulePass = foundNew
End Function

' The semi-naive round loop: round 1 evaluates every rule once, fully
' (deltaPos=0); every later round evaluates every rule once PER body
' position, substituting the previous round's delta at that one
' position - the standard construction, correct for both simple and
' mutually-recursive rule sets, and guaranteed to terminate because
' function-free Horn clauses (this engine's own parse-time-enforced
' restriction) bound the whole universe of derivable tuples. MAX_ROUNDS
' is a named, loud safety valve for a buggy rule set, not a normal exit.
'
' Renamed from RunFixpoint (DATALOG.0) to RunFixpointForRules once
' stratified negation needed to call this same flat algorithm once per
' stratum rather than once over the whole program - this sub itself
' still knows nothing about strata; RunStratifiedFixpoint below owns
' that. A negated body position is never chosen as deltaPos (skipped in
' the round loop below) - stratification guarantees the predicate it
' negates already finished, in full, in an earlier call to this same
' sub, so it has no delta of its own to contribute here.
Private Sub RunFixpointForRules(ByVal rules As Collection, ByVal relations As Object)
    Const MAX_ROUNDS As Long = 10000
    Dim emptyDeltas As Object
    Set emptyDeltas = VLA_Runtime.VlaDictNew()

    Dim newDeltas As Object
    Set newDeltas = VLA_Runtime.VlaDictNew()
    Dim r As Variant, ruleRec As Collection
    Dim anyNew As Boolean
    anyNew = False
    For Each r In rules
        Set ruleRec = r
        If RunOneRulePass(ruleRec, 0, relations, emptyDeltas, newDeltas) Then anyNew = True
    Next r

    Dim round As Long
    round = 1
    Do While anyNew
        round = round + 1
        If round > MAX_ROUNDS Then VLA_Messages.RaiseMsg "datalog-round-ceiling", "rounds", MAX_ROUNDS
        Dim deltas As Object
        Set deltas = newDeltas
        Set newDeltas = VLA_Runtime.VlaDictNew()
        anyNew = False
        For Each r In rules
            Set ruleRec = r
            Dim bodyItems As Collection
            Set bodyItems = ruleRec.Item(2)
            Dim bi As Long
            For bi = 1 To bodyItems.Count
                ' BI_POS only - DATALOG.4's note above BodyItemNeedsFullRelation
                ' has the reasoning: since it added BI_CMP/BI_LET, "does not
                ' need the full relation" no longer implies "is a valid
                ' delta position" (a comparison/arithmetic built-in reads no
                ' predicate's relations()/deltas() dict at all, by name or
                ' otherwise), so this checks the narrower, correct condition
                ' directly rather than "Not BodyItemNeedsFullRelation(...)".
                If BodyItemKind(bodyItems.Item(bi)) = BI_POS Then
                    If RunOneRulePass(ruleRec, bi, relations, deltas, newDeltas) Then anyNew = True
                End If
            Next bi
        Next r
    Loop
End Sub

' Builds the predicate dependency graph (an edge bodyPred -> headPred per
' body item, tagged strict/non-strict via BodyItemNeedsFullRelation - a
' BI_NOT, BI_COUNT, or BI_SUM body item all demand the SAME "fully,
' finally computed first" guarantee, so all three tag their own edge
' identically) and solves for each predicate's stratum number via the
' standard Bellman-Ford-shaped relaxation: a non-strict (BI_POS) edge
' requires stratum(head) >= stratum(body); a strict edge requires
' stratum(head) > stratum(body) (the dependency must be fully computed
' in an earlier stratum before the rule using it may begin). Every
' predicate not otherwise constrained (a bare EDB fact/table predicate,
' or a rule head with no `not`/`count`/`sum` anywhere upstream of it)
' settles at stratum 0.
'
' preds.Count relaxation passes are always enough to reach a fixed point
' for any STRATIFIABLE program (the longest possible dependency chain
' visits every predicate at most once); one further verification pass
' that still finds a required increase is therefore proof of a cycle
' carrying at least one strict edge - some predicate negates or
' aggregates itself, directly or through a chain of rules - refused by
' name rather than looped forever or silently mis-evaluated.
Private Sub ComputeStrata(ByVal rules As Collection, ByRef strataOf As Object)
    Dim preds As Object
    Set preds = VLA_Runtime.VlaDictNew()
    Dim edgeFrom As New Collection, edgeTo As New Collection, edgeNeg As New Collection

    Dim r As Variant, ruleRec As Collection
    For Each r In rules
        Set ruleRec = r
        Dim headAtom As Collection, bodyItems As Collection
        Set headAtom = ruleRec.Item(1)
        Set bodyItems = ruleRec.Item(2)
        Dim headPred As String
        headPred = AtomPred(headAtom)
        VLA_Runtime.VlaDictSet preds, headPred, True
        Dim bi As Variant
        For Each bi In bodyItems
            Dim item As Collection
            Set item = bi
            Dim bik As Long
            bik = BodyItemKind(item)
            If bik = BI_CMP Or bik = BI_LET Then
                ' DATALOG.4's own comparison/arithmetic built-ins
                ' contribute NO edge here: BodyItemAtom's own "predicate"
                ' for either kind is an OPERATOR symbol (">", "+", ...),
                ' never a real relation - counting it as one would risk
                ' colliding with an actual predicate someone separately
                ' chose to name "+" elsewhere in the same program
                ' (ParseAtom itself never forbids that), however unlikely
                ' in practice. Consistent with BodyItemNeedsFullRelation's
                ' own False for both - neither kind touches stratification
                ' at all, this item's own roadmap words.
            Else
                Dim bodyPred As String
                bodyPred = AtomPred(BodyItemAtom(item))
                VLA_Runtime.VlaDictSet preds, bodyPred, True
                edgeFrom.Add bodyPred
                edgeTo.Add headPred
                edgeNeg.Add BodyItemNeedsFullRelation(bik)
            End If
        Next bi
    Next r

    Set strataOf = VLA_Runtime.VlaDictNew()
    Dim k As Variant
    For Each k In VLA_Runtime.VlaDictKeys(preds)
        VLA_Runtime.VlaDictSet strataOf, CStr(k), 0&
    Next k

    Dim nPreds As Long
    nPreds = VLA_Runtime.VlaDictKeys(preds).Count

    Dim pass As Long, changed As Boolean
    For pass = 1 To nPreds
        changed = False
        Dim ei As Long
        For ei = 1 To edgeFrom.Count
            Dim need As Long
            need = CLng(VLA_Runtime.VlaDictGet(strataOf, edgeFrom.Item(ei)))
            If edgeNeg.Item(ei) Then need = need + 1
            If CLng(VLA_Runtime.VlaDictGet(strataOf, edgeTo.Item(ei))) < need Then
                VLA_Runtime.VlaDictSet strataOf, edgeTo.Item(ei), need
                changed = True
            End If
        Next ei
        If Not changed Then Exit For
    Next pass

    ' Verification pass: a stratifiable program is already stable by
    ' here (nPreds passes is provably enough); anything still wanting to
    ' rise proves a cycle carrying at least one negative edge.
    Dim ej As Long
    For ej = 1 To edgeFrom.Count
        Dim need2 As Long
        need2 = CLng(VLA_Runtime.VlaDictGet(strataOf, edgeFrom.Item(ej)))
        If edgeNeg.Item(ej) Then need2 = need2 + 1
        If CLng(VLA_Runtime.VlaDictGet(strataOf, edgeTo.Item(ej))) < need2 Then
            VLA_Messages.RaiseMsg "datalog-negation-not-stratifiable", "predicate", CStr(edgeTo.Item(ej))
        End If
    Next ej
End Sub

' Runs every rule to a full fixpoint, one stratum at a time, in
' ascending stratum order, sharing ONE relations dict across all of
' them - by the time stratum S's own rules run, every predicate any of
' them negates (necessarily a strictly lower stratum, ComputeStrata's
' own guarantee) is already sitting in relations() fully and finally
' computed, so EvalRuleBody's negation branch never observes a partial
' answer. Replaces the flat, single-call RunFixpoint (DATALOG.0) that
' had no notion of strata at all - a program with no `not` anywhere
' folds every rule into stratum 0, so this degenerates back to exactly
' one RunFixpointForRules call, identical to the old behavior.
Private Sub RunStratifiedFixpoint(ByVal rules As Collection, ByVal relations As Object)
    If rules.Count = 0 Then Exit Sub

    Dim strataOf As Object
    ComputeStrata rules, strataOf

    Dim byStratum As Object
    Set byStratum = VLA_Runtime.VlaDictNew()
    Dim maxStratum As Long
    maxStratum = 0

    Dim r As Variant, ruleRec As Collection
    For Each r In rules
        Set ruleRec = r
        Dim s As Long
        s = CLng(VLA_Runtime.VlaDictGet(strataOf, AtomPred(ruleRec.Item(1))))
        If s > maxStratum Then maxStratum = s
        Dim key As String
        key = CStr(s)
        Dim bucket As Collection
        If VLA_Runtime.VlaDictHas(byStratum, key) Then
            Set bucket = VLA_Runtime.VlaDictGet(byStratum, key)
        Else
            Set bucket = New Collection
            VLA_Runtime.VlaDictSet byStratum, key, bucket
        End If
        bucket.Add ruleRec
    Next r

    Dim s2 As Long
    For s2 = 0 To maxStratum
        Dim key2 As String
        key2 = CStr(s2)
        If VLA_Runtime.VlaDictHas(byStratum, key2) Then
            RunFixpointForRules VLA_Runtime.VlaDictGet(byStratum, key2), relations
        End If
    Next s2
End Sub

' The pure core, with no Excel dependency at all: parses rulesText,
' merges its (fact ...) forms and any supplied baseRelations (a VlaDict
' of predicate-name -> Relation, typically built via
' VLA_Relation.RelFromRange) into one working set, runs the fixpoint,
' and returns a 4-item Collection: (1) the queried predicate's own
' folded name, (2) the full VlaDict of every relation (EDB and IDB
' alike) after evaluation, (3) that predicate's own head-variable
' names as a Variant array (e.g. "X","Y" for a rule-derived predicate),
' or Empty when no rule defines it (a raw fact/table predicate has no
' variable names to offer - DATALOG falls back to "Col1".."ColN"),
' (4) whether the program's own (headless) directive was present.
' VLA_Tests_Query.TestDatalog is this function's own first caller.
'
' Checks "baseRelations Is Nothing", not IsMissing(baseRelations) - a
' real, live-caught VBA trap: IsMissing only reliably detects an
' omitted argument for an Optional parameter typed As Variant. For a
' typed Optional Object parameter like this one, IsMissing returns
' False even when the caller omitted it entirely, so the omitted
' parameter's own default value (Nothing, for any omitted Object
' parameter) is the only thing that can be checked correctly. An
' omitted-so-False IsMissing here fell through to `Set relations =
' baseRelations` with baseRelations still Nothing, and the first
' VlaDictKeys call on it raised runtime error 424 ("Object required").
'
' DATALOG.5: headerMap (predicate's own folded name -> VLA_Relation.
' RangeColumnNames' own (folded, original) pair Collection) is built by
' the caller (DATALOG() UDF; VLA_Tests_Query's own pure tests build one
' by hand) BEFORE this function ever runs, the same call-ordering
' baseRelations already needed - threaded straight through to
' ParseProgram, which is where a keyed rule-body atom actually gets
' desugared. Omitted (Nothing) is treated as "no table has a header at
' all" (an empty VlaDict), never a crash - the same Is-Nothing-not-
' IsMissing discipline baseRelations above already established, for the
' identical reason (a typed Optional Object parameter).
Public Function DatalogRun(ByVal rulesText As String, Optional ByVal baseRelations As Object, Optional ByVal headerMap As Object) As Collection
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

    Dim predArity As Object
    Set predArity = VLA_Runtime.VlaDictNew()
    Dim k As Variant
    For Each k In VLA_Runtime.VlaDictKeys(relations)
        RecordArity predArity, CStr(k), VLA_Relation.RelArity(VLA_Runtime.VlaDictGet(relations, k))
    Next k

    Dim facts As Collection, rules As Collection, queryName As String
    Dim headless As Boolean
    ParseProgram rulesText, facts, rules, queryName, headless, hMap

    Dim fa As Variant, factAtom As Collection
    For Each fa In facts
        Set factAtom = fa
        RecordArity predArity, AtomPred(factAtom), AtomArity(factAtom)
    Next fa
    Dim rr As Variant, ruleRec As Collection
    For Each rr In rules
        Set ruleRec = rr
        RecordArity predArity, AtomPred(ruleRec.Item(1)), AtomArity(ruleRec.Item(1))
        Dim ba As Variant
        For Each ba In ruleRec.Item(2)
            Dim baItem As Collection
            Set baItem = ba
            RecordArity predArity, AtomPred(BodyItemAtom(baItem)), AtomArity(BodyItemAtom(baItem))
        Next ba
    Next rr

    For Each fa In facts
        Set factAtom = fa
        Dim rel As Collection
        Set rel = GetOrCreateRelation(relations, AtomPred(factAtom), AtomArity(factAtom))
        Dim t() As Variant
        ReDim t(1 To AtomArity(factAtom))
        Dim ti As Long
        For ti = 1 To AtomArity(factAtom)
            t(ti) = ArgText(AtomArgAt(factAtom, ti))
        Next ti
        VLA_Relation.RelTryAdd rel, t
    Next fa

    ' every IDB head predicate is queryable as "zero rows", not
    ' "undefined", even before the fixpoint derives anything for it -
    ' GetOrCreateRelation is idempotent, so calling it again once rows
    ' actually exist is safe.
    For Each rr In rules
        Set ruleRec = rr
        GetOrCreateRelation relations, AtomPred(ruleRec.Item(1)), AtomArity(ruleRec.Item(1))
    Next rr

    RunStratifiedFixpoint rules, relations

    If Not VLA_Runtime.VlaDictHas(relations, queryName) Then
        VLA_Messages.RaiseMsg "datalog-query-unknown-predicate", "predicate", queryName
    End If

    ' The first rule (encounter order) whose own head names a given
    ' predicate lends that predicate its column names - X, Y, not
    ' Col1, Col2 - for the one case Datalog's own positional-only
    ' predicates actually have real names to offer. A predicate with
    ' no defining rule (a raw fact or table relation, queried
    ' directly) has no such names; VlaDictHas below correctly stays
    ' False for it, and DATALOG falls back to its own generic headers.
    Dim headNamesByPred As Object
    Set headNamesByPred = VLA_Runtime.VlaDictNew()
    For Each rr In rules
        Set ruleRec = rr
        Dim hAtom As Collection
        Set hAtom = ruleRec.Item(1)
        Dim hp As String
        hp = AtomPred(hAtom)
        If Not VLA_Runtime.VlaDictHas(headNamesByPred, hp) Then
            Dim names() As Variant
            ReDim names(1 To AtomArity(hAtom))
            Dim hi3 As Long
            For hi3 = 1 To AtomArity(hAtom)
                names(hi3) = ArgText(AtomArgAt(hAtom, hi3))
            Next hi3
            VLA_Runtime.VlaDictSet headNamesByPred, hp, names
        End If
    Next rr

    Dim outp As New Collection
    outp.Add queryName
    outp.Add relations
    If VLA_Runtime.VlaDictHas(headNamesByPred, queryName) Then
        outp.Add VLA_Runtime.VlaDictGet(headNamesByPred, queryName)
    Else
        outp.Add Empty
    End If
    outp.Add headless
    Set DatalogRun = outp
End Function

' A table argument's own name: an Excel Table's ListObject.Name, or a
' plain Range's own defined name (its sheet-qualifying "Sheet1!"
' prefix, if any, stripped) - never a name the caller chooses in the
' formula. This is what makes "an Employees table IS a relation with
' no authoring step" true, and it is also why renaming what a table
' means to DATALOG is a one-line (rule (reports_to X Y) (Employees X Y))
' rather than a second naming mechanism.
'
' SQL.1: the actual resolution now lives in VLA_Relation.TableArgResolve,
' shared with VLA_Sql.bas the moment a second engine needed the identical
' mechanism (this module's own header note, "the substrate/engine split
' doing its job"). VLA_Relation.bas cannot itself call VLA_Messages (its
' own LAYER 0 contract), so the shared function returns a category code
' rather than raising - this thin wrapper is what turns that code back
' into DATALOG's own, already-shipped wording, unchanged.
Private Function TableArgName(ByVal v As Variant) As String
    Dim ok As Boolean, reason As String
    TableArgName = VLA_Relation.TableArgResolve(v, ok, reason)
    If ok Then Exit Function
    Select Case reason
    Case "not-a-range"
        VLA_Messages.RaiseMsg "datalog-table-not-a-range"
    Case "noncontiguous"
        VLA_Messages.RaiseMsg "datalog-table-noncontiguous-columns"
    Case "needs-a-name"
        VLA_Messages.RaiseMsg "datalog-table-needs-a-name"
    End Select
End Function

' The worksheet-facing entry point - =DATALOG(rules, table1, table2, ...).
' rulesText: a single cell or a literal string; for a multi-cell rule
' set, join it in the formula itself first (TEXTJOIN(CHAR(10), TRUE,
' rules_range)), the ordinary Excel idiom for a text-taking function -
' DATALOG does not flatten a range on its own. Each further argument is
' a cell range: an Excel Table (name = its own ListObject name, header
' row auto-stripped) or a plain named range (name = its defined name,
' every row a fact, no header assumed). Passing the WHOLE table is still
' the common case, but a NARROWER, contiguous block of the table's own
' columns is honored too (DATALOG.3): the predicate name stays the
' table's own ListObject.Name regardless, only the arity and values
' narrow to exactly those columns, positionally - VLA_Relation.
' SourceToArray's own header has the full mechanism. A multi-area
' (Ctrl-selected, non-contiguous) column selection is refused by name
' (datalog-table-noncontiguous-columns), not silently guessed.
'
' Errors return as readable TEXT in the cell ("#DATALOG! ...") rather
' than Excel's own opaque #VALUE! - a worksheet function has no other
' visible surface for a refusal's words to reach the person who wrote
' the formula, and swallowing them into #VALUE! would undo LX.8's own
' "name the problem" doctrine at the one surface this engine has.
Public Function DATALOG(ByVal rulesText As String, ParamArray tables() As Variant) As Variant
    On Error GoTo fail
    Dim relations As Object
    Set relations = VLA_Runtime.VlaDictNew()
    ' DATALOG.5: alongside each table argument's own Relation, also
    ' capture its header row (VLA_Relation.RangeColumnNames - Nothing
    ' added for a plain named range, which has no header of its own to
    ' offer) so a keyed rule-body atom can resolve a column NAME against
    ' it. Built here, before DatalogRun/ParseProgram ever run, the same
    ' way baseRelations already had to be.
    Dim headerMap As Object
    Set headerMap = VLA_Runtime.VlaDictNew()
    Dim i As Long
    For i = LBound(tables) To UBound(tables)
        Dim nm As String
        nm = TableArgName(tables(i))
        Dim rel As Collection
        Set rel = VLA_Relation.RelFromRange(tables(i))
        VLA_Runtime.VlaDictSet relations, nm, rel
        Dim colsOk As Boolean
        Dim cols As Collection
        Set cols = VLA_Relation.RangeColumnNames(tables(i), colsOk)
        If colsOk Then VLA_Runtime.VlaDictSet headerMap, nm, cols
    Next i

    Dim result As Collection
    Set result = DatalogRun(rulesText, relations, headerMap)
    Dim queryName As String
    queryName = result.Item(1)
    Dim allRel As Object
    Set allRel = result.Item(2)
    Dim queried As Collection
    Set queried = VLA_Runtime.VlaDictGet(allRel, queryName)
    ' result.Item(3) is the queried predicate's own rule-head variable
    ' names (e.g. "X","Y"), or Empty when no rule defines it - DatalogRun's
    ' own header comment. A raw fact/table predicate, queried directly
    ' with no defining rule, still falls back to RelToSpilledArray's own
    ' generic "Col1".."ColN". result.Item(4) is the program's own
    ' (headless) directive - when present, RelToSpilledArray ignores
    ' whichever header source result.Item(3) would otherwise have named.
    Dim isHeadless As Boolean
    isHeadless = result.Item(4)
    If IsEmpty(result.Item(3)) Then
        DATALOG = VLA_Relation.RelToSpilledArray(queried, , isHeadless)
    Else
        DATALOG = VLA_Relation.RelToSpilledArray(queried, result.Item(3), isHeadless)
    End If
    Exit Function
fail:
    DATALOG = "#DATALOG! " & Err.Description
End Function
