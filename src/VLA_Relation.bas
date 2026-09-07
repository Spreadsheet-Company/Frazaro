Attribute VB_Name = "VLA_Relation"
Option Explicit
Public Const VLA_RELATION_VERSION As String = "SQL.4"
' SQL.4: RelGroupBy - the shared grouping kernel BETA_ROADMAP2.md's own
' SQL.4 text asked for by name ("the third appearance of grouping in
' this codebase..., so it generalizes existing key-hashing rather than
' inventing"), generalizing VLA_Datalog.bas's own BoundPositionPairs/
' KeyFromPositions/ComputeAggregateGroups/ApplyAggregate (DATALOG.2's own
' grouping mechanism) to plain, ALREADY-RESOLVED 1-based column
' POSITIONS rather than atom variables. What's genuinely shared is only
' the group-key-building plus per-group-accumulator mechanism itself -
' DATALOG's own bound/unbound-variable consistency checking
' (BoundPositionPairs' own colOf/ArgIsVar walk) is real Datalog
' unification-lite semantics, not generic grouping, and stays exactly
' where it is; VLA_Sql.bas's own caller has already resolved every GROUP
' BY item and every aggregate operand to an absolute row position (via
' its own colMap) before this function ever runs, so there is no
' unification layer left to share. VLA_Datalog.bas itself is NOT
' refactored to consume this (considered, per this item's own scoping
' note - not required, and its own consistency-checking walk has no
' equivalent here to hook into cleanly; left as a documented option for
' a future item rather than done speculatively).
'
' Five aggregate kinds (Public consts below) - no separate COUNT(*) kind:
' COUNT(*) and COUNT(col) are IDENTICAL here (both just the group's own
' row count) because this engine has no NULL representation anywhere
' (real SQL's COUNT(col) skips NULL values; this one has nothing to
' skip) - a deliberate, named simplification, not an oversight; the
' isStar/operand distinction that DOES still matter (whether an operand
' needs to be evaluated and validated at all) stays entirely on
' VLA_Sql.bas's own side, in its own AST.
'
' MIN/MAX/AVG over a GROUP with zero matching tuples cannot happen for a
' REAL GROUP BY group (a group key only ever exists because at least one
' row produced it) - only the caller's own "alwaysOneGroup" forced
' trivial group (SQL's own "an ungrouped aggregate is a real answer even
' over zero rows" stance, matching DATALOG.2's own COUNT/SUM precedent)
' can be empty, and MIN/MAX/AVG have no sensible numeric default over
' nothing - ok=False/reason="empty-aggregate" there, rather than
' inventing a NULL sentinel this engine has no representation for
' anywhere else. COUNT/SUM over the SAME empty case are real, meaningful
' zeros (DATALOG.2's own precedent, unchanged). SUM/AVG additionally
' refuse (ok=False/reason="not-numeric") the moment a non-numeric value
' reaches their own accumulator - MIN/MAX never do, since CompareValues
' already degrades to a text comparison when a group's own values aren't
' both numeric, exactly like EvalBool's own NK_CMP case. Message WORDING
' stays VLA_Sql.bas's own job throughout (this module's own LAYER 0.5
' contract, TableArgResolve's own header) - this function only ever
' returns a reason code plus, via failSpecIndex, WHICH aggSpec failed.
'
' DATALOG.5: RelFromRange gains an optional hasHeader flag (default
' False, every existing caller unchanged) - a pure-test-only seam so
' VLA_Datalog.bas's own keyed-atom desugaring (named-column atoms,
' (staffing (name X) (salary S)) instead of (staffing X _ S)) can be
' pinned against a hand-built array, with no live workbook, the same
' way DATALOG.3's own column-scoping needed a live Table and had to
' stay host-required - this item's own has-header case does not.
' RangeColumnNames now narrows to src's own passed column span exactly
' the way SourceToArray's own DATALOG.3 narrowing already does (colOffset/
' passedColCount, repeated rather than shared since the two read
' different things off .ListObject through different Range shapes) -
' DATALOG.5's own keyed names resolve within a narrowed slice only,
' since the slice is all VLA_Datalog.bas ever reads; every existing SQL
' caller always passes the whole table, so this is unchanged for SQL.
' Neither change touches this module's own actual join/storage/spill
' substrate - the real desugaring walk (recognizing a keyed atom's own
' S-expression shape, resolving it against a header map, building the
' positional atom it stands for) is bound to VLA_Datalog.bas's own
' ParseAtom/atom-record conventions, not DSL-ignorant substrate, and
' lives entirely there - this module owns no fact/rule parsing, no
' variables, no unification, unchanged.
'
' DATALOG.4: added the shared value-comparison/arithmetic primitive
' this section's own SQL.1 sequencing note asked for ("DATALOG.4
' should consume [SQL.1's scalar evaluator] rather than growing a
' second one") - ValueIsNumericType/CompareValues are VLA_Sql.bas's own
' former Private IsNumericValue/CompareValues, hoisted here UNCHANGED
' the moment a second engine needed the identical mechanical compare -
' TableArgResolve's own SQL.1 precedent, repeated: VLA_Sql.bas's own
' EvalBool now calls the shared CompareValues directly rather than
' keeping a second copy. What could NOT simply be reused as-is:
' deciding whether a value COUNTS as numeric in the first place. SQL's
' own WHERE clause trusts only a real Excel Value2 numeric type (a text
' cell that merely LOOKS numeric, "007", must never be silently
' promoted into a numeric comparison the column's own typing didn't ask
' for) - but DATALOG's own rule-text constants and fact-block values
' are ALWAYS plain VBA Strings, never a real Excel cell type at all
' (unlike a table-sourced value, which DOES carry a real Value2 type
' through RelFromRange's own ValueOf), so a filter like (> Salary
' 50000) could never fire on a hand-written fact under SQL's own
' stricter policy alone. Numeric-ness is therefore left to the CALLER
' (a bothNumeric Boolean each of CompareValues/ComputeArithmetic takes)
' rather than decided inside the shared function: VLA_Sql.bas keeps its
' own strict policy (ValueIsNumericType alone), while VLA_Datalog.bas's
' own policy also treats a numeric-LOOKING string as numeric
' (IsInvariantNumericString - locale-invariant, always '.' as the
' decimal separator, this section's own SD-4 freeze, never VBA's own
' locale-AWARE IsNumeric()). ComputeArithmetic is new (SQL.1 never
' needed arithmetic) but shares the identical ok/reason shape
' TableArgResolve already established for this module's own LAYER 0.5
' contract (never VLA_Messages - refusal WORDING stays each engine's
' own job; reason is "not-numeric" or "divide-by-zero", turned back
' into VLA_Datalog.bas's own worded refusals by its own caller).
'
' SQL.1: added TableArgResolve - VLA_Datalog.bas's own TableArgName
' hoisted here (as a message-agnostic resolver; see its own header) the
' moment VLA_Sql.bas needed the identical table-argument-naming
' mechanism - the substrate/engine split doing its job, per BETA_
' ROADMAP2.md's own SQL.1 sequencing note. This is the module's FIRST
' real VLA_* dependency (VLA_Identity.Fold) - LAYER moves from 0 to 0.5
' below; still never VLA_Messages (see TableArgResolve's own header for
' why that boundary holds regardless).
'
' Also added RangeToRows (a real, load-bearing correction, not a
' convenience twin of RelFromRange: SQL rows are a BAG, a Relation is a
' SET, and RelFromRange's own RelTryAdd would silently collapse two
' identical source rows into one - wrong for SQL, correct for Datalog
' facts, never interchangeable) and RangeColumnNames (the "new substrate
' DATALOG never needed" the SQL.1 sequencing note names - a folded/
' original column-name map, DATALOG's own Relations having no concept
' of a column NAME at all).
'
' DATALOG.2: SourceToArray narrows to a passed Range's own columns when
' it covers fewer than the table's full ListColumns.Count, rather than
' always reading the whole DataBodyRange once .ListObject resolves at
' all - closing the MVP's own stated "table argument must be the whole
' table, not one column of it" limit (BETA_ROADMAP2.md's own DATALOG.3
' history note, VLA_Datalog.bas, has the caller-side half: the early
' multi-area refusal). Always the table's full row span - only columns
' narrow, row-scoping was never part of this item's own stated limit.
'
' DATALOG.1: added RelContainsTuple - an exact-tuple membership probe,
' the one new substrate primitive VLA_Datalog.bas's stratified negation
' (`not`) needed and the MVP's own RelTryAdd/RelJoin/RelToSpilledArray
' set didn't already provide.

' =====================================================================
'  VLA_Relation - the shared substrate BETA_ROADMAP2.md's QUERY AND
'  LOGIC section promises: a Relation (a named-elsewhere, fixed-arity
'  SET of Variant tuples - duplicates silently absorbed, relational,
'  not a list), a real hash join (index built on whichever side is
'  smaller, probed with the larger - never nested-loop), and the two
'  boundary conversions every one of the four engines needs (a live
'  Excel Range OR a plain 2D array, in; a spilled 2D array, out).
'  VLA_Datalog.bas is the first caller; the roadmap's own SQL/PROLOG/
'  SOLVE items are written to reuse this file rather than re-deriving
'  a join apiece.
'
'  Deliberately ignorant of every DSL's own syntax: no fact/rule
'  parsing, no variables, no unification lives here - that is
'  VLA_Datalog.bas's job (and, later, each sibling engine's own). This
'  module only ever sees already-resolved tuples of constants, indexed
'  by 1-based column POSITION - it has no concept of a column name.
'
'  A Relation is a plain Collection (this codebase's own ad-hoc-record
'  convention - VLA_Interpreter.bas's mProcs entries, VLA_Messages.bas's
'  catalogue rows), not a class: Item(1)=arity As Long, Item(2)=tuples
'  As Collection (each element a 1-based Variant() of length arity),
'  Item(3)=a membership index (Object - Scripting.Dictionary, or
'  Nothing on a host with no Scripting runtime, e.g. Mac Excel).
'
'  LAYER:     0.5 (VLA_Identity only, as of SQL.1's TableArgResolve -
'             a foundational, dependency-free fold utility, not another
'             DSL's own machinery, so this stays well short of LAYER 1)
'  MAY CALL:  VLA_Identity (Fold - identifier normalization only, never
'             tuple data, SD-8's own distinction, unchanged from before).
'             Never VLA_Messages - refusal WORDING stays each engine's
'             own job (TableArgResolve's own header has the mechanism).
'  SHIPS:     add-in only - a Relation never appears in emitted VBA and
'             this module is never injected into a user workbook.
'  PAYS INTO: QUERY AND LOGIC's own "shared substrate, not four
'             engines" claim; DATALOG is the first proof of it.
'  REASON:    building the substrate once, generically, against
'             DATALOG's honest MVP-sized needs, is cheaper than
'             discovering its shape mid-way through SQL's full-
'             relational tier - the section's own sequencing note,
'             made real.
' =====================================================================

' SQL.4's own aggregate-kind vocabulary (RelGroupBy, near the end of
' this file) - declared here, in the module's own declarations section,
' because VBA refuses a module-level Const appearing AFTER any
' Sub/Function ("Only comments may appear after End Sub, End Function,
' or End Property") - every other Const in this codebase's own modules
' (VLA_Sql.bas's own TK_*/NK_* included) already lives up here for
' exactly this reason.
Public Const AGG_COUNT As Long = 0
Public Const AGG_SUM As Long = 1
Public Const AGG_MIN As Long = 2
Public Const AGG_MAX As Long = 3
Public Const AGG_AVG As Long = 4

' Tuple identity for dedup/index keys is DELIBERATELY case-SENSITIVE
' and NOT folded through VLA_Identity.Fold: a Relation's tuples are a
' program's DATA (a person's name, a part number), never an
' identifier - SD-8's invariant fold governs identifiers and keywords,
' not data, and folding "Bob" onto "bob" here would silently merge two
' different facts. Chr$(31) (ASCII Unit Separator) joins a tuple's
' values into one key string; it is not a printable character a real
' spreadsheet cell is likely to contain, so a false collision between
' ("a", "b") and ("a" & Chr$(31) & "b") is a theoretical, not a
' practical, risk.
Private Function TupleKey(ByRef t() As Variant) As String
    Dim i As Long, r As String
    For i = LBound(t) To UBound(t)
        r = r & CStr(t(i)) & Chr$(31)
    Next i
    TupleKey = r
End Function

' Same construction over a chosen SUBSET of a tuple's columns (1-based
' positions in cols()) - the join key builder, shared by RelJoin's own
' index build and probe sides so a key built one way always compares
' equal to a key built the other way for the same underlying values.
Private Function PartialKey(ByRef t() As Variant, ByRef cols As Variant) As String
    Dim i As Long, r As String
    For i = LBound(cols) To UBound(cols)
        r = r & CStr(t(cols(i))) & Chr$(31)
    Next i
    PartialKey = r
End Function

Private Function NewIndex() As Object
    On Error Resume Next
    Set NewIndex = CreateObject("Scripting.Dictionary")
    On Error GoTo 0
    ' Nothing on Mac - RelTryAdd's own fallback branch (a linear,
    ' case-sensitive scan over the relation's own tuples) handles that
    ' case directly; no substitute structure is allocated here.
End Function

Public Function RelNew(ByVal arity As Long) As Collection
    Dim rec As New Collection
    rec.Add arity
    rec.Add New Collection      ' tuples
    rec.Add NewIndex()          ' Nothing on Mac
    Set RelNew = rec
End Function

Public Function RelArity(ByVal rel As Collection) As Long
    RelArity = rel.Item(1)
End Function

Public Function RelTuples(ByVal rel As Collection) As Collection
    Set RelTuples = rel.Item(2)
End Function

Public Function RelCount(ByVal rel As Collection) As Long
    RelCount = rel.Item(2).Count
End Function

' A Relation of arity 0 holding exactly one (empty) tuple - the
' identity element for RelJoin: joining anything against RelUnit()
' with no join columns on either side reproduces the other side
' unchanged. VLA_Datalog.bas uses this to fold "the first body atom"
' into the same join-then-project loop as every later one, instead of
' special-casing it.
Public Function RelUnit() As Collection
    Dim rel As Collection
    Set rel = RelNew(0)
    ' Array() (zero-argument) is the reliable way to build a genuinely
    ' empty array in VBA - a manual ReDim straight to inverted bounds
    ' (1 To 0) on a freshly-declared array is not reliable across
    ' every host (live-caught: runtime error 9, "Subscript out of
    ' range", on a real Windows Excel session, at the ReDim itself).
    Dim t() As Variant
    t = Array()
    RelTryAdd rel, t
    Set RelUnit = rel
End Function

' Adds tuple t if no equal tuple (by TupleKey) is already present;
' returns True iff it was new. The one mutator this module has -
' every relation grows only through this, so "was this tuple already
' known" and "add it" can never drift apart (the exact bug class a
' separate Contains-then-Add pair invites).
Public Function RelTryAdd(ByVal rel As Collection, ByRef t() As Variant) As Boolean
    Dim key As String
    key = TupleKey(t)
    Dim idx As Object
    Set idx = rel.Item(3)
    If Not idx Is Nothing Then
        If idx.Exists(key) Then Exit Function
        idx.Add key, True
    Else
        Dim existing As Variant, exArr() As Variant
        For Each existing In RelTuples(rel)
            exArr = existing
            If StrComp(TupleKey(exArr), key, vbBinaryCompare) = 0 Then Exit Function
        Next existing
    End If
    Dim tCopy() As Variant
    tCopy = t                          ' snapshot - VBA array assignment
                                        ' copies the data, so the caller
                                        ' is free to reuse/ReDim t after
                                        ' this call returns.
    RelTuples(rel).Add tCopy
    RelTryAdd = True
End Function

' Exact-tuple membership test - the read-only counterpart to RelTryAdd's
' own "was this tuple already known" half, added for VLA_Datalog's
' stratified negation (`not`): a negated atom is only ever probed once
' every one of its own variables is already bound to a concrete value
' (the safety condition CheckRuleSafety enforces at parse time), so t()
' here is always fully constant - never itself unified or filtered.
Public Function RelContainsTuple(ByVal rel As Collection, ByRef t() As Variant) As Boolean
    Dim key As String
    key = TupleKey(t)
    Dim idx As Object
    Set idx = rel.Item(3)
    If Not idx Is Nothing Then
        RelContainsTuple = idx.Exists(key)
    Else
        Dim existing As Variant, exArr() As Variant
        For Each existing In RelTuples(rel)
            exArr = existing
            If StrComp(TupleKey(exArr), key, vbBinaryCompare) = 0 Then
                RelContainsTuple = True
                Exit Function
            End If
        Next existing
    End If
End Function

' A table argument's own name: an Excel Table's ListObject.Name, or a
' plain Range's own defined name (its sheet-qualifying "Sheet1!" prefix,
' if any, stripped) - never a name the caller chooses in the formula.
' Hoisted here from VLA_Datalog.bas (SQL.1) the moment a second engine
' needed the identical mechanism - this is what makes "an Employees
' table IS a relation with no authoring step" true for every engine that
' reads one, not just DATALOG.
'
' Returns "" and sets ok=False on failure, with reason set to one of
' "not-a-range" / "noncontiguous" / "needs-a-name" - NEVER raises,
' because this module cannot call VLA_Messages (its own LAYER 0.5
' contract, this file's own header) and refusal WORDING is deliberately
' each engine's own job (VLA_Datalog.bas's own TableArgName and
' VLA_Sql.bas's own equivalent both turn a reason code back into their
' OWN already-established wording - datalog-table-not-a-range vs.
' sql-table-not-a-range name the identical failure differently on
' purpose, matching how each engine already owns its own message
' catalogue entries for everything else).
Public Function TableArgResolve(ByVal v As Variant, ByRef ok As Boolean, ByRef reason As String) As String
    ok = False
    reason = ""
    TableArgResolve = ""
    If Not IsObject(v) Then
        reason = "not-a-range"
        Exit Function
    End If
    Dim rng As Object
    Set rng = v
    ' Checked before .ListObject even gets a chance to resolve (a
    ' multi-area Range's own .ListObject read is unpredictable) - a
    ' Ctrl-selected, non-contiguous column pick is refused, not silently
    ' guessed, the same discipline column-scoped table arguments
    ' (SourceToArray, below) is built on.
    If rng.Areas.Count > 1 Then
        reason = "noncontiguous"
        Exit Function
    End If
    Dim lo As Object
    On Error Resume Next
    Set lo = rng.ListObject
    On Error GoTo 0
    If Not lo Is Nothing Then
        TableArgResolve = VLA_Identity.Fold(lo.Name)
        ok = True
        Exit Function
    End If
    Dim nm As String
    On Error Resume Next
    nm = rng.Name.Name
    On Error GoTo 0
    If Len(nm) = 0 Then
        reason = "needs-a-name"
        Exit Function
    End If
    Dim bangPos As Long
    bangPos = InStr(nm, "!")
    If bangPos > 0 Then nm = Mid$(nm, bangPos + 1)
    TableArgResolve = VLA_Identity.Fold(nm)
    ok = True
End Function

' Excel Range -> its own Value2 2D array (ListObject-aware: a Table's
' header row is stripped automatically, arity = ListColumns.Count even
' when DataBodyRange Is Nothing, i.e. a header with zero data rows). A
' plain Range or a plain 2D array carries NO header - every row is a
' fact, positionally (the classical Datalog convention: predicates
' have arity, never column names). A 1-cell range's own Value2 quirk
' (a scalar, not a 1x1 array) is wrapped before use.
'
' DATALOG.3: column-scoped table arguments. When the passed range covers
' FEWER columns than the table's own ListColumns.Count, only those
' columns' own data (still every data row - row-scoping was never part
' of this item's own stated limit, "not one column of it", and stays out
' of scope) is read, narrowing the resulting arity to match - the
' predicate's own NAME is untouched by this (VLA_Datalog.TableArgName's
' own job, reading .ListObject.Name off the very same passed range,
' unaffected by how many of its columns were actually selected). A
' multi-area (Ctrl-selected, non-contiguous) range is refused before
' reaching here (VLA_Datalog.TableArgName's own earlier check) for any
' caller that routes through it; SourceToArray itself has no message
' catalog to raise a named refusal from (VLA_Relation.bas's own LAYER 0 -
' "MAY CALL: (nothing)" - precludes depending on VLA_Messages), so a
' multi-area range reaching this function directly (bypassing
' TableArgName) raises a plain, undressed error instead of silently
' guessing which area to read.
Private Function SourceToArray(ByVal src As Variant) As Variant
    If Not IsObject(src) Then
        If IsArray(src) Then
            SourceToArray = src
        Else
            Dim one(1 To 1, 1 To 1) As Variant
            one(1, 1) = src
            SourceToArray = one
        End If
        Exit Function
    End If
    Dim rng As Object
    Set rng = src
    Dim lo As Object
    On Error Resume Next
    Set lo = rng.ListObject
    On Error GoTo 0
    If Not lo Is Nothing Then
        If rng.Areas.Count > 1 Then
            VLA_Messages.RaiseMsg "relation-table-noncontiguous-areas"
        End If
        Dim tableColCount As Long
        tableColCount = lo.ListColumns.Count
        Dim passedColCount As Long
        passedColCount = rng.Columns.Count
        If passedColCount < tableColCount Then
            Dim colOffset As Long
            colOffset = rng.Column - lo.Range.Column + 1
            If lo.DataBodyRange Is Nothing Then
                Dim emptyN() As Variant
                ReDim emptyN(1 To 1, 1 To passedColCount)
                SourceToArray = emptyN
            Else
                Dim narrowed As Object
                Set narrowed = lo.DataBodyRange.Worksheet.Range( _
                    lo.DataBodyRange.Cells(1, colOffset), _
                    lo.DataBodyRange.Cells(lo.DataBodyRange.Rows.Count, colOffset + passedColCount - 1))
                SourceToArray = ValueOf(narrowed)
            End If
        ElseIf lo.DataBodyRange Is Nothing Then
            ' A Table with headers but zero data rows still needs the
            ' right ARITY reported (lo.ListColumns.Count), but a 2D
            ' array with an empty first dimension hits the identical
            ' inverted-bounds ReDim hazard RelUnit's own header
            ' explains - sidestepped the same way that hazard is
            ' avoided everywhere else in this module: never construct
            ' the empty case directly. One row of default-Empty values
            ' reports the same arity and is skipped by RelFromRange's
            ' own all-blank-row filter, landing at the identical
            ' zero-tuple result through a normal, non-inverted ReDim.
            Dim empty2() As Variant
            ReDim empty2(1 To 1, 1 To lo.ListColumns.Count)
            SourceToArray = empty2
        Else
            SourceToArray = ValueOf(lo.DataBodyRange)
        End If
    Else
        SourceToArray = ValueOf(rng)
    End If
End Function

Private Function ValueOf(ByVal rng As Object) As Variant
    If rng.Cells.Count = 1 Then
        Dim one(1 To 1, 1 To 1) As Variant
        one(1, 1) = rng.Value2
        ValueOf = one
    Else
        ValueOf = rng.Value2
    End If
End Function

' Builds a Relation from src (a live Range, or a plain 2D array shaped
' like one - the seam that makes this fully testable with no live
' workbook: VLA_Tests.TestDatalog is this function's own first caller,
' off a hand-built array, never a live Range). Arity is whatever the
' source's own column count is; a row where every cell is blank is
' skipped (a selection slightly taller than its real data is the
' common case, not a phantom all-empty fact).
'
' DATALOG.5: hasHeader (default False, every existing caller unchanged)
' skips src's own first row as data - purely a PURE-TEST seam for the
' ARRAY path, so TestDSLs can pin keyed-atom desugaring against a
' hand-built array with no live workbook, mirroring what a live Table's
' own header row already gets (stripped automatically via
' SourceToArray's own ListObject/DataBodyRange path, unaffected by this
' flag - never pass hasHeader:=True against a live Table argument, or
' its first DATA row would be skipped a second time). The caller reads
' that same first row's own text independently to build its own
' header-name map (VLA_Relation.RangeColumnNames' own (folded,
' original) pair shape) - no new function needed here, since the test
' already holds the raw array before ever calling this one.
Public Function RelFromRange(ByVal src As Variant, Optional ByVal hasHeader As Boolean = False) As Collection
    Dim arr As Variant
    arr = SourceToArray(src)
    Dim rLo As Long, rHi As Long, cLo As Long, cHi As Long
    rLo = LBound(arr, 1): rHi = UBound(arr, 1)
    cLo = LBound(arr, 2): cHi = UBound(arr, 2)
    If hasHeader And rLo <= rHi Then rLo = rLo + 1
    Dim arity As Long
    arity = cHi - cLo + 1
    Dim rel As Collection
    Set rel = RelNew(arity)
    Dim r As Long, c As Long
    For r = rLo To rHi
        Dim t() As Variant
        ReDim t(1 To arity)
        Dim allBlank As Boolean
        allBlank = True
        For c = cLo To cHi
            t(c - cLo + 1) = arr(r, c)
            If Len(Trim$(CStr(arr(r, c)))) > 0 Then allBlank = False
        Next c
        If Not allBlank Then RelTryAdd rel, t
    Next r
    Set RelFromRange = rel
End Function

' SQL.1: the bag/multiset counterpart to RelFromRange, and NOT
' interchangeable with it - a real, load-bearing distinction, not a
' style choice. A Relation is a SET (RelTryAdd silently absorbs an
' already-seen tuple), correct for Datalog facts (two identical facts
' assert nothing a single one didn't already), but WRONG for a SQL
' table's own rows: a table can legitimately contain two identical
' rows, and SELECTing it must return both, not silently collapse them
' to one. RangeToRows shares SourceToArray's own header-stripping/
' column-narrowing logic (the exact same "an Excel Table's header row
' auto-stripped" mechanism RelFromRange's own header documents) but
' accumulates into a plain Collection - every row kept, in order,
' duplicates included - never a Relation, never RelTryAdd.
Public Function RangeToRows(ByVal src As Variant) As Collection
    Dim arr As Variant
    arr = SourceToArray(src)
    Dim rLo As Long, rHi As Long, cLo As Long, cHi As Long
    rLo = LBound(arr, 1): rHi = UBound(arr, 1)
    cLo = LBound(arr, 2): cHi = UBound(arr, 2)
    Dim arity As Long
    arity = cHi - cLo + 1
    Dim rows As New Collection
    Dim r As Long, c As Long
    For r = rLo To rHi
        Dim t() As Variant
        ReDim t(1 To arity)
        Dim allBlank As Boolean
        allBlank = True
        For c = cLo To cHi
            t(c - cLo + 1) = arr(r, c)
            If Len(Trim$(CStr(arr(r, c)))) > 0 Then allBlank = False
        Next c
        If Not allBlank Then rows.Add t
    Next r
    Set RangeToRows = rows
End Function

' SQL.1's own "new substrate DATALOG never needed" (BETA_ROADMAP2.md's
' own SQL.1 sequencing note): DATALOG is deliberately positional - a
' Relation has arity but no concept of a column NAME - while SQL selects
' and filters BY name. Returns a 1-based Collection of 2-item pairs,
' (foldedName, originalName), positionally matching RangeToRows' own row
' arrays - foldedName (VLA_Identity.Fold, SD-8's law, never a second
' LCase - the neutrality audit's own Turkish-I warning) is what a
' query's own identifiers resolve against; originalName is kept
' separately so a `SELECT *` header can echo the table's own natural
' casing instead of leaking the fold outward. ok=False (and an empty
' Collection) when src isn't backed by a real ListObject - a plain
' range/array has no column names of its own to report, and this
' function has no message catalog to raise a named refusal from (this
' module's own LAYER 0.5 contract); the caller (VLA_Sql.bas) turns that
' into its own sql-table-needs-real-table.
'
' DATALOG.5: narrows to src's own passed column span when it covers
' FEWER columns than the table's own ListColumns.Count - the identical
' colOffset/passedColCount computation SourceToArray's own DATALOG.3
' narrowing already uses, repeated here rather than shared, since the
' two read different things off lo (values vs. names) through different
' Range shapes (DataBodyRange vs. ListColumns). Column-scoped table
' arguments were always DATALOG-only (SQL has no such feature yet), so
' every existing SQL caller always passes the WHOLE table - passedColCount
' then equals tableColCount and colOffset is always 1, identical to this
' function's own pre-DATALOG.5 behavior, unchanged.
Public Function RangeColumnNames(ByVal src As Variant, ByRef ok As Boolean) As Collection
    ok = False
    Dim names As New Collection
    Set RangeColumnNames = names
    If Not IsObject(src) Then Exit Function
    Dim rng As Object
    Set rng = src
    Dim lo As Object
    On Error Resume Next
    Set lo = rng.ListObject
    On Error GoTo 0
    If lo Is Nothing Then Exit Function
    Dim tableColCount As Long
    tableColCount = lo.ListColumns.Count
    Dim passedColCount As Long
    passedColCount = rng.Columns.Count
    Dim colOffset As Long
    If passedColCount < tableColCount Then
        colOffset = rng.Column - lo.Range.Column + 1
    Else
        colOffset = 1
        passedColCount = tableColCount
    End If
    ' `Dim pair As New Collection` INSIDE this loop is a live-caught VBA
    ' trap this codebase already has a name for: `As New` only auto-
    ' instantiates when the variable is currently Nothing, which after
    ' the FIRST iteration it never is again - every later iteration
    ' would silently APPEND onto that SAME shared pair object instead of
    ' starting fresh, so `names` would end up holding three REFERENCES
    ' to one ever-growing Collection rather than three distinct pairs -
    ' live-caught exactly this way: colMap ended up with only ONE usable
    ' key ("name", the first-ever item added), every other column
    ' silently unresolvable. Explicit Set each iteration is the fix,
    ' ParseAtom's own `rec` (VLA_Datalog.bas) is the precedent.
    Dim i As Long
    For i = 1 To passedColCount
        Dim pair As Collection
        Set pair = New Collection
        Dim colName As String
        colName = CStr(lo.ListColumns(colOffset + i - 1).Name)
        pair.Add VLA_Identity.Fold(colName)
        pair.Add colName
        names.Add pair
    Next i
    ok = True
End Function

' idx is either a real Dictionary (bucket-valued) or, on a host with
' no Scripting runtime, a plain Collection of 2-element (key, bucket)
' Collections scanned linearly and case-sensitively (StrComp,
' vbBinaryCompare) - VLA_Runtime.VlaDictSet/Get's own fallback shape,
' NOT reused directly here because that helper's fallback keys on a
' case-INSENSITIVE fold (correct for identifiers, wrong for tuple
' data - this file's own header explains why).
Private Function BuildJoinIndex(ByVal rel As Collection, ByRef cols As Variant) As Object
    Dim idx As Object
    On Error Resume Next
    Set idx = CreateObject("Scripting.Dictionary")
    On Error GoTo 0
    If idx Is Nothing Then Set idx = New Collection
    Dim t As Variant, arr() As Variant
    For Each t In RelTuples(rel)
        arr = t
        JoinIndexPut idx, PartialKey(arr, cols), arr
    Next t
    Set BuildJoinIndex = idx
End Function

Private Sub JoinIndexPut(ByVal idx As Object, ByVal key As String, ByRef arr() As Variant)
    Dim bucket As Collection
    If TypeName(idx) = "Dictionary" Then
        If idx.Exists(key) Then
            Set bucket = idx.Item(key)
        Else
            Set bucket = New Collection
            idx.Add key, bucket
        End If
    Else
        Set bucket = JoinIndexFindBucket(idx, key)
        If bucket Is Nothing Then
            Set bucket = New Collection
            Dim pair As New Collection
            pair.Add key
            pair.Add bucket
            idx.Add pair
        End If
    End If
    bucket.Add arr
End Sub

Private Function JoinIndexFindBucket(ByVal idx As Collection, ByVal key As String) As Collection
    Dim pair As Variant
    For Each pair In idx
        If StrComp(CStr(pair.Item(1)), key, vbBinaryCompare) = 0 Then
            Set JoinIndexFindBucket = pair.Item(2)
            Exit Function
        End If
    Next pair
End Function

Private Function JoinIndexGet(ByVal idx As Object, ByVal key As String) As Collection
    If TypeName(idx) = "Dictionary" Then
        If idx.Exists(key) Then Set JoinIndexGet = idx.Item(key)
    Else
        Set JoinIndexGet = JoinIndexFindBucket(idx, key)
    End If
End Function

' Natural-join building block: every matching (leftTuple, rightTuple)
' pair, concatenated left-then-right (output arity = leftArity +
' rightArity; duplicate join columns are included on both sides -
' projecting them away is the caller's job, the same separation of
' concerns a real query planner keeps between join and project).
' leftCols()/rightCols() are 1-based column POSITIONS, equal length;
' position i on each side must match for a pair to join. Zero-length
' leftCols()/rightCols() means "no shared columns" and correctly
' degenerates to the full cross product (every key collapses to the
' same empty string, so every build-side tuple lands in one bucket
' every probe-side tuple matches) - not a special case, an emergent
' property of the key construction.
'
' Builds its hash index on whichever side has FEWER tuples right now
' (RelCount is O(1)) and probes with the other - "always a hash join
' ... build a Dictionary on the smaller side" the QUERY AND LOGIC
' section's own SQL item requires, generalized here since all four
' engines need the identical operation.
' Returns a Variant (holding an array), never a bare Long() - the
' c.Count = 0 case needs a genuinely empty array, and the reliable way
' to construct one in VBA is the zero-argument Array() function; a
' manual ReDim straight to inverted bounds (1 To 0) on a freshly-
' declared array is not reliable across every host (live-caught,
' VLA_Datalog.bas's own original copy of this function: runtime error
' 9, "Subscript out of range", on a real Windows Excel session that
' never got past this construction). Array() produces a 0-based Variant
' array, which cannot be assigned into a Long()-typed variable directly
' - every caller holds the result in a plain Variant instead; LBound/
' UBound/indexing all work identically on a Variant that holds an
' array, so nothing downstream changes.
'
' SQL.3: hoisted here from VLA_Datalog.bas (its own BuildJoinPlan, the
' identical Collection-of-Longs -> RelJoin's own leftCols()/rightCols()
' shape) the moment a second engine needed it - the same second-consumer
' move as TableArgResolve/RangeColumnNames/CompareValues before it.
' VLA_Datalog.bas's own former Private copy now calls this directly;
' VLA_Sql.bas's own equi-condition split (SplitOnExpr) is this
' function's second real caller.
Public Function CollToLongArray(ByVal c As Collection) As Variant
    If c.Count = 0 Then
        CollToLongArray = Array()
    Else
        Dim r() As Variant
        ReDim r(1 To c.Count)
        Dim i As Long
        For i = 1 To c.Count: r(i) = CLng(c.Item(i)): Next i
        CollToLongArray = r
    End If
End Function

' A Relation-SHAPED wrapper (RelNew's own 3-item record: arity, tuples,
' index) around an ALREADY-BUILT rows Collection, WITHOUT deduping it -
' deliberately never RelTryAdd, since SQL's own rows are a BAG (SQL.1's
' own RangeToRows/RelFromRange header note: two identical source rows
' are two real rows, not one) while RelTryAdd's own membership index
' exists precisely to COLLAPSE duplicates, which would silently drop a
' legitimate join match. Safe specifically because RelJoin (and its own
' BuildJoinIndex helper) reads a Relation ONLY through RelArity/
' RelTuples/RelCount - all three read Item(1)/Item(2) directly and
' never touch Item(3) - so a Nothing index here is never dereferenced;
' Item(3) matters only to RelTryAdd/RelContainsTuple, this wrapper's
' own mutators, which SQL.3 never calls on either side of a join. NOT a
' general-purpose Relation constructor - do not RelTryAdd into a
' Relation built this way (it has no index to update, so RelTryAdd
' would silently fall back to Mac's own O(n) linear-scan path instead
' of crashing, but would ALSO stay a bag, not a set, defeating the
' whole point of RelTryAdd's own dedup contract) - build one with
' RelNew instead for that.
Public Function RelWrapBag(ByVal arity As Long, ByVal rows As Collection) As Collection
    Dim rec As New Collection
    rec.Add arity
    rec.Add rows
    rec.Add Nothing
    Set RelWrapBag = rec
End Function

Public Function RelJoin(ByVal leftRel As Collection, ByRef leftCols As Variant, _
                         ByVal rightRel As Collection, ByRef rightCols As Variant) As Collection
    Dim out As New Collection
    Dim buildOnLeft As Boolean
    buildOnLeft = (RelCount(leftRel) <= RelCount(rightRel))
    Dim buildRel As Collection, probeRel As Collection
    Dim buildCols As Variant, probeCols As Variant
    If buildOnLeft Then
        Set buildRel = leftRel: buildCols = leftCols
        Set probeRel = rightRel: probeCols = rightCols
    Else
        Set buildRel = rightRel: buildCols = rightCols
        Set probeRel = leftRel: probeCols = leftCols
    End If
    Dim idx As Object
    Set idx = BuildJoinIndex(buildRel, buildCols)
    Dim leftArity As Long, rightArity As Long
    leftArity = RelArity(leftRel): rightArity = RelArity(rightRel)
    Dim probeT As Variant, buildT As Variant
    Dim probeArr() As Variant, buildArr() As Variant
    For Each probeT In RelTuples(probeRel)
        probeArr = probeT
        Dim k As String
        k = PartialKey(probeArr, probeCols)
        Dim bucket As Collection
        Set bucket = JoinIndexGet(idx, k)
        If Not bucket Is Nothing Then
            For Each buildT In bucket
                buildArr = buildT
                Dim combined() As Variant
                ReDim combined(1 To leftArity + rightArity)
                Dim i As Long
                If buildOnLeft Then
                    For i = 1 To leftArity: combined(i) = buildArr(i): Next i
                    For i = 1 To rightArity: combined(leftArity + i) = probeArr(i): Next i
                Else
                    For i = 1 To leftArity: combined(i) = probeArr(i): Next i
                    For i = 1 To rightArity: combined(leftArity + i) = buildArr(i): Next i
                End If
                out.Add combined
            Next buildT
        End If
    Next probeT
    Set RelJoin = out
End Function

' A 2D Variant array, header row first (SELECTROWS' own spilled-array
' shape - scripts/spreadsheet.lisp - reused rather than invented
' fresh): headers, if supplied (any 1-based or 0-based array of arity
' strings), name each column; otherwise "Col1".."ColN". A relation with
' zero rows still returns a 1-row, header-only array, never Empty - a
' cell showing only its own column names is legible; a blank cell is
' indistinguishable from an error.
' Stated limit: rel's arity must be >= 1. Zero rows is fine (only the
' first array dimension shrinks, never inverted); zero COLUMNS is not
' handled - it would need this function's own hdr/out arrays built
' with an empty second dimension, the identical inverted-bounds ReDim
' hazard RelUnit's own header documents, live-caught elsewhere in this
' file. Not hardened here because nothing reaches it at arity 0 today:
' VLA_Datalog.bas's own ParseAtom refuses a zero-argument predicate
' before a Relation of that arity could ever exist. Revisit if a
' future caller (SQL's own aggregate-to-scalar case, perhaps) ever
' legitimately needs one.
'
' headless skips the header row entirely (headers is then ignored) so
' the CALLER can type its own column names into the cell above the
' spill range, DATALOG's own (headless) directive. A zero-row headless
' result returns "" (a blank-looking cell) rather than a 2D array with
' an empty first dimension - the same inverted-bounds ReDim hazard as
' arity 0, sidestepped the same way: never construct the empty case
' directly.
Public Function RelToSpilledArray(ByVal rel As Collection, Optional ByVal headers As Variant, Optional ByVal headless As Boolean = False) As Variant
    Dim arity As Long
    arity = RelArity(rel)
    Dim n As Long
    n = RelCount(rel)

    If headless Then
        If n = 0 Then
            RelToSpilledArray = ""
            Exit Function
        End If
        Dim outD() As Variant
        ReDim outD(1 To n, 1 To arity)
        Dim rD As Long
        rD = 0
        Dim tD As Variant, arrD() As Variant
        Dim iD As Long
        For Each tD In RelTuples(rel)
            arrD = tD
            rD = rD + 1
            For iD = 1 To arity
                outD(rD, iD) = arrD(iD)
            Next iD
        Next tD
        RelToSpilledArray = outD
        Exit Function
    End If

    Dim hdr() As String
    ReDim hdr(1 To arity)
    Dim i As Long
    If Not IsMissing(headers) Then
        For i = 1 To arity
            hdr(i) = CStr(headers(LBound(headers) + i - 1))
        Next i
    Else
        For i = 1 To arity
            hdr(i) = "Col" & i
        Next i
    End If
    Dim out() As Variant
    ReDim out(1 To n + 1, 1 To arity)
    For i = 1 To arity
        out(1, i) = hdr(i)
    Next i
    Dim r As Long
    r = 1
    Dim t As Variant, arr() As Variant
    For Each t In RelTuples(rel)
        arr = t
        r = r + 1
        For i = 1 To arity
            out(r, i) = arr(i)
        Next i
    Next t
    RelToSpilledArray = out
End Function

' =====================================================================
'  DATALOG.4's own shared value-comparison/arithmetic primitive - the
'  ONE thing SQL.1's own scalar evaluator (VLA_Sql.EvalScalar/EvalBool)
'  and DATALOG's new (> X 50000)/(let Z (+ X Y)) built-ins both
'  genuinely need identically: a mechanical "given two already-resolved
'  values and an operator, compare or compute them." What is
'  DELIBERATELY NOT shared: deciding whether a value counts as numeric
'  in the first place - this module's own header above has the full
'  reasoning; every function below that needs that verdict takes it as
'  a bothNumeric argument rather than deciding it internally.
' =====================================================================

' Whether v's own VBA type is one of the REAL Excel Value2 numeric
' kinds a live cell read can produce (Integer/Long/Single/Double/
' Currency/Date/Byte) - SQL.1's own former Private IsNumericValue,
' hoisted here unchanged: a text cell that merely LOOKS numeric ("007")
' must never be silently promoted into a numeric comparison the
' column's own typing didn't ask for. Deliberately NOT IsNumeric(v) -
' IsInvariantNumericString's own header below has the reason that VBA
' builtin is never used anywhere in this section.
Public Function ValueIsNumericType(ByVal v As Variant) As Boolean
    Select Case VarType(v)
    Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDate, vbByte
        ValueIsNumericType = True
    End Select
End Function

' Whether s, taken as a WHOLE string, is a valid locale-invariant
' number: an optional leading sign, at least one digit, at most one '.'
' (always the decimal separator regardless of the caller's own Excel
' locale - this section's own SD-4 freeze, "decided now rather than
' inherited by accident"). Deliberately NOT VBA's own IsNumeric() - that
' builtin is locale-AWARE (a comma-decimal Excel install would accept
' "50000,5" as numeric), which would silently reintroduce exactly the
' locale trap InvariantVal/Val() exists to avoid. Also deliberately NOT
' a bare "InvariantVal(s) <> 0" check - Val() parses only a LEADING
' numeric prefix and silently ignores trailing junk ("50000abc" ->
' 50000), so every character of s must be accounted for here, not just
' its first few.
Public Function IsInvariantNumericString(ByVal s As String) As Boolean
    Dim n As Long
    n = Len(s)
    If n = 0 Then Exit Function
    Dim i As Long
    i = 1
    Dim c As String
    c = Mid$(s, i, 1)
    If c = "-" Or c = "+" Then i = i + 1
    If i > n Then Exit Function
    Dim sawDigit As Boolean, sawDot As Boolean
    Do While i <= n
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
    IsInvariantNumericString = sawDigit
End Function

' Val() is genuinely locale-invariant in VBA - always '.' as the
' decimal separator, unlike CDbl/CSng, which respect the user's own
' regional settings. VLA_Sql.bas's own former Private ParseInvariantNumber,
' hoisted here unchanged the moment a second engine needed identical
' text -> number parsing.
Public Function InvariantVal(ByVal s As String) As Double
    InvariantVal = Val(s)
End Function

' A value already known/decided to be numeric (bothNumeric at the call
' site, below) as a Double: a real Excel numeric type widens/narrows
' via CDbl with no string parsing involved at all; anything else (a
' numeric-LOOKING string, DATALOG's own lenient half of bothNumeric) is
' read through InvariantVal's own locale-invariant text -> number path
' instead. Private - only ComputeArithmetic needs a value AS a Double;
' CompareValues compares two already-typed values directly and never
' needs this conversion.
Private Function AsInvariantDouble(ByVal v As Variant) As Double
    If ValueIsNumericType(v) Then
        AsInvariantDouble = CDbl(v)
    Else
        AsInvariantDouble = InvariantVal(CStr(v))
    End If
End Function

' Compares l and r under one shared six-operator vocabulary (=, <>, <,
' <=, >, >=) - numerically (AsInvariantDouble on both) if bothNumeric is
' True, else as case-SENSITIVE text (CStr, StrComp, vbBinaryCompare - a
' Relation's own tuple-identity discipline, this file's own header,
' applied here to a comparison operand instead of a whole tuple).
' VLA_Sql.bas's own former Private CompareValues, byte-for-byte,
' parameterized on bothNumeric instead of recomputing IsNumericValue
' internally - SQL.1's own caller passes ValueIsNumericType(l) And
' ValueIsNumericType(r) (its own STRICT policy, unchanged); VLA_Datalog's
' own caller passes its own, more lenient verdict instead (this file's
' own header note above has the reasoning).
Public Function CompareValues(ByVal op As String, ByVal l As Variant, ByVal r As Variant, ByVal bothNumeric As Boolean) As Boolean
    If bothNumeric Then
        Dim ln As Double, rn As Double
        ln = AsInvariantDouble(l): rn = AsInvariantDouble(r)
        Select Case op
        Case "=": CompareValues = (ln = rn)
        Case "<>": CompareValues = (ln <> rn)
        Case "<": CompareValues = (ln < rn)
        Case "<=": CompareValues = (ln <= rn)
        Case ">": CompareValues = (ln > rn)
        Case ">=": CompareValues = (ln >= rn)
        End Select
    Else
        Dim c As Long
        c = StrComp(CStr(l), CStr(r), vbBinaryCompare)
        Select Case op
        Case "=": CompareValues = (c = 0)
        Case "<>": CompareValues = (c <> 0)
        Case "<": CompareValues = (c < 0)
        Case "<=": CompareValues = (c <= 0)
        Case ">": CompareValues = (c > 0)
        Case ">=": CompareValues = (c >= 0)
        End Select
    End If
End Function

' The mechanical half of DATALOG.4's own (let Z (op X Y)) - the four
' basic arithmetic operators (+, -, *, /), given two operands ALREADY
' decided to both be numeric (bothNumeric, the identical caller-decided
' verdict CompareValues above takes). Never raises (this module's own
' LAYER 0.5 contract, TableArgResolve's own header - refusal WORDING
' stays each engine's own job): ok=False and reason "not-numeric" when
' bothNumeric is False, or "divide-by-zero" for op="/" when r evaluates
' to exactly 0 - never a silent 0, never VBA's own raw runtime
' "Division by zero" error leaking out undressed.
Public Function ComputeArithmetic(ByVal op As String, ByVal l As Variant, ByVal r As Variant, ByVal bothNumeric As Boolean, ByRef ok As Boolean, ByRef reason As String) As Variant
    ok = False
    reason = ""
    If Not bothNumeric Then
        reason = "not-numeric"
        Exit Function
    End If
    Dim ln As Double, rn As Double
    ln = AsInvariantDouble(l): rn = AsInvariantDouble(r)
    Select Case op
    Case "+": ComputeArithmetic = ln + rn
    Case "-": ComputeArithmetic = ln - rn
    Case "*": ComputeArithmetic = ln * rn
    Case "/"
        If rn = 0 Then
            reason = "divide-by-zero"
            Exit Function
        End If
        ComputeArithmetic = ln / rn
    End Select
    ok = True
End Function

' =====================================================================
'  SQL.4's own shared grouping kernel - see this file's own header note
'  above for the full design (what generalizes from DATALOG.2's own
'  grouping mechanism, what deliberately doesn't, and the MIN/MAX/AVG-
'  over-zero-rows policy). The five AGG_* kind constants themselves live
'  in this module's own declarations section, near the top (VBA's own
'  "Const after a procedure" restriction, this file's own header note
'  there has the reason).
' =====================================================================

' One aggregate output column's own spec: kind, plus WHICH position in
' each already-built "eval row" (VLA_Sql.bas's own construction, below)
' holds the value to accumulate - unused (pass 0) for AGG_COUNT, which
' never reads a value column at all (this file's own header note has
' the reason).
Public Function MakeAggSpec(ByVal kind As Long, ByVal valuePos As Long) As Collection
    Dim s As New Collection
    s.Add kind
    s.Add valuePos
    Set MakeAggSpec = s
End Function

Private Function AggSpecKind(ByVal s As Collection) As Long
    AggSpecKind = s.Item(1)
End Function

Private Function AggSpecValuePos(ByVal s As Collection) As Long
    AggSpecValuePos = s.Item(2)
End Function

' DATALOG.2's own KeyFromPositions, byte-for-byte, generalized only in
' what "positions" means: an already-resolved Long array (RelJoin's own
' leftCols()/rightCols() shape, built via CollToLongArray) instead of a
' Collection of atom positions paired against a colOf dictionary - the
' caller has already done all the resolving DATALOG's own version did
' inline.
Private Function GroupKeyFromPositions(ByRef arr() As Variant, ByVal positions As Variant) As String
    Dim r As String
    Dim i As Long
    For i = LBound(positions) To UBound(positions)
        r = r & CStr(arr(CLng(positions(i)))) & Chr$(31)
    Next i
    GroupKeyFromPositions = r
End Function

' Folds one row's own aggregate-operand value into acc(gk)'s own running
' state. AGG_SUM/AGG_AVG share one running Double total (AVG divides by
' the group's own row count only at output time, RelGroupBy's own job) -
' ok=False/reason="not-numeric" the moment a non-numeric value reaches
' either, this engine's own strict SUM/AVG policy (never DATALOG's more
' lenient numeric-looking-string one - SQL's values are always real
' Excel Value2 reads). AGG_MIN/AGG_MAX keep a running "best so far" via
' CompareValues itself, recomputing bothNumeric per comparison exactly
' as EvalBool's own NK_CMP case does - never refuses, degrading to a
' text comparison when the values aren't both numeric.
Private Sub AccumulateAggValue(ByVal kind As Long, ByVal acc As Object, ByVal gk As String, _
                                ByVal v As Variant, ByRef ok As Boolean, ByRef reason As String)
    Select Case kind
    Case AGG_SUM, AGG_AVG
        If Not ValueIsNumericType(v) Then
            ok = False
            reason = "not-numeric"
            Exit Sub
        End If
        Dim addend As Double
        addend = CDbl(v)
        If VLA_Runtime.VlaDictHas(acc, gk) Then
            VLA_Runtime.VlaDictSet acc, gk, CDbl(VLA_Runtime.VlaDictGet(acc, gk)) + addend
        Else
            VLA_Runtime.VlaDictSet acc, gk, addend
        End If
    Case AGG_MIN, AGG_MAX
        If VLA_Runtime.VlaDictHas(acc, gk) Then
            Dim cur As Variant
            cur = VLA_Runtime.VlaDictGet(acc, gk)
            Dim bothNum As Boolean
            bothNum = ValueIsNumericType(v) And ValueIsNumericType(cur)
            Dim takeNew As Boolean
            If kind = AGG_MIN Then
                takeNew = CompareValues("<", v, cur, bothNum)
            Else
                takeNew = CompareValues(">", v, cur, bothNum)
            End If
            If takeNew Then VLA_Runtime.VlaDictSet acc, gk, v
        Else
            VLA_Runtime.VlaDictSet acc, gk, v
        End If
    End Select
End Sub

' The kernel itself: groups rows (a Collection of already-evaluated
' Variant() "eval rows" - VLA_Sql.bas's own ComputeGroupedRows builds
' these, one per source row that survived WHERE, each holding its own
' GROUP BY key values followed by whatever aggregate operand values are
' needed) by keyPositions (a Long array, possibly zero-length - an empty
' keyPositions means every row belongs to the SAME one group), computing
' aggSpecs in order, and returns one row per DISTINCT key: the key's own
' values (in keyPositions order) followed by each aggSpec's own result,
' in aggSpecs order.
'
' alwaysOneGroup forces exactly one output group to exist even when rows
' is entirely empty - SQL's own "an ungrouped aggregate is a real answer
' even over zero matching rows" stance (this file's own header note);
' pass this True only when keyPositions is itself empty (no real GROUP
' BY at all) AND the caller actually has an aggregate to report - a real
' GROUP BY with zero matching rows correctly produces zero groups either
' way, alwaysOneGroup or not, since keyPositions being non-empty is what
' makes a group's own existence conditional on a real row producing it.
'
' ok/reason/failSpecIndex: ok=False on either of the two named aggregate
' failures (this file's own header note) - reason is "not-numeric" or
' "empty-aggregate", failSpecIndex the 1-based aggSpecs position (and,
' by construction, VLA_Sql.bas's own parallel aggList position) that
' failed, so the caller can name WHICH function in its own refusal
' wording. Never raises itself (LAYER 0.5, TableArgResolve's own header).
Public Function RelGroupBy(ByVal rows As Collection, ByVal keyPositions As Variant, _
                            ByVal aggSpecs As Collection, ByVal alwaysOneGroup As Boolean, _
                            ByRef ok As Boolean, ByRef reason As String, ByRef failSpecIndex As Long) As Collection
    ok = True
    reason = ""
    failSpecIndex = 0

    Dim keyArity As Long
    keyArity = UBound(keyPositions) - LBound(keyPositions) + 1

    Dim groupOrder As New Collection
    Dim groupSeen As Object
    Set groupSeen = VLA_Runtime.VlaDictNew()
    Dim keyVals As Object
    Set keyVals = VLA_Runtime.VlaDictNew()
    Dim countAcc As Object
    Set countAcc = VLA_Runtime.VlaDictNew()

    Dim nAgg As Long
    nAgg = aggSpecs.Count
    Dim aggAcc() As Object
    If nAgg > 0 Then ReDim aggAcc(1 To nAgg)
    Dim ai As Long
    For ai = 1 To nAgg
        Set aggAcc(ai) = VLA_Runtime.VlaDictNew()
    Next ai

    Dim t As Variant, arr() As Variant
    Dim gk As String
    For Each t In rows
        arr = t
        gk = GroupKeyFromPositions(arr, keyPositions)
        If Not VLA_Runtime.VlaDictHas(groupSeen, gk) Then
            groupOrder.Add gk
            VLA_Runtime.VlaDictSet groupSeen, gk, True
            Dim kv() As Variant
            If keyArity > 0 Then
                ReDim kv(1 To keyArity)
                Dim ki As Long
                For ki = 1 To keyArity
                    kv(ki) = arr(CLng(keyPositions(LBound(keyPositions) + ki - 1)))
                Next ki
            End If
            VLA_Runtime.VlaDictSet keyVals, gk, kv
            VLA_Runtime.VlaDictSet countAcc, gk, 0&
        End If
        VLA_Runtime.VlaDictSet countAcc, gk, CLng(VLA_Runtime.VlaDictGet(countAcc, gk)) + 1

        For ai = 1 To nAgg
            Dim spec As Collection
            Set spec = aggSpecs.Item(ai)
            If AggSpecKind(spec) <> AGG_COUNT Then
                Dim vOk As Boolean, vReason As String
                vOk = True
                vReason = ""
                AccumulateAggValue AggSpecKind(spec), aggAcc(ai), gk, arr(AggSpecValuePos(spec)), vOk, vReason
                If Not vOk Then
                    ok = False
                    reason = vReason
                    failSpecIndex = ai
                    Set RelGroupBy = New Collection
                    Exit Function
                End If
            End If
        Next ai
    Next t

    If alwaysOneGroup And groupOrder.Count = 0 Then
        Dim emptyKey As String
        emptyKey = ""
        groupOrder.Add emptyKey
        Dim kv0() As Variant
        If keyArity > 0 Then ReDim kv0(1 To keyArity)
        VLA_Runtime.VlaDictSet keyVals, emptyKey, kv0
        VLA_Runtime.VlaDictSet countAcc, emptyKey, 0&
    End If

    Dim outp As New Collection
    Dim outArity As Long
    outArity = keyArity + nAgg
    Dim gi As Long
    For gi = 1 To groupOrder.Count
        Dim gkey As String
        gkey = groupOrder.Item(gi)
        Dim outRow() As Variant
        ReDim outRow(1 To outArity)
        Dim kvArr As Variant
        kvArr = VLA_Runtime.VlaDictGet(keyVals, gkey)
        Dim p As Long
        For p = 1 To keyArity
            outRow(p) = kvArr(p)
        Next p
        Dim rowCount As Long
        rowCount = CLng(VLA_Runtime.VlaDictGet(countAcc, gkey))
        For ai = 1 To nAgg
            Dim spec2 As Collection
            Set spec2 = aggSpecs.Item(ai)
            Dim k2 As Long
            k2 = AggSpecKind(spec2)
            Dim result As Variant
            Select Case k2
            Case AGG_COUNT
                result = rowCount
            Case AGG_SUM
                If VLA_Runtime.VlaDictHas(aggAcc(ai), gkey) Then
                    result = VLA_Runtime.VlaDictGet(aggAcc(ai), gkey)
                Else
                    result = 0#
                End If
            Case AGG_MIN, AGG_MAX
                If VLA_Runtime.VlaDictHas(aggAcc(ai), gkey) Then
                    result = VLA_Runtime.VlaDictGet(aggAcc(ai), gkey)
                Else
                    ok = False
                    reason = "empty-aggregate"
                    failSpecIndex = ai
                    Set RelGroupBy = New Collection
                    Exit Function
                End If
            Case AGG_AVG
                If VLA_Runtime.VlaDictHas(aggAcc(ai), gkey) Then
                    result = CDbl(VLA_Runtime.VlaDictGet(aggAcc(ai), gkey)) / rowCount
                Else
                    ok = False
                    reason = "empty-aggregate"
                    failSpecIndex = ai
                    Set RelGroupBy = New Collection
                    Exit Function
                End If
            End Select
            outRow(keyArity + ai) = result
        Next ai
        outp.Add outRow
    Next gi
    Set RelGroupBy = outp
End Function
