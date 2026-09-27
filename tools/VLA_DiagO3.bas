Attribute VB_Name = "VLA_DiagO3"
Option Explicit

' Standalone diagnostic for OPTIMIZE.3 - NOT part of the VLA project. Import
' it into the dev workbook (VLA.xlsm), type one line in the Immediate window,
' and delete the module afterwards; nothing in Frazaro calls it. The steps are
' in archive/optimize3_live_steps.md.
'
' SLICE 1 - THE INTEGER GROUNDER, TIMED. O3Ground runs OPTIMIZE.0's grounding
' shapes two ways over the SAME rule and the SAME in-memory relations:
'   - DATALOG's own evaluator, VLA_Datalog.DatalogRunForms (the fixpoint,
'     round two included), which is what the pre-flight timed through a
'     worksheet formula; and
'   - the integer grounder, VLA_Datalog.DatalogGroundRules (one pass, Long
'     arrays), which is what OPTIMIZE.3 grounds with.
' Both are timed as VBA calls, with no worksheet written and no formula
' recalculated, so the ratio is the two evaluators' alone. Each rung checks
' that both produced the count the fixture's closed form says, that the two
' agree row for row IN ORDER (the count and an order-sensitive hash of every
' cell's spelling), and that the integer grounder really took the integer
' path. A wrong or disagreeing rung stops the step and is never timed again.
'
' THE FIXTURE is OPTIMIZE.0's (docs/OPTIMIZATION.md, Entry 1 section 3a):
' shifts s = 1..21W, people p = 1..P, every (shift, person) pair eligible
' (the O0 harness's PoolRows with no leave pruned), Next(s, s+1), and a
' person-week's shifts carrying their week. Ids are Doubles, as a worksheet
' hands them over.
'
' Results print as O3G| lines. The pre-flight's own DATALOG times for the
' same rungs, through a worksheet formula, are in the steps file for
' comparison; this harness re-times DATALOG in the same session rather than
' trusting them.
'
' SLICE 2 - THE SEARCH, TIMED. O3Search times two things, and prints O3S|
' lines:
'   - the search alone (VLA_OptimizeSearch.OptSearchRun) on the TWO-SHIFT
'     FAMILY - n people, two shifts, exactly k on each, nobody on both, with
'     2k > n. It has no schedule, but not by counting (one choice form and
'     no cap), so the search must try every way; built as the search's own
'     problem, with no grounding, no guard and a budget far past its need.
'   - the whole run (VLA_Optimize.OptimizeRun) on the same family written as
'     an OPTIMIZE program, and on OPTIMIZE's toy scaled to 20 people and 21
'     shifts, four a shift - grounding and search together.
' Every line checks its outcome and its decisions and dead ends against the
' counts a line-for-line transliteration of the search computed before the
' pass, and says WRONG rather than print a time for a different search.
'
' SLICE 3 - THE CEILINGS, TIMED. O3Ceiling prints O3C| lines: the integer
' grounder alone (VLA_Datalog.DatalogGroundRules) on the two shapes a
' formula's ceilings bound, each sized to lay out just under 100,000 rows -
' a member rule pairing 250 items with 198 slots (pass 2's shape, its
' largest step 49,500) and the roster's "never two in a row", planned, at 60
' people and 553 shifts (pass 3's shape); then that rule in its own written
' order, refused, which times the first step and the COUNTING of the step
' it refuses; and a whole OPTIMIZE run at the ceiling, grounding and search
' together. Every line checks its rows against the counts a transliteration
' of the grounder computed before the pass, and says WRONG rather than time
' a different grounding.
'
' SLICE 4 - THE LADDER. O3Ladder prints O3L| lines: OPTIMIZE.0's reference
' roster (docs/OPTIMIZATION.md, Entry 1 section 3a - tight needs, one day's
' leave a person a week, a senior on every night, at most five a week, never
' two in a row) at 5, 10, 20 and 50 people over 1 and 4 weeks, written as
' an OPTIMIZE program over its four Tables and run whole by OptimizeRun:
' DATALOG's pass, the grounding, the search and the answer. Every rung
' checks its counts against a transliteration computed before the pass, and
' every schedule is checked against the fixture's own definition by this
' module, not by OPTIMIZE; a wrong rung stops the ladder and is never timed.
' Then the calibration: the same roster at 50 x 4 with ELEVEN a shift, where
' the transliteration says every effort level runs out, timed at (effort 1),
' normal and thorough - a unit of search work at a real roster's widths,
' which is what the effort levels are set from - and where the reference
' roster's own run spends its time: the memo key, DATALOG's pass over its
' rules, and OPTIMIZE's own work.

Private Const GUARD_SECONDS As Double = 15#

' step 6: never two in a row, through the pair relation Next (read-bound).
' step 8: at most k per shift as canonical (k+1)-subsets (produce-bound).
' step 10: at most k a week as (k+1)-subsets of a person-week (produce-bound).
Public Sub O3Ground(ByVal stepNo As Long, Optional ByVal skipDatalog As Boolean = False)
    Dim rungs As Variant
    Select Case stepNo
    Case 6
        rungs = Split("10,1 20,4 50,4", " ")
    Case 8
        rungs = Split("10,1,2 14,1,2 20,1,2 10,4,2 14,4,2 20,4,2", " ")
    Case 10
        rungs = Split("5,1,2 10,1,2 5,1,3 10,1,3", " ")
    Case Else
        Debug.Print "O3G: the steps are 6, 8 and 10."
        Exit Sub
    End Select
    Debug.Print "=== O3G step " & stepNo & ": " & ShapeText(stepNo) & " ==="
    Debug.Print "    Excel " & Application.Version & ", " & Format$(Now, "yyyy-mm-dd hh:nn")
    Debug.Print "O3G|step|people|weeks|k|expected rows|datalog s|integer s|datalog/integer|datalog rows|integer rows|same rows, same order|path|verdict"

    Dim ri As Long
    Dim stopAt As String
    For ri = LBound(rungs) To UBound(rungs)
        Dim bits As Variant
        bits = Split(CStr(rungs(ri)), ",")
        Dim nP As Long, nW As Long, k As Long
        nP = CLng(bits(0))
        nW = CLng(bits(1))
        k = 2
        If UBound(bits) >= 2 Then k = CLng(bits(2))
        If Len(stopAt) > 0 Then
            Debug.Print "O3G|" & stepNo & "|" & nP & "|" & nW & "|" & k & "|||||||||not run: " & stopAt
        Else
            Dim verdict As String
            verdict = RunRung(stepNo, nP, nW, k, skipDatalog)
            If Left$(verdict, 2) <> "ok" Then stopAt = "an earlier rung said " & verdict
        End If
    Next ri
    Debug.Print "=== O3G step " & stepNo & " done ==="
End Sub

Private Function RunRung(ByVal stepNo As Long, ByVal nP As Long, ByVal nW As Long, ByVal k As Long, _
                         ByVal skipDatalog As Boolean) As String
    Dim expected As Double
    expected = ExpectedRows(stepNo, nP, nW, k)
    Dim ruleText As String
    ruleText = RuleOf(stepNo, k)

    ' --- DATALOG's own evaluator, on its own copy of the relations ---------
    Dim tD As Double, dRows As Long, dHash As Long
    tD = -1
    dRows = -1
    If Not skipDatalog Then
        Dim dRels As Object
        Set dRels = FixtureRelations(stepNo, nP, nW)
        Dim t0 As Double
        t0 = Timer
        Dim dRes As Collection
        Set dRes = VLA_Datalog.DatalogRunForms(VLA.VlaReadForms(ruleText & " (query o3g)"), dRels, Nothing)
        tD = SecondsSince(t0)
        Dim dRel As Collection
        Set dRel = VLA_Runtime.VlaDictGet(dRes.Item(2), "o3g")
        dRows = VLA_Relation.RelCount(dRel)
        dHash = HashOfRelation(dRel)
    End If

    ' --- the integer grounder, on a fresh copy ------------------------------
    Dim gRels As Object
    Set gRels = FixtureRelations(stepNo, nP, nW)
    Dim syms As VlaSymbols
    VLA_Relation.VlaSymInit syms
    Dim t1 As Double
    t1 = Timer
    Dim gRes As Collection
    Set gRes = VLA_Datalog.DatalogGroundRules(VLA.VlaReadForms(ruleText), gRels, Nothing, syms)
    Dim tG As Double
    tG = SecondsSince(t1)
    Dim path As String
    If VLA_Datalog.DatalogGroundIntegerRules() = 1 Then path = "integer" Else path = "EvalRuleBody"
    Dim g As Variant
    g = gRes.Item(1)
    Dim gRows As Long
    gRows = CLng(g(2))
    Dim gIds() As Long
    gIds = g(3)
    Dim gHash As Long
    gHash = HashOfIds(gIds, gRows, CLng(g(1)), syms)

    Dim same As String
    If skipDatalog Then
        same = "not compared"
    ElseIf dRows = gRows And dHash = gHash Then
        same = "yes"
    Else
        same = "NO"
    End If
    Dim verdict As String
    If gRows <> expected Then
        verdict = "WRONG: the integer grounder made " & gRows & " rows where " & Format$(expected, "#,##0") & " were expected"
    ElseIf Not skipDatalog And dRows <> expected Then
        verdict = "WRONG: DATALOG made " & dRows & " rows where " & Format$(expected, "#,##0") & " were expected"
    ElseIf same = "NO" Then
        verdict = "DISAGREE: the same count, but not the same rows in the same order"
    ElseIf path <> "integer" Then
        verdict = "FELL BACK: the rule went to EvalRuleBody, so the integer path was not timed"
    Else
        verdict = "ok"
    End If
    Dim ratioText As String
    If tD >= 0 And tG > 0 Then
        ratioText = Format$(tD / tG, "0.0") & "x"
    ElseIf tD >= 0 Then
        ratioText = "(integer under the timer's floor)"
    End If
    Debug.Print "O3G|" & stepNo & "|" & nP & "|" & nW & "|" & k & "|" & Format$(expected, "#,##0") & "|" & _
                SecsText(tD) & "|" & SecsText(tG) & "|" & ratioText & "|" & CountText(dRows) & "|" & _
                Format$(gRows, "#,##0") & "|" & same & "|" & path & "|" & verdict
    If tD > GUARD_SECONDS Then
        RunRung = "DATALOG took " & Format$(tD, "0.0") & "s, over the " & GUARD_SECONDS & "s guard"
    Else
        RunRung = verdict
    End If
End Function

' ---------------------------------------------------------------------------
' The fixture, as DATALOG relations built directly - no worksheet.
' ---------------------------------------------------------------------------

Private Function FixtureRelations(ByVal stepNo As Long, ByVal nP As Long, ByVal nW As Long) As Object
    Dim rels As Object
    Set rels = VLA_Runtime.VlaDictNew()
    Dim nS As Long
    nS = 21 * nW
    Dim s As Long, p As Long
    Dim t() As Variant
    Select Case stepNo
    Case 6
        Dim nextRel As Collection
        Set nextRel = VLA_Relation.RelNew(2)
        ReDim t(1 To 2)
        For s = 1 To nS - 1
            t(1) = CDbl(s)
            t(2) = CDbl(s + 1)
            VLA_Relation.RelTryAdd nextRel, t
        Next s
        VLA_Runtime.VlaDictSet rels, "nexto3", nextRel
        VLA_Runtime.VlaDictSet rels, "eligo3", PoolRelation(nS, nP, False)
    Case 8
        VLA_Runtime.VlaDictSet rels, "eligo3", PoolRelation(nS, nP, False)
    Case 10
        VLA_Runtime.VlaDictSet rels, "eligo3", PoolRelation(nS, nP, True)
    End Select
    Set FixtureRelations = rels
End Function

' Every (shift, person) pair, shift-major, as the O0 harness's PoolRows
' writes them with no leave pruned; with the week as a third column when the
' shape groups by person-week.
Private Function PoolRelation(ByVal nS As Long, ByVal nP As Long, ByVal withWeek As Boolean) As Collection
    Dim rel As Collection
    Dim t() As Variant
    If withWeek Then
        Set rel = VLA_Relation.RelNew(3)
        ReDim t(1 To 3)
    Else
        Set rel = VLA_Relation.RelNew(2)
        ReDim t(1 To 2)
    End If
    Dim s As Long, p As Long
    For s = 1 To nS
        For p = 1 To nP
            t(1) = CDbl(s)
            t(2) = CDbl(p)
            If withWeek Then t(3) = CDbl(((s - 1) \ 3) \ 7 + 1)
            VLA_Relation.RelTryAdd rel, t
        Next p
    Next s
    Set PoolRelation = rel
End Function

Private Function RuleOf(ByVal stepNo As Long, ByVal k As Long) As String
    Select Case stepNo
    Case 6
        RuleOf = "(rule (o3g A B P) (nexto3 A B) (eligo3 A P) (eligo3 B P))"
    Case 8
        RuleOf = SubsetRule("(o3g S", "S #", k)
    Case 10
        RuleOf = SubsetRule("(o3g P W", "# P W", k)
    End Select
End Function

' OPTIMIZE.0's SubsetRule: k+1 members in canonical order, each join followed
' at once by the comparison it makes possible.
Private Function SubsetRule(ByVal headStart As String, ByVal argPattern As String, ByVal k As Long) As String
    Dim t As String
    Dim j As Long
    t = "(rule " & headStart
    For j = 1 To k + 1
        t = t & " X" & j
    Next j
    t = t & ")"
    For j = 1 To k + 1
        t = t & " (eligo3 " & Replace(argPattern, "#", "X" & j) & ")"
        If j > 1 Then t = t & " (< X" & (j - 1) & " X" & j & ")"
    Next j
    SubsetRule = t & ")"
End Function

Private Function ExpectedRows(ByVal stepNo As Long, ByVal nP As Long, ByVal nW As Long, ByVal k As Long) As Double
    Dim nS As Double
    nS = 21# * nW
    Select Case stepNo
    Case 6
        ExpectedRows = (nS - 1) * nP
    Case 8
        ExpectedRows = nS * Comb(nP, k + 1)
    Case 10
        ExpectedRows = CDbl(nP) * nW * Comb(21, k + 1)
    End Select
End Function

Private Function ShapeText(ByVal stepNo As Long) As String
    Select Case stepNo
    Case 6: ShapeText = "never two in a row, through the pair relation (read-bound)"
    Case 8: ShapeText = "at most k per shift, as canonical (k+1)-subsets (produce-bound)"
    Case 10: ShapeText = "at most k a week, as (k+1)-subsets of a person-week (produce-bound)"
    End Select
End Function

Private Function Comb(ByVal n As Long, ByVal r As Long) As Double
    Dim i As Long, v As Double
    If r < 0 Or r > n Then Exit Function
    v = 1#
    For i = 1 To r
        v = v * (n - r + i) / i
    Next i
    Comb = Int(v + 0.5)
End Function

' ---------------------------------------------------------------------------
' Comparing the two answers: the same rows in the same order, as spellings.
' ---------------------------------------------------------------------------

Private Function HashOfRelation(ByVal rel As Collection) As Long
    Dim h As Long
    h = 5381
    Dim t As Variant, arr() As Variant
    Dim c As Long
    For Each t In VLA_Relation.RelTuples(rel)
        arr = t
        For c = LBound(arr) To UBound(arr)
            h = MixText(h, CStr(arr(c)))
        Next c
    Next t
    HashOfRelation = h
End Function

Private Function HashOfIds(ByRef ids() As Long, ByVal nRows As Long, ByVal arity As Long, ByRef syms As VlaSymbols) As Long
    Dim h As Long
    h = 5381
    Dim r As Long, c As Long
    For r = 0 To nRows - 1
        For c = 1 To arity
            h = MixText(h, syms.spell(ids(r * arity + c)))
        Next c
    Next r
    HashOfIds = h
End Function

' Folds one cell's text, and a separator, into a running 23-bit hash that
' cannot overflow a Long - the same arithmetic as VLA_Relation's SymHash.
Private Function MixText(ByVal h As Long, ByVal s As String) As Long
    Dim b() As Byte
    b = s
    Dim i As Long
    For i = 0 To UBound(b)
        h = ((h * 33) + b(i)) And &H7FFFFF
    Next i
    MixText = ((h * 33) + 31) And &H7FFFFF
End Function

Private Function SecondsSince(ByVal t0 As Double) As Double
    Dim t As Double
    t = Timer - t0
    If t < 0 Then t = t + 86400#
    SecondsSince = t
End Function

Private Function SecsText(ByVal t As Double) As String
    If t < 0 Then
        SecsText = "-"
    Else
        SecsText = Format$(t, "0.000")
    End If
End Function

Private Function CountText(ByVal n As Long) As String
    If n < 0 Then
        CountText = "-"
    Else
        CountText = Format$(n, "#,##0")
    End If
End Function

' ---------------------------------------------------------------------------
' SLICE 2: the search, timed.
' ---------------------------------------------------------------------------

Public Sub O3Search()
    Debug.Print "=== O3S: the search, timed ==="
    Debug.Print "    Excel " & Application.Version & ", " & Format$(Now, "yyyy-mm-dd hh:nn")
    Debug.Print "O3S|run|people|k|atoms|decisions|dead ends|expected decisions|expected dead ends|outcome|seconds|us per unit of work|verdict"
    ' The expected counts, from the transliteration: (people, k, decisions).
    ' Dead ends are always one more than decisions on this family.
    Dim rungs As Variant
    rungs = Split("14,8,923 16,9,3431 18,10,12869 20,11,48619", " ")
    Dim i As Long
    Dim stopAt As String
    For i = LBound(rungs) To UBound(rungs)
        Dim bits As Variant
        bits = Split(CStr(rungs(i)), ",")
        If Len(stopAt) = 0 Then
            Dim v As String
            v = SearchRung(CLng(bits(0)), CLng(bits(1)), CLng(bits(2)))
            If v <> "ok" Then stopAt = v
        Else
            Debug.Print "O3S|search alone|" & bits(0) & "|" & bits(1) & "||||||||not run: " & stopAt
        End If
    Next i
    If Len(stopAt) = 0 Then
        ProgramRung "family through OPTIMIZE", FamilyProgram(16, 9), 16, 9, 32, 3431, 3432, VLA_Optimize.VLA_OPTIMIZE_NO_SCHEDULE
        ProgramRung "toy scaled, through OPTIMIZE", RowProgram(20, 21, 4), 20, 4, 420, 84, 0, VLA_Optimize.VLA_OPTIMIZE_PROVEN_BEST
    End If
    Debug.Print "=== O3S done ==="
End Sub

' One family rung through the search alone.
Private Function SearchRung(ByVal n As Long, ByVal k As Long, ByVal wantD As Long) As String
    Dim prob As OptSearchProblem
    VLA_OptimizeSearch.OptProblemInit prob, 2 * n
    Dim buf() As Long
    ReDim buf(1 To 2 * n)
    Dim q As Long
    For q = 1 To n
        buf(1) = -q
        buf(2) = -(n + q)
        VLA_OptimizeSearch.OptProblemAddClause prob, buf, 2, q
    Next q
    For q = 1 To n
        buf(q) = q
    Next q
    VLA_OptimizeSearch.OptProblemAddCounter prob, buf, n, k, k, 1
    For q = 1 To n
        buf(q) = n + q
    Next q
    VLA_OptimizeSearch.OptProblemAddCounter prob, buf, n, k, k, 2
    Dim res As OptSearchResult
    Dim t0 As Double
    t0 = Timer
    VLA_OptimizeSearch.OptSearchRun prob, 100000000, 0, res
    Dim secs As Double
    secs = SecondsSince(t0)
    Dim verdict As String
    If res.outcome <> VLA_OptimizeSearch.OPT_SEARCH_NONE Then
        verdict = "WRONG: the outcome was " & res.outcome & ", not none (2)"
    ElseIf res.decisions <> wantD Or res.conflicts <> wantD + 1 Then
        verdict = "WRONG: a different search - the counts differ from the transliteration's"
    Else
        verdict = "ok"
    End If
    Debug.Print "O3S|search alone|" & n & "|" & k & "|" & (2 * n) & "|" & res.decisions & "|" & res.conflicts & "|" & _
                wantD & "|" & (wantD + 1) & "|" & OutcomeWord(res.outcome) & "|" & SecsText(secs) & "|" & _
                PerUnitText(secs, res.work) & "|" & verdict
    SearchRung = verdict
End Function

' A whole OPTIMIZE run: grounding and search, timed together.
Private Sub ProgramRung(ByVal label As String, ByVal prog As String, ByVal n As Long, ByVal k As Long, _
                        ByVal wantAtoms As Long, ByVal wantD As Long, ByVal wantC As Long, ByVal wantState As Long)
    VLA_Optimize.OptimizeMemoClear
    Dim t0 As Double
    t0 = Timer
    Dim r As Collection
    Set r = VLA_Optimize.OptimizeRun(prog)
    Dim secs As Double
    secs = SecondsSince(t0)
    Dim st As Variant
    st = r.Item(9)
    Dim verdict As String
    If CLng(r.Item(6)) <> wantState Then
        verdict = "WRONG: the state was " & r.Item(6) & ", not " & wantState & " - " & r.Item(7)
    ElseIf CLng(st(0)) <> wantAtoms Or CLng(st(3)) <> wantD Or CLng(st(4)) <> wantC Then
        verdict = "WRONG: a different grounding or search - atoms, decisions or dead ends differ"
    Else
        verdict = "ok"
    End If
    Debug.Print "O3S|" & label & "|" & n & "|" & k & "|" & st(0) & "|" & st(3) & "|" & st(4) & "|" & _
                wantD & "|" & wantC & "|" & OutcomeWord(CLng(st(6))) & "|" & SecsText(secs) & "|" & _
                PerUnitText(secs, CLng(st(5))) & "|" & verdict
End Sub

' The family as an OPTIMIZE program: its atoms are numbered exactly as
' SearchRung numbers them (shift a's people in order, then shift b's).
Private Function FamilyProgram(ByVal n As Long, ByVal k As Long) As String
    Dim s As String
    Dim q As Long
    For q = 1 To n
        s = s & "(fact (person p" & q & ")) "
    Next q
    FamilyProgram = s & "(fact (shift a)) (fact (shift b)) (rule (elig S P) (shift S) (person P)) " & _
        "(choose-exactly " & k & " (on S P) (elig S P) (per (shift S))) " & _
        "(forbid (on a P) (on b P)) (effort thorough) (query on)"
End Function

' OPTIMIZE's toy, scaled: m people, s shifts in a row, exactly k a shift,
' never two in a row.
Private Function RowProgram(ByVal m As Long, ByVal nS As Long, ByVal k As Long) As String
    Dim s As String
    Dim q As Long
    For q = 1 To m
        s = s & "(fact (person p" & q & ")) "
    Next q
    For q = 1 To nS
        s = s & "(fact (shift s" & q & ")) "
    Next q
    For q = 1 To nS - 1
        s = s & "(fact (next s" & q & " s" & (q + 1) & ")) "
    Next q
    RowProgram = s & "(rule (elig S P) (shift S) (person P)) " & _
        "(choose-exactly " & k & " (assign S P) (elig S P) (per (shift S))) " & _
        "(forbid (assign S P) (assign T P) (next S T)) (query assign)"
End Function

Private Function OutcomeWord(ByVal o As Long) As String
    Select Case o
    Case 1: OutcomeWord = "found"
    Case 2: OutcomeWord = "none"
    Case 3: OutcomeWord = "budget"
    Case 4: OutcomeWord = "guard"
    Case Else: OutcomeWord = "(" & o & ")"
    End Select
End Function

Private Function PerUnitText(ByVal secs As Double, ByVal work As Long) As String
    If work <= 0 Or secs <= 0 Then
        PerUnitText = "(under the timer's floor)"
    Else
        PerUnitText = Format$(secs * 1000000# / work, "0.0")
    End If
End Function

' ---------------------------------------------------------------------------
' SLICE 3: the ceilings, timed.
' ---------------------------------------------------------------------------

Public Sub O3Ceiling()
    Debug.Print "=== O3C: the ceilings, timed ==="
    Debug.Print "    Excel " & Application.Version & ", " & Format$(Now, "yyyy-mm-dd hh:nn")
    Debug.Print "O3C|run|rows laid out|largest step|expected|expected largest|step refused|seconds|us per row laid out|verdict"
    Dim rels As Object
    ' Pass 2's shape: 250 items each in any of 198 slots - the items' 250,
    ' 250 x 198 = 49,500 paired, then the 49,500 members: 99,250.
    Set rels = CeilItemsSlots(250, 198)
    GroundRung "member rule, 250 items x 198 slots", "(rule (m k S X S) (item X) (slot S))", rels, False, _
               99250, 49500, 0
    ' Pass 3's shape, planned: next's 552 pairs, each pair's first shift for
    ' 60 people (33,120), its second (33,120), the rule's own: 99,912.
    Set rels = CeilRoster(60, 553)
    GroundRung "never two in a row, planned, 60 x 553", "(rule (h S T P) (assign S P) (assign T P) (next S T))", rels, _
               True, 99912, 33120, 0
    ' The same rule in its own written order: assign's 33,180 rows, then
    ' every pair of one person's shifts, 60 x 553 x 553 = 18,348,540 -
    ' counted from the keys, and refused before a row of it is made.
    GroundRung "never two in a row, written, 60 x 553", "(rule (h S T P) (assign S P) (assign T P) (next S T))", rels, _
               False, 33180, 33180, 18348540
    ' A whole run at the ceiling: slot's 198 groups (396) and the member
    ' rule's 99,250 - 99,646 laid out - then a search deciding all 49,500
    ' atoms in, and the answer built.
    WholeRung "OPTIMIZE, 250 items x 198 slots", CeilProgram(250, 198), 99646, 49500, 49500
    Debug.Print "=== O3C done ==="
End Sub

' One rule through the grounder at a formula's own ceilings, 50,000 rows in
' one step and 100,000 in all. wantRefused > 0: the rows of the step it
' must be refused at.
Private Sub GroundRung(ByVal label As String, ByVal ruleText As String, ByVal rels As Object, _
                       ByVal planned As Boolean, ByVal wantLaid As Double, ByVal wantPeak As Double, _
                       ByVal wantRefused As Double)
    Dim syms As VlaSymbols
    VLA_Relation.VlaSymInit syms
    Dim forms As Collection
    Set forms = VLA.VlaReadForms(ruleText)
    Dim t0 As Double
    t0 = Timer
    Dim res As Collection
    Set res = VLA_Datalog.DatalogGroundRules(forms, rels, Nothing, syms, False, planned, 50000, 100000, 0)
    Dim secs As Double
    secs = SecondsSince(t0)
    Dim over As Variant
    over = VLA_Datalog.DatalogGroundOverflow()
    Dim laid As Double, peak As Double, refused As Double
    laid = VLA_Datalog.DatalogGroundRowsLaid()
    peak = VLA_Datalog.DatalogGroundPeakStep()
    If IsArray(over) Then refused = CDbl(over(3))
    Dim verdict As String
    If refused <> wantRefused Then
        verdict = "WRONG: the step refused was " & refused & ", not " & wantRefused
    ElseIf laid <> wantLaid Or peak <> wantPeak Then
        verdict = "WRONG: a different grounding - the rows laid out differ from the transliteration's"
    Else
        verdict = "ok"
    End If
    Dim refusedText As String
    If refused > 0 Then refusedText = CountText(CLng(refused)) Else refusedText = "-"
    Debug.Print "O3C|" & label & "|" & CountText(CLng(laid)) & "|" & CountText(CLng(peak)) & "|" & _
                CountText(CLng(wantLaid)) & "|" & CountText(CLng(wantPeak)) & "|" & refusedText & "|" & _
                SecsText(secs) & "|" & PerUnitText(secs, CLng(laid)) & "|" & verdict
End Sub

' A whole OptimizeRun: pass 1, the grounding, the search and the answer.
Private Sub WholeRung(ByVal label As String, ByVal prog As String, ByVal wantLaid As Long, _
                      ByVal wantPeak As Long, ByVal wantAtoms As Long)
    VLA_Optimize.OptimizeMemoClear
    Dim t0 As Double
    t0 = Timer
    Dim r As Collection
    Set r = VLA_Optimize.OptimizeRun(prog)
    Dim secs As Double
    secs = SecondsSince(t0)
    Dim st As Variant
    st = r.Item(9)
    Dim verdict As String
    If CLng(st(6)) <> VLA_OptimizeSearch.OPT_SEARCH_FOUND Then
        verdict = "WRONG: the search's outcome was " & OutcomeWord(CLng(st(6))) & ", not found"
    ElseIf CLng(st(8)) <> wantLaid Or CLng(st(9)) <> wantPeak Or CLng(st(0)) <> wantAtoms Then
        verdict = "WRONG: a different grounding - the rows laid out or the atoms differ"
    Else
        verdict = "ok"
    End If
    Debug.Print "O3C|" & label & "|" & CountText(CLng(st(8))) & "|" & CountText(CLng(st(9))) & "|" & _
                CountText(wantLaid) & "|" & CountText(wantPeak) & "|-|" & SecsText(secs) & "|" & _
                PerUnitText(secs, CLng(st(8))) & "|" & verdict & " (" & CountText(CLng(st(3))) & " decisions)"
End Sub

' item i1..iN and slot t1..tM, as plain rows.
Private Function CeilItemsSlots(ByVal nItems As Long, ByVal nSlots As Long) As Object
    Dim rels As Object
    Set rels = VLA_Runtime.VlaDictNew()
    Dim items As Collection, slots As Collection
    Set items = VLA_Relation.RelNew(1)
    Set slots = VLA_Relation.RelNew(1)
    Dim one(1 To 1) As Variant
    Dim k As Long
    For k = 1 To nItems
        one(1) = "i" & k
        VLA_Relation.RelTryAdd items, one
    Next k
    For k = 1 To nSlots
        one(1) = "t" & k
        VLA_Relation.RelTryAdd slots, one
    Next k
    VLA_Runtime.VlaDictSet rels, "item", items
    VLA_Runtime.VlaDictSet rels, "slot", slots
    Set CeilItemsSlots = rels
End Function

' The roster's two relations, shift by shift: assign, every shift with every
' person, and next, each shift with the one after it.
Private Function CeilRoster(ByVal nPeople As Long, ByVal nShifts As Long) As Object
    Dim rels As Object
    Set rels = VLA_Runtime.VlaDictNew()
    Dim asn As Collection, nx As Collection
    Set asn = VLA_Relation.RelNew(2)
    Set nx = VLA_Relation.RelNew(2)
    Dim pair(1 To 2) As Variant
    Dim s As Long, p As Long
    For s = 1 To nShifts
        For p = 1 To nPeople
            pair(1) = "s" & s
            pair(2) = "p" & p
            VLA_Relation.RelTryAdd asn, pair
        Next p
        If s < nShifts Then
            pair(1) = "s" & s
            pair(2) = "s" & (s + 1)
            VLA_Relation.RelTryAdd nx, pair
        End If
    Next s
    VLA_Runtime.VlaDictSet rels, "assign", asn
    VLA_Runtime.VlaDictSet rels, "next", nx
    Set CeilRoster = rels
End Function

' 250 items in any of the slots, at most 250 a slot - which binds nothing.
Private Function CeilProgram(ByVal nItems As Long, ByVal nSlots As Long) As String
    Dim s As String
    Dim k As Long
    For k = 1 To nItems
        s = s & "(fact (item i" & k & ")) "
    Next k
    For k = 1 To nSlots
        s = s & "(fact (slot t" & k & ")) "
    Next k
    CeilProgram = s & "(choose-at-most " & nItems & " (pick X S) (item X) (per (slot S))) (query pick)"
End Function

' ---------------------------------------------------------------------------
' SLICE 4: the ladder.
' ---------------------------------------------------------------------------

' Each rung: people and weeks, then the counts a transliteration computed
' before the pass - atoms, counters, clauses, rows laid out, largest step,
' decisions and dead ends - and the answer: S a schedule, N no schedule by
' the counting pre-check. The need is the fixture's tight one, a fifth of
' the people and at least one.
Public Sub O3Ladder()
    Debug.Print "=== O3L: the ladder - OPTIMIZE.0's reference roster, grounded and searched ==="
    Debug.Print "    Excel " & Application.Version & ", " & Format$(Now, "yyyy-mm-dd hh:nn")
    Debug.Print "O3L|rung|a shift|atoms|counters|clauses|rows laid out|largest step|decisions|dead ends|answer|runs|seconds a run|verdict"
    Dim rungs As Variant
    rungs = Array("5,1,90,33,80,889,90,0,0,N", "5,4,360,132,335,3604,360,0,0,N", _
                  "10,1,180,38,162,1707,180,35,2,S", "10,4,720,152,672,6906,720,140,8,S", _
                  "20,1,360,48,325,3358,360,80,2,S", "20,4,1440,192,1345,13579,1440,320,8,S", _
                  "50,1,900,78,814,8277,900,200,0,S", "50,4,3600,312,3364,33456,3600,800,0,S")
    Dim i As Long
    Dim stopAt As String
    For i = LBound(rungs) To UBound(rungs)
        Dim b As Variant
        b = Split(CStr(rungs(i)), ",")
        If Len(stopAt) = 0 Then
            Dim v As String
            v = LadderRung(b)
            If v <> "ok" Then stopAt = "an earlier rung said " & v
        Else
            Debug.Print "O3L|" & b(0) & " x " & b(1) & String$(12, "|") & "not run: " & stopAt
        End If
    Next i
    If Len(stopAt) = 0 Then
        LadderCalibration
    Else
        Debug.Print "O3L|the calibration" & String$(12, "|") & "not run: " & stopAt
    End If
    Debug.Print "=== O3L done ==="
End Sub

' One rung: the program run once and checked, then timed.
Private Function LadderRung(ByVal want As Variant) As String
    Dim nP As Long, nW As Long, need As Long
    nP = CLng(want(0))
    nW = CLng(want(1))
    need = nP \ 5
    If need < 1 Then need = 1
    Dim rels As Object
    Set rels = LadderRelations(nP, nW, need)
    Dim prog As String
    prog = LadderProgram("")
    VLA_Optimize.OptimizeMemoClear
    Dim r As Collection
    Set r = VLA_Optimize.OptimizeRun(prog, rels)
    Dim st As Variant
    st = r.Item(9)
    Dim verdict As String
    If CLng(st(0)) <> CLng(want(2)) Or CLng(st(2)) <> CLng(want(3)) Or CLng(st(1)) <> CLng(want(4)) Or _
       CLng(st(8)) <> CLng(want(5)) Or CLng(st(9)) <> CLng(want(6)) Then
        verdict = "WRONG: a different grounding - the atoms, counters, clauses or rows laid out differ"
    ElseIf CLng(st(3)) <> CLng(want(7)) Or CLng(st(4)) <> CLng(want(8)) Then
        verdict = "WRONG: a different search - the decisions or dead ends differ"
    ElseIf want(9) = "N" Then
        verdict = LadderNoneCheck(r)
    Else
        verdict = LadderAnswerCheck(r, nP, nW, need)
    End If
    Dim runs As Long, secs As Double
    secs = -1
    If verdict = "ok" Then LadderTime prog, rels, runs, secs
    Debug.Print "O3L|" & nP & " x " & nW & "|" & need & "|" & CountText(CLng(st(0))) & "|" & st(2) & "|" & _
                CountText(CLng(st(1))) & "|" & CountText(CLng(st(8))) & "|" & CountText(CLng(st(9))) & "|" & _
                st(3) & "|" & st(4) & "|" & LadderAnswerWord(r) & "|" & runs & "|" & SecsText(secs) & "|" & verdict
    If verdict = "ok" And secs > 10# Then verdict = "a run took " & Format$(secs, "0.0") & " s, over 10 s"
    LadderRung = verdict
End Function

' The 5-people rungs: no schedule, by counting, before any search. Their one
' senior, P4, is on leave on day 3, so the night of day 3 - shift 9 - has no
' senior free to take it (Entry 1 section 3a, optimize-roster-senior).
Private Function LadderNoneCheck(ByVal r As Collection) As String
    Dim wantWords As String
    wantWords = "no schedule satisfies every rule: night S = 9 needs 1 row from 'senior-free', " & _
                "and only 0 rows can fill it."
    If CLng(r.Item(6)) <> VLA_Optimize.VLA_OPTIMIZE_NO_SCHEDULE Then
        LadderNoneCheck = "WRONG: the state was " & r.Item(6) & ", not no schedule (1)"
    ElseIf CStr(r.Item(7)) <> wantWords Then
        LadderNoneCheck = "WRONG: the status said: " & r.Item(7)
    ElseIf VLA_Relation.RelCount(VLA_Runtime.VlaDictGet(r.Item(2), "assign")) <> 0 Then
        LadderNoneCheck = "WRONG: an answer with no schedule holds rows"
    Else
        LadderNoneCheck = "ok"
    End If
End Function

' A schedule, checked against the fixture's own definition by this module,
' not by OPTIMIZE: every row a shift and a person of the fixture, the person
' not on leave that day; every shift holding exactly its need; a senior on
' every night; nobody on two shifts in a row; nobody on more than five in a
' week.
Private Function LadderAnswerCheck(ByVal r As Collection, ByVal nP As Long, ByVal nW As Long, _
                                   ByVal need As Long) As String
    If CLng(r.Item(6)) <> VLA_Optimize.VLA_OPTIMIZE_PROVEN_BEST Then
        LadderAnswerCheck = "WRONG: the state was " & r.Item(6) & ", not proven best (4)"
        Exit Function
    End If
    Dim nS As Long
    nS = 21 * nW
    Dim rel As Collection
    Set rel = VLA_Runtime.VlaDictGet(r.Item(2), "assign")
    If VLA_Relation.RelCount(rel) <> nS * need Then
        LadderAnswerCheck = "WRONG: the schedule has " & VLA_Relation.RelCount(rel) & " rows, not " & (nS * need)
        Exit Function
    End If
    Dim onShift() As Boolean
    ReDim onShift(1 To nS + 1, 1 To nP)
    Dim t As Variant, arr() As Variant
    Dim s As Long, p As Long, w As Long, cnt As Long
    Dim senior As Boolean
    For Each t In VLA_Relation.RelTuples(rel)
        arr = t
        s = CLng(arr(LBound(arr)))
        p = CLng(arr(LBound(arr) + 1))
        If s < 1 Or s > nS Or p < 1 Or p > nP Then
            LadderAnswerCheck = "WRONG: a row outside the fixture, shift " & s & " person " & p
            Exit Function
        End If
        If (p + LadderDay(s)) Mod 7 = 0 Then
            LadderAnswerCheck = "WRONG: P" & p & " is on shift " & s & " while on leave"
            Exit Function
        End If
        onShift(s, p) = True
    Next t
    For s = 1 To nS
        cnt = 0
        senior = False
        For p = 1 To nP
            If onShift(s, p) Then
                cnt = cnt + 1
                If p Mod 4 = 0 Then senior = True
                If onShift(s + 1, p) Then
                    LadderAnswerCheck = "WRONG: P" & p & " works shifts " & s & " and " & (s + 1)
                    Exit Function
                End If
            End If
        Next p
        If cnt <> need Then
            LadderAnswerCheck = "WRONG: shift " & s & " has " & cnt & ", not " & need
            Exit Function
        End If
        If (s - 1) Mod 3 = 2 And Not senior Then
            LadderAnswerCheck = "WRONG: the night of shift " & s & " has no senior"
            Exit Function
        End If
    Next s
    For p = 1 To nP
        For w = 1 To nW
            cnt = 0
            For s = 21 * (w - 1) + 1 To 21 * w
                If onShift(s, p) Then cnt = cnt + 1
            Next s
            If cnt > 5 Then
                LadderAnswerCheck = "WRONG: P" & p & " works " & cnt & " shifts in week " & w
                Exit Function
            End If
        Next w
    Next p
    LadderAnswerCheck = "ok"
End Function

' The same program again and again, the memo cleared each time, until two
' seconds have passed or fifty runs: a small rung is timed well clear of
' Timer's steps, and the mean is what is printed.
Private Sub LadderTime(ByVal prog As String, ByVal rels As Object, ByRef runs As Long, ByRef secs As Double)
    Dim t0 As Double, total As Double
    Dim r As Collection
    runs = 0
    t0 = Timer
    Do
        VLA_Optimize.OptimizeMemoClear
        Set r = VLA_Optimize.OptimizeRun(prog, rels)
        runs = runs + 1
        total = SecondsSince(t0)
    Loop While total < 2# And runs < 50
    secs = total / runs
End Sub

Private Function LadderAnswerWord(ByVal r As Collection) As String
    Select Case CLng(r.Item(6))
    Case VLA_Optimize.VLA_OPTIMIZE_PROVEN_BEST
        LadderAnswerWord = "a schedule"
    Case VLA_Optimize.VLA_OPTIMIZE_NO_SCHEDULE
        LadderAnswerWord = "none"
    Case VLA_Optimize.VLA_OPTIMIZE_NONE_IN_BUDGET
        LadderAnswerWord = "none within the effort"
    Case Else
        LadderAnswerWord = "(state " & r.Item(6) & ")"
    End Select
End Function

' THE CALIBRATION. The same roster at 50 x 4 with ELEVEN a shift, not ten: a
' tightness of 0.92 against the reference's 0.84. The transliteration says
' the search never leaves week 1 there - 175 decisions in it meets its first
' dead end, on day 6, and backtracking one decision at a time never reaches
' back past the 147th - so every effort level runs out, and a run's search is
' exactly its effort's units. (effort 1) times everything else: DATALOG's
' pass, the grounding, the problem, one decision and the answer. What is left
' when that is taken away is the search alone, at a real roster's widths.
' Round 2 runs the levels slice 4 set, 25,000 and 250,000 units; round 1 ran
' slice 2's 50,000 and 500,000 (25,084 and 250,084 decisions).
Private Sub LadderCalibration()
    Dim rels As Object
    Set rels = LadderRelations(50, 4, 11)
    Dim t1 As Double, tN As Double, tT As Double
    If Not CalibRung("(effort 1)", rels, 1, 0, t1) Then Exit Sub
    If Not CalibRung("(effort normal)", rels, 12583, 12417, tN) Then Exit Sub
    If Not CalibRung("(effort thorough)", rels, 125087, 124913, tT) Then Exit Sub
    Debug.Print "O3L|a unit of search work, 50 x 4 at 11 a shift: " & _
                UsPerUnit(tN - t1, VLA_Optimize.VLA_OPTIMIZE_WORK_NORMAL - 1) & " us at normal, " & _
                UsPerUnit(tT - t1, VLA_Optimize.VLA_OPTIMIZE_WORK_THOROUGH - 1) & " us at thorough; " & _
                "everything but the search " & SecsText(t1) & " s"
    LadderSplit
End Sub

' WHERE A RUN GOES, at the reference roster itself (50 x 4, 10 a shift): its
' memo key, a hash of the program and every Table; DATALOG's own pass over
' the program's six rules alone, which is OPTIMIZE's first pass less its
' stubs; and the whole run at (effort 1), which is everything but the
' search. What is left is OPTIMIZE's own work: reading the program, the
' grounding the ceilings count, the counters, clauses and checks, the
' problem, one decision and the answer.
Private Sub LadderSplit()
    Dim rels As Object
    Set rels = LadderRelations(50, 4, 10)
    Dim prog As String
    prog = LadderProgram("(effort 1) ")
    Dim hdr As Object
    Set hdr = VLA_Runtime.VlaDictNew()
    Dim runs As Long
    Dim t0 As Double, total As Double
    Dim tKey As Double, tPass As Double, tRun As Double
    Dim key As String
    t0 = Timer
    Do
        key = VLA_Optimize.OptimizeMemoKey(prog, rels, hdr)
        runs = runs + 1
        total = SecondsSince(t0)
    Loop While total < 2# And runs < 50
    tKey = total / runs
    Dim rulesOnly As String
    rulesOnly = LadderRules() & "(query free)"
    Dim res As Collection, copyRels As Object
    Dim k As Variant
    runs = 0
    t0 = Timer
    Do
        Set copyRels = VLA_Runtime.VlaDictNew()
        For Each k In VLA_Runtime.VlaDictKeys(rels)
            VLA_Runtime.VlaDictSet copyRels, CStr(k), VLA_Runtime.VlaDictGet(rels, k)
        Next k
        Set res = VLA_Datalog.DatalogRunForms(VLA.VlaReadForms(rulesOnly), copyRels, Nothing)
        runs = runs + 1
        total = SecondsSince(t0)
    Loop While total < 2# And runs < 50
    tPass = total / runs
    LadderTime prog, rels, runs, tRun
    Debug.Print "O3L|where a 50 x 4 run goes: the memo key " & SecsText(tKey) & " s, DATALOG's pass over the six rules " & _
                SecsText(tPass) & " s, OPTIMIZE's own work " & SecsText(tRun - tKey - tPass) & " s; " & _
                SecsText(tRun) & " s in all at (effort 1)"
End Sub

Private Function CalibRung(ByVal effortText As String, ByVal rels As Object, ByVal wantD As Long, _
                           ByVal wantC As Long, ByRef secs As Double) As Boolean
    Dim prog As String
    prog = LadderProgram(effortText & " ")
    VLA_Optimize.OptimizeMemoClear
    Dim r As Collection
    Set r = VLA_Optimize.OptimizeRun(prog, rels)
    Dim st As Variant
    st = r.Item(9)
    Dim verdict As String
    If CLng(st(0)) <> 3600 Or CLng(st(2)) <> 312 Or CLng(st(1)) <> 3364 Or CLng(st(8)) <> 33456 Then
        verdict = "WRONG: a different grounding - the atoms, counters, clauses or rows laid out differ"
    ElseIf CLng(st(6)) = VLA_OptimizeSearch.OPT_SEARCH_GUARD Then
        verdict = "GUARD: the 10 s guard stopped the search at " & CountText(CLng(st(5))) & _
                  " units, before its effort ran out"
    ElseIf CLng(st(6)) <> VLA_OptimizeSearch.OPT_SEARCH_BUDGET Then
        verdict = "WRONG: the search's outcome was " & OutcomeWord(CLng(st(6))) & ", not budget"
    ElseIf CLng(st(3)) <> wantD Or CLng(st(4)) <> wantC Then
        verdict = "WRONG: a different search - the decisions or dead ends differ"
    ElseIf VLA_Relation.RelCount(VLA_Runtime.VlaDictGet(r.Item(2), "assign")) <> 0 Then
        verdict = "WRONG: an answer with no schedule holds rows"
    Else
        verdict = "ok"
    End If
    Dim runs As Long
    secs = -1
    If verdict = "ok" Then LadderTime prog, rels, runs, secs
    Debug.Print "O3L|50 x 4 at 11, " & effortText & "|11|" & CountText(CLng(st(0))) & "|" & st(2) & "|" & _
                CountText(CLng(st(1))) & "|" & CountText(CLng(st(8))) & "|" & CountText(CLng(st(9))) & "|" & _
                CountText(CLng(st(3))) & "|" & CountText(CLng(st(4))) & "|" & LadderAnswerWord(r) & "|" & _
                runs & "|" & SecsText(secs) & "|" & verdict
    CalibRung = (verdict = "ok")
End Function

Private Function UsPerUnit(ByVal secs As Double, ByVal units As Long) As String
    UsPerUnit = Format$(secs * 1000000# / units, "0.00")
End Function

' The reference roster's four Tables as DATALOG relations built directly, in
' OPTIMIZE.0's O0Fixture columns: people (Id, Name, Senior, Contract), shifts
' (Id, Week, Day, Slot, Need, Skill), leave (Person, Day) and next (From,
' To). Ids and numbers are Doubles, as a worksheet hands them over.
Private Function LadderRelations(ByVal nP As Long, ByVal nW As Long, ByVal need As Long) As Object
    Dim rels As Object
    Set rels = VLA_Runtime.VlaDictNew()
    Dim people As Collection, shifts As Collection, leave As Collection, nx As Collection
    Set people = VLA_Relation.RelNew(4)
    Set shifts = VLA_Relation.RelNew(6)
    Set leave = VLA_Relation.RelNew(2)
    Set nx = VLA_Relation.RelNew(2)
    Dim t4(1 To 4) As Variant, t6(1 To 6) As Variant, t2(1 To 2) As Variant
    Dim p As Long, s As Long, d As Long, nS As Long
    For p = 1 To nP
        t4(1) = CDbl(p)
        t4(2) = "P" & p
        If p Mod 4 = 0 Then t4(3) = "yes" Else t4(3) = "no"
        If p Mod 2 = 1 Then t4(4) = 4# Else t4(4) = 5#
        VLA_Relation.RelTryAdd people, t4
    Next p
    nS = 21 * nW
    For s = 1 To nS
        t6(1) = CDbl(s)
        t6(2) = CDbl((LadderDay(s) - 1) \ 7 + 1)
        t6(3) = CDbl(LadderDay(s))
        t6(4) = Choose((s - 1) Mod 3 + 1, "Early", "Late", "Night")
        t6(5) = CDbl(need)
        If (s - 1) Mod 3 = 2 Then t6(6) = "senior" Else t6(6) = "any"
        VLA_Relation.RelTryAdd shifts, t6
    Next s
    For p = 1 To nP
        For d = 1 To 7 * nW
            If (p + d) Mod 7 = 0 Then
                t2(1) = CDbl(p)
                t2(2) = CDbl(d)
                VLA_Relation.RelTryAdd leave, t2
            End If
        Next d
    Next p
    For s = 1 To nS - 1
        t2(1) = CDbl(s)
        t2(2) = CDbl(s + 1)
        VLA_Relation.RelTryAdd nx, t2
    Next s
    VLA_Runtime.VlaDictSet rels, "people", people
    VLA_Runtime.VlaDictSet rels, "shifts", shifts
    VLA_Runtime.VlaDictSet rels, "leave", leave
    VLA_Runtime.VlaDictSet rels, "next", nx
    Set LadderRelations = rels
End Function

Private Function LadderDay(ByVal s As Long) As Long
    LadderDay = (s - 1) \ 3 + 1
End Function

' The reference roster's rules, less its objective and its kept roster
' (OPTIMIZE.6 and .7): free is who may work a shift, being on leave that day
' neither; every shift gets exactly its need; nobody works more than five a
' week; every night gets at least one senior; nobody works two in a row.
Private Function LadderProgram(ByVal effortText As String) As String
    LadderProgram = LadderRules() & _
        "(choose-exactly N (assign S P) (free S P) (per (shifts S W D Slot N K))) " & _
        "(choose-at-most 5 (assign S P) (free-week S P W) (per (person-week P W))) " & _
        "(choose-at-least 1 (assign S P) (senior-free S P) (per (night S))) " & _
        "(forbid (assign S P) (assign T P) (next S T)) " & effortText & "(query assign)"
End Function

' The six rules over the four Tables that every choice above reads.
Private Function LadderRules() As String
    LadderRules = "(rule (free S P) (shifts S W D Slot N K) (people P Nm Sr C) (not (leave P D))) " & _
        "(rule (week W) (shifts S W D Slot N K)) " & _
        "(rule (person-week P W) (people P Nm Sr C) (week W)) " & _
        "(rule (free-week S P W) (free S P) (shifts S W D Slot N K)) " & _
        "(rule (night S) (shifts S W D Slot N ""senior"")) " & _
        "(rule (senior-free S P) (free S P) (night S) (people P Nm ""yes"" C)) "
End Function
