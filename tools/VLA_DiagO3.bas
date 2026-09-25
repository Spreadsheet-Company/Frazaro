Attribute VB_Name = "VLA_DiagO3"
Option Explicit

' Standalone diagnostic for OPTIMIZE.3 - NOT part of the VLA project. Import
' it into the dev workbook (VLA.xlsm), type one line in the Immediate window,
' and delete the module afterwards; nothing in Frazaro calls it. The steps are
' in tools/optimize3_live_steps.md.
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
