Attribute VB_Name = "VLA_DiagO0"
Option Explicit

' Standalone diagnostic for OPTIMIZE.0 - NOT part of the VLA project. Import it
' into the dev workbook (VLA.xlsm), type one step's line in the Immediate
' window, and delete the module afterwards; nothing in Frazaro calls it. The
' steps are in archive/optimize0_live_steps.md.
'
' WHAT IT CALLS IN FRAZARO, AND WHY: nothing from VBA. It writes a
' =DATALOG(...) worksheet formula and times Excel's recalculation of it, because
' OPTIMIZE will ground through DATALOG's own evaluator and the measurement is
' what that evaluator costs today. Everything else here - the fixtures, the
' expected answers, the model - is this module's own.
'
' WHAT THIS MEASURES. OPTIMIZE does not exist. What it will pay first is
' GROUNDING: instantiating each rule over the Tables. This harness writes each
' constraint shape a staff roster needs as a DATALOG rule whose rows ARE that
' shape's ground instances, grows the fixture up a ladder of sizes, times the
' answer, and checks every row against what the fixture's definition says it
' must be. Each step builds its own fresh sheet (O0T<n>) and its own Tables.
'
' THE MODEL, predicted before anything was measured (docs/OPTIMIZATION.md,
' Entry 1): time = a*F + b*J, where F counts source rows filtered (one
' Scripting.Dictionary created per row in AtomMatches, a = 0.16 ms) and J counts
' rows a join, comparison, let or head produces (plain arrays, b = 0.01 ms,
' band 0.005-0.03), both over every fixpoint round - round 2 re-joins the atoms
' ahead of each delta position before it finds the empty delta. Peak memory is
' the largest single join output. Every rung prints the model beside the
' measurement, and the ratio between them.
'
' HOW A LADDER STOPS, so no rung holds Excel for long:
'   - a rung whose model peak passes PEAK_GUARD live rows is never run (the
'     combination shapes hit memory before time; the line still prints the
'     model, so the paper model is in the output for every rung);
'   - a rung whose time, projected as model x the largest measured/model ratio
'     so far in this step, passes GUARD_SECONDS is not run;
'   - after a rung over LIMIT_SECONDS, the rest are not run;
'   - after a refusal, a wrong answer or a VBA error, the step stops.
' Calculation is manual while a step runs and restored after. Timer resolves
' about 16 ms. Application.Calculate also recalculates volatile formulas in any
' open workbook: noise, not DATALOG. The "control" column times a bare UDF
' returning an array of the answer's shape, so the cost of handing the cells
' back is measured beside each answer rather than assumed.
'
' A recalculation after one cell changes is not timed here: DATALOG.12 measured
' it as costing what a first answer costs, for this engine, in six shapes.
'
' Results print as O0| lines and are appended to sheet O0Results.

Private Const LIMIT_SECONDS As Double = 10#
Private Const GUARD_SECONDS As Double = 15#
Private Const PEAK_GUARD As Double = 500000#
Private Const A_MS As Double = 0.16
Private Const B_MS As Double = 0.01
Private Const B_LO_MS As Double = 0.005
Private Const B_HI_MS As Double = 0.03
Private Const RULES_ADDR As String = "N1"
Private Const ANSWER_ADDR As String = "N2"
Private Const CTRL_ROWS_ADDR As String = "Z1"
Private Const CTRL_COLS_ADDR As String = "AA1"
Private Const CTRL_ADDR As String = "Z2"
Private Const RESULTS_SHEET As String = "O0Results"
Private Const LAST_STEP As Long = 12

' Model item kinds: a positive atom over a Table, a positive atom over a
' derived predicate (it has a delta in round 2), and anything else.
Private Const IK_BASE As Long = 1
Private Const IK_DERIVED As Long = 2
Private Const IK_OTHER As Long = 3
Private Const MAX_ITEMS As Long = 32

Private mKind(1 To MAX_ITEMS) As Long
Private mSrc(1 To MAX_ITEMS) As Double
Private mOut(1 To MAX_ITEMS) As Double
Private mItems As Long
Private mF As Double
Private mJ As Double
Private mPeak As Double

' ---------------------------------------------------------------------------
' The ladder
' ---------------------------------------------------------------------------

' onlyRung (added 2026-09-19 after step 7's top rung ran 1.76x its model): run
' just that rung (1 = the first in the step's list), without the TIME guard -
' for a rung the guard skipped that is worth one deliberate stall. The memory
' guard still applies. The step still needs a fresh sheet O0T<n>.
Public Sub O0Ladder(ByVal stepNo As Long, Optional ByVal onlyRung As Long = 0)
    Dim rungs As Variant
    Dim ri As Long
    Dim ws As Worksheet, ans As Range, ctrl As Range
    Dim prevCalc As Long, prevScreen As Boolean, calcSaved As Boolean
    Dim nP As Long, nW As Long, k As Long
    Dim modelS As Double, loS As Double, hiS As Double, projected As Double
    Dim peakRows As Double, ratio As Double, rMax As Double
    Dim ground As Double
    Dim tFirst As Double, tCtrl As Double
    Dim verdictText As String, note As String, outLine As String
    Dim stoppedAt As String
    Dim resultRow As Long
    Dim errNo As Long, errText As String
    Dim bits As String

    If stepNo < 1 Or stepNo > LAST_STEP Then
        Debug.Print "O0: there is no step " & stepNo & " - the steps are 1 to " & LAST_STEP & "."
        Exit Sub
    End If
    If ActiveWorkbook Is Nothing Then
        Debug.Print "O0: no workbook is active."
        Exit Sub
    End If
    If SheetExists(ActiveWorkbook, "O0T" & stepNo) Then
        Debug.Print "O0: sheet O0T" & stepNo & " already exists - delete it first, so the step builds on a fresh sheet."
        Exit Sub
    End If
    rungs = RungList(stepNo)
    If onlyRung < 0 Or onlyRung > UBound(rungs) - LBound(rungs) + 1 Then
        Debug.Print "O0: step " & stepNo & " has rungs 1 to " & (UBound(rungs) - LBound(rungs) + 1) & "."
        Exit Sub
    End If

    #If Win64 Then
        bits = "64-bit"
    #Else
        bits = "32-bit"
    #End If

    On Error GoTo fail
    prevCalc = Application.Calculation
    prevScreen = Application.ScreenUpdating
    calcSaved = True
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    Set ws = ActiveWorkbook.Worksheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
    ws.Name = "O0T" & stepNo
    ParseRung CStr(rungs(LBound(rungs))), nP, nW, k
    SetupStep stepNo, ws, k
    Set ans = ws.Range(ANSWER_ADDR)
    Set ctrl = ws.Range(CTRL_ADDR)
    Application.Calculate
    resultRow = NextResultRow()

    Debug.Print "=== O0 step " & stepNo & ": " & ShapeName(stepNo) & " ==="
    Debug.Print "    Excel " & Application.Version & " " & bits & ", workbook " & ActiveWorkbook.Name & ", " & Format$(Now, "yyyy-mm-dd hh:nn")
    If stepNo < LAST_STEP Then
        Debug.Print "    formula: " & ans.Formula2 & "   with " & RULES_ADDR & " = " & RulesText(stepNo, k)
    End If
    Debug.Print "    model: a = " & A_MS & " ms per source row, b = " & B_MS & " ms per produced row (band " & B_LO_MS & "-" & B_HI_MS & ")"
    Debug.Print "O0|step|people|weeks|k|ground rows|first s|control s|model s (band)|measured/model|peak rows|verdict|note"

    For ri = LBound(rungs) To UBound(rungs)
      If onlyRung = 0 Or ri - LBound(rungs) + 1 = onlyRung Then
        ParseRung CStr(rungs(ri)), nP, nW, k
        loS = ModelSeconds(stepNo, nP, nW, k, B_LO_MS)
        hiS = ModelSeconds(stepNo, nP, nW, k, B_HI_MS)
        modelS = ModelSeconds(stepNo, nP, nW, k, B_MS)
        peakRows = mPeak
        ground = ExpectedRows(stepNo, nP, nW, k)
        tFirst = -1
        tCtrl = -1
        ratio = -1
        verdictText = ""
        note = ""

        If Len(stoppedAt) > 0 Then
            verdictText = "not run"
            note = stoppedAt
        ElseIf peakRows > PEAK_GUARD Then
            verdictText = "not run"
            note = "model peak " & Format$(peakRows, "#,##0") & " live rows, over the " & Format$(PEAK_GUARD, "#,##0") & " memory guard"
        Else
            projected = modelS
            If rMax > 0 Then projected = modelS * rMax
            If projected > GUARD_SECONDS And onlyRung = 0 Then
                verdictText = "not run"
                note = "projected " & Format$(projected, "0.0") & "s, over the " & GUARD_SECONDS & "s guard"
            Else
                Application.StatusBar = "O0 step " & stepNo & ": writing " & RungText(stepNo, nP, nW, k)
                FillStep stepNo, ws, nP, nW, k
                Application.StatusBar = "O0 step " & stepNo & ": calculating " & RungText(stepNo, nP, nW, k) & " - Excel will not respond until it is done"
                If stepNo = LAST_STEP Then
                    ws.Range(CTRL_ROWS_ADDR).Value2 = nP
                    ws.Range(CTRL_COLS_ADDR).Value2 = nW
                    tFirst = SecondsToCalculate(ctrl, False)
                    verdictText = VerifyControl(ctrl, nP, nW)
                Else
                    tFirst = SecondsToCalculate(ans, True)
                    verdictText = VerifyAnswer(stepNo, ans, nP, nW, k, ground)
                    If verdictText = "ok" Then
                        ws.Range(CTRL_ROWS_ADDR).Value2 = ground + 1
                        ws.Range(CTRL_COLS_ADDR).Value2 = AnswerCols(stepNo, k)
                        tCtrl = SecondsToCalculate(ctrl, False)
                        If VerifyControl(ctrl, CLng(ground + 1), AnswerCols(stepNo, k)) <> "ok" Then note = "the control did not spill in full"
                    End If
                End If
                If modelS >= 0.05 And tFirst >= 0 Then
                    ratio = tFirst / modelS
                    If ratio > rMax Then rMax = ratio
                End If
                If tFirst > LIMIT_SECONDS Then
                    stoppedAt = "a smaller size took " & Format$(tFirst, "0.0") & "s, over the " & LIMIT_SECONDS & "s limit"
                End If
            End If
        End If

        outLine = "O0|" & stepNo & "|" & nP & "|" & nW & "|" & KText(stepNo, k) & "|" & Format$(ground, "#,##0") & "|" & _
                  Secs(tFirst) & "|" & Secs(tCtrl) & "|" & Format$(modelS, "0.00") & " (" & Format$(loS, "0.00") & "-" & Format$(hiS, "0.00") & ")|" & _
                  RatioText(ratio) & "|" & Format$(peakRows, "#,##0") & "|" & verdictText & "|" & note
        Debug.Print outLine
        ActiveWorkbook.Worksheets(RESULTS_SHEET).Cells(resultRow, 1).Resize(1, 14).Value2 = _
            Array(stepNo, nP, nW, KText(stepNo, k), ground, Secs(tFirst), Secs(tCtrl), Format$(modelS, "0.00"), _
                  Format$(loS, "0.00") & "-" & Format$(hiS, "0.00"), RatioText(ratio), peakRows, verdictText, note, Format$(Now, "yyyy-mm-dd hh:nn:ss"))
        resultRow = resultRow + 1
        If verdictText <> "ok" And verdictText <> "not run" Then Exit For
      End If
    Next ri

cleanup:
    On Error Resume Next
    Application.StatusBar = False
    If calcSaved Then
        Application.ScreenUpdating = prevScreen
        ' Restoring automatic calculation recalculates anything still dirty.
        Application.Calculation = prevCalc
    End If
    On Error GoTo 0
    Debug.Print "=== O0 step " & stepNo & " done ==="
    Exit Sub
fail:
    errNo = Err.Number
    errText = Err.Description
    Debug.Print "O0|" & stepNo & "|" & nP & "|" & nW & "|" & KText(stepNo, k) & "||||||||VBA error " & errNo & "|" & errText
    Resume cleanup
End Sub

' The model alone, for one step (or every step with 0): nothing is written, no
' Excel state changes. What the ladder will skip by the memory guard is exact
' from this; what it skips by the time guard depends on the measured ratio.
Public Sub O0Model(Optional ByVal stepNo As Long = 0)
    Dim s As Long, sFrom As Long, sTo As Long
    Dim rungs As Variant, ri As Long
    Dim nP As Long, nW As Long, k As Long
    Dim modelS As Double, loS As Double, hiS As Double, peakRows As Double, fRows As Double, jRows As Double
    Dim guardText As String
    If stepNo = 0 Then
        sFrom = 1
        sTo = LAST_STEP - 1
    Else
        sFrom = stepNo
        sTo = stepNo
    End If
    Debug.Print "O0M|step|people|weeks|k|ground rows|F source rows|J produced rows|peak rows|model s (band)|guard"
    For s = sFrom To sTo
        rungs = RungList(s)
        For ri = LBound(rungs) To UBound(rungs)
            ParseRung CStr(rungs(ri)), nP, nW, k
            loS = ModelSeconds(s, nP, nW, k, B_LO_MS)
            hiS = ModelSeconds(s, nP, nW, k, B_HI_MS)
            modelS = ModelSeconds(s, nP, nW, k, B_MS)
            peakRows = mPeak
            fRows = mF
            jRows = mJ
            guardText = ""
            If peakRows > PEAK_GUARD Then
                guardText = "never run: memory"
            ElseIf modelS > GUARD_SECONDS Then
                guardText = "skipped unless measured faster than the model"
            End If
            Debug.Print "O0M|" & s & "|" & nP & "|" & nW & "|" & KText(s, k) & "|" & Format$(ExpectedRows(s, nP, nW, k), "#,##0") & "|" & _
                        Format$(fRows, "#,##0") & "|" & Format$(jRows, "#,##0") & "|" & Format$(peakRows, "#,##0") & "|" & _
                        Format$(modelS, "0.00") & " (" & Format$(loS, "0.00") & "-" & Format$(hiS, "0.00") & ")|" & guardText
        Next ri
    Next s
End Sub

' Added 2026-09-19 after step 7: its per-row cost rose by half at the one rung
' whose relation held ~350,000 keyed rows. The suspect, named and not measured:
' a Scripting.Dictionary slowing as it grows. This times RelTryAdd's own shape
' with no engine in the way - a four-column TupleKey (CStr per cell, Chr$(31)
' after each), Exists, Add, and the tuple kept in a Collection - at growing
' sizes, beside the key build alone and the teardown. A flat "add" column
' clears the dictionary; a rising one convicts it.
Public Sub O0DictCost(Optional ByVal maxKeys As Long = 400000)
    Dim sizes As Variant, si As Long, n As Long, i As Long
    Dim d As Object, tupleBag As Collection
    Dim t() As Variant
    Dim key As String, sep As String
    Dim t0 As Double, tKeys As Double, tAdd As Double, tFree As Double
    sizes = Array(10000, 50000, 100000, 200000, 400000)
    sep = Chr$(31)
    Debug.Print "O0D|keys|key build us/key|exists+add+keep us/key|teardown s"
    For si = LBound(sizes) To UBound(sizes)
        n = CLng(sizes(si))
        If n > maxKeys Then Exit For
        t0 = Timer
        For i = 1 To n
            key = CStr(i \ 7056) & sep & CStr((i \ 84) Mod 84) & sep & CStr(i Mod 84) & sep & "1" & sep
        Next i
        tKeys = ElapsedSince(t0)
        Set d = CreateObject("Scripting.Dictionary")
        Set tupleBag = New Collection
        ReDim t(1 To 4)
        t0 = Timer
        For i = 1 To n
            t(1) = i \ 7056
            t(2) = (i \ 84) Mod 84
            t(3) = i Mod 84
            t(4) = 1
            key = CStr(t(1)) & sep & CStr(t(2)) & sep & CStr(t(3)) & sep & CStr(t(4)) & sep
            If Not d.Exists(key) Then
                d.Add key, True
                tupleBag.Add t
            End If
        Next i
        tAdd = ElapsedSince(t0)
        t0 = Timer
        Set d = Nothing
        Set tupleBag = Nothing
        tFree = ElapsedSince(t0)
        Debug.Print "O0D|" & Format$(n, "#,##0") & "|" & Format$(tKeys / n * 1000000#, "0.00") & "|" & _
                    Format$(tAdd / n * 1000000#, "0.00") & "|" & Format$(tFree, "0.000")
        If tKeys + tAdd + tFree > LIMIT_SECONDS Then
            Debug.Print "O0D|stopped: that size took over " & LIMIT_SECONDS & "s"
            Exit For
        End If
    Next si
End Sub

Private Function ElapsedSince(ByVal t0 As Double) As Double
    ElapsedSince = Timer - t0
    If ElapsedSince < 0 Then ElapsedSince = ElapsedSince + 86400#
End Function

' A worksheet function that does nothing but hand Excel an array of numbers of
' the given shape: the control beside every answer, and step 12's ladder.
Public Function O0Cells(ByVal nRows As Long, ByVal nCols As Long) As Variant
    Dim a() As Variant
    Dim r As Long, c As Long
    If nRows < 1 Then nRows = 1
    If nCols < 1 Then nCols = 1
    If nRows > 1048575 Then nRows = 1048575
    ReDim a(1 To nRows, 1 To nCols)
    For r = 1 To nRows
        For c = 1 To nCols
            a(r, c) = r
        Next c
    Next r
    O0Cells = a
End Function

' ---------------------------------------------------------------------------
' The fixture generator: the reference roster's Tables, for every later item
' ---------------------------------------------------------------------------

' Writes People, Shifts, Leave, Next and Kept for `people` over `weeks` on a
' fresh sheet O0F<people>x<weeks> (suffix L for the loose variant), and prints
' what the definitions say about them. Tight (the reference, settled
' 2026-09-19): every shift needs max(1, people \ 5), a tightness of 0.84 when
' people is a multiple of 5. Loose: Early 2, Late 3, Night 4 (the entry's
' first proposal, tightness 63 / (5 x people) a week).
Public Sub O0Fixture(ByVal people As Long, ByVal weeks As Long, Optional ByVal loose As Boolean = False)
    Dim sfx As String
    Dim ws As Worksheet
    Dim nS As Long, s As Long, p As Long, d As Long, w As Long, j As Long, r As Long, n As Long
    Dim a() As Variant
    Dim kept() As Variant
    Dim onShift() As Boolean
    Dim perWeek() As Long
    Dim leaveClash As Long, inARow As Long, overFive As Long, nightNoSenior As Long, overtimePW As Long, leaveN As Long
    Dim demand As Double, seniorsHere As Long, hasSenior As Boolean
    Dim prevScreen As Boolean
    Dim errNo As Long, errText As String

    If people < 1 Or weeks < 1 Then
        Debug.Print "O0F: people and weeks must both be at least 1."
        Exit Sub
    End If
    sfx = people & "x" & weeks
    If loose Then sfx = sfx & "L"
    If SheetExists(ActiveWorkbook, "O0F" & sfx) Then
        Debug.Print "O0F: sheet O0F" & sfx & " already exists - delete it first."
        Exit Sub
    End If
    nS = 21 * weeks

    prevScreen = Application.ScreenUpdating
    On Error GoTo failFix
    Application.ScreenUpdating = False
    Set ws = ActiveWorkbook.Worksheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
    ws.Name = "O0F" & sfx

    ReDim a(1 To people, 1 To 4)
    For p = 1 To people
        a(p, 1) = p
        a(p, 2) = "P" & p
        If IsSenior(p) Then a(p, 3) = "yes" Else a(p, 3) = "no"
        a(p, 4) = ContractOf(p)
    Next p
    WriteTable NewTable(ws, "A1", "PeopleO0f" & sfx, Array("Id", "Name", "Senior", "Contract")), a, people

    ReDim a(1 To nS, 1 To 6)
    For s = 1 To nS
        a(s, 1) = s
        a(s, 2) = ShiftWeek(s)
        a(s, 3) = ShiftDay(s)
        a(s, 4) = SlotName(ShiftSlot(s))
        a(s, 5) = NeedOf(s, people, loose)
        If ShiftSlot(s) = 2 Then a(s, 6) = "senior" Else a(s, 6) = "any"
        demand = demand + NeedOf(s, people, loose)
    Next s
    WriteTable NewTable(ws, "F1", "ShiftsO0f" & sfx, Array("Id", "Week", "Day", "Slot", "Need", "Skill")), a, nS

    ReDim a(1 To people * weeks, 1 To 2)
    r = 0
    For p = 1 To people
        For d = 1 To 7 * weeks
            If OnLeave(p, d) Then
                r = r + 1
                a(r, 1) = p
                a(r, 2) = d
            End If
        Next d
    Next p
    leaveN = r
    WriteTable NewTable(ws, "M1", "LeaveO0f" & sfx, Array("Person", "Day")), a, leaveN

    If nS > 1 Then
        ReDim a(1 To nS - 1, 1 To 2)
        For s = 1 To nS - 1
            a(s, 1) = s
            a(s, 2) = s + 1
        Next s
        WriteTable NewTable(ws, "P1", "NextO0f" & sfx, Array("From", "To")), a, nS - 1
    End If

    kept = KeptRows(people, weeks, loose, n)
    WriteTable NewTable(ws, "S1", "KeptO0f" & sfx, Array("Shift", "Person", "Week")), kept, n

    ' What the kept schedule breaks, counted from its rows.
    ReDim onShift(1 To nS, 1 To people)
    ReDim perWeek(1 To people, 1 To weeks)
    For r = 1 To n
        s = kept(r, 1)
        p = kept(r, 2)
        onShift(s, p) = True
        perWeek(p, ShiftWeek(s)) = perWeek(p, ShiftWeek(s)) + 1
        If OnLeave(p, ShiftDay(s)) Then leaveClash = leaveClash + 1
    Next r
    For s = 1 To nS
        If s < nS Then
            For p = 1 To people
                If onShift(s, p) And onShift(s + 1, p) Then inARow = inARow + 1
            Next p
        End If
        If ShiftSlot(s) = 2 Then
            hasSenior = False
            For p = 1 To people
                If onShift(s, p) And IsSenior(p) Then hasSenior = True
            Next p
            If Not hasSenior Then nightNoSenior = nightNoSenior + 1
        End If
    Next s
    For p = 1 To people
        For w = 1 To weeks
            If perWeek(p, w) > 5 Then overFive = overFive + 1
            If perWeek(p, w) > ContractOf(p) Then overtimePW = overtimePW + 1
        Next w
    Next p
    seniorsHere = people \ 4

    Debug.Print "=== O0F fixture " & sfx & " on sheet " & ws.Name & " ==="
    Debug.Print "O0F|people|" & people
    Debug.Print "O0F|weeks|" & weeks
    Debug.Print "O0F|shifts|" & nS
    If loose Then
        Debug.Print "O0F|need|Early 2, Late 3, Night 4 (loose)"
    Else
        Debug.Print "O0F|need|" & TightNeed(people) & " on every shift (tight)"
    End If
    Debug.Print "O0F|weekly demand|" & demand / weeks
    Debug.Print "O0F|weekly capacity at 5 each|" & 5 * people
    Debug.Print "O0F|tightness|" & Format$(demand / weeks / (5# * people), "0.00")
    Debug.Print "O0F|seniors|" & seniorsHere
    Debug.Print "O0F|leave rows|" & leaveN
    Debug.Print "O0F|pool, full|" & people * nS
    Debug.Print "O0F|pool, minus leave|" & 18 * people * weeks
    Debug.Print "O0F|kept rows|" & n
    Debug.Print "O0F|kept on leave|" & leaveClash
    Debug.Print "O0F|kept two in a row|" & inARow
    Debug.Print "O0F|kept over five a week|" & overFive
    Debug.Print "O0F|kept nights with no senior|" & nightNoSenior
    Debug.Print "O0F|kept person-weeks over contract|" & overtimePW
    If demand / weeks > 5# * people Then
        Debug.Print "O0F|a schedule exists|NO, provably: the shifts need " & demand / weeks & " people a week, and " & people & " people may work at most 5 each"
    ElseIf 5 * seniorsHere < 7 Then
        Debug.Print "O0F|a schedule exists|NO, provably: 7 nights a week each need a senior, and " & seniorsHere & " senior(s) may work at most 5 a week each"
    Else
        Debug.Print "O0F|a schedule exists|not decided by hand - tools/optimize0_lp.ps1 exports this fixture for clingo"
    End If
    Debug.Print "=== O0F done ==="

cleanupFix:
    On Error Resume Next
    Application.ScreenUpdating = prevScreen
    On Error GoTo 0
    Exit Sub
failFix:
    errNo = Err.Number
    errText = Err.Description
    Debug.Print "O0F: stopped by VBA error " & errNo & ": " & errText
    Resume cleanupFix
End Sub

' ---------------------------------------------------------------------------
' Steps: their shapes, rules, Tables and ladders
' ---------------------------------------------------------------------------

Private Function ShapeName(ByVal stepNo As Long) As String
    Select Case stepNo
    Case 1: ShapeName = "the pool - anyone, any shift (a Cartesian body)"
    Case 2: ShapeName = "the pool minus vacations (not)"
    Case 3: ShapeName = "the pool through a skill join"
    Case 4: ShapeName = "a native counter per shift (count)"
    Case 5: ShapeName = "a native counter per person-week (count)"
    Case 6: ShapeName = "never two in a row, through the pair relation Next"
    Case 7: ShapeName = "never two in a row, by arithmetic over all pairs"
    Case 8: ShapeName = "at most k per shift, as (k+1)-subsets in canonical order"
    Case 9: ShapeName = "at most 2 per shift, as triples WITHOUT canonical order"
    Case 10: ShapeName = "at most k a week, as (k+1)-subsets of a person's week"
    Case 11: ShapeName = "overtime per person-week, against the kept schedule"
    Case 12: ShapeName = "control: a bare UDF returning an array, no engine"
    End Select
End Function

' Each rung is "people,weeks" or "people,weeks,k"; step 12's is "rows,cols".
' Ordered by the model's own cost, so a rung over the limit ends the step.
Private Function RungList(ByVal stepNo As Long) As Variant
    Select Case stepNo
    Case 1, 2, 3, 4, 5, 6, 7, 11
        RungList = Split("5,1 10,1 5,4 20,1 10,4 50,1 20,4 50,4", " ")
    Case 8
        RungList = Split("5,1 7,1 10,1 5,4 14,1 7,4 20,1 10,4 14,4 20,4 50,1 50,4 50,4,10", " ")
    Case 9
        RungList = Split("5,1 7,1 10,1 5,4 14,1 7,4 20,1 10,4 14,4 20,4 50,1 50,4", " ")
    Case 10
        RungList = Split("1,1,1 5,1,1 1,1,2 10,1,1 20,1,1 5,1,2 1,1,3 10,1,2 20,1,2 5,1,3 10,1,3 1,1,5 20,1,3 5,1,5 50,4,5", " ")
    Case 12
        RungList = Split("1000,4 10000,4 100000,4 100000,8 430000,4 430000,8", " ")
    End Select
End Function

Private Sub ParseRung(ByVal spec As String, ByRef nP As Long, ByRef nW As Long, ByRef k As Long)
    Dim bitsOf As Variant
    bitsOf = Split(spec, ",")
    nP = CLng(bitsOf(0))
    nW = CLng(bitsOf(1))
    k = 2
    If UBound(bitsOf) >= 2 Then k = CLng(bitsOf(2))
End Sub

Private Function RungText(ByVal stepNo As Long, ByVal nP As Long, ByVal nW As Long, ByVal k As Long) As String
    If stepNo = LAST_STEP Then
        RungText = Format$(nP, "#,##0") & " rows x " & nW & " columns"
    Else
        RungText = nP & " people x " & nW & " week(s)"
        If stepNo = 8 Or stepNo = 10 Then RungText = RungText & ", k = " & k
    End If
End Function

Private Function KText(ByVal stepNo As Long, ByVal k As Long) As String
    If stepNo = 8 Or stepNo = 10 Then KText = CStr(k)
End Function

Private Function RulesText(ByVal stepNo As Long, ByVal k As Long) As String
    Select Case stepNo
    Case 1: RulesText = "(rule (o0g S P) (shiftso0x1 S) (peopleo0x1 P)) (query o0g)"
    Case 2: RulesText = "(rule (o0g S P) (shiftso0x2 S D) (peopleo0x2 P) (not (leaveo0x2 P D))) (query o0g)"
    Case 3: RulesText = "(rule (o0g S P) (shiftso0x3 S K) (skillso0x3 P K)) (query o0g)"
    Case 4: RulesText = "(rule (o0g S N) (shiftso0x4 S) (count N (eligo0x4 S P))) (query o0g)"
    Case 5: RulesText = "(rule (o0g P W N) (pwo0x5 P W) (count N (eligo0x5 S P W))) (query o0g)"
    Case 6: RulesText = "(rule (o0g A B P) (nexto0x6 A B) (eligo0x6 A P) (eligo0x6 B P)) (query o0g)"
    Case 7: RulesText = "(rule (o0g A B P) (eligo0x7 A P) (eligo0x7 B P) (let D (- B A)) (= D 1)) (query o0g)"
    Case 8: RulesText = SubsetRule("(o0g S", "eligo0x8", "S #", k)
    Case 9: RulesText = "(rule (o0g S A B C) (eligo0x9 S A) (eligo0x9 S B) (<> A B) (eligo0x9 S C) (<> A C) (<> B C)) (query o0g)"
    Case 10: RulesText = SubsetRule("(o0g P W", "eligo0x10", "# P W", k)
    Case 11: RulesText = "(rule (o0w P W N) (pwo0x11 P W) (count N (kepto0x11 S P W))) " & _
                         "(rule (o0g P W O) (o0w P W N) (peopleo0x11 P C) (let O (- N C)) (> O 0)) (query o0g)"
    End Select
End Function

' (rule <head X1 .. Xm>) (pred <args with # = Xj>) ... (< Xj-1 Xj) ...) with
' m = k + 1 members in canonical order: each join is followed at once by the
' comparison it makes possible, the cheapest order DATALOG can be given.
Private Function SubsetRule(ByVal headStart As String, ByVal pred As String, ByVal argPattern As String, _
                            ByVal k As Long) As String
    Dim t As String, j As Long
    t = "(rule " & headStart
    For j = 1 To k + 1
        t = t & " X" & j
    Next j
    t = t & ")"
    For j = 1 To k + 1
        t = t & " (" & pred & " " & Replace(argPattern, "#", "X" & j) & ")"
        If j > 1 Then t = t & " (< X" & (j - 1) & " X" & j & ")"
    Next j
    SubsetRule = t & ") (query o0g)"
End Function

Private Function AnswerCols(ByVal stepNo As Long, ByVal k As Long) As Long
    Select Case stepNo
    Case 1, 2, 3, 4: AnswerCols = 2
    Case 5, 6, 7, 11: AnswerCols = 3
    Case 8: AnswerCols = k + 2
    Case 9: AnswerCols = 4
    Case 10: AnswerCols = k + 3
    End Select
End Function

' Creates the step's Tables (headers plus one placeholder row), its rules cell,
' its formula and the control, so the first timed rung calculates from clean.
Private Sub SetupStep(ByVal stepNo As Long, ByVal ws As Worksheet, ByVal k As Long)
    Dim refs As String
    Select Case stepNo
    Case 1
        NewTable ws, "A1", "ShiftsO0x1", Array("Shift")
        NewTable ws, "E1", "PeopleO0x1", Array("Person")
        refs = "ShiftsO0x1,PeopleO0x1"
    Case 2
        NewTable ws, "A1", "ShiftsO0x2", Array("Shift", "Day")
        NewTable ws, "E1", "PeopleO0x2", Array("Person")
        NewTable ws, "I1", "LeaveO0x2", Array("Person", "Day")
        refs = "ShiftsO0x2,PeopleO0x2,LeaveO0x2"
    Case 3
        NewTable ws, "A1", "ShiftsO0x3", Array("Shift", "Skill")
        NewTable ws, "E1", "SkillsO0x3", Array("Person", "Skill")
        refs = "ShiftsO0x3,SkillsO0x3"
    Case 4
        NewTable ws, "A1", "ShiftsO0x4", Array("Shift")
        NewTable ws, "E1", "EligO0x4", Array("Shift", "Person")
        refs = "ShiftsO0x4,EligO0x4"
    Case 5
        NewTable ws, "A1", "PwO0x5", Array("Person", "Week")
        NewTable ws, "E1", "EligO0x5", Array("Shift", "Person", "Week")
        refs = "PwO0x5,EligO0x5"
    Case 6
        NewTable ws, "A1", "NextO0x6", Array("From", "To")
        NewTable ws, "E1", "EligO0x6", Array("Shift", "Person")
        refs = "NextO0x6,EligO0x6"
    Case 7, 8, 9
        NewTable ws, "A1", "EligO0x" & stepNo, Array("Shift", "Person")
        refs = "EligO0x" & stepNo
    Case 10
        NewTable ws, "A1", "EligO0x10", Array("Shift", "Person", "Week")
        refs = "EligO0x10"
    Case 11
        NewTable ws, "A1", "KeptO0x11", Array("Shift", "Person", "Week")
        NewTable ws, "E1", "PwO0x11", Array("Person", "Week")
        NewTable ws, "I1", "PeopleO0x11", Array("Person", "Contract")
        refs = "KeptO0x11,PwO0x11,PeopleO0x11"
    End Select
    ws.Range(CTRL_ROWS_ADDR).Value2 = 1
    ws.Range(CTRL_COLS_ADDR).Value2 = 1
    ws.Range(CTRL_ADDR).Formula2 = "=O0Cells(" & CTRL_ROWS_ADDR & "," & CTRL_COLS_ADDR & ")"
    If stepNo < LAST_STEP Then
        ws.Range(RULES_ADDR).Value2 = RulesText(stepNo, k)
        ws.Range(ANSWER_ADDR).Formula2 = "=DATALOG(" & RULES_ADDR & "," & refs & ")"
    End If
End Sub

' Rewrites the step's Tables for one rung, and its rules cell (k can change).
Private Sub FillStep(ByVal stepNo As Long, ByVal ws As Worksheet, ByVal nP As Long, ByVal nW As Long, ByVal k As Long)
    Dim a() As Variant
    Dim nS As Long, n As Long, r As Long, s As Long, p As Long, d As Long, w As Long
    nS = 21 * nW
    Select Case stepNo
    Case 1
        a = IdColumn(nS)
        WriteTable TableOn(ws, "ShiftsO0x1"), a, nS
        a = IdColumn(nP)
        WriteTable TableOn(ws, "PeopleO0x1"), a, nP
    Case 2
        ReDim a(1 To nS, 1 To 2)
        For s = 1 To nS
            a(s, 1) = s
            a(s, 2) = ShiftDay(s)
        Next s
        WriteTable TableOn(ws, "ShiftsO0x2"), a, nS
        a = IdColumn(nP)
        WriteTable TableOn(ws, "PeopleO0x2"), a, nP
        ReDim a(1 To nP * nW, 1 To 2)
        r = 0
        For p = 1 To nP
            For d = 1 To 7 * nW
                If OnLeave(p, d) Then
                    r = r + 1
                    a(r, 1) = p
                    a(r, 2) = d
                End If
            Next d
        Next p
        WriteTable TableOn(ws, "LeaveO0x2"), a, r
    Case 3
        ReDim a(1 To nS, 1 To 2)
        For s = 1 To nS
            a(s, 1) = s
            If ShiftSlot(s) = 2 Then a(s, 2) = "senior" Else a(s, 2) = "any"
        Next s
        WriteTable TableOn(ws, "ShiftsO0x3"), a, nS
        ReDim a(1 To nP + nP \ 4, 1 To 2)
        r = 0
        For p = 1 To nP
            r = r + 1
            a(r, 1) = p
            a(r, 2) = "any"
        Next p
        For p = 1 To nP
            If IsSenior(p) Then
                r = r + 1
                a(r, 1) = p
                a(r, 2) = "senior"
            End If
        Next p
        WriteTable TableOn(ws, "SkillsO0x3"), a, r
    Case 4
        a = IdColumn(nS)
        WriteTable TableOn(ws, "ShiftsO0x4"), a, nS
        a = PoolRows(nP, nW, True, False, n)
        WriteTable TableOn(ws, "EligO0x4"), a, n
    Case 5
        ReDim a(1 To nP * nW, 1 To 2)
        r = 0
        For p = 1 To nP
            For w = 1 To nW
                r = r + 1
                a(r, 1) = p
                a(r, 2) = w
            Next w
        Next p
        WriteTable TableOn(ws, "PwO0x5"), a, r
        a = PoolRows(nP, nW, True, True, n)
        WriteTable TableOn(ws, "EligO0x5"), a, n
    Case 6
        ReDim a(1 To nS - 1, 1 To 2)
        For s = 1 To nS - 1
            a(s, 1) = s
            a(s, 2) = s + 1
        Next s
        WriteTable TableOn(ws, "NextO0x6"), a, nS - 1
        a = PoolRows(nP, nW, False, False, n)
        WriteTable TableOn(ws, "EligO0x6"), a, n
    Case 7, 8, 9
        a = PoolRows(nP, nW, False, False, n)
        WriteTable TableOn(ws, "EligO0x" & stepNo), a, n
    Case 10
        a = PoolRows(nP, nW, False, True, n)
        WriteTable TableOn(ws, "EligO0x10"), a, n
    Case 11
        a = KeptRows(nP, nW, False, n)
        WriteTable TableOn(ws, "KeptO0x11"), a, n
        ReDim a(1 To nP * nW, 1 To 2)
        r = 0
        For p = 1 To nP
            For w = 1 To nW
                r = r + 1
                a(r, 1) = p
                a(r, 2) = w
            Next w
        Next p
        WriteTable TableOn(ws, "PwO0x11"), a, r
        ReDim a(1 To nP, 1 To 2)
        For p = 1 To nP
            a(p, 1) = p
            a(p, 2) = ContractOf(p)
        Next p
        WriteTable TableOn(ws, "PeopleO0x11"), a, nP
    Case 12
        Exit Sub
    End Select
    ws.Range(RULES_ADDR).Value2 = RulesText(stepNo, k)
End Sub

' ---------------------------------------------------------------------------
' The fixture's definitions - every one a closed form, no generator of chance
' ---------------------------------------------------------------------------

' Shift s (1-based) is on day (s-1)\3+1, in slot (s-1) mod 3 (0 Early, 1 Late,
' 2 Night), in week (day-1)\7+1.
Private Function ShiftDay(ByVal s As Long) As Long
    ShiftDay = (s - 1) \ 3 + 1
End Function

Private Function ShiftSlot(ByVal s As Long) As Long
    ShiftSlot = (s - 1) Mod 3
End Function

Private Function ShiftWeek(ByVal s As Long) As Long
    ShiftWeek = (ShiftDay(s) - 1) \ 7 + 1
End Function

Private Function SlotName(ByVal slot As Long) As String
    Select Case slot
    Case 0: SlotName = "Early"
    Case 1: SlotName = "Late"
    Case Else: SlotName = "Night"
    End Select
End Function

' Person i is on leave on day d when (i + d) mod 7 = 0: one day a week each.
Private Function OnLeave(ByVal person As Long, ByVal dayNo As Long) As Boolean
    OnLeave = ((person + dayNo) Mod 7 = 0)
End Function

Private Function OffCount(ByVal nP As Long, ByVal dayNo As Long) As Long
    Dim p As Long
    For p = 1 To nP
        If OnLeave(p, dayNo) Then OffCount = OffCount + 1
    Next p
End Function

Private Function IsSenior(ByVal person As Long) As Boolean
    IsSenior = (person Mod 4 = 0)
End Function

Private Function ContractOf(ByVal person As Long) As Long
    If person Mod 2 = 1 Then ContractOf = 4 Else ContractOf = 5
End Function

Private Function TightNeed(ByVal nP As Long) As Long
    TightNeed = nP \ 5
    If TightNeed < 1 Then TightNeed = 1
End Function

Private Function NeedOf(ByVal s As Long, ByVal nP As Long, ByVal loose As Boolean) As Long
    If loose Then
        NeedOf = 2 + ShiftSlot(s)
    Else
        NeedOf = TightNeed(nP)
    End If
End Function

' The kept schedule: a round-robin. Shift s takes the next NeedOf(s) people in
' Id order after the shift before it, wrapping at nP.
Private Function KeptStart(ByVal s As Long, ByVal nP As Long, ByVal loose As Boolean) As Long
    If loose Then
        Select Case ShiftSlot(s)
        Case 0: KeptStart = 9 * ((s - 1) \ 3)
        Case 1: KeptStart = 9 * ((s - 1) \ 3) + 2
        Case Else: KeptStart = 9 * ((s - 1) \ 3) + 5
        End Select
    Else
        KeptStart = TightNeed(nP) * (s - 1)
    End If
End Function

Private Function KeptRows(ByVal nP As Long, ByVal nW As Long, ByVal loose As Boolean, ByRef n As Long) As Variant()
    Dim a() As Variant
    Dim nS As Long, s As Long, j As Long, total As Long
    nS = 21 * nW
    For s = 1 To nS
        total = total + NeedOf(s, nP, loose)
    Next s
    ReDim a(1 To total, 1 To 3)
    n = 0
    For s = 1 To nS
        For j = 0 To NeedOf(s, nP, loose) - 1
            n = n + 1
            a(n, 1) = s
            a(n, 2) = ((KeptStart(s, nP, loose) + j) Mod nP) + 1
            a(n, 3) = ShiftWeek(s)
        Next j
    Next s
    KeptRows = a
End Function

' The pool, shift-major then person, so Table order is the canonical order.
Private Function PoolRows(ByVal nP As Long, ByVal nW As Long, ByVal minusLeave As Boolean, _
                          ByVal withWeek As Boolean, ByRef n As Long) As Variant()
    Dim a() As Variant
    Dim nS As Long, s As Long, p As Long, nC As Long
    nS = 21 * nW
    nC = 2
    If withWeek Then nC = 3
    ReDim a(1 To nS * nP, 1 To nC)
    n = 0
    For s = 1 To nS
        For p = 1 To nP
            If Not (minusLeave And OnLeave(p, ShiftDay(s))) Then
                n = n + 1
                a(n, 1) = s
                a(n, 2) = p
                If withWeek Then a(n, 3) = ShiftWeek(s)
            End If
        Next p
    Next s
    PoolRows = a
End Function

Private Function IdColumn(ByVal n As Long) As Variant()
    Dim a() As Variant, i As Long
    ReDim a(1 To n, 1 To 1)
    For i = 1 To n
        a(i, 1) = i
    Next i
    IdColumn = a
End Function

' Overtime per person-week in the tight kept schedule: shifts worked minus the
' contract, as the step-11 rule computes it (a row only where it is above 0).
Private Function OvertimeTable(ByVal nP As Long, ByVal nW As Long) As Long()
    Dim t() As Long
    Dim a() As Variant
    Dim n As Long, r As Long, p As Long, w As Long
    ReDim t(1 To nP, 1 To nW)
    a = KeptRows(nP, nW, False, n)
    For r = 1 To n
        t(a(r, 2), a(r, 3)) = t(a(r, 2), a(r, 3)) + 1
    Next r
    For p = 1 To nP
        For w = 1 To nW
            t(p, w) = t(p, w) - ContractOf(p)
        Next w
    Next p
    OvertimeTable = t
End Function

Private Function OvertimeRows(ByVal nP As Long, ByVal nW As Long) As Long
    Dim t() As Long, p As Long, w As Long
    t = OvertimeTable(nP, nW)
    For p = 1 To nP
        For w = 1 To nW
            If t(p, w) > 0 Then OvertimeRows = OvertimeRows + 1
        Next w
    Next p
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

' The answer's data rows (header not counted), from the definitions.
Private Function ExpectedRows(ByVal stepNo As Long, ByVal nP As Long, ByVal nW As Long, ByVal k As Long) As Double
    Dim nS As Double
    nS = 21# * nW
    Select Case stepNo
    Case 1: ExpectedRows = nP * nS
    Case 2: ExpectedRows = 18# * nP * nW
    Case 3: ExpectedRows = 14# * nW * nP + 7# * nW * (nP \ 4)
    Case 4: ExpectedRows = nS
    Case 5: ExpectedRows = CDbl(nP) * nW
    Case 6, 7: ExpectedRows = nP * (nS - 1)
    Case 8: ExpectedRows = nS * Comb(nP, k + 1)
    Case 9: ExpectedRows = nS * nP * (nP - 1) * (nP - 2)
    Case 10: ExpectedRows = CDbl(nP) * nW * Comb(21, k + 1)
    Case 11: ExpectedRows = OvertimeRows(nP, nW)
    Case 12: ExpectedRows = nP
    End Select
End Function

' ---------------------------------------------------------------------------
' The model: time = a*F + b*J, peak = the largest single output
' ---------------------------------------------------------------------------

Private Function ModelSeconds(ByVal stepNo As Long, ByVal nP As Long, ByVal nW As Long, ByVal k As Long, ByVal bMs As Double) As Double
    BuildModel stepNo, nP, nW, k
    ModelSeconds = (A_MS * mF + bMs * mJ) / 1000#
End Function

Private Sub BuildModel(ByVal stepNo As Long, ByVal nP As Long, ByVal nW As Long, ByVal k As Long)
    Dim nS As Double, pool As Double, pruned As Double, pw As Double, j As Long
    Dim g As Double, groups As Double, outRows As Double, ot As Double, keptN As Double
    mF = 0
    mJ = 0
    mPeak = 0
    nS = 21# * nW
    pool = nP * nS
    pruned = 18# * nP * nW
    pw = CDbl(nP) * nW
    Select Case stepNo
    Case 1
        TablesRead nS + nP
        RuleBegin
        RuleItem IK_BASE, nS, nS
        RuleItem IK_BASE, nP, pool
        RuleEnd pool
    Case 2
        TablesRead nS + nP + pw
        RuleBegin
        RuleItem IK_BASE, nS, nS
        RuleItem IK_BASE, nP, pool
        RuleItem IK_OTHER, pw, pruned
        RuleEnd pruned
    Case 3
        outRows = ExpectedRows(3, nP, nW, k)
        TablesRead nS + nP + (nP \ 4)
        RuleBegin
        RuleItem IK_BASE, nS, nS
        RuleItem IK_BASE, nP + (nP \ 4), outRows
        RuleEnd outRows
    Case 4
        TablesRead nS + pruned
        RuleBegin
        RuleItem IK_BASE, nS, nS
        RuleItem IK_OTHER, pruned, nS
        RuleEnd nS
    Case 5
        TablesRead pw + pruned
        RuleBegin
        RuleItem IK_BASE, pw, pw
        RuleItem IK_OTHER, pruned, pw
        RuleEnd pw
    Case 6
        TablesRead nS - 1 + pool
        RuleBegin
        RuleItem IK_BASE, nS - 1, nS - 1
        RuleItem IK_BASE, pool, (nS - 1) * nP
        RuleItem IK_BASE, pool, (nS - 1) * nP
        RuleEnd (nS - 1) * nP
    Case 7
        TablesRead pool
        RuleBegin
        RuleItem IK_BASE, pool, pool
        RuleItem IK_BASE, pool, nP * nS * nS
        RuleItem IK_OTHER, 0, nP * nS * nS
        RuleItem IK_OTHER, 0, nP * (nS - 1)
        RuleEnd nP * (nS - 1)
    Case 8, 10
        If stepNo = 8 Then
            groups = nS
            g = nP
        Else
            groups = pw
            g = 21
        End If
        TablesRead pool
        RuleBegin
        RuleItem IK_BASE, pool, groups * g
        For j = 2 To k + 1
            RuleItem IK_BASE, pool, groups * g * Comb(CLng(g), j - 1)
            RuleItem IK_OTHER, 0, groups * Comb(CLng(g), j)
        Next j
        RuleEnd groups * Comb(CLng(g), k + 1)
    Case 9
        TablesRead pool
        RuleBegin
        RuleItem IK_BASE, pool, pool
        RuleItem IK_BASE, pool, nS * nP * nP
        RuleItem IK_OTHER, 0, nS * nP * (nP - 1)
        RuleItem IK_BASE, pool, nS * nP * (nP - 1) * nP
        RuleItem IK_OTHER, 0, nS * nP * (nP - 1) * (nP - 1)
        RuleItem IK_OTHER, 0, nS * nP * (nP - 1) * (nP - 2)
        RuleEnd nS * nP * (nP - 1) * (nP - 2)
    Case 11
        ot = OvertimeRows(nP, nW)
        keptN = nS * TightNeed(nP)
        TablesRead keptN + pw + nP
        RuleBegin
        RuleItem IK_BASE, pw, pw
        RuleItem IK_OTHER, keptN, pw
        RuleEnd pw
        RuleBegin
        RuleItem IK_DERIVED, pw, pw
        RuleItem IK_BASE, nP, pw
        RuleItem IK_OTHER, 0, pw
        RuleItem IK_OTHER, 0, ot
        RuleEnd ot
    End Select
End Sub

' Reading a Table argument puts every row through RelTryAdd: produced rows.
Private Sub TablesRead(ByVal nRows As Double)
    mJ = mJ + nRows
End Sub

Private Sub RuleBegin()
    mItems = 0
End Sub

Private Sub RuleItem(ByVal kind As Long, ByVal srcRows As Double, ByVal outRows As Double)
    mItems = mItems + 1
    mKind(mItems) = kind
    mSrc(mItems) = srcRows
    mOut(mItems) = outRows
End Sub

' Round 1 evaluates the rule once. Round 2 evaluates it once per positive
' body position: at a Table's position the items ahead of it run again and
' then the missing delta ends it; at a derived predicate's position its delta
' is everything round 1 made, so the whole rule runs again. The spill of the
' head's rows is counted with the head.
Private Sub RuleEnd(ByVal headRows As Double)
    Dim i As Long, p As Long
    For i = 1 To mItems
        mF = mF + mSrc(i)
        mJ = mJ + mOut(i)
        If mOut(i) > mPeak Then mPeak = mOut(i)
    Next i
    mJ = mJ + 2# * headRows
    For p = 1 To mItems
        If mKind(p) = IK_BASE Then
            For i = 1 To p - 1
                mF = mF + mSrc(i)
                mJ = mJ + mOut(i)
            Next i
        ElseIf mKind(p) = IK_DERIVED Then
            For i = 1 To mItems
                mF = mF + mSrc(i)
                mJ = mJ + mOut(i)
            Next i
            mJ = mJ + headRows
        End If
    Next p
End Sub

' ---------------------------------------------------------------------------
' Checking an answer against the definitions
' ---------------------------------------------------------------------------

Private Function VerifyAnswer(ByVal stepNo As Long, ByVal ans As Range, ByVal nP As Long, ByVal nW As Long, _
                              ByVal k As Long, ByVal ground As Double) As String
    Dim v As Variant, arr As Variant
    Dim shown As String
    v = ans.Value2
    If IsError(v) Then
        VerifyAnswer = "WRONG: Excel shows " & ans.Text
        Exit Function
    End If
    If Not ans.HasSpill Then
        shown = CStr(v)
        If Left$(shown, 9) = "#DATALOG!" Then
            VerifyAnswer = "REFUSED: " & Left$(shown, 300)
        Else
            VerifyAnswer = "WRONG: no spill; the cell shows " & Left$(shown, 200)
        End If
        Exit Function
    End If
    arr = ans.SpillingToRange.Value2
    If UBound(arr, 1) - 1 <> ground Then
        VerifyAnswer = "WRONG: " & Format$(UBound(arr, 1) - 1, "#,##0") & " rows, expected " & Format$(ground, "#,##0")
        Exit Function
    End If
    If UBound(arr, 2) <> AnswerCols(stepNo, k) Then
        VerifyAnswer = "WRONG: " & UBound(arr, 2) & " columns, expected " & AnswerCols(stepNo, k)
        Exit Function
    End If
    VerifyAnswer = VerifyRows(stepNo, arr, nP, nW, k)
End Function

Private Function VerifyControl(ByVal ctrl As Range, ByVal nRows As Long, ByVal nCols As Long) As String
    If Not ctrl.HasSpill Then
        If nRows = 1 And nCols = 1 Then
            VerifyControl = "ok"
        Else
            VerifyControl = "WRONG: the control did not spill"
        End If
        Exit Function
    End If
    If ctrl.SpillingToRange.Rows.Count <> nRows Or ctrl.SpillingToRange.Columns.Count <> nCols Then
        VerifyControl = "WRONG: the control spilled " & ctrl.SpillingToRange.Rows.Count & " x " & ctrl.SpillingToRange.Columns.Count
    Else
        VerifyControl = "ok"
    End If
End Function

' Every data row (arr's row 1 is the header) is checked against the fixture's
' definition, and every group's tally against what the definition says it is.
' Row count and width are already checked; DATALOG's answers are sets, so a
' right tally of valid rows in every group is a right answer.
Private Function VerifyRows(ByVal stepNo As Long, ByRef arr As Variant, ByVal nP As Long, ByVal nW As Long, ByVal k As Long) As String
    Dim nS As Long, r As Long, c As Long, i As Long, idx As Long, m As Long
    Dim tally() As Long, ot() As Long
    Dim x As Long, y As Long, z As Long, prevX As Long
    Dim ok As Boolean
    Dim want As Double
    Dim seen(1 To 3) As Long
    nS = 21 * nW
    Select Case stepNo
    Case 1, 2, 3, 4
        ReDim tally(1 To nS)
        For r = 2 To UBound(arr, 1)
            x = CellLong(arr(r, 1), ok)
            If Not ok Or x < 1 Or x > nS Then
                VerifyRows = RowFault(r, 1, arr(r, 1))
                Exit Function
            End If
            y = CellLong(arr(r, 2), ok)
            If Not ok Then
                VerifyRows = RowFault(r, 2, arr(r, 2))
                Exit Function
            End If
            If stepNo = 4 Then
                If y <> nP - OffCount(nP, ShiftDay(x)) Then
                    VerifyRows = "WRONG: shift " & x & " counts " & y & ", expected " & (nP - OffCount(nP, ShiftDay(x)))
                    Exit Function
                End If
            Else
                If y < 1 Or y > nP Then
                    VerifyRows = RowFault(r, 2, arr(r, 2))
                    Exit Function
                End If
                If stepNo = 2 Then
                    If OnLeave(y, ShiftDay(x)) Then
                        VerifyRows = "WRONG: row " & r & " puts person " & y & " on shift " & x & ", a leave day"
                        Exit Function
                    End If
                End If
                If stepNo = 3 Then
                    If ShiftSlot(x) = 2 And Not IsSenior(y) Then
                        VerifyRows = "WRONG: row " & r & " puts person " & y & ", not a senior, on night shift " & x
                        Exit Function
                    End If
                End If
            End If
            tally(x) = tally(x) + 1
        Next r
        For i = 1 To nS
            Select Case stepNo
            Case 1: want = nP
            Case 2: want = nP - OffCount(nP, ShiftDay(i))
            Case 3
                If ShiftSlot(i) = 2 Then want = nP \ 4 Else want = nP
            Case 4: want = 1
            End Select
            If tally(i) <> want Then
                VerifyRows = "WRONG: shift " & i & " has " & tally(i) & " rows, expected " & want
                Exit Function
            End If
        Next i
    Case 5, 11
        ReDim tally(1 To nP * nW)
        If stepNo = 11 Then ot = OvertimeTable(nP, nW)
        For r = 2 To UBound(arr, 1)
            x = CellLong(arr(r, 1), ok)
            If Not ok Or x < 1 Or x > nP Then
                VerifyRows = RowFault(r, 1, arr(r, 1))
                Exit Function
            End If
            y = CellLong(arr(r, 2), ok)
            If Not ok Or y < 1 Or y > nW Then
                VerifyRows = RowFault(r, 2, arr(r, 2))
                Exit Function
            End If
            z = CellLong(arr(r, 3), ok)
            If Not ok Then
                VerifyRows = RowFault(r, 3, arr(r, 3))
                Exit Function
            End If
            If stepNo = 5 Then
                want = 18
            Else
                want = ot(x, y)
            End If
            If z <> want Or z <= 0 Then
                VerifyRows = "WRONG: person " & x & ", week " & y & " shows " & z & ", expected " & want
                Exit Function
            End If
            idx = (x - 1) * nW + y
            tally(idx) = tally(idx) + 1
            If tally(idx) > 1 Then
                VerifyRows = "WRONG: person " & x & ", week " & y & " appears twice"
                Exit Function
            End If
        Next r
    Case 6, 7
        ReDim tally(1 To nP)
        For r = 2 To UBound(arr, 1)
            x = CellLong(arr(r, 1), ok)
            If Not ok Or x < 1 Or x >= nS Then
                VerifyRows = RowFault(r, 1, arr(r, 1))
                Exit Function
            End If
            y = CellLong(arr(r, 2), ok)
            If Not ok Or y <> x + 1 Then
                VerifyRows = RowFault(r, 2, arr(r, 2))
                Exit Function
            End If
            z = CellLong(arr(r, 3), ok)
            If Not ok Or z < 1 Or z > nP Then
                VerifyRows = RowFault(r, 3, arr(r, 3))
                Exit Function
            End If
            tally(z) = tally(z) + 1
        Next r
        For i = 1 To nP
            If tally(i) <> nS - 1 Then
                VerifyRows = "WRONG: person " & i & " has " & tally(i) & " pairs, expected " & (nS - 1)
                Exit Function
            End If
        Next i
    Case 8, 9
        ReDim tally(1 To nS)
        m = AnswerCols(stepNo, k) - 1
        For r = 2 To UBound(arr, 1)
            x = CellLong(arr(r, 1), ok)
            If Not ok Or x < 1 Or x > nS Then
                VerifyRows = RowFault(r, 1, arr(r, 1))
                Exit Function
            End If
            prevX = 0
            For c = 2 To m + 1
                y = CellLong(arr(r, c), ok)
                If Not ok Or y < 1 Or y > nP Then
                    VerifyRows = RowFault(r, c, arr(r, c))
                    Exit Function
                End If
                If stepNo = 8 Then
                    If y <= prevX Then
                        VerifyRows = "WRONG: row " & r & " is not in canonical order"
                        Exit Function
                    End If
                    prevX = y
                Else
                    seen(c - 1) = y
                End If
            Next c
            If stepNo = 9 Then
                If seen(1) = seen(2) Or seen(1) = seen(3) Or seen(2) = seen(3) Then
                    VerifyRows = "WRONG: row " & r & " names one person twice"
                    Exit Function
                End If
            End If
            tally(x) = tally(x) + 1
        Next r
        If stepNo = 8 Then
            want = Comb(nP, k + 1)
        Else
            want = CDbl(nP) * (nP - 1) * (nP - 2)
        End If
        For i = 1 To nS
            If tally(i) <> want Then
                VerifyRows = "WRONG: shift " & i & " has " & tally(i) & " rows, expected " & want
                Exit Function
            End If
        Next i
    Case 10
        ReDim tally(1 To nP * nW)
        For r = 2 To UBound(arr, 1)
            x = CellLong(arr(r, 1), ok)
            If Not ok Or x < 1 Or x > nP Then
                VerifyRows = RowFault(r, 1, arr(r, 1))
                Exit Function
            End If
            y = CellLong(arr(r, 2), ok)
            If Not ok Or y < 1 Or y > nW Then
                VerifyRows = RowFault(r, 2, arr(r, 2))
                Exit Function
            End If
            prevX = 0
            For c = 3 To k + 3
                z = CellLong(arr(r, c), ok)
                If Not ok Or z <= prevX Or z > nS Then
                    VerifyRows = RowFault(r, c, arr(r, c))
                    Exit Function
                End If
                If ShiftWeek(z) <> y Then
                    VerifyRows = "WRONG: row " & r & " puts shift " & z & " in week " & y
                    Exit Function
                End If
                prevX = z
            Next c
            idx = (x - 1) * nW + y
            tally(idx) = tally(idx) + 1
        Next r
        want = Comb(21, k + 1)
        For i = 1 To nP * nW
            If tally(i) <> want Then
                VerifyRows = "WRONG: person-week " & i & " has " & tally(i) & " rows, expected " & want
                Exit Function
            End If
        Next i
    End Select
    VerifyRows = "ok"
End Function

Private Function CellLong(ByVal v As Variant, ByRef ok As Boolean) As Long
    ok = False
    If IsError(v) Then Exit Function
    Select Case VarType(v)
    Case vbDouble, vbLong, vbInteger, vbCurrency, vbSingle
        If Abs(v) < 2000000000# Then
            If v = Int(v) Then
                CellLong = CLng(v)
                ok = True
            End If
        End If
    End Select
End Function

Private Function RowFault(ByVal r As Long, ByVal c As Long, ByVal v As Variant) As String
    Dim t As String
    If IsError(v) Then
        t = "an Excel error"
    Else
        t = "'" & Left$(CStr(v), 60) & "'"
    End If
    RowFault = "WRONG: row " & r & ", column " & c & " holds " & t
End Function

' ---------------------------------------------------------------------------
' Sheets, Tables, timing
' ---------------------------------------------------------------------------

Private Function SheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim sh As Object
    For Each sh In wb.Sheets
        If StrComp(sh.Name, sheetName, vbTextCompare) = 0 Then
            SheetExists = True
            Exit Function
        End If
    Next sh
End Function

Private Function NewTable(ByVal ws As Worksheet, ByVal addr As String, ByVal tableName As String, ByVal headers As Variant) As ListObject
    Dim n As Long, i As Long
    Dim rng As Range
    Dim lo As ListObject
    n = UBound(headers) - LBound(headers) + 1
    Set rng = ws.Range(addr).Resize(2, n)
    For i = 1 To n
        rng.Cells(1, i).Value2 = headers(LBound(headers) + i - 1)
        rng.Cells(2, i).Value2 = 1
    Next i
    Set lo = ws.ListObjects.Add(xlSrcRange, rng, , xlYes)
    lo.Name = tableName
    Set NewTable = lo
End Function

Private Function TableOn(ByVal ws As Worksheet, ByVal tableName As String) As ListObject
    Dim lo As ListObject
    On Error Resume Next
    Set lo = ws.ListObjects(tableName)
    On Error GoTo 0
    If lo Is Nothing Then
        Err.Raise vbObjectError + 1400, "VLA_DiagO0", "sheet " & ws.Name & " has no Table called " & tableName
    End If
    Set TableOn = lo
End Function

' Empties the Table, resizes it to nRows data rows and writes them in one go.
Private Sub WriteTable(ByVal lo As ListObject, ByRef data() As Variant, ByVal nRows As Long)
    Dim hdr As Range
    Dim block() As Variant
    Dim r As Long, c As Long
    If nRows < 1 Then
        Err.Raise vbObjectError + 1401, "VLA_DiagO0", "Table " & lo.Name & " would have no rows"
    End If
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.ClearContents
    Set hdr = lo.HeaderRowRange
    lo.Resize hdr.Worksheet.Range(hdr.Cells(1, 1), hdr.Cells(1, hdr.Columns.Count).Offset(nRows, 0))
    If nRows = UBound(data, 1) Then
        lo.DataBodyRange.Value2 = data
    Else
        ReDim block(1 To nRows, 1 To UBound(data, 2))
        For r = 1 To nRows
            For c = 1 To UBound(data, 2)
                block(r, c) = data(r, c)
            Next c
        Next r
        lo.DataBodyRange.Value2 = block
    End If
End Sub

Private Function SecondsToCalculate(ByVal target As Range, ByVal markDirty As Boolean) As Double
    Dim t0 As Double, dt As Double
    If markDirty Then target.Dirty
    DoEvents
    t0 = Timer
    Application.Calculate
    dt = Timer - t0
    If dt < 0 Then dt = dt + 86400#
    SecondsToCalculate = dt
End Function

Private Function Secs(ByVal t As Double) As String
    If t < 0 Then
        Secs = ""
    Else
        Secs = Format$(t, "0.000")
    End If
End Function

Private Function RatioText(ByVal ratio As Double) As String
    If ratio < 0 Then
        RatioText = ""
    Else
        RatioText = Format$(ratio, "0.00")
    End If
End Function

Private Function NextResultRow() As Long
    Dim rs As Worksheet
    If Not SheetExists(ActiveWorkbook, RESULTS_SHEET) Then
        Set rs = ActiveWorkbook.Worksheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
        rs.Name = RESULTS_SHEET
        rs.Range("A1").Resize(1, 14).Value2 = Array("step", "people", "weeks", "k", "ground rows", "first s", "control s", _
                                                     "model s", "model band", "measured/model", "peak rows", "verdict", "note", "when")
        NextResultRow = 2
    Else
        Set rs = ActiveWorkbook.Worksheets(RESULTS_SHEET)
        NextResultRow = rs.Cells(rs.Rows.Count, 1).End(xlUp).Row + 1
    End If
End Function
