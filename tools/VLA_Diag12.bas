Attribute VB_Name = "VLA_Diag12"
Option Explicit

' Standalone diagnostic for DATALOG.12 - NOT part of the VLA project, and it
' calls nothing in it. It writes worksheet cells and times Excel's own
' recalculation of a =DATALOG(...) formula that a sentence program already
' wrote. Import it into any open project, run one step's program through the
' IDE, then type that step's line in the Immediate window. Delete the module
' afterwards; nothing in Frazaro calls it. The steps are in
' archive/datalog12_live_steps.md.
'
' WHAT THIS MEASURES. No DATALOG answer over more than a few dozen rows had
' been timed, and a transliteration cannot time VBA (PROLOG.28's live pass
' found two costs no model showed). Each step's program builds a small
' fixture and asks the grammar's own question, so the formula timed is the one
' the grammar wrote. D12Ladder then grows that step's Tables up a ladder of
' sizes, in place, and at each size reads:
'   first      - seconds for Excel to recalculate the answer cell after the
'                Tables were rewritten. DATALOG keeps no cache, so this is
'                what a first answer costs;
'   change     - seconds to recalculate after ONE Table cell changes, a change
'                chosen to move the answer by a known amount;
'   unrelated  - seconds to recalculate after a cell nothing reads changes.
'                DATALOG declares no Application.Volatile, so this should read
'                about 0.000;
' and checks the answer against what the generator knows it must be (the row
' count, plus one group's value where a count alone could hide a wrong
' answer), so a fast wrong answer is never recorded as a result.
'
' HOW A LADDER STOPS, so no step holds Excel for long:
'   - after a size whose first answer or change took over LIMIT_SECONDS (10,
'     the owner's "a limit" line, 2026-09-14): that size is the limit;
'   - before a size whose time, projected from the sizes before it, passes
'     GUARD_SECONDS (15): it is skipped and the projection printed;
'   - after a refusal, a wrong answer, or a VBA error.
' A projection assumes the cost grows at least as fast as the step's own
' exponent, read from the code (a closure over one chain is quadratic), and
' faster if the last two sizes measured faster growth.
'
' Calculation is set to manual while a ladder runs, so rewriting a Table does
' not recalculate on every write, and restored at the end. Timer resolves
' about 16 ms, so a reading under 0.02 is too fast to see. Application.Calculate
' also recalculates any volatile formula in any open workbook: noise, not
' DATALOG. Results print to the Immediate window as D12| lines and are kept on
' the step's own sheet from column AD.

' The ladder and the guard were cut after the first live pass (2026-09-18):
' step 1, the simplest shape there is, measured about 2.5 ms per ROW, so
' 10,000 rows took 25s and Excel could not even be switched to. The sizes
' below keep every step inside a few seconds; the guard stops a size
' projected past 15s rather than the 60s the first pass allowed.
Private Const LIMIT_SECONDS As Double = 10#
Private Const GUARD_SECONDS As Double = 15#
Private Const RESULTS_COL As Long = 30
Private Const UNRELATED_ADDR As String = "AP1"

' extraRows (DATALOG.13, 2026-09-18): one more size appended after the
' step's own ladder, when larger than its last size - "D12Ladder 1, , 10000"
' re-measures pass 1's 10,000-row scan after the fix. The limit and the
' guard still apply to it, so a fix that falls short skips it by projection.
Public Sub D12Ladder(ByVal stepNo As Long, Optional ByVal maxRows As Long = 0, Optional ByVal extraRows As Long = 0)
    Dim answerAddr As String, shapeName As String, sizes As Variant, minExp As Double
    Dim ws As Worksheet, ans As Range
    Dim bits As String
    Dim prevCalc As Long, prevScreen As Boolean
    Dim resultRow As Long
    Dim si As Long, n As Long
    Dim lastN As Long, lastT As Double, prevN As Long, prevT As Double
    Dim projected As Double
    Dim expected As Long, gotRows As Long
    Dim sentinelKey As String, sentinelWant As String, wantRefusal As Boolean
    Dim note As String, shown As String, what As String, verdictText As String
    Dim tFirst As Double, tChange As Double, tUnrelated As Double
    Dim outLine As String
    Dim errNo As Long, errText As String

    If Not D12Spec(stepNo, answerAddr, shapeName, sizes, minExp) Then
        Debug.Print "D12: there is no step " & stepNo & " - the steps are 1 to 8."
        Exit Sub
    End If
    If extraRows > CLng(sizes(UBound(sizes))) Then
        Dim grown() As Variant, gi As Long
        ReDim grown(LBound(sizes) To UBound(sizes) + 1)
        For gi = LBound(sizes) To UBound(sizes)
            grown(gi) = sizes(gi)
        Next gi
        grown(UBound(grown)) = extraRows
        sizes = grown
    End If
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets("D12T" & stepNo)
    On Error GoTo 0
    If ws Is Nothing Then
        Debug.Print "D12: the active workbook has no sheet D12T" & stepNo & " - run step " & stepNo & "'s program first, with that workbook active."
        Exit Sub
    End If
    Set ans = ws.Range(answerAddr)
    If InStr(1, ans.Formula2, "DATALOG(", vbTextCompare) = 0 Then
        Debug.Print "D12: " & ws.Name & "!" & answerAddr & " holds no DATALOG formula - run step " & stepNo & "'s program first."
        Exit Sub
    End If

    #If Win64 Then
        bits = "64-bit"
    #Else
        bits = "32-bit"
    #End If
    Debug.Print "=== D12 step " & stepNo & ": " & shapeName & " ==="
    Debug.Print "    Excel " & Application.Version & " " & bits & ", workbook " & ActiveWorkbook.Name & ", " & Format$(Now, "yyyy-mm-dd hh:nn")
    Debug.Print "    formula: " & Left$(ans.Formula2, 600)
    Debug.Print "D12|step|rows|first s|change s|unrelated s|answer rows|expected|verdict|note"

    prevCalc = Application.Calculation
    prevScreen = Application.ScreenUpdating
    On Error GoTo fail
    resultRow = NextResultRow(ws)
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    For si = LBound(sizes) To UBound(sizes)
        n = CLng(sizes(si))
        If maxRows > 0 And n > maxRows Then
            Debug.Print "D12|" & stepNo & "|" & n & "||||||not run|above maxRows " & maxRows
            Exit For
        End If
        If lastN > 0 Then
            If lastT > LIMIT_SECONDS Then
                Debug.Print "D12|" & stepNo & "|" & n & "||||||not run|the size before took " & Format$(lastT, "0.0") & "s, over the " & LIMIT_SECONDS & "s limit"
                Exit For
            End If
            projected = Projection(n, lastN, lastT, prevN, prevT, minExp)
            If projected > GUARD_SECONDS Then
                Debug.Print "D12|" & stepNo & "|" & n & "||||||not run|projected " & Format$(projected, "0") & "s, over the " & GUARD_SECONDS & "s guard"
                Exit For
            End If
        End If

        Application.StatusBar = "D12 step " & stepNo & ": writing " & Format$(n, "#,##0") & " rows"
        expected = D12Fill(stepNo, ws, n, sentinelKey, sentinelWant, wantRefusal, note)
        Application.StatusBar = "D12 step " & stepNo & ": calculating " & Format$(n, "#,##0") & " rows - Excel will not respond until it is done"
        tFirst = SecondsToCalculate(ans, True)
        verdictText = Verdict(ans, expected, sentinelKey, sentinelWant, wantRefusal, gotRows, shown)
        tChange = -1
        tUnrelated = -1
        If verdictText = "ok" Then
            expected = D12Change(stepNo, ws, n, expected, sentinelWant, what)
            tChange = SecondsToCalculate(ans, False)
            verdictText = Verdict(ans, expected, sentinelKey, sentinelWant, False, gotRows, shown)
            If verdictText = "ok" Then
                ws.Range(UNRELATED_ADDR).Value2 = n
                tUnrelated = SecondsToCalculate(ans, False)
            Else
                verdictText = verdictText & " (after " & what & ")"
            End If
        End If
        If verdictText <> "ok" Then
            If Len(note) > 0 Then note = note & "; "
            note = note & "cell shows: " & Left$(shown, 300)
        End If

        outLine = "D12|" & stepNo & "|" & n & "|" & Secs(tFirst) & "|" & Secs(tChange) & "|" & Secs(tUnrelated) & "|" & _
                  gotRows & "|" & expected & "|" & verdictText & "|" & note
        Debug.Print outLine
        ws.Cells(resultRow, RESULTS_COL).Resize(1, 10).Value2 = Array(stepNo, n, Secs(tFirst), Secs(tChange), Secs(tUnrelated), _
                                                                     gotRows, expected, verdictText, note, Format$(Now, "yyyy-mm-dd hh:nn:ss"))
        resultRow = resultRow + 1
        If verdictText <> "ok" Then Exit For

        prevN = lastN
        prevT = lastT
        lastN = n
        lastT = tFirst
        If tChange > lastT Then lastT = tChange
    Next si

cleanup:
    On Error Resume Next
    Application.StatusBar = False
    Application.ScreenUpdating = prevScreen
    ' Restoring automatic calculation recalculates anything still dirty -
    ' after a VBA error mid-size, that can be a whole DATALOG answer.
    Application.Calculation = prevCalc
    On Error GoTo 0
    Debug.Print "=== D12 step " & stepNo & " done ==="
    Exit Sub
fail:
    errNo = Err.Number
    errText = Err.Description
    Debug.Print "D12|" & stepNo & "|" & n & "||||||VBA error " & errNo & "|" & errText
    Resume cleanup
End Sub

' One step's answer cell, what it measures, its ladder, and the slowest growth
' the code says to assume when projecting the next size.
Private Function D12Spec(ByVal stepNo As Long, ByRef answerAddr As String, ByRef shapeName As String, _
                         ByRef sizes As Variant, ByRef minExp As Double) As Boolean
    minExp = 1#
    sizes = Array(100, 300, 1000, 3000)
    Select Case stepNo
    Case 1
        answerAddr = "J2"
        shapeName = "a scan - which bill is big"
    Case 2
        answerAddr = "N2"
        shapeName = "a join - who can-cover S1, the README rule"
    Case 3
        answerAddr = "H2"
        shapeName = "count per group - how many people works-in each dept that is listed"
    Case 4
        answerAddr = "H2"
        shapeName = "textjoin per group - which people works-in each dept that is listed as one list"
    Case 5
        answerAddr = "I2"
        shapeName = "not through a projection - which person is available"
    Case 6
        answerAddr = "F2"
        shapeName = "otherwise - who has-tier Bronze"
    Case 7
        answerAddr = "F2"
        shapeName = "closure over a wide org - who reports-to E1 directly or not"
    Case 8
        answerAddr = "F2"
        shapeName = "closure over one chain - who reports-to E1 directly or not"
        sizes = Array(100, 250, 500, 1000)
        minExp = 2#
    Case Else
        Exit Function
    End Select
    D12Spec = True
End Function

' Rewrites step stepNo's Tables at n rows and returns the answer's expected
' data rows, the header not counted. sentinelKey/sentinelWant name one answer
' row whose second column must read sentinelWant ("" key when unused);
' wantRefusal is True where the generator knows DATALOG must refuse; note
' carries anything worth printing beside the timings.
Private Function D12Fill(ByVal stepNo As Long, ByVal ws As Worksheet, ByVal n As Long, _
                         ByRef sentinelKey As String, ByRef sentinelWant As String, _
                         ByRef wantRefusal As Boolean, ByRef note As String) As Long
    sentinelKey = ""
    sentinelWant = ""
    wantRefusal = False
    note = ""
    Select Case stepNo
    Case 1
        D12Fill = FillBills(ws, n)
    Case 2
        D12Fill = FillCanCover(ws, n)
    Case 3
        D12Fill = FillDepts(ws, n, "12x3", 100, False, sentinelKey, sentinelWant, wantRefusal, note)
    Case 4
        D12Fill = FillDepts(ws, n, "12x4", 10, True, sentinelKey, sentinelWant, wantRefusal, note)
    Case 5
        D12Fill = FillLeave(ws, n)
    Case 6
        D12Fill = FillTiers(ws, n)
    Case 7
        D12Fill = FillOrg(ws, n, "Reports12x7", False, note)
    Case 8
        D12Fill = FillOrg(ws, n, "Chain12x8", True, note)
    End Select
End Function

' Changes ONE Table cell and returns the answer rows expected after it;
' sentinelWant is updated where the change moves that row instead.
Private Function D12Change(ByVal stepNo As Long, ByVal ws As Worksheet, ByVal n As Long, ByVal expected As Long, _
                           ByRef sentinelWant As String, ByRef what As String) As Long
    Dim i As Long, m As Long
    D12Change = expected
    Select Case stepNo
    Case 1
        ' B1's amount is 7400, not big.
        NeedTable(ws, "Bills12x1").DataBodyRange.Cells(1, 3).Value2 = 20000
        what = "B1's amount set to 20000"
        D12Change = expected + 1
    Case 2
        ' The first person holding C1 at level 1 who is not on leave from S1.
        m = LeaveCount(n)
        For i = 1 To n
            If i Mod 10 = 1 And i Mod 3 = 0 And Not OnLeaveFromS1(i, m) Then
                NeedTable(ws, "Staff12x2").DataBodyRange.Cells(i, 3).Value2 = 3
                what = "P" & i & "'s level set to 3"
                D12Change = expected + 1
                Exit Function
            End If
        Next i
        what = "no change available"
    Case 3
        ' P1 moves from D1 to D100, the listed dept nobody works in.
        NeedTable(ws, "Staff12x3").DataBodyRange.Cells(1, 2).Value2 = "D100"
        what = "P1's dept set to D100"
        sentinelWant = "1"
    Case 4
        NeedTable(ws, "Staff12x4").DataBodyRange.Cells(1, 2).Value2 = "D10"
        what = "P1's dept set to D10"
        sentinelWant = "P1"
    Case 5
        ' P3 is the first person on leave.
        NeedTable(ws, "Leave12x5").DataBodyRange.Cells(1, 1).Value2 = "X0"
        what = "the first leave row's name set to X0"
        D12Change = expected + 1
    Case 6
        ' K1 spends 3700, so is Bronze.
        NeedTable(ws, "Customers12x6").DataBodyRange.Cells(1, 2).Value2 = 20000
        what = "K1's spend set to 20000"
        D12Change = expected - 1
    Case 7, 8
        ' The last employee manages nobody, so leaves the answer alone.
        If stepNo = 7 Then
            NeedTable(ws, "Reports12x7").DataBodyRange.Cells(n, 2).Value2 = "X0"
        Else
            NeedTable(ws, "Chain12x8").DataBodyRange.Cells(n, 2).Value2 = "X0"
        End If
        what = "E" & (n + 1) & "'s manager set to X0"
        D12Change = expected - 1
    End Select
End Function

' Bills(Bill, Vendor, Amount). (i * 37) Mod 100 visits every residue once per
' hundred rows, so 49 rows in a hundred are over 10000.
Private Function FillBills(ByVal ws As Worksheet, ByVal n As Long) As Long
    Dim data() As Variant
    Dim i As Long, amount As Long, big As Long
    ReDim data(1 To n, 1 To 3)
    For i = 1 To n
        amount = ((i * 37) Mod 100) * 200
        data(i, 1) = "B" & i
        data(i, 2) = "V" & (i Mod 50)
        data(i, 3) = amount
        If amount > 10000 Then big = big + 1
    Next i
    WriteTable NeedTable(ws, "Bills12x1"), data, n
    FillBills = big
End Function

' Staff(Name, Cert, Level) at n rows; Shifts(Shift, Needs, MinLevel) at 20;
' Leave(Name, Shift) at n \ 20, every twentieth person from P1 on leave from
' S1. The join Shifts x Staff on the cert makes about 2n pairs.
Private Function FillCanCover(ByVal ws As Worksheet, ByVal n As Long) As Long
    Dim staffData() As Variant, shiftData() As Variant, leaveData() As Variant
    Dim i As Long, m As Long, cover As Long
    ReDim staffData(1 To n, 1 To 3)
    For i = 1 To n
        staffData(i, 1) = "P" & i
        staffData(i, 2) = "C" & (i Mod 10)
        staffData(i, 3) = (i Mod 3) + 1
    Next i
    ReDim shiftData(1 To 20, 1 To 3)
    For i = 1 To 20
        shiftData(i, 1) = "S" & i
        shiftData(i, 2) = "C" & (i Mod 10)
        shiftData(i, 3) = (i Mod 3) + 1
    Next i
    m = LeaveCount(n)
    ReDim leaveData(1 To m, 1 To 2)
    For i = 1 To m
        leaveData(i, 1) = "P" & (20 * i - 19)
        leaveData(i, 2) = "S1"
    Next i
    WriteTable NeedTable(ws, "Staff12x2"), staffData, n
    WriteTable NeedTable(ws, "Shifts12x2"), shiftData, 20
    WriteTable NeedTable(ws, "Leave12x2"), leaveData, m
    ' S1 needs C1 at level 2 or more.
    For i = 1 To n
        If i Mod 10 = 1 And (i Mod 3) + 1 >= 2 And Not OnLeaveFromS1(i, m) Then cover = cover + 1
    Next i
    FillCanCover = cover
End Function

Private Function LeaveCount(ByVal n As Long) As Long
    LeaveCount = n \ 20
    If LeaveCount < 1 Then LeaveCount = 1
End Function

Private Function OnLeaveFromS1(ByVal i As Long, ByVal m As Long) As Boolean
    OnLeaveFromS1 = (i Mod 20 = 1 And (i + 19) \ 20 <= m)
End Function

' Staff(Name, Dept) at n rows over `groups` depts; Depts(Dept) lists D0 to
' D<groups>, the last with nobody in it: a count of 0, or a blank list.
Private Function FillDepts(ByVal ws As Worksheet, ByVal n As Long, ByVal suffix As String, ByVal groups As Long, _
                           ByVal asList As Boolean, ByRef sentinelKey As String, ByRef sentinelWant As String, _
                           ByRef wantRefusal As Boolean, ByRef note As String) As Long
    Dim staffData() As Variant, deptData() As Variant
    Dim lens() As Long, counts() As Long
    Dim i As Long, g As Long, longest As Long
    ReDim staffData(1 To n, 1 To 2)
    ReDim lens(0 To groups - 1)
    ReDim counts(0 To groups - 1)
    For i = 1 To n
        g = i Mod groups
        staffData(i, 1) = "P" & i
        staffData(i, 2) = "D" & g
        If counts(g) > 0 Then lens(g) = lens(g) + 2
        lens(g) = lens(g) + Len("P" & i)
        counts(g) = counts(g) + 1
    Next i
    ReDim deptData(1 To groups + 1, 1 To 1)
    For g = 0 To groups
        deptData(g + 1, 1) = "D" & g
    Next g
    WriteTable NeedTable(ws, "Staff" & suffix), staffData, n
    WriteTable NeedTable(ws, "Depts" & suffix), deptData, groups + 1
    sentinelKey = "D" & groups
    If asList Then
        sentinelWant = ""
        For g = 0 To groups - 1
            If lens(g) > longest Then longest = lens(g)
        Next g
        note = "longest list " & longest & " characters"
        wantRefusal = (longest > 32767)
    Else
        sentinelWant = "0"
    End If
    FillDepts = groups + 1
End Function

' Staff(Name, Dept) at n rows; Leave(Name, Day) at n \ 3, every third person,
' so the rule's `not` reads a projection that drops Day.
Private Function FillLeave(ByVal ws As Worksheet, ByVal n As Long) As Long
    Dim staffData() As Variant, leaveData() As Variant
    Dim i As Long, m As Long
    ReDim staffData(1 To n, 1 To 2)
    For i = 1 To n
        staffData(i, 1) = "P" & i
        staffData(i, 2) = "D" & (i Mod 100)
    Next i
    m = n \ 3
    If m < 1 Then m = 1
    ReDim leaveData(1 To m, 1 To 2)
    For i = 1 To m
        leaveData(i, 1) = "P" & (3 * i)
        leaveData(i, 2) = "Day" & (i Mod 5)
    Next i
    WriteTable NeedTable(ws, "Staff12x5"), staffData, n
    WriteTable NeedTable(ws, "Leave12x5"), leaveData, m
    FillLeave = n - m
End Function

' Customers(Customer, Spend). (i * 37) Mod 150 visits every residue once per
' 150 rows, a third of them under 5000: Bronze.
Private Function FillTiers(ByVal ws As Worksheet, ByVal n As Long) As Long
    Dim data() As Variant
    Dim i As Long, spend As Long, bronze As Long
    ReDim data(1 To n, 1 To 2)
    For i = 1 To n
        spend = ((i * 37) Mod 150) * 100
        data(i, 1) = "K" & i
        data(i, 2) = spend
        If spend < 5000 Then bronze = bronze + 1
    Next i
    WriteTable NeedTable(ws, "Customers12x6"), data, n
    FillTiers = bronze
End Function

' Reports(Employee, Manager) at n rows: E<i+1> reports to E<i> (one chain) or
' to E<(i-1)\10+1> (ten reports each). Everyone reaches E1, so the answer has n
' rows; the note gives the closure's size, which DATALOG builds whole.
Private Function FillOrg(ByVal ws As Worksheet, ByVal n As Long, ByVal tableName As String, _
                         ByVal oneChain As Boolean, ByRef note As String) As Long
    Dim data() As Variant
    Dim depth() As Long
    Dim i As Long, boss As Long, deepest As Long
    Dim pairs As Double
    ReDim data(1 To n, 1 To 2)
    ReDim depth(1 To n + 1)
    For i = 1 To n
        If oneChain Then
            boss = i
        Else
            boss = (i - 1) \ 10 + 1
        End If
        data(i, 1) = "E" & (i + 1)
        data(i, 2) = "E" & boss
        depth(i + 1) = depth(boss) + 1
        pairs = pairs + depth(i + 1)
        If depth(i + 1) > deepest Then deepest = depth(i + 1)
    Next i
    WriteTable NeedTable(ws, tableName), data, n
    note = "closure " & Format$(pairs, "#,##0") & " pairs, deepest " & deepest & " links"
    FillOrg = n
End Function

Private Function NeedTable(ByVal ws As Worksheet, ByVal nm As String) As ListObject
    Dim lo As ListObject
    On Error Resume Next
    Set lo = ws.ListObjects(nm)
    On Error GoTo 0
    If lo Is Nothing Then
        Err.Raise vbObjectError + 1212, "VLA_Diag12", "sheet " & ws.Name & " has no Table called " & nm & " - run this step's program first"
    End If
    Set NeedTable = lo
End Function

' Empties the Table, resizes it to nRows data rows and writes them in one go.
' Emptying first means a Table shrunk after a longer run leaves nothing behind.
Private Sub WriteTable(ByVal lo As ListObject, ByRef data() As Variant, ByVal nRows As Long)
    Dim hdr As Range
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.ClearContents
    Set hdr = lo.HeaderRowRange
    lo.Resize hdr.Worksheet.Range(hdr.Cells(1, 1), hdr.Cells(1, hdr.Columns.Count).Offset(nRows, 0))
    lo.DataBodyRange.Value2 = data
End Sub

Private Function SecondsToCalculate(ByVal ans As Range, ByVal markDirty As Boolean) As Double
    Dim t0 As Double, dt As Double
    If markDirty Then ans.Dirty
    DoEvents
    t0 = Timer
    Application.Calculate
    dt = Timer - t0
    If dt < 0 Then dt = dt + 86400#
    SecondsToCalculate = dt
End Function

' "ok", "refused, as predicted", or what is wrong. gotRows is the spill's data
' rows (-1 when the cell is a refusal or an error); shown is the cell's text.
Private Function Verdict(ByVal ans As Range, ByVal expected As Long, ByVal sentinelKey As String, _
                         ByVal sentinelWant As String, ByVal wantRefusal As Boolean, _
                         ByRef gotRows As Long, ByRef shown As String) As String
    Dim v As Variant, cellsV As Variant
    Dim r As Long, found As Boolean
    Dim got As String
    gotRows = -1
    v = ans.Value2
    If IsError(v) Then
        shown = ans.Text
        Verdict = "WRONG: Excel shows " & shown
        Exit Function
    End If
    shown = CStr(v)
    If Left$(shown, 9) = "#DATALOG!" Then
        If wantRefusal Then
            Verdict = "refused, as predicted"
        Else
            Verdict = "REFUSED"
        End If
        Exit Function
    End If
    If ans.HasSpill Then
        gotRows = ans.SpillingToRange.Rows.Count - 1
    Else
        gotRows = 0
    End If
    If wantRefusal Then
        Verdict = "WRONG: answered where a refusal was predicted"
        Exit Function
    End If
    If gotRows <> expected Then
        Verdict = "WRONG: " & gotRows & " rows, expected " & expected
        Exit Function
    End If
    If Len(sentinelKey) > 0 Then
        cellsV = ans.SpillingToRange.Value2
        For r = 2 To UBound(cellsV, 1)
            If Not IsError(cellsV(r, 1)) Then
                If CStr(cellsV(r, 1)) = sentinelKey Then
                    found = True
                    If IsError(cellsV(r, 2)) Then
                        got = "(an Excel error)"
                    Else
                        got = CStr(cellsV(r, 2))
                    End If
                    If got <> sentinelWant Then
                        Verdict = "WRONG: " & sentinelKey & " shows '" & got & "', expected '" & sentinelWant & "'"
                        Exit Function
                    End If
                End If
            End If
        Next r
        If Not found Then
            Verdict = "WRONG: no row for " & sentinelKey
            Exit Function
        End If
    End If
    Verdict = "ok"
End Function

' Seconds the next size should take, from the last one: at least the step's
' own exponent, or the growth the last two sizes measured if that is steeper.
' Readings under 0.2s are too coarse to measure growth from.
Private Function Projection(ByVal n As Long, ByVal lastN As Long, ByVal lastT As Double, _
                            ByVal prevN As Long, ByVal prevT As Double, ByVal minExp As Double) As Double
    Dim k As Double, measured As Double, fromT As Double
    k = minExp
    If prevN > 0 And prevT >= 0.2 And lastT >= 0.2 Then
        measured = Log(lastT / prevT) / Log(CDbl(lastN) / CDbl(prevN))
        If measured > k Then k = measured
    End If
    fromT = lastT
    If fromT < 0.02 Then fromT = 0.02
    Projection = fromT * (CDbl(n) / CDbl(lastN)) ^ k
End Function

Private Function Secs(ByVal t As Double) As String
    If t < 0 Then
        Secs = ""
    Else
        Secs = Format$(t, "0.000")
    End If
End Function

' A DECOMPOSITION, added after live pass 1 (2026-09-18). Step 1 measured
' about 2.5 ms a row and tools/VLA_Diag12b.bas accounted for only an eighth
' of it, so this asks the engine itself where the rest goes: four formulas
' over the SAME Bills12x1 Table, each adding one stage to the one before,
' timed separately at one size. The difference between two lines is what that
' stage costs per row. Needs step 1's sheet; it writes its probes to spare
' columns and clears them afterwards.
'
'   1  (query bills12x1)             - read the Table, spill it whole
'   2  + one rule, no test           - filter, join, project, derive
'   3  + the comparison              - DATALOG.4's own per-row arm
'   4  + the question's narrowing rule - the whole step-1 program
Public Sub D12Parts(Optional ByVal rowsWanted As Long = 1000)
    Dim ws As Worksheet, probe As Range
    Dim prevCalc As Long, prevScreen As Boolean
    Dim progs As Variant, addrs As Variant, names As Variant
    Dim i As Long
    Dim seconds As Double, last As Double
    Dim rowsSeen As Long
    Dim shown As String, added As String
    Dim errNo As Long, errText As String

    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets("D12T1")
    On Error GoTo 0
    If ws Is Nothing Then
        Debug.Print "D12P: the active workbook has no sheet D12T1 - run step 1's program first."
        Exit Sub
    End If

    progs = Array( _
        "(query bills12x1)", _
        "(rule (r Bill Amount) (bills12x1 (bill Bill) (amount Amount))) (query r)", _
        "(rule (r Bill Amount) (bills12x1 (bill Bill) (amount Amount)) (> Amount 10000)) (query r)", _
        "(rule (big Bill) (bills12x1 (bill Bill) (amount Amount)) (> Amount 10000)) (rule (vla-ask-big Bill) (big Bill)) (query vla-ask-big)")
    addrs = Array("L2", "P2", "S2", "V2")
    names = Array("read the Table and spill it whole", "one rule, no test", "one rule and the test", "the whole step-1 program")

    prevCalc = Application.Calculation
    prevScreen = Application.ScreenUpdating
    On Error GoTo failParts
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Debug.Print "=== D12 parts: where a row's time goes, at " & rowsWanted & " rows ==="
    FillBills ws, rowsWanted
    For i = LBound(progs) To UBound(progs)
        ws.Range(CStr(addrs(i))).Formula2 = "=DATALOG(""" & progs(i) & """, bills12x1)"
    Next i
    ' One untimed pass, so every formula on the sheet - step 1's own answer
    ' included - is clean before anything is timed.
    Application.Calculate
    Debug.Print "D12P|probe|rows|seconds|ms per row|answer rows|what it adds"
    last = 0
    For i = LBound(progs) To UBound(progs)
        Set probe = ws.Range(CStr(addrs(i)))
        seconds = SecondsToCalculate(probe, True)
        ReadRows probe, rowsSeen, shown
        added = ""
        If i > LBound(progs) Then added = "  (+" & Format$((seconds - last) / rowsWanted * 1000#, "0.000") & " ms/row)"
        Debug.Print "D12P|" & (i + 1) & "|" & rowsWanted & "|" & Format$(seconds, "0.000") & "|" & _
                    Format$(seconds / rowsWanted * 1000#, "0.000") & "|" & rowsSeen & "|" & names(i) & added
        If Left$(shown, 9) = "#DATALOG!" Then Debug.Print "   refused: " & Left$(shown, 200)
        last = seconds
    Next i
    For i = LBound(progs) To UBound(progs)
        ws.Range(CStr(addrs(i))).ClearContents
    Next i

cleanupParts:
    On Error Resume Next
    Application.StatusBar = False
    Application.ScreenUpdating = prevScreen
    Application.Calculation = prevCalc
    On Error GoTo 0
    Debug.Print "=== D12 parts done ==="
    Exit Sub
failParts:
    errNo = Err.Number
    errText = Err.Description
    Debug.Print "D12P: stopped by VBA error " & errNo & ": " & errText
    Resume cleanupParts
End Sub

' ONE THING AT A TIME, added after live pass 3 (2026-09-18). D12Parts found
' that a query with no rule at all cost the most per row, and that adding a
' rule made it faster - but every one of its probes varied the work AND the
' answer's width together, so it cannot say which. These five vary one thing:
' three DATALOG formulas over the SAME rows, each doing the same per-row work
' (the same three-column atom, the same join) and differing only in how many
' columns the answer HAS, beside a bare VBA UDF that hands Excel an array of
' the same shape and touches no engine at all.
'
' The differences between the first three are the cost of one more column per
' row; the last two say what that costs with no DATALOG in the picture. Run
' it at two sizes (D12Width 1000, then D12Width 3000) to separate the per-row
' cost from the per-cell one. Needs step 1's sheet; probes go to spare
' columns out beyond the results block and are cleared afterwards.
Public Sub D12Width(Optional ByVal rowsWanted As Long = 1000)
    Dim ws As Worksheet, probe As Range
    Dim prevCalc As Long, prevScreen As Boolean
    Dim progs As Variant, addrs As Variant, names As Variant, widths As Variant
    Dim secs() As Double
    Dim i As Long
    Dim seconds As Double
    Dim rowsSeen As Long
    Dim shown As String
    Dim errNo As Long, errText As String

    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets("D12T1")
    On Error GoTo 0
    If ws Is Nothing Then
        Debug.Print "D12W: the active workbook has no sheet D12T1 - run step 1's program first."
        Exit Sub
    End If

    progs = Array( _
        "(rule (r Bill) (bills12x1 (bill Bill))) (query r)", _
        "(rule (r Bill Vendor) (bills12x1 (bill Bill) (vendor Vendor))) (query r)", _
        "(rule (r Bill Vendor Amount) (bills12x1 (bill Bill) (vendor Vendor) (amount Amount))) (query r)", _
        "", "")
    addrs = Array("AR2", "AT2", "AW2", "BB2", "BD2")
    widths = Array(1, 2, 3, 1, 3)
    names = Array("DATALOG, 1 column out", "DATALOG, 2 columns out", "DATALOG, 3 columns out", _
                  "a bare UDF, 1 column out", "a bare UDF, 3 columns out")
    ReDim secs(LBound(progs) To UBound(progs))

    prevCalc = Application.Calculation
    prevScreen = Application.ScreenUpdating
    On Error GoTo failWidth
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Debug.Print "=== D12 width: the same rows, a wider answer, at " & rowsWanted & " rows ==="
    FillBills ws, rowsWanted
    For i = LBound(progs) To UBound(progs)
        If Len(CStr(progs(i))) > 0 Then
            ws.Range(CStr(addrs(i))).Formula2 = "=DATALOG(""" & progs(i) & """, bills12x1)"
        Else
            ws.Range(CStr(addrs(i))).Formula2 = "=D12Cells(" & rowsWanted & "," & widths(i) & ")"
        End If
    Next i
    ' One untimed pass, so everything on the sheet is clean before timing.
    Application.Calculate
    Debug.Print "D12W|probe|rows|cols|seconds|ms per row|ms per cell|answer rows|what it is"
    For i = LBound(progs) To UBound(progs)
        Set probe = ws.Range(CStr(addrs(i)))
        seconds = SecondsToCalculate(probe, True)
        secs(i) = seconds
        ReadRows probe, rowsSeen, shown
        Debug.Print "D12W|" & (i + 1) & "|" & rowsWanted & "|" & widths(i) & "|" & Format$(seconds, "0.000") & "|" & _
                    Format$(seconds / rowsWanted * 1000#, "0.000") & "|" & _
                    Format$(seconds / rowsWanted / widths(i) * 1000#, "0.000") & "|" & rowsSeen & "|" & names(i)
        If Left$(shown, 9) = "#DATALOG!" Then Debug.Print "   refused: " & Left$(shown, 200)
    Next i
    Debug.Print "   -> one more column through DATALOG costs " & _
                Format$((secs(LBound(progs) + 1) - secs(LBound(progs))) / rowsWanted * 1000#, "0.000") & " and " & _
                Format$((secs(LBound(progs) + 2) - secs(LBound(progs) + 1)) / rowsWanted * 1000#, "0.000") & " ms per row."
    Debug.Print "   -> two more columns through a bare UDF cost " & _
                Format$((secs(LBound(progs) + 4) - secs(LBound(progs) + 3)) / rowsWanted * 1000#, "0.000") & " ms per row."
    For i = LBound(progs) To UBound(progs)
        ws.Range(CStr(addrs(i))).ClearContents
    Next i

cleanupWidth:
    On Error Resume Next
    Application.StatusBar = False
    Application.ScreenUpdating = prevScreen
    Application.Calculation = prevCalc
    On Error GoTo 0
    Debug.Print "=== D12 width done ==="
    Exit Sub
failWidth:
    errNo = Err.Number
    errText = Err.Description
    Debug.Print "D12W: stopped by VBA error " & errNo & ": " & errText
    Resume cleanupWidth
End Sub

' A worksheet function that does NOTHING but hand Excel an array of the shape
' a DATALOG answer has - no Table read, no rules, no engine. What it costs is
' what returning that many cells from VBA costs, and nothing else.
Public Function D12Cells(ByVal nRows As Long, ByVal nCols As Long) As Variant
    Dim a() As Variant
    Dim r As Long, c As Long
    ReDim a(1 To nRows, 1 To nCols)
    For r = 1 To nRows
        For c = 1 To nCols
            a(r, c) = "B" & r & "-" & c
        Next c
    Next r
    D12Cells = a
End Function

Private Sub ReadRows(ByVal ans As Range, ByRef rowsSeen As Long, ByRef shown As String)
    Dim v As Variant
    v = ans.Value2
    If IsError(v) Then
        shown = ans.Text
        rowsSeen = -1
        Exit Sub
    End If
    shown = CStr(v)
    If ans.HasSpill Then
        rowsSeen = ans.SpillingToRange.Rows.Count - 1
    Else
        rowsSeen = 0
    End If
End Sub

Private Function NextResultRow(ByVal ws As Worksheet) As Long
    If IsEmpty(ws.Cells(1, RESULTS_COL).Value2) Then
        ws.Cells(1, RESULTS_COL).Resize(1, 10).Value2 = Array("step", "rows", "first s", "change s", "unrelated s", _
                                                            "answer rows", "expected", "verdict", "note", "when")
        NextResultRow = 2
    Else
        NextResultRow = ws.Cells(ws.Rows.Count, RESULTS_COL).End(xlUp).Row + 1
    End If
End Function
