Attribute VB_Name = "VLA_DiagEmpty"
Option Explicit

' Standalone diagnostic for KERNEL.46, a write that empties a cell - NOT part
' of the VLA project, and called by nothing in Frazaro. It measures Excel's own
' empty, Range.ClearContents beside Range.Clear, so that the scoping's
' decisions rest on a measurement and not a memory: what each leaves of a
' styled cell, what the cell's readers compute after it, what each does to a
' spill, to a legacy array formula's block, to a shared formula's master, to
' a conditional format and to the used range, and what the file then saves.
'
' WHERE IT RUNS. Import it into a blank workbook that is never saved (File >
' New > Blank workbook, then Alt+F11, File > Import File); never into the dev
' workbook. Run K46Probe from Excel's Macros dialog (Alt+F8, Macros in: All
' Open Workbooks), never from the Immediate window: it adds and opens
' workbooks (the VBA traps, 19). Excel 365: SEQUENCE, Formula2, HasSpill.
'
' WHAT IT DOES. It builds a workbook of its own (Workbooks.Add), lays each
' case on a sheet of its own, and records every case at five stages: laid;
' cleared, after the clears; saved, after the workbook is saved; reopened,
' the saved file opened again; and recomputed, after a full calculation of
' it. The laid workbook is saved as k46_before.xlsx before any clear, and the
' cleared one as k46_after.xlsx. Every fact is a line, K46|<n>|<stage>|<what>|
' <measured>, in k46_probe.txt; the three files are in %TEMP%\Frazaro_K46,
' which the message at the end names. Nothing else is touched, and both
' workbooks it makes are closed at the end.
'
' THE STOP RULE, stated before the first save. One run. A second only to
' repair this instrument, where it fails to record a line, and never to add
' a case. A fact Excel answers differently in two cells meant to agree, or
' between the session and the reopened file where no save should move it,
' is labelled in the scoping and never chased.
'
' THE SHEETS.
'   Clear   column A, a cell a row: 1 a styled 5 by ClearContents; 2 by
'           Clear; 3 a text of no characters (what Alonzo's Delete writes
'           today), never cleared; 4 a styled formula by ClearContents; 5 a
'           cell never written; 6, 7 and 8 a styled 5 by Value = "", by
'           Value = Empty and by Formula = "". Columns B to H read it:
'           =A1+1, =A1&"x", ISBLANK, COUNTA, COUNTBLANK, =A1="" and =A1=0.
'   Spill   =SEQUENCE(3) five times: A1 blocked by a 7 at A3, which is
'           cleared; C1, its spilled C2 cleared; E1, its anchor cleared;
'           G1, G1:G3 cleared; I1 blocked by a styled 5 at I2, cleared.
'   Arr     legacy array formulas (Ctrl+Shift+Enter) over A1:A3, C1:C3 and
'           E1:E3: A2 cleared, then A1:A3; C1:C4 cleared; E1 cleared.
'   Shared  =A1*2, =A1*3 and =A1*4 filled down B, D and F from 1 to 5:
'           B1 cleared (the master), D3 (one inside), F1:F2 (the master and
'           the next).
'   Db1-4   a database in column A: "k" at A1, 1 to 4 at A2:A5, and at A8
'           a styled 9 by ClearContents (Db1), a styled 9 by Clear (Db2),
'           a plain 9 by ClearContents (Db3), and a format with no value,
'           never cleared (Db4).
'   CF      a rule "value above 3" over A1:A2, each a 5, A1 by ClearContents
'           and A2 by Clear; a rule =ISBLANK(C1) over C1:C2, each a 5, C1 by
'           ClearContents and C2 by Clear.
'   Read    the databases read: DCOUNTA of each whole column with no field
'           and a blank criterion (Read!A1:A2), COUNTA of the column, ROWS
'           of it, and DCOUNTA of A1:A20, each a row.

Private mFile As Integer
Private mN As Long

Public Sub K46Probe()
    Dim dirPath As String, wb As Workbook, oldAlerts As Boolean, oldCalc As Long
    Dim n As Long, d As String
    dirPath = Environ$("TEMP") & "\Frazaro_K46"
    On Error Resume Next
    MkDir dirPath
    Kill dirPath & "\k46_before.xlsx"
    Kill dirPath & "\k46_after.xlsx"
    Kill dirPath & "\k46_probe.txt"
    On Error GoTo 0
    mN = 0
    mFile = FreeFile
    Open dirPath & "\k46_probe.txt" For Output As #mFile
    oldAlerts = Application.DisplayAlerts
    oldCalc = Application.Calculation
    On Error GoTo Failed
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationAutomatic
    Fact "excel", "version", Application.Version & " build " & Application.Build
    Set wb = Workbooks.Add(xlWBATWorksheet)
    LayCases wb
    RecordAll wb, "laid"
    wb.SaveAs Filename:=dirPath & "\k46_before.xlsx", FileFormat:=xlOpenXMLWorkbook
    ApplyClears wb
    RecordAll wb, "cleared"
    wb.SaveAs Filename:=dirPath & "\k46_after.xlsx", FileFormat:=xlOpenXMLWorkbook
    RecordAll wb, "saved"
    wb.Close SaveChanges:=False
    Set wb = Workbooks.Open(Filename:=dirPath & "\k46_after.xlsx")
    RecordAll wb, "reopened"
    Application.CalculateFull
    RecordAll wb, "recomputed"
    wb.Close SaveChanges:=False
    Set wb = Nothing
    Fact "end", "facts", CStr(mN + 1)
    Close #mFile
    Application.Calculation = oldCalc
    Application.DisplayAlerts = oldAlerts
    MsgBox "K46 done: " & mN & " facts in " & dirPath & "\k46_probe.txt", vbInformation, "K46Probe"
    Exit Sub
Failed:
    n = Err.Number
    d = Err.Description
    On Error Resume Next
    Fact "failed", "error", n & ": " & d
    Close #mFile
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Application.Calculation = oldCalc
    Application.DisplayAlerts = oldAlerts
    MsgBox "K46 stopped at fact " & mN & ", error " & n & ": " & d, vbExclamation, "K46Probe"
End Sub

' ---- the cases ----------------------------------------------------------

Private Sub LayCases(ByVal wb As Workbook)
    Dim wsC As Worksheet, wsS As Worksheet, wsA As Worksheet, wsF As Worksheet
    Dim wsK As Worksheet, wsR As Worksheet, ws As Worksheet, fc As Object
    Dim i As Long, k As Long
    Set wsC = wb.Worksheets(1)
    wsC.Name = "Clear"
    Set wsS = wb.Worksheets.Add(After:=wsC)
    wsS.Name = "Spill"
    Set wsA = wb.Worksheets.Add(After:=wsS)
    wsA.Name = "Arr"
    Set wsF = wb.Worksheets.Add(After:=wsA)
    wsF.Name = "Shared"
    Set ws = wsF
    For k = 1 To 4
        Set ws = wb.Worksheets.Add(After:=ws)
        ws.Name = "Db" & k
    Next k
    Set wsK = wb.Worksheets.Add(After:=ws)
    wsK.Name = "CF"
    Set wsR = wb.Worksheets.Add(After:=wsK)
    wsR.Name = "Read"
    For Each ws In wb.Worksheets
        ws.Columns("A:I").ColumnWidth = 16
    Next ws

    ' Clear: a cell a row in column A, read across B to H.
    For i = 1 To 8
        If i <> 3 And i <> 5 Then
            wsC.Cells(i, 1).Value = 5
            Dress wsC.Cells(i, 1)
        End If
    Next i
    wsC.Cells(4, 1).Formula = "=10*2"
    wsC.Activate
    wsC.Cells(3, 1).Formula = "="""""
    wsC.Cells(3, 1).Copy
    wsC.Cells(3, 1).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    For i = 1 To 8
        wsC.Cells(i, 2).Formula = "=A" & i & "+1"
        wsC.Cells(i, 3).Formula = "=A" & i & "&""x"""
        wsC.Cells(i, 4).Formula = "=ISBLANK(A" & i & ")"
        wsC.Cells(i, 5).Formula = "=COUNTA(A" & i & ")"
        wsC.Cells(i, 6).Formula = "=COUNTBLANK(A" & i & ")"
        wsC.Cells(i, 7).Formula = "=A" & i & "="""""
        wsC.Cells(i, 8).Formula = "=A" & i & "=0"
    Next i

    ' Spill: five spills of three, two of them blocked.
    wsS.Range("A1").Formula2 = "=SEQUENCE(3)"
    wsS.Range("A3").Value = 7
    wsS.Range("C1").Formula2 = "=SEQUENCE(3)"
    wsS.Range("E1").Formula2 = "=SEQUENCE(3)"
    wsS.Range("G1").Formula2 = "=SEQUENCE(3)"
    wsS.Range("I1").Formula2 = "=SEQUENCE(3)"
    wsS.Range("I2").Value = 5
    Dress wsS.Range("I2")

    ' Arr: three legacy array formulas over blocks of three.
    wsA.Range("A1:A3").FormulaArray = "=ROW(A1:A3)*10"
    wsA.Range("C1:C3").FormulaArray = "=ROW(C1:C3)*10"
    wsA.Range("E1:E3").FormulaArray = "=ROW(E1:E3)*10"

    ' Shared: three formulas filled down, which Excel saves shared.
    For i = 1 To 5
        wsF.Cells(i, 1).Value = i
    Next i
    wsF.Range("B1").Formula = "=A1*2"
    wsF.Range("B1:B5").FillDown
    wsF.Range("D1").Formula = "=A1*3"
    wsF.Range("D1:D5").FillDown
    wsF.Range("F1").Formula = "=A1*4"
    wsF.Range("F1:F5").FillDown

    ' Db1 to Db4: a database in column A, its last cell A8.
    For k = 1 To 4
        Set ws = wb.Worksheets("Db" & k)
        ws.Range("A1").Value = "k"
        For i = 2 To 5
            ws.Cells(i, 1).Value = i - 1
        Next i
        If k <> 4 Then ws.Range("A8").Value = 9
        If k <> 3 Then Dress ws.Range("A8")
    Next k

    ' CF: a rule of a value and a rule of a blank, each over two cells. A
    ' rule's relative reference is read from the active cell, so C1 is
    ' selected before its rule is added.
    wsK.Range("A1:A2").Value = 5
    Set fc = wsK.Range("A1:A2").FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="=3")
    fc.Interior.Color = RGB(255, 0, 0)
    wsK.Range("C1:C2").Value = 5
    wsK.Activate
    wsK.Range("C1").Select
    Set fc = wsK.Range("C1:C2").FormatConditions.Add(Type:=xlExpression, Formula1:="=ISBLANK(C1)")
    fc.Interior.Color = RGB(0, 176, 80)

    ' Read: the databases' readers, a row a database; its criterion A1:A2,
    ' the field's name over a blank.
    wsR.Range("A1").Value = "k"
    For k = 1 To 4
        wsR.Cells(k, 2).Formula = "=DCOUNTA(Db" & k & "!A:A,,$A$1:$A$2)"
        wsR.Cells(k, 3).Formula = "=COUNTA(Db" & k & "!A:A)"
        wsR.Cells(k, 4).Formula = "=ROWS(Db" & k & "!A:A)"
        wsR.Cells(k, 5).Formula = "=DCOUNTA(Db" & k & "!A1:A20,,$A$1:$A$2)"
    Next k
    wsC.Activate
    wsC.Range("A1").Select
End Sub

Private Sub ApplyClears(ByVal wb As Workbook)
    Dim ws As Worksheet
    Set ws = wb.Worksheets("Clear")
    ClearBy "Clear!A1 by ClearContents, a styled 5", ws.Range("A1"), "contents"
    ClearBy "Clear!A2 by Clear, a styled 5", ws.Range("A2"), "all"
    ClearBy "Clear!A4 by ClearContents, a styled formula", ws.Range("A4"), "contents"
    ClearBy "Clear!A6 by Value = """", a styled 5", ws.Range("A6"), "value-text"
    ClearBy "Clear!A7 by Value = Empty, a styled 5", ws.Range("A7"), "value-empty"
    ClearBy "Clear!A8 by Formula = """", a styled 5", ws.Range("A8"), "formula-text"

    Set ws = wb.Worksheets("Spill")
    ClearBy "Spill!A3 by ClearContents, the 7 blocking A1", ws.Range("A3"), "contents"
    ClearBy "Spill!C2 by ClearContents, a cell C1 spills into", ws.Range("C2"), "contents"
    ClearBy "Spill!E1 by ClearContents, a spill's anchor", ws.Range("E1"), "contents"
    ClearBy "Spill!G1:G3 by ClearContents, a spill whole", ws.Range("G1:G3"), "contents"
    ClearBy "Spill!I2 by ClearContents, the styled 5 blocking I1", ws.Range("I2"), "contents"

    Set ws = wb.Worksheets("Arr")
    ClearBy "Arr!A2 by ClearContents, part of a block", ws.Range("A2"), "contents"
    Fact "clear", "Arr!A1:A3 after A2's clear", Spell(ws.Range("A1")) & "; " & Spell(ws.Range("A2")) & "; " & Spell(ws.Range("A3"))
    ClearBy "Arr!A1:A3 by ClearContents, the block whole", ws.Range("A1:A3"), "contents"
    ClearBy "Arr!C1:C4 by ClearContents, over a block", ws.Range("C1:C4"), "contents"
    ClearBy "Arr!E1 by ClearContents, a block's anchor alone", ws.Range("E1"), "contents"

    Set ws = wb.Worksheets("Shared")
    ClearBy "Shared!B1 by ClearContents, a fill's first cell", ws.Range("B1"), "contents"
    ClearBy "Shared!D3 by ClearContents, a cell inside a fill", ws.Range("D3"), "contents"
    ClearBy "Shared!F1:F2 by ClearContents, a fill's first two", ws.Range("F1:F2"), "contents"

    ClearBy "Db1!A8 by ClearContents, a styled 9", wb.Worksheets("Db1").Range("A8"), "contents"
    ClearBy "Db2!A8 by Clear, a styled 9", wb.Worksheets("Db2").Range("A8"), "all"
    ClearBy "Db3!A8 by ClearContents, a plain 9", wb.Worksheets("Db3").Range("A8"), "contents"

    Set ws = wb.Worksheets("CF")
    ClearBy "CF!A1 by ClearContents, under a value rule", ws.Range("A1"), "contents"
    ClearBy "CF!A2 by Clear, under a value rule", ws.Range("A2"), "all"
    ClearBy "CF!C1 by ClearContents, under a blank rule", ws.Range("C1"), "contents"
    ClearBy "CF!C2 by Clear, under a blank rule", ws.Range("C2"), "all"
End Sub

Private Sub ClearBy(ByVal what As String, ByVal r As Range, ByVal how As String)
    Dim n As Long, d As String
    On Error Resume Next
    Select Case how
        Case "contents"
            r.ClearContents
        Case "all"
            r.Clear
        Case "value-text"
            r.Value = ""
        Case "value-empty"
            r.Value = Empty
        Case "formula-text"
            r.Formula = ""
    End Select
    n = Err.Number
    d = Err.Description
    On Error GoTo 0
    If n = 0 Then
        Fact "clear", what, "done"
    Else
        Fact "clear", what, "refused, error " & n & ": " & d
    End If
End Sub

' ---- the record -----------------------------------------------------------

Private Sub RecordAll(ByVal wb As Workbook, ByVal stage As String)
    Dim ws As Worksheet, r As Range, a As Variant, i As Long, j As Long, k As Long
    For Each ws In wb.Worksheets
        UsedOf ws, stage
    Next ws

    Set ws = wb.Worksheets("Clear")
    For i = 1 To 8
        Set r = ws.Cells(i, 1)
        Fact stage, "Clear!A" & i, Spell(r) & "; formula [" & r.Formula & "]; " & Looks(r)
        For j = 2 To 8
            Set r = ws.Cells(i, j)
            Fact stage, "Clear!" & r.Address(False, False) & " " & r.Formula, Spell(r)
        Next j
    Next i

    Set ws = wb.Worksheets("Spill")
    For Each a In Array("A1", "A2", "A3", "C1", "C2", "C3", "E1", "E2", "E3", "G1", "G2", "G3", "I1", "I2", "I3")
        Set r = ws.Range(a)
        Fact stage, "Spill!" & a, Spell(r) & "; formula [" & r.Formula2 & "]; " & SpillOf(r) & "; " & Looks(r)
    Next a

    Set ws = wb.Worksheets("Arr")
    For Each a In Array("A1", "A2", "A3", "C1", "C2", "C3", "C4", "E1", "E2", "E3")
        Set r = ws.Range(a)
        Fact stage, "Arr!" & a, Spell(r) & "; formula [" & r.Formula & "]; " & ArrayOf(r)
    Next a

    Set ws = wb.Worksheets("Shared")
    For Each a In Array("B1", "B2", "B3", "B4", "B5", "D1", "D2", "D3", "D4", "D5", "F1", "F2", "F3", "F4", "F5")
        Set r = ws.Range(a)
        Fact stage, "Shared!" & a, Spell(r) & "; formula [" & r.Formula & "]"
    Next a

    For k = 1 To 4
        Set r = wb.Worksheets("Db" & k).Range("A8")
        Fact stage, "Db" & k & "!A8", Spell(r) & "; " & Looks(r)
    Next k

    Set ws = wb.Worksheets("Read")
    For k = 1 To 4
        For j = 2 To 5
            Set r = ws.Cells(k, j)
            Fact stage, "Read!" & r.Address(False, False) & " " & r.Formula, Spell(r)
        Next j
    Next k

    Set ws = wb.Worksheets("CF")
    For Each a In Array("A1", "A2", "C1", "C2")
        Set r = ws.Range(a)
        Fact stage, "CF!" & a, Spell(r) & "; " & RulesOf(r) & "; " & Looks(r)
    Next a
    For k = 1 To ws.Cells.FormatConditions.Count
        Fact stage, "CF rule " & k, "applies to " & ws.Cells.FormatConditions(k).AppliesTo.Address(False, False)
    Next k
End Sub

' A sheet's used range, with its last cell (Ctrl+End) read before and after
' the used range is, since reading it may move it.
Private Sub UsedOf(ByVal ws As Worksheet, ByVal stage As String)
    Dim atFirst As String, usedAt As String, atLast As String
    atFirst = ws.Cells.SpecialCells(xlCellTypeLastCell).Address(False, False)
    usedAt = ws.UsedRange.Address(False, False)
    atLast = ws.Cells.SpecialCells(xlCellTypeLastCell).Address(False, False)
    Fact stage, ws.Name & " used range", usedAt & "; last cell " & atFirst & " before it was read, " & atLast & " after"
End Sub

Private Function Spell(ByVal r As Range) As String
    Dim v As Variant
    v = r.Value
    If IsEmpty(v) Then
        Spell = "empty"
    ElseIf IsError(v) Then
        Spell = "error " & r.Text & " (" & CStr(v) & ")"
    ElseIf VarType(v) = vbString Then
        If Len(v) = 0 Then
            Spell = "a text of no characters"
        Else
            Spell = "text """ & v & """"
        End If
    ElseIf VarType(v) = vbBoolean Then
        Spell = "truth " & CStr(v)
    Else
        Spell = "number " & CStr(v)
    End If
End Function

Private Function Looks(ByVal r As Range) As String
    Looks = "fill " & r.Interior.ColorIndex & " " & r.Interior.Color & ", bold " & r.Font.Bold & _
        ", code " & r.NumberFormat & ", bottom edge " & r.Borders(xlEdgeBottom).LineStyle
End Function

Private Function SpillOf(ByVal r As Range) As String
    Dim s As String, has As Boolean, n As Long, p As String
    On Error Resume Next
    has = r.HasSpill
    n = Err.Number
    Err.Clear
    If n <> 0 Then
        s = "no HasSpill (" & n & ")"
    ElseIf has Then
        s = "spills over " & r.SpillingToRange.Address(False, False)
    Else
        s = "spills over nothing"
    End If
    Err.Clear
    p = r.SpillParent.Address(False, False)
    n = Err.Number
    Err.Clear
    On Error GoTo 0
    If n = 0 Then s = s & ", spilled from " & p
    SpillOf = s
End Function

Private Function ArrayOf(ByVal r As Range) As String
    Dim has As Boolean, n As Long, p As String
    On Error Resume Next
    has = r.HasArray
    p = r.CurrentArray.Address(False, False)
    n = Err.Number
    Err.Clear
    On Error GoTo 0
    If n = 0 Then
        ArrayOf = "array " & has & " over " & p
    Else
        ArrayOf = "array " & has
    End If
End Function

Private Function RulesOf(ByVal r As Range) As String
    Dim n As Long, c As Long
    On Error Resume Next
    n = r.FormatConditions.Count
    c = r.DisplayFormat.Interior.Color
    On Error GoTo 0
    RulesOf = "rules " & n & ", drawn fill " & c
End Function

Private Sub Dress(ByVal r As Range)
    r.Interior.Color = RGB(224, 208, 248)
    r.Font.Bold = True
    r.NumberFormat = "0.00"
    r.Borders(xlEdgeBottom).LineStyle = xlContinuous
End Sub

Private Sub Fact(ByVal stage As String, ByVal what As String, ByVal said As String)
    mN = mN + 1
    Print #mFile, "K46|" & mN & "|" & stage & "|" & what & "|" & said
End Sub
