Attribute VB_Name = "VLA_DiagSheetHelpers"
Option Explicit

' Standalone diagnostic for L-SHEET-HELPERS - NOT part of the VLA project, and
' called by nothing in Frazaro. It measures the Excel facts the sheet helpers
' (VLA_Runtime.bas, VlaAddSheetAt and the rest) are built on, so that each one
' is a measurement and not a memory, and it records what Undo Last Run does to
' a workbook's sheets after each helper, which the helpers leave alone and
' U.32 is filed for.
'
' WHERE IT RUNS. Import it into a NEW BLANK WORKBOOK that is never saved. Run
' every entry point from Excel's Macros dialog (Alt+F8, Macros in: All Open
' Workbooks); its lines print to the Immediate window (Alt+F11, then Ctrl+G).
' A macro that adds a workbook can move the VBE's active project (the VBA
' traps, 19), so the Macros dialog, never the Immediate window, for these.
'
' ENTRY POINTS, each a macro with no arguments:
'   LSHProbe     builds its own scratch workbooks (Workbooks.Add, closed
'                without saving at the end) and prints one LSH| line per
'                fact, numbered, with what was measured. Nothing else is
'                touched; the dev workbook may be open beside it.
'   LSHRecord    records the ACTIVE workbook's sheets - name, position,
'                visibility, which is active - on this workbook's own first
'                sheet (column A from row 3; B1 the book, B2 the active
'                sheet), never in module memory: a Run under Compile adds
'                code to the dev workbook's project, and that resets every
'                open project's variables (the owner's run, 2026-10-07: the
'                record vanished under Compile and survived under
'                Interpret). Make the dev workbook active first (click it),
'                then Alt+F8, LSHRecord.
'   LSHCompare   prints what changed in the active workbook since LSHRecord:
'                each recorded sheet's new position and visibility or GONE,
'                every NEW sheet, and the active sheet before and after.
'                Run it after a Run (what the helper did), and again after
'                Undo Last Run (what Undo put back, and what it did not).
'                LSHCompare does not reset the record; the next LSHRecord does.
'
' THE FACTS LSHProbe MEASURES (the helpers' rules, one line each):
'   1-2   Worksheet.Copy and Worksheet.Move with neither Before nor After
'         make a NEW WORKBOOK (Workbooks.Count rises by one), so the helpers
'         always pass one.
'   3-9   which acts change ActiveSheet: Worksheets.Add, Copy, Move, a
'         rename, Visible = xlSheetHidden on another sheet, Visible =
'         xlSheetVisible, Cells.Clear. The helpers give the active sheet
'         back after the ones that change it.
'   10-11 a copy made After a very-hidden last sheet lands BEFORE it (the
'         VBA traps, 12), and whether a Move does the same.
'   12    a copy of a hidden sheet is hidden or visible.
'   13-18 Excel's own error for each act the helpers refuse first in words:
'         a rename to a taken name, a rename that changes only case, hiding
'         the only visible worksheet, deleting the only worksheet, a move
'         after itself, and each act on a workbook whose structure is
'         protected.
'   19-20 whether a visible chart sheet lets the only worksheet be hidden
'         or deleted (the helpers count worksheets alone).
' Every line is LSH|<n>|<what>|<measured>. Paste them all back.

' No module-level state: the roster lives on this workbook's first sheet
' (RosterSheet), since a Run under Compile resets every project's variables.
' Every line LSHRecord and LSHCompare print carries this version and the
' place the roster was written to or looked for, so an old copy of this
' module left in a project can never be mistaken for this one.
Private Const LSH_VERSION As String = "v2"

Public Sub LSHProbe()
    Dim scr As Boolean, da As Boolean
    scr = Application.ScreenUpdating
    da = Application.DisplayAlerts
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Dim wb As Workbook
    Dim wb2 As Workbook
    Dim nBefore As Long
    Dim n As Long, d As String
    On Error GoTo failed
    Set wb = FreshBook()
    AddAfter wb, "Beta", "Alpha"
    AddAfter wb, "Gamma", "Beta"
    wb.Worksheets("Alpha").Activate
    Debug.Print "LSH|0|Excel " & Application.Version & ", scratch workbook " & wb.Name & "|" & OrderOf(wb) & "; active " & wb.ActiveSheet.Name

    ' 1-2: Copy and Move with neither Before nor After.
    nBefore = Workbooks.Count
    wb.Worksheets("Beta").Copy
    Debug.Print "LSH|1|Worksheet.Copy with neither Before nor After|workbooks " & nBefore & " -> " & Workbooks.Count & "; active workbook " & ActiveWorkbook.Name
    If ActiveWorkbook.Name <> wb.Name Then ActiveWorkbook.Close SaveChanges:=False
    wb.Activate
    nBefore = Workbooks.Count
    wb.Worksheets("Gamma").Move
    Debug.Print "LSH|2|Worksheet.Move with neither Before nor After|workbooks " & nBefore & " -> " & Workbooks.Count & "; Gamma still in the scratch workbook: " & HasSheet(wb, "Gamma")
    If ActiveWorkbook.Name <> wb.Name Then ActiveWorkbook.Close SaveChanges:=False
    wb.Activate
    If Not HasSheet(wb, "Gamma") Then AddAfter wb, "Gamma", "Beta"

    ' 3-9: which acts change the active sheet.
    wb.Worksheets("Alpha").Activate
    Dim ws As Worksheet
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets("Gamma"))
    ws.Name = "Delta"
    Debug.Print "LSH|3|Worksheets.Add After:=Gamma, named Delta|active " & wb.ActiveSheet.Name & " (was Alpha); " & OrderOf(wb)
    wb.Worksheets("Alpha").Activate
    wb.Worksheets("Beta").Copy After:=wb.Worksheets("Delta")
    Debug.Print "LSH|4|Copy After:=Delta|active " & wb.ActiveSheet.Name & " (was Alpha); " & OrderOf(wb)
    wb.Worksheets("Alpha").Activate
    wb.Worksheets("Beta (2)").Move Before:=wb.Worksheets("Alpha")
    Debug.Print "LSH|5|Move Before:=Alpha|active " & wb.ActiveSheet.Name & " (was Alpha); " & OrderOf(wb)
    wb.Worksheets("Alpha").Activate
    wb.Worksheets("Beta (2)").Name = "Copy"
    Debug.Print "LSH|6|.Name = Copy|active " & wb.ActiveSheet.Name & " (was Alpha); " & OrderOf(wb)
    wb.Worksheets("Copy").Visible = xlSheetHidden
    Debug.Print "LSH|7|Visible = xlSheetHidden on Copy|active " & wb.ActiveSheet.Name & " (was Alpha); " & OrderOf(wb)
    wb.Worksheets("Copy").Visible = xlSheetVisible
    Debug.Print "LSH|8|Visible = xlSheetVisible on Copy|active " & wb.ActiveSheet.Name & " (was Alpha); " & OrderOf(wb)
    wb.Worksheets("Gamma").Range("A1").Value = "x"
    wb.Worksheets("Gamma").Cells.Clear
    Debug.Print "LSH|9|Cells.Clear on Gamma|active " & wb.ActiveSheet.Name & " (was Alpha); Gamma!A1 now '" & CStr(wb.Worksheets("Gamma").Range("A1").Value) & "'"

    ' 10-11: After a very-hidden last sheet.
    wb.Worksheets("Delta").Move After:=wb.Worksheets(wb.Worksheets.Count)
    wb.Worksheets("Delta").Visible = xlSheetVeryHidden
    wb.Worksheets("Alpha").Activate
    wb.Worksheets("Beta").Copy After:=wb.Worksheets("Delta")
    Debug.Print "LSH|10|Copy After:=Delta (very hidden, last)|" & OrderOf(wb) & "; the copy lands before or after Delta as shown"
    wb.Worksheets("Gamma").Move After:=wb.Worksheets("Delta")
    Debug.Print "LSH|11|Move Gamma After:=Delta (very hidden, last)|" & OrderOf(wb)
    wb.Worksheets("Delta").Visible = xlSheetVisible

    ' 12: a copy of a hidden sheet.
    wb.Worksheets("Copy").Visible = xlSheetHidden
    wb.Worksheets("Alpha").Activate
    wb.Worksheets("Copy").Copy After:=wb.Worksheets("Alpha")
    Debug.Print "LSH|12|Copy of a hidden sheet (Copy) After:=Alpha|" & OrderOf(wb) & "; active " & wb.ActiveSheet.Name
    wb.Worksheets("Copy").Visible = xlSheetVisible

    ' 13-18: Excel's own errors.
    On Error Resume Next
    Err.Clear
    wb.Worksheets("Gamma").Name = "Alpha"
    n = Err.Number: d = Err.Description
    On Error GoTo failed
    Debug.Print "LSH|13|Rename Gamma to Alpha (taken)|error " & n & ": " & d
    On Error Resume Next
    Err.Clear
    wb.Worksheets("Gamma").Name = "GAMMA"
    n = Err.Number: d = Err.Description
    On Error GoTo failed
    Debug.Print "LSH|14|Rename Gamma to GAMMA (case only)|error " & n & ": " & d & "; name now " & NameOf(wb, "gamma")
    HideAllBut wb, "Alpha"
    On Error Resume Next
    Err.Clear
    wb.Worksheets("Alpha").Visible = xlSheetHidden
    n = Err.Number: d = Err.Description
    On Error GoTo failed
    Debug.Print "LSH|15|Hide the only visible worksheet|error " & n & ": " & d & "; " & OrderOf(wb)
    ShowAll wb
    Set wb2 = FreshBook()
    On Error Resume Next
    Err.Clear
    wb2.Worksheets(1).Delete
    n = Err.Number: d = Err.Description
    On Error GoTo failed
    Debug.Print "LSH|16|Delete the only worksheet|error " & n & ": " & d & "; sheets now " & wb2.Worksheets.Count
    wb2.Close SaveChanges:=False
    Set wb2 = Nothing
    wb.Activate
    On Error Resume Next
    Err.Clear
    wb.Worksheets("Beta").Move After:=wb.Worksheets("Beta")
    n = Err.Number: d = Err.Description
    On Error GoTo failed
    Debug.Print "LSH|17|Move Beta After:=Beta (itself)|error " & n & ": " & d & "; " & OrderOf(wb)
    wb.Protect Structure:=True
    Debug.Print "LSH|18|structure protected: Add|" & TryAdd(wb)
    Debug.Print "LSH|18|structure protected: Copy After|" & TryCopy(wb)
    Debug.Print "LSH|18|structure protected: Move After|" & TryMove(wb)
    Debug.Print "LSH|18|structure protected: Rename|" & TryRename(wb)
    Debug.Print "LSH|18|structure protected: Hide|" & TryHide(wb)
    Debug.Print "LSH|18|structure protected: Delete|" & TryDelete(wb)
    wb.Unprotect
    Debug.Print "LSH|18|structure unprotected again|" & OrderOf(wb)

    ' 19-20: a chart sheet beside the only worksheet.
    Set wb2 = FreshBook()
    On Error Resume Next
    Err.Clear
    wb2.Charts.Add
    n = Err.Number: d = Err.Description
    On Error GoTo failed
    If n <> 0 Then
        Debug.Print "LSH|19|Charts.Add in a blank workbook|error " & n & ": " & d & " (19 and 20 not measured)"
    Else
        Debug.Print "LSH|19|Charts.Add|sheets " & wb2.Sheets.Count & ", worksheets " & wb2.Worksheets.Count & ", active " & wb2.ActiveSheet.Name
        On Error Resume Next
        Err.Clear
        wb2.Worksheets(1).Visible = xlSheetHidden
        n = Err.Number: d = Err.Description
        On Error GoTo failed
        Debug.Print "LSH|19|Hide the only worksheet while a chart sheet shows|error " & n & ": " & d & "; worksheet visible = " & wb2.Worksheets(1).Visible
        wb2.Worksheets(1).Visible = xlSheetVisible
        On Error Resume Next
        Err.Clear
        wb2.Worksheets(1).Delete
        n = Err.Number: d = Err.Description
        On Error GoTo failed
        Debug.Print "LSH|20|Delete the only worksheet while a chart sheet exists|error " & n & ": " & d & "; worksheets now " & wb2.Worksheets.Count
    End If
    wb2.Close SaveChanges:=False
    Set wb2 = Nothing

    wb.Close SaveChanges:=False
    Application.ScreenUpdating = scr
    Application.DisplayAlerts = da
    Debug.Print "LSH|done|every scratch workbook closed without saving"
    Exit Sub
failed:
    d = Err.Description
    n = Err.Number
    Debug.Print "LSH|ERR|stopped: error " & n & ": " & d
    On Error Resume Next
    If Not wb2 Is Nothing Then wb2.Close SaveChanges:=False
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Application.ScreenUpdating = scr
    Application.DisplayAlerts = da
End Sub

' The roster's home: this workbook's first sheet. Row 1 "book" and the
' workbook's name, row 2 "active" and the active sheet, then one row per
' worksheet from row 3: the name (behind an apostrophe, so a name shaped
' like a number or a formula stays text), its position, its visibility.
Private Function RosterSheet() As Worksheet
    Set RosterSheet = ThisWorkbook.Worksheets(1)
End Function

' Record the active workbook's sheets: name, position, visibility, which
' is active.
Public Sub LSHRecord()
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    If wb.Name = ThisWorkbook.Name Then
        Debug.Print "LSH|record|the active workbook is this diagnostic's own - click the dev workbook first, then Alt+F8"
        Exit Sub
    End If
    Dim rs As Worksheet
    Set rs = RosterSheet()
    rs.Cells.Clear
    rs.Cells(1, 1).Value = "book"
    rs.Cells(1, 2).Value = "'" & wb.Name
    rs.Cells(2, 1).Value = "active"
    rs.Cells(2, 2).Value = "'" & ActiveNameOf(wb)
    Dim i As Long
    For i = 1 To wb.Worksheets.Count
        rs.Cells(2 + i, 1).Value = "'" & wb.Worksheets(i).Name
        rs.Cells(2 + i, 2).Value = i
        rs.Cells(2 + i, 3).Value = VisOf(wb.Worksheets(i))
    Next
    Debug.Print "LSH|record|" & LSH_VERSION & "|" & wb.Name & "|" & OrderOf(wb) & "; active " & ActiveNameOf(wb) & _
                "|roster written to " & ThisWorkbook.Name & "!" & rs.Name & " rows 1-" & (2 + wb.Worksheets.Count)
End Sub

' What changed since LSHRecord in the active workbook.
Public Sub LSHCompare()
    Dim rs As Worksheet
    Set rs = RosterSheet()
    If CStr(rs.Cells(1, 1).Value) <> "book" Then
        Debug.Print "LSH|compare|" & LSH_VERSION & "|nothing recorded on " & ThisWorkbook.Name & "!" & rs.Name & " (A1 is '" & CStr(rs.Cells(1, 1).Value) & "') - run LSHRecord first"
        Exit Sub
    End If
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    Dim recordedBook As String, recordedActive As String
    recordedBook = CStr(rs.Cells(1, 2).Value)
    recordedActive = CStr(rs.Cells(2, 2).Value)
    Dim note As String
    If wb.Name <> recordedBook Then note = " (recorded " & recordedBook & ")"
    Debug.Print "LSH|compare|" & LSH_VERSION & "|" & wb.Name & note & "|" & OrderOf(wb) & "|roster read from " & ThisWorkbook.Name & "!" & rs.Name
    Dim recorded As Collection
    Set recorded = New Collection
    Dim r As Long
    Dim nm As String
    r = 3
    Do While Len(CStr(rs.Cells(r, 1).Value)) > 0
        nm = CStr(rs.Cells(r, 1).Value)
        recorded.Add nm, nm
        If HasSheet(wb, nm) Then
            Debug.Print "LSH|compare|" & nm & "|was " & CStr(rs.Cells(r, 2).Value) & " " & CStr(rs.Cells(r, 3).Value) & "|now " & PosOf(wb, nm) & " " & VisOf(wb.Worksheets(nm))
        Else
            Debug.Print "LSH|compare|" & nm & "|was " & CStr(rs.Cells(r, 2).Value) & " " & CStr(rs.Cells(r, 3).Value) & "|GONE"
        End If
        r = r + 1
    Loop
    Dim i As Long
    For i = 1 To wb.Worksheets.Count
        If Not HasKey(recorded, wb.Worksheets(i).Name) Then
            Debug.Print "LSH|compare|" & wb.Worksheets(i).Name & "|NEW|now " & i & " " & VisOf(wb.Worksheets(i))
        End If
    Next
    Debug.Print "LSH|compare|active|was " & recordedActive & "|now " & ActiveNameOf(wb)
End Sub

' --- helpers --------------------------------------------------------------

' A new workbook with one worksheet, Alpha, whatever Excel's new-workbook
' sheet count is.
Private Function FreshBook() As Workbook
    Dim wb As Workbook
    Set wb = Workbooks.Add
    Do While wb.Worksheets.Count > 1
        wb.Worksheets(wb.Worksheets.Count).Delete
    Loop
    wb.Worksheets(1).Name = "Alpha"
    Set FreshBook = wb
End Function

Private Sub AddAfter(ByVal wb As Workbook, ByVal nm As String, ByVal afterName As String)
    Dim ws As Worksheet
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(afterName))
    ws.Name = nm
End Sub

Private Function HasSheet(ByVal wb As Workbook, ByVal nm As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(nm)
    On Error GoTo 0
    HasSheet = Not ws Is Nothing
End Function

Private Function NameOf(ByVal wb As Workbook, ByVal nm As String) As String
    On Error Resume Next
    NameOf = wb.Worksheets(nm).Name
    On Error GoTo 0
End Function

Private Function ActiveNameOf(ByVal wb As Workbook) As String
    On Error Resume Next
    ActiveNameOf = wb.ActiveSheet.Name
    On Error GoTo 0
End Function

Private Function VisOf(ByVal ws As Worksheet) As String
    Select Case ws.Visible
        Case xlSheetVisible: VisOf = "visible"
        Case xlSheetHidden: VisOf = "hidden"
        Case xlSheetVeryHidden: VisOf = "very hidden"
        Case Else: VisOf = "visible=" & ws.Visible
    End Select
End Function

' The worksheet names in tab order, a hidden one marked.
Private Function OrderOf(ByVal wb As Workbook) As String
    Dim i As Long
    Dim r As String
    For i = 1 To wb.Worksheets.Count
        If i > 1 Then r = r & ","
        r = r & wb.Worksheets(i).Name
        If wb.Worksheets(i).Visible <> xlSheetVisible Then r = r & "(" & VisOf(wb.Worksheets(i)) & ")"
    Next
    OrderOf = r
End Function

Private Function PosOf(ByVal wb As Workbook, ByVal nm As String) As Long
    Dim i As Long
    For i = 1 To wb.Worksheets.Count
        If StrComp(wb.Worksheets(i).Name, nm, vbTextCompare) = 0 Then
            PosOf = i
            Exit Function
        End If
    Next
End Function

Private Function HasKey(ByVal col As Collection, ByVal k As String) As Boolean
    Dim v As Variant
    On Error Resume Next
    v = col.Item(k)
    HasKey = (Err.Number = 0)
    On Error GoTo 0
End Function

Private Sub HideAllBut(ByVal wb As Workbook, ByVal keepName As String)
    Dim i As Long
    For i = 1 To wb.Worksheets.Count
        If StrComp(wb.Worksheets(i).Name, keepName, vbTextCompare) <> 0 Then wb.Worksheets(i).Visible = xlSheetHidden
    Next
End Sub

Private Sub ShowAll(ByVal wb As Workbook)
    Dim i As Long
    For i = 1 To wb.Worksheets.Count
        wb.Worksheets(i).Visible = xlSheetVisible
    Next
End Sub

' Each act on a structure-protected workbook, its own error scope; the
' error number and words, or "no error" with what happened.
Private Function TryAdd(ByVal wb As Workbook) As String
    Dim n As Long, d As String
    On Error Resume Next
    Err.Clear
    wb.Worksheets.Add After:=wb.Worksheets(wb.Worksheets.Count)
    n = Err.Number: d = Err.Description
    On Error GoTo 0
    TryAdd = ErrText(n, d) & "; " & OrderOf(wb)
End Function

Private Function TryCopy(ByVal wb As Workbook) As String
    Dim n As Long, d As String
    On Error Resume Next
    Err.Clear
    wb.Worksheets("Beta").Copy After:=wb.Worksheets("Beta")
    n = Err.Number: d = Err.Description
    On Error GoTo 0
    TryCopy = ErrText(n, d) & "; " & OrderOf(wb)
End Function

Private Function TryMove(ByVal wb As Workbook) As String
    Dim n As Long, d As String
    On Error Resume Next
    Err.Clear
    wb.Worksheets("Beta").Move After:=wb.Worksheets(wb.Worksheets.Count)
    n = Err.Number: d = Err.Description
    On Error GoTo 0
    TryMove = ErrText(n, d) & "; " & OrderOf(wb)
End Function

Private Function TryRename(ByVal wb As Workbook) As String
    Dim n As Long, d As String
    On Error Resume Next
    Err.Clear
    wb.Worksheets("Beta").Name = "Beta renamed"
    n = Err.Number: d = Err.Description
    On Error GoTo 0
    TryRename = ErrText(n, d) & "; " & OrderOf(wb)
    On Error Resume Next
    wb.Worksheets("Beta renamed").Name = "Beta"
    On Error GoTo 0
End Function

Private Function TryHide(ByVal wb As Workbook) As String
    Dim n As Long, d As String
    On Error Resume Next
    Err.Clear
    wb.Worksheets("Beta").Visible = xlSheetHidden
    n = Err.Number: d = Err.Description
    On Error GoTo 0
    TryHide = ErrText(n, d) & "; " & OrderOf(wb)
    On Error Resume Next
    wb.Worksheets("Beta").Visible = xlSheetVisible
    On Error GoTo 0
End Function

Private Function TryDelete(ByVal wb As Workbook) As String
    Dim n As Long, d As String
    On Error Resume Next
    Err.Clear
    wb.Worksheets("Copy").Delete
    n = Err.Number: d = Err.Description
    On Error GoTo 0
    TryDelete = ErrText(n, d) & "; " & OrderOf(wb)
End Function

Private Function ErrText(ByVal n As Long, ByVal d As String) As String
    If n = 0 Then
        ErrText = "no error"
    Else
        ErrText = "error " & n & ": " & d
    End If
End Function
