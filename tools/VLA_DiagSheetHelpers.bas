Attribute VB_Name = "VLA_DiagSheetHelpers"
Option Explicit

' Standalone diagnostic for L-SHEET-HELPERS - NOT part of the VLA project, and
' called by nothing in Frazaro. It measures the Excel facts the sheet helpers
' (VLA_Runtime.bas, VlaAddSheetAt and the rest) are built on, so that each one
' is a measurement and not a memory, and it records what Undo Last Run does to
' a workbook's sheets after each helper, which the helpers leave alone and
' U.32 is filed for.
'
' WHERE IT RUNS. Import it into any open workbook that is never saved, the
' workbook under test included; never into the dev workbook, which is saved.
' Run every entry point from Excel's Macros dialog (Alt+F8, Macros in: All
' Open Workbooks); its lines print to the Immediate window (Alt+F11, then
' Ctrl+G).
' A macro that adds a workbook can move the VBE's active project (the VBA
' traps, 19), so the Macros dialog, never the Immediate window, for these.
'
' ENTRY POINTS, each a macro with no arguments:
'   LSHProbe     builds its own scratch workbooks (Workbooks.Add, closed
'                without saving at the end) and prints one LSH| line per
'                fact, numbered, with what was measured. Nothing else is
'                touched; the dev workbook may be open beside it.
'   LSHRecord    records the ACTIVE workbook's sheets - name, place,
'                visibility, which is active - in a text file in the user's
'                temp folder (RosterPath), never in module memory: a Run
'                under Compile adds code to a project, and that resets every
'                open project's variables (the owner's run, 2026-10-07: the
'                record vanished under Compile and survived under
'                Interpret). Click a cell on the Frazaro tab first, so the
'                workbook that holds it is the active one, then Alt+F8,
'                LSHRecord; the line it prints names the workbook recorded.
'                A workbook with no Frazaro tab is refused, by name.
'                v5 (2026-10-09): v4 kept the record on its own workbook's
'                first sheet and refused when that workbook was the active
'                one, which is exactly the case when the module was
'                imported into the workbook under test (the owner's run,
'                2026-10-09); a file serves wherever the module sits.
'   LSHCompare   prints what changed in the active workbook since LSHRecord:
'                each recorded sheet's place and visibility now, or GONE,
'                every NEW sheet, and the active sheet before and after,
'                each line that differs marked DIFFERS, then one verdict
'                line: "as recorded", or how many lines differ. Run it after
'                a Run (what the helper did), and again after Undo Last Run
'                (what Undo put back, and what it did not). LSHCompare does
'                not reset the record; the next LSHRecord does.
'                v4 (U.32's close, 2026-10-09): Frazaro's snapshot sheets
'                (VLAu_, VLAd_, VLAn_, VLAs_) are its bookkeeping, remade at
'                every Run, so they are left out of the record and the
'                compare; a sheet's place is counted among the worksheets
'                that are neither those nor very hidden, since every sweep
'                of old snapshot sheets shifts an absolute index ("-" for a
'                very-hidden sheet); and names are compared exactly, so a
'                sheet that comes back as "sales" for "Sales" DIFFERS.
'   LSHProbeMoves (v3, for U.32) measures the facts Undo's roster is built
'                on: moving a hidden or a very-hidden sheet, a Move After or
'                Before a hidden or a very-hidden sheet in the middle of the
'                tab order, activating a hidden sheet, hiding and deleting
'                the active sheet, what deleting a very-hidden sheet does to
'                the positions after it, a Move beside a chart sheet, and a
'                Move beside a hidden neighbour. Its own scratch workbook,
'                closed without saving. Facts 21-33.
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
' THE FACTS LSHProbeMoves MEASURES (the roster's rules, U.32):
'   21-22 whether a hidden, then a very-hidden, sheet can be moved at all,
'         and whether the move changes the active sheet.
'   23-26 where a visible sheet lands when moved After, then Before, a
'         hidden sheet in the MIDDLE of the tab order, and the same beside a
'         very-hidden one (10-11 measured only the very-hidden LAST sheet).
'   27    Excel's error on activating a hidden sheet.
'   28-29 hiding, then deleting, the active sheet: which sheet is active after.
'   30    deleting a very-hidden sheet ahead of others: a later sheet's Index
'         and its position among the worksheets, before and after (a
'         tombstone sweep shifts every later tab, so the roster keeps
'         relative order, never absolute positions).
'   31    a Move After a worksheet that a chart sheet follows: whether the
'         moved sheet lands before or after the chart.
'   32-33 a Move After, then Before, a visible sheet whose neighbour on that
'         side is hidden: whether the moved sheet lands beside the anchor
'         or past the hidden sheet (the host suite found a Copy doing the
'         latter, 2026-10-08).
' Every line is LSH|<n>|<what>|<measured>. Paste them all back.

' No module-level state: the record lives in a file (RosterPath), since a Run
' under Compile resets every project's variables. Every line LSHRecord and
' LSHCompare print carries this version and the file the record was written
' to or read from, so an old copy of this module left in a project can never
' be mistaken for this one.
Private Const LSH_VERSION As String = "v5"

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

' U.32: the facts Undo's roster put-back is built on, each measured on a
' scratch workbook of five sheets, Alpha to Eps, Alpha active, and reported
' with Excel's error (or none), the order afterwards and the active sheet -
' so the put-back's order of steps (show, remove, move, hide, activate) is a
' measurement and not a memory. Every fact starts from the same state
' (Reorder); nothing outside the scratch workbook is touched.
Public Sub LSHProbeMoves()
    Dim scr As Boolean, da As Boolean
    scr = Application.ScreenUpdating
    da = Application.DisplayAlerts
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Dim wb As Workbook
    Dim n As Long, d As String
    Dim idxBefore As Long, posBefore As Long
    On Error GoTo failed
    Set wb = FreshBook()
    AddAfter wb, "Beta", "Alpha"
    AddAfter wb, "Gamma", "Beta"
    AddAfter wb, "Delta", "Gamma"
    AddAfter wb, "Eps", "Delta"
    wb.Worksheets("Alpha").Activate
    Debug.Print "LSH|0|" & LSH_VERSION & " Excel " & Application.Version & ", scratch workbook " & wb.Name & "|" & OrderOf(wb) & "; active " & ActiveNameOf(wb)

    ' 21: a hidden sheet moved Before a visible one.
    wb.Worksheets("Delta").Visible = xlSheetHidden
    On Error Resume Next
    Err.Clear
    wb.Worksheets("Delta").Move Before:=wb.Worksheets("Beta")
    n = Err.Number: d = Err.Description
    On Error GoTo failed
    Debug.Print "LSH|21|Move hidden Delta Before:=Beta|" & ErrText(n, d) & "; " & OrderOf(wb) & "; active " & ActiveNameOf(wb) & " (was Alpha)"

    ' 22: a very-hidden sheet moved After a visible one.
    wb.Worksheets("Delta").Visible = xlSheetVeryHidden
    On Error Resume Next
    Err.Clear
    wb.Worksheets("Delta").Move After:=wb.Worksheets("Gamma")
    n = Err.Number: d = Err.Description
    On Error GoTo failed
    Debug.Print "LSH|22|Move very-hidden Delta After:=Gamma|" & ErrText(n, d) & "; " & OrderOf(wb) & "; active " & ActiveNameOf(wb) & " (was Alpha)"

    ' 23-24: a visible sheet moved After, then Before, a hidden sheet in the middle.
    Reorder wb
    wb.Worksheets("Gamma").Visible = xlSheetHidden
    wb.Worksheets("Eps").Move After:=wb.Worksheets("Gamma")
    Debug.Print "LSH|23|Move Eps After:=Gamma (hidden, in the middle)|" & OrderOf(wb) & "; active " & ActiveNameOf(wb) & " (was Alpha)"
    Reorder wb
    wb.Worksheets("Gamma").Visible = xlSheetHidden
    wb.Worksheets("Alpha").Move Before:=wb.Worksheets("Gamma")
    Debug.Print "LSH|24|Move Alpha Before:=Gamma (hidden, in the middle)|" & OrderOf(wb) & "; active " & ActiveNameOf(wb) & " (was Alpha)"

    ' 25-26: the same beside a very-hidden sheet in the middle.
    Reorder wb
    wb.Worksheets("Gamma").Visible = xlSheetVeryHidden
    wb.Worksheets("Eps").Move After:=wb.Worksheets("Gamma")
    Debug.Print "LSH|25|Move Eps After:=Gamma (very hidden, in the middle)|" & OrderOf(wb) & "; active " & ActiveNameOf(wb) & " (was Alpha)"
    Reorder wb
    wb.Worksheets("Gamma").Visible = xlSheetVeryHidden
    wb.Worksheets("Alpha").Move Before:=wb.Worksheets("Gamma")
    Debug.Print "LSH|26|Move Alpha Before:=Gamma (very hidden, in the middle)|" & OrderOf(wb) & "; active " & ActiveNameOf(wb) & " (was Alpha)"

    ' 27: activating a hidden sheet.
    Reorder wb
    wb.Worksheets("Gamma").Visible = xlSheetHidden
    On Error Resume Next
    Err.Clear
    wb.Worksheets("Gamma").Activate
    n = Err.Number: d = Err.Description
    On Error GoTo failed
    Debug.Print "LSH|27|Activate hidden Gamma|" & ErrText(n, d) & "; active " & ActiveNameOf(wb) & " (was Alpha)"

    ' 28: hiding the active sheet while others show.
    Reorder wb
    wb.Worksheets("Gamma").Activate
    On Error Resume Next
    Err.Clear
    wb.Worksheets("Gamma").Visible = xlSheetHidden
    n = Err.Number: d = Err.Description
    On Error GoTo failed
    Debug.Print "LSH|28|Hide the active sheet Gamma|" & ErrText(n, d) & "; " & OrderOf(wb) & "; active " & ActiveNameOf(wb) & " (was Gamma)"

    ' 29: deleting the active sheet.
    Reorder wb
    wb.Worksheets("Gamma").Activate
    wb.Worksheets("Gamma").Delete
    Debug.Print "LSH|29|Delete the active sheet Gamma|" & OrderOf(wb) & "; active " & ActiveNameOf(wb) & " (was Gamma)"
    AddAfter wb, "Gamma", "Beta"

    ' 30: a very-hidden sheet deleted ahead of others, as a tombstone sweep does.
    Reorder wb
    wb.Worksheets("Beta").Visible = xlSheetVeryHidden
    idxBefore = wb.Worksheets("Eps").Index
    posBefore = PosOf(wb, "Eps")
    wb.Worksheets("Beta").Visible = xlSheetVisible
    wb.Worksheets("Beta").Delete
    Debug.Print "LSH|30|Delete very-hidden Beta ahead of Eps|Eps Index " & idxBefore & " -> " & wb.Worksheets("Eps").Index & ", position among worksheets " & posBefore & " -> " & PosOf(wb, "Eps") & "; " & OrderOf(wb)
    AddAfter wb, "Beta", "Alpha"

    ' 31: a chart sheet right after Beta, then Eps moved After:=Beta.
    Reorder wb
    On Error Resume Next
    Err.Clear
    wb.Charts.Add
    n = Err.Number: d = Err.Description
    On Error GoTo failed
    If n <> 0 Then
        Debug.Print "LSH|31|Charts.Add|" & ErrText(n, d) & " (31 not measured)"
    Else
        wb.Charts(1).Move After:=wb.Worksheets("Beta")
        Debug.Print "LSH|31|chart sheet moved After:=Beta|" & SheetsOrderOf(wb) & "; worksheets " & OrderOf(wb)
        wb.Worksheets("Eps").Move After:=wb.Worksheets("Beta")
        Debug.Print "LSH|31|Move Eps After:=Beta with the chart right after Beta|" & SheetsOrderOf(wb) & "; Eps Index " & wb.Worksheets("Eps").Index & ", position among worksheets " & PosOf(wb, "Eps")
        wb.Charts(1).Delete
    End If

    ' 32-33: a visible anchor whose neighbour is hidden. The host suite's
    ' fixture found a Copy After:=Gamma landing past the hidden Delta beside
    ' it (2026-10-08): Excel places by the visible tabs. Measured here for
    ' Move, After and Before.
    Reorder wb
    wb.Worksheets("Gamma").Visible = xlSheetHidden
    wb.Worksheets("Eps").Move After:=wb.Worksheets("Beta")
    Debug.Print "LSH|32|Move Eps After:=Beta, Gamma hidden right after Beta|" & OrderOf(wb)
    Reorder wb
    wb.Worksheets("Gamma").Visible = xlSheetHidden
    wb.Worksheets("Alpha").Move Before:=wb.Worksheets("Delta")
    Debug.Print "LSH|33|Move Alpha Before:=Delta, Gamma hidden right before Delta|" & OrderOf(wb)

    wb.Close SaveChanges:=False
    Application.ScreenUpdating = scr
    Application.DisplayAlerts = da
    Debug.Print "LSH|done|the scratch workbook closed without saving"
    Exit Sub
failed:
    d = Err.Description
    n = Err.Number
    Debug.Print "LSH|ERR|stopped: error " & n & ": " & d
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Application.ScreenUpdating = scr
    Application.DisplayAlerts = da
End Sub

' The state every fact of LSHProbeMoves starts from: Alpha, Beta, Gamma,
' Delta, Eps in that order, every one visible, Alpha active. Shown first,
' since fact 21 is what says whether a hidden sheet can be moved.
Private Sub Reorder(ByVal wb As Workbook)
    ShowAll wb
    wb.Worksheets("Beta").Move After:=wb.Worksheets("Alpha")
    wb.Worksheets("Gamma").Move After:=wb.Worksheets("Beta")
    wb.Worksheets("Delta").Move After:=wb.Worksheets("Gamma")
    wb.Worksheets("Eps").Move After:=wb.Worksheets("Delta")
    wb.Worksheets("Alpha").Activate
End Sub

' Every sheet's name in tab order, worksheets and chart sheets alike, a
' chart sheet marked.
Private Function SheetsOrderOf(ByVal wb As Workbook) As String
    Dim i As Long
    Dim r As String
    For i = 1 To wb.Sheets.Count
        If i > 1 Then r = r & ","
        r = r & wb.Sheets(i).Name
        If TypeName(wb.Sheets(i)) = "Chart" Then r = r & "[chart]"
    Next
    SheetsOrderOf = r
End Function

' The record's home: a text file in the user's temp folder, the same path
' from any project, so the module may sit in any open workbook. One line a
' fact, fields split by a tab, which no sheet name can hold: "book" and the
' workbook's name, "active" and the active sheet, then "sheet", the name,
' its place (PlaceText) and its visibility for every worksheet in tab order,
' Frazaro's snapshot sheets left out. Written and read as the system's code
' page, which serves the ASCII names a test uses.
Private Function RosterPath() As String
    RosterPath = Environ$("TEMP") & "\VLA_DiagSheetHelpers_record.txt"
End Function

' Whether the workbook holds a program tab, "Frazaro" or "Frazaro (<name>)":
' the workbook an Undo test measures always does.
Private Function HasFrazaroTab(ByVal wb As Workbook) As Boolean
    Dim i As Long
    Dim k As String
    For i = 1 To wb.Worksheets.Count
        k = LCase$(wb.Worksheets(i).Name)
        If k = "frazaro" Then
            HasFrazaroTab = True
            Exit Function
        End If
        If Left$(k, 9) = "frazaro (" And Right$(k, 1) = ")" Then
            HasFrazaroTab = True
            Exit Function
        End If
    Next
End Function

' Record the active workbook's sheets: name, place, visibility, which is
' active.
Public Sub LSHRecord()
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    If Not HasFrazaroTab(wb) Then
        Debug.Print "LSH|record|" & LSH_VERSION & "|" & wb.Name & " has no Frazaro tab - click a cell on the Frazaro tab of the workbook under test, then Alt+F8"
        Exit Sub
    End If
    Dim f As Integer
    Dim i As Long, n As Long
    f = FreeFile
    Open RosterPath() For Output As #f
    Print #f, "book" & vbTab & wb.Name
    Print #f, "active" & vbTab & ActiveNameOf(wb)
    For i = 1 To wb.Worksheets.Count
        If Not IsSnapshotName(wb.Worksheets(i).Name) Then
            Print #f, "sheet" & vbTab & wb.Worksheets(i).Name & vbTab & PlaceText(wb, wb.Worksheets(i).Name) & vbTab & VisOf(wb.Worksheets(i))
            n = n + 1
        End If
    Next
    Close #f
    Debug.Print "LSH|record|" & LSH_VERSION & "|" & wb.Name & "|" & UserOrderOf(wb) & "; active " & ActiveNameOf(wb) & _
                "|" & n & " sheets recorded in " & RosterPath()
End Sub

' What changed since LSHRecord in the active workbook, each line that
' differs marked, then the verdict.
Public Sub LSHCompare()
    If Len(Dir$(RosterPath())) = 0 Then
        Debug.Print "LSH|compare|" & LSH_VERSION & "|nothing recorded: " & RosterPath() & " is not there - run LSHRecord first"
        Exit Sub
    End If
    Dim recordedBook As String, recordedActive As String
    Dim recNames As Collection, recWas As Collection
    Set recNames = New Collection
    Set recWas = New Collection
    Dim f As Integer
    Dim ln As String
    Dim fields() As String
    f = FreeFile
    Open RosterPath() For Input As #f
    Do While Not EOF(f)
        Line Input #f, ln
        If Len(ln) > 0 Then
            fields = Split(ln, vbTab)
            If fields(0) = "book" And UBound(fields) >= 1 Then recordedBook = fields(1)
            If fields(0) = "active" And UBound(fields) >= 1 Then recordedActive = fields(1)
            If fields(0) = "sheet" And UBound(fields) >= 3 Then
                recNames.Add fields(1)
                recWas.Add fields(2) & " " & fields(3)
            End If
        End If
    Loop
    Close #f
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    Dim differ As Long
    Dim note As String
    If wb.Name <> recordedBook Then
        note = " (RECORDED " & recordedBook & ", ANOTHER WORKBOOK)"
        differ = differ + 1
    End If
    Debug.Print "LSH|compare|" & LSH_VERSION & "|" & wb.Name & note & "|" & UserOrderOf(wb) & "|record read from " & RosterPath()
    Dim recorded As Collection
    Set recorded = New Collection
    Dim r As Long
    Dim nm As String, wasText As String, nowText As String
    For r = 1 To recNames.Count
        nm = CStr(recNames.Item(r))
        recorded.Add nm, nm
        wasText = CStr(recWas.Item(r))
        If HasSheet(wb, nm) Then
            nowText = PlaceText(wb, nm) & " " & VisOf(wb.Worksheets(nm))
            If wb.Worksheets(nm).Name <> nm Then nowText = nowText & ", named " & wb.Worksheets(nm).Name
            If nowText = wasText Then
                Debug.Print "LSH|compare|" & nm & "|was " & wasText & "|now " & nowText
            Else
                Debug.Print "LSH|compare|" & nm & "|was " & wasText & "|now " & nowText & "|DIFFERS"
                differ = differ + 1
            End If
        Else
            Debug.Print "LSH|compare|" & nm & "|was " & wasText & "|GONE"
            differ = differ + 1
        End If
    Next
    Dim i As Long
    For i = 1 To wb.Worksheets.Count
        If Not IsSnapshotName(wb.Worksheets(i).Name) Then
            If Not HasKey(recorded, wb.Worksheets(i).Name) Then
                Debug.Print "LSH|compare|" & wb.Worksheets(i).Name & "|NEW|now " & PlaceText(wb, wb.Worksheets(i).Name) & " " & VisOf(wb.Worksheets(i))
                differ = differ + 1
            End If
        End If
    Next
    If ActiveNameOf(wb) = recordedActive Then
        Debug.Print "LSH|compare|active|was " & recordedActive & "|now " & ActiveNameOf(wb)
    Else
        Debug.Print "LSH|compare|active|was " & recordedActive & "|now " & ActiveNameOf(wb) & "|DIFFERS"
        differ = differ + 1
    End If
    If differ = 0 Then
        Debug.Print "LSH|compare|verdict|as recorded"
    ElseIf differ = 1 Then
        Debug.Print "LSH|compare|verdict|1 line differs"
    Else
        Debug.Print "LSH|compare|verdict|" & differ & " lines differ"
    End If
End Sub

' Frazaro's snapshot sheets, in any case: an Undo copy, a tombstone, a
' Run's staging and its roster (VLAu_, VLAd_, VLAn_, VLAs_).
Private Function IsSnapshotName(ByVal nm As String) As Boolean
    Select Case LCase$(Left$(nm, 5))
        Case "vlau_", "vlad_", "vlan_", "vlas_"
            IsSnapshotName = True
    End Select
End Function

' A sheet's place among the worksheets a person can see or unhide - the
' ones that are neither snapshot sheets nor very hidden - as text; "-" for
' a very-hidden sheet, "" for a name the workbook does not have.
Private Function PlaceText(ByVal wb As Workbook, ByVal nm As String) As String
    Dim i As Long, n As Long
    For i = 1 To wb.Worksheets.Count
        If Not IsSnapshotName(wb.Worksheets(i).Name) Then
            If wb.Worksheets(i).Visible = xlSheetVeryHidden Then
                If StrComp(wb.Worksheets(i).Name, nm, vbTextCompare) = 0 Then
                    PlaceText = "-"
                    Exit Function
                End If
            Else
                n = n + 1
                If StrComp(wb.Worksheets(i).Name, nm, vbTextCompare) = 0 Then
                    PlaceText = CStr(n)
                    Exit Function
                End If
            End If
        End If
    Next
End Function

' The worksheet names in tab order, the snapshot sheets left out, a hidden
' one marked.
Private Function UserOrderOf(ByVal wb As Workbook) As String
    Dim i As Long
    Dim r As String
    For i = 1 To wb.Worksheets.Count
        If Not IsSnapshotName(wb.Worksheets(i).Name) Then
            If Len(r) > 0 Then r = r & ","
            r = r & wb.Worksheets(i).Name
            If wb.Worksheets(i).Visible <> xlSheetVisible Then r = r & "(" & VisOf(wb.Worksheets(i)) & ")"
        End If
    Next
    UserOrderOf = r
End Function

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
