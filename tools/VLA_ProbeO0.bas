Attribute VB_Name = "VLA_ProbeO0"
Option Explicit

' Standalone probe for OPTIMIZE.0 - NOT part of the VLA project, and it calls
' nothing in it. Import it into a NEW, BLANK workbook saved as .xlsm (not the
' dev workbook: a hang here should cost nothing), then follow
' archive/optimize0_live_steps.md, part C. Delete the workbook afterwards.
'
' WHAT IT ASKS. OPTIMIZE's search may run long inside a worksheet formula, and
' nobody knows what the host allows there. Each worksheet function below asks
' one question: does Esc stop a looping UDF (and a looping macro), and does
' Application.EnableCancelKey change that; does a MsgBox or a modal UserForm
' show from inside a UDF, and what the Function Wizard does with one; can a UDF
' set the status bar, or schedule a macro with Application.OnTime; and what
' DoEvents inside a UDF does - in particular whether Excel re-enters the UDF
' while it waits. (A ribbon cannot be invalidated here: Frazaro ships no custom
' ribbon, so there is nothing to invalidate.)
'
' SAFETY. Every loop is bounded by Timer (midnight handled), 30 seconds at
' most. Every modal (O0Ask, O0ShowForm) and every scheduled macro (O0Later) is
' capped at 3 per session, so a recalculation storm or the Function Wizard
' cannot stack dialogs; O0ProbeReset clears the caps. Every call prints an
' O0P| line to the Immediate window, so what happened is on record even when
' the sheet shows nothing.

Private Const MAX_SPIN As Double = 30#
Private Const MAX_PUMP As Double = 15#
Private Const MAX_MODALS As Long = 3

Private mCalls As Long
Private mAsks As Long
Private mForms As Long
Private mLaters As Long
Private mDepth As Long
Private mMaxDepth As Long
Private mLaterBook As String
Private mLaterSheet As String
Private mLaterAddr As String

' ---------------------------------------------------------------------------
' Setup
' ---------------------------------------------------------------------------

Public Sub O0ProbeInfo()
    Dim bits As String
    #If Win64 Then
        bits = "64-bit"
    #Else
        bits = "32-bit"
    #End If
    Debug.Print "O0P|info|Excel " & Application.Version & " " & bits & ", workbook " & ActiveWorkbook.Name & ", " & Format$(Now, "yyyy-mm-dd hh:nn")
    Debug.Print "O0P|info|EnableCancelKey " & Application.EnableCancelKey & "  (0 disabled, 1 interrupt, 2 error handler)"
    Debug.Print "O0P|info|CalculationInterruptKey " & Application.CalculationInterruptKey & "  (0 no key, 1 Esc, 2 any key)"
    Debug.Print "O0P|info|Calculation " & Application.Calculation & "  (-4105 automatic, -4135 manual)"
    Debug.Print "O0P|info|MultiThreadedCalculation " & Application.MultiThreadedCalculation.Enabled
End Sub

Public Sub O0ProbeReset()
    mCalls = 0
    mAsks = 0
    mForms = 0
    mLaters = 0
    mDepth = 0
    mMaxDepth = 0
    mLaterBook = ""
    mLaterSheet = ""
    mLaterAddr = ""
    Application.StatusBar = False
    Debug.Print "O0P|reset|every counter and cap cleared"
End Sub

' Builds a fresh sheet O0P<n> naming the step and what to type where. The
' formula is shown as TEXT: the owner types it, because a formula a macro
' writes is calculated inside that macro, which is a different question.
Public Sub O0ProbeSheet(ByVal stepNo As Long)
    Dim ws As Worksheet, sh As Object
    Dim title As String, typeThis As String, thenDo As String
    Select Case stepNo
    Case 1
        title = "Esc against a looping UDF, EnableCancelKey untouched"
        typeThis = "=O0Spin(20)"
        thenDo = "After about 3 seconds, press Esc once. Wait until Excel responds."
    Case 2
        title = "Esc against a looping UDF, EnableCancelKey = xlErrorHandler"
        typeThis = "=O0Spin(20,2)"
        thenDo = "After about 3 seconds, press Esc once. Wait until Excel responds."
    Case 3
        title = "Esc against a looping UDF, EnableCancelKey = xlDisabled"
        typeThis = "=O0Spin(20,0)"
        thenDo = "After about 3 seconds, press Esc once. Wait until Excel responds."
    Case 4
        title = "Esc against a looping MACRO (the command form's path)"
        typeThis = "(Immediate window)  O0SpinMacro 20   then   O0SpinMacro 20, 2   then   O0SpinMacro 20, 0"
        thenDo = "Each time, press Esc once after about 3 seconds."
    Case 5
        title = "A MsgBox from inside a UDF"
        typeThis = "=O0Ask(B1)"
        thenDo = "Answer the box. Then press F9. Then press Ctrl+Alt+F9."
    Case 6
        title = "The Function Wizard: how often it evaluates a UDF"
        typeThis = "=O0Calls(B1)"
        thenDo = "Select the cell, click fx beside the formula bar, then change the argument B1 to B12 one keystroke at a time. Cancel."
    Case 7
        title = "A modal UserForm from inside a UDF (insert an empty UserForm named O0Form first)"
        typeThis = "=O0ShowForm(B1)"
        thenDo = "Close the form with its X."
    Case 8
        title = "The status bar, set from inside a UDF"
        typeThis = "=O0Status(""O0 probe: a better roster may exist - search longer"")"
        thenDo = "Look at the status bar. Then run O0StatusClear in the Immediate window."
    Case 9
        title = "A macro scheduled from inside a UDF by Application.OnTime"
        typeThis = "=O0Later(B1)"
        thenDo = "Wait two seconds and look at the cell to the right of it."
    Case 10
        title = "DoEvents inside a UDF, and re-entrancy (SAVE FIRST - the riskiest step)"
        typeThis = "=O0Pump(10,B1)"
        thenDo = "While it runs: type 5 in B1 and press Enter; type x in D9 and press Enter; press Esc."
    Case Else
        Debug.Print "O0P: there is no probe step " & stepNo & " - the steps are 1 to 10."
        Exit Sub
    End Select
    For Each sh In ActiveWorkbook.Sheets
        If StrComp(sh.Name, "O0P" & stepNo, vbTextCompare) = 0 Then
            Debug.Print "O0P: sheet O0P" & stepNo & " already exists - delete it first, so the step runs on a fresh sheet."
            Exit Sub
        End If
    Next sh
    Set ws = ActiveWorkbook.Worksheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
    ws.Name = "O0P" & stepNo
    ws.Range("A1").Value2 = "O0 probe step " & stepNo & ": " & title
    ws.Range("B1").Value2 = 1
    ws.Range("A3").Value2 = "Type this in B5 and press Enter:"
    ws.Range("A4").Value2 = "'" & typeThis
    ws.Range("A6").Value2 = "Then: " & thenDo
    ws.Range("A7").Value2 = "The Immediate window (Ctrl+G in the VBA editor) records every call as an O0P| line."
    ws.Range("B5").Select
    Debug.Print "O0P|sheet|" & ws.Name & " ready: " & title
End Sub

' ---------------------------------------------------------------------------
' Probes 1-4: Esc
' ---------------------------------------------------------------------------

' Loops for secondsWanted (at most 30), doing arithmetic. cancelMode -1 leaves
' Application.EnableCancelKey alone; 0, 1 or 2 sets it for the loop and puts
' the old value back after. The cell says whether the loop ran to the end or
' was interrupted, when, and by which error.
Public Function O0Spin(ByVal secondsWanted As Double, Optional ByVal cancelMode As Long = -1) As Variant
    Dim t0 As Double, elapsed As Double, loops As Double, x As Double
    Dim prevKey As Long, keySet As Boolean, setNote As String
    Dim errNo As Long, errText As String
    If secondsWanted > MAX_SPIN Then secondsWanted = MAX_SPIN
    If secondsWanted < 0 Then secondsWanted = 0
    Debug.Print "O0P|spin|start|" & secondsWanted & "s|cancel mode " & cancelMode & "|" & Format$(Now, "hh:nn:ss")
    If cancelMode >= 0 Then
        prevKey = Application.EnableCancelKey
        On Error Resume Next
        Application.EnableCancelKey = cancelMode
        errNo = Err.Number
        errText = Err.Description
        On Error GoTo 0
        If errNo = 0 Then
            keySet = True
            setNote = "; EnableCancelKey read back as " & Application.EnableCancelKey
        Else
            setNote = "; EnableCancelKey could not be set: error " & errNo & " " & errText
        End If
    End If
    On Error GoTo trap
    t0 = Timer
    Do
        x = x + Sqr(loops + 1#)
        loops = loops + 1#
        elapsed = Timer - t0
        If elapsed < 0 Then elapsed = elapsed + 86400#
    Loop While elapsed < secondsWanted
    O0Spin = "ran " & Format$(elapsed, "0.0") & "s, not interrupted (" & Format$(loops, "#,##0") & " loops" & setNote & ")"
    Debug.Print "O0P|spin|end|ran to the end|" & Format$(elapsed, "0.0") & "s|" & Format$(Now, "hh:nn:ss")
    GoTo done
trap:
    errNo = Err.Number
    errText = Err.Description
    elapsed = Timer - t0
    If elapsed < 0 Then elapsed = elapsed + 86400#
    O0Spin = "interrupted at " & Format$(elapsed, "0.0") & "s (error " & errNo & ": " & errText & setNote & ")"
    Debug.Print "O0P|spin|end|interrupted|" & Format$(elapsed, "0.0") & "s|error " & errNo & "|" & Format$(Now, "hh:nn:ss")
    Resume done
done:
    If keySet Then
        On Error Resume Next
        Application.EnableCancelKey = prevKey
        On Error GoTo 0
    End If
End Function

' The same loop as a macro, run from the Immediate window: the path the
' command form ("Optimize and keep") would take.
Public Sub O0SpinMacro(Optional ByVal secondsWanted As Double = 20, Optional ByVal cancelMode As Long = -1)
    Dim t0 As Double, elapsed As Double, loops As Double, x As Double
    Dim prevKey As Long
    Dim errNo As Long, errText As String
    If secondsWanted > MAX_SPIN Then secondsWanted = MAX_SPIN
    If secondsWanted < 0 Then secondsWanted = 0
    prevKey = Application.EnableCancelKey
    If cancelMode >= 0 Then Application.EnableCancelKey = cancelMode
    Debug.Print "O0P|macro|start|" & secondsWanted & "s|EnableCancelKey " & Application.EnableCancelKey & "|" & Format$(Now, "hh:nn:ss")
    On Error GoTo trap
    t0 = Timer
    Do
        x = x + Sqr(loops + 1#)
        loops = loops + 1#
        elapsed = Timer - t0
        If elapsed < 0 Then elapsed = elapsed + 86400#
    Loop While elapsed < secondsWanted
    Debug.Print "O0P|macro|end|ran to the end|" & Format$(elapsed, "0.0") & "s, " & Format$(loops, "#,##0") & " loops"
    GoTo done
trap:
    errNo = Err.Number
    errText = Err.Description
    elapsed = Timer - t0
    If elapsed < 0 Then elapsed = elapsed + 86400#
    Debug.Print "O0P|macro|end|interrupted|" & Format$(elapsed, "0.0") & "s|error " & errNo & ": " & errText
    Resume done
done:
    On Error Resume Next
    Application.EnableCancelKey = prevKey
    On Error GoTo 0
End Sub

' ---------------------------------------------------------------------------
' Probes 5-7: modals, and the Function Wizard
' ---------------------------------------------------------------------------

Public Function O0Ask(Optional ByVal trigger As Variant) As Variant
    Dim answer As VbMsgBoxResult
    mAsks = mAsks + 1
    Debug.Print "O0P|ask|call " & mAsks & "|caller " & CallerText() & "|" & Format$(Now, "hh:nn:ss")
    If mAsks > MAX_MODALS Then
        O0Ask = "not asked: the cap of " & MAX_MODALS & " is reached (O0ProbeReset clears it)"
        Exit Function
    End If
    On Error GoTo trap
    answer = MsgBox("O0 probe: a MsgBox from inside a worksheet function." & vbCrLf & vbCrLf & "Continue?", vbYesNo + vbQuestion, "O0 probe")
    If answer = vbYes Then
        O0Ask = "answered Yes (ask " & mAsks & ")"
    Else
        O0Ask = "answered No (ask " & mAsks & ")"
    End If
    Debug.Print "O0P|ask|answered|" & O0Ask
    Exit Function
trap:
    O0Ask = "error " & Err.Number & ": " & Err.Description
    Debug.Print "O0P|ask|error|" & Err.Number & "|" & Err.Description
End Function

' Counts every evaluation, so the Function Wizard's own evaluations show up as
' O0P|calls lines - on opening it, and on each keystroke in the argument box.
Public Function O0Calls(Optional ByVal trigger As Variant) As Variant
    mCalls = mCalls + 1
    Debug.Print "O0P|calls|" & mCalls & "|caller " & CallerText() & "|" & Format$(Timer, "0.000")
    O0Calls = mCalls
End Function

' Needs an empty UserForm named O0Form in this workbook's project (Insert >
' UserForm, then set its (Name) to O0Form). Loaded by name, so this module
' compiles without it.
Public Function O0ShowForm(Optional ByVal trigger As Variant) As Variant
    Dim frm As Object
    mForms = mForms + 1
    Debug.Print "O0P|form|call " & mForms & "|caller " & CallerText() & "|" & Format$(Now, "hh:nn:ss")
    If mForms > MAX_MODALS Then
        O0ShowForm = "not shown: the cap of " & MAX_MODALS & " is reached (O0ProbeReset clears it)"
        Exit Function
    End If
    On Error GoTo trap
    Set frm = VBA.UserForms.Add("O0Form")
    frm.Caption = "O0 probe: a modal form from inside a worksheet function"
    frm.Show
    O0ShowForm = "form shown and closed (show " & mForms & ")"
    Debug.Print "O0P|form|closed"
    Exit Function
trap:
    O0ShowForm = "error " & Err.Number & ": " & Err.Description
    Debug.Print "O0P|form|error|" & Err.Number & "|" & Err.Description
End Function

' ---------------------------------------------------------------------------
' Probes 8-9: reaching past the cell without blocking
' ---------------------------------------------------------------------------

Public Function O0Status(ByVal message As String) As Variant
    On Error GoTo trap
    Application.StatusBar = message
    O0Status = "set; the status bar reads back: " & CStr(Application.StatusBar)
    Debug.Print "O0P|status|" & O0Status
    Exit Function
trap:
    O0Status = "error " & Err.Number & ": " & Err.Description
    Debug.Print "O0P|status|error|" & Err.Number & "|" & Err.Description
End Function

Public Sub O0StatusClear()
    Application.StatusBar = False
    Debug.Print "O0P|status|cleared"
End Sub

' Schedules O0LaterMacro one second from now, which writes into the cell to
' the right of this formula - a UDF causing a write, by the side door.
Public Function O0Later(Optional ByVal trigger As Variant) As Variant
    Dim callerCell As Range
    mLaters = mLaters + 1
    Debug.Print "O0P|later|call " & mLaters & "|caller " & CallerText() & "|" & Format$(Now, "hh:nn:ss")
    If mLaters > MAX_MODALS Then
        O0Later = "not scheduled: the cap of " & MAX_MODALS & " is reached (O0ProbeReset clears it)"
        Exit Function
    End If
    On Error GoTo trap
    If TypeName(Application.Caller) = "Range" Then
        Set callerCell = Application.Caller
        mLaterBook = callerCell.Worksheet.Parent.Name
        mLaterSheet = callerCell.Worksheet.Name
        mLaterAddr = callerCell.Offset(0, 1).Address
    End If
    Application.OnTime Now + TimeSerial(0, 0, 1), "'" & ThisWorkbook.Name & "'!O0LaterMacro"
    O0Later = "scheduled O0LaterMacro for one second from now (" & mLaters & ")"
    Exit Function
trap:
    O0Later = "error " & Err.Number & ": " & Err.Description
    Debug.Print "O0P|later|error|" & Err.Number & "|" & Err.Description
End Function

Public Sub O0LaterMacro()
    Debug.Print "O0P|later|the scheduled macro ran|" & Format$(Now, "hh:nn:ss")
    If Len(mLaterAddr) = 0 Then Exit Sub
    On Error GoTo trap
    Workbooks(mLaterBook).Worksheets(mLaterSheet).Range(mLaterAddr).Value2 = "written by a macro a UDF scheduled, at " & Format$(Now, "hh:nn:ss")
    Debug.Print "O0P|later|wrote " & mLaterSheet & "!" & mLaterAddr
    Exit Sub
trap:
    Debug.Print "O0P|later|could not write|" & Err.Number & "|" & Err.Description
End Sub

' ---------------------------------------------------------------------------
' Probe 10: DoEvents inside a UDF, and re-entrancy
' ---------------------------------------------------------------------------

' Calls DoEvents on every pass of a loop of secondsWanted (at most 15). Tracks
' how deep it is: an "enter|depth 2" line means Excel called this UDF again
' while the first call was still running.
Public Function O0Pump(ByVal secondsWanted As Double, Optional ByVal trigger As Variant) As Variant
    Dim t0 As Double, elapsed As Double, loops As Double
    Dim errNo As Long, errText As String
    Dim result As String
    mDepth = mDepth + 1
    If mDepth > mMaxDepth Then mMaxDepth = mDepth
    Debug.Print "O0P|pump|enter|depth " & mDepth & "|caller " & CallerText() & "|" & Format$(Now, "hh:nn:ss")
    If secondsWanted > MAX_PUMP Then secondsWanted = MAX_PUMP
    If secondsWanted < 0 Then secondsWanted = 0
    On Error GoTo trap
    t0 = Timer
    Do
        DoEvents
        loops = loops + 1#
        elapsed = Timer - t0
        If elapsed < 0 Then elapsed = elapsed + 86400#
    Loop While elapsed < secondsWanted
    result = "pumped " & Format$(elapsed, "0.0") & "s, " & Format$(loops, "#,##0") & " DoEvents, deepest call seen " & mMaxDepth
    GoTo done
trap:
    errNo = Err.Number
    errText = Err.Description
    result = "interrupted at " & Format$(Timer - t0, "0.0") & "s (error " & errNo & ": " & errText & "), deepest call seen " & mMaxDepth
    Resume done
done:
    Debug.Print "O0P|pump|exit|depth " & mDepth & "|" & result
    mDepth = mDepth - 1
    O0Pump = result
End Function

' ---------------------------------------------------------------------------

' Where a call came from: a cell address, or what Excel handed back instead
' (the Function Wizard may not say).
Private Function CallerText() As String
    Dim c As Variant
    On Error GoTo trap
    If TypeName(Application.Caller) = "Range" Then
        CallerText = Application.Caller.Worksheet.Name & "!" & Application.Caller.Address(False, False)
    Else
        c = Application.Caller
        If IsError(c) Then
            CallerText = "(an error value)"
        Else
            CallerText = "(" & TypeName(c) & ")"
        End If
    End If
    Exit Function
trap:
    CallerText = "(unknown: error " & Err.Number & ")"
End Function
