Attribute VB_Name = "VLA_Diag2"
Option Explicit

' Standalone diagnostic - NOT part of the VLA project, no dependency on
' VLA_Interpreter.bas/VLA_Runtime.bas/anything else in this workbook.
' Import as its own module, run DiagRunAll (F5, or from the Immediate
' window), read the Immediate window output.
'
' What this isolates: IN.12's fix added a new VLA_Runtime.bas Sub that
' raises Err.Raise 5 on a business-rule condition (a sheet already
' exists). The interpreter reaches VLA_Runtime helpers through
' TryRuntimeHelper (VLA_Interpreter.bas), which calls them via
' Application.Run(target, ...) wrapped in its OWN local "On Error
' Resume Next / On Error GoTo 0", and THAT is called from
' ExecStmtTrapped, which has ITS OWN outer "On Error Resume Next" for
' Try:'s own error-catching emulation. Two nested On-Error-Resume-Next
' scopes, one hop apart, with Application.Run in the middle - a shape
' that has never actually carried a real Err.Raise before now, because
' no VLA_Runtime validation helper had ever fired while wrapped in a
' live Try: until VlaCheckSheetAbsent (IN.12).
'
' The observed symptom: a hard, unhandled "Run-time error '5'" break
' showing the INNER Sub's exact original message - which should be
' impossible if EITHER Resume Next were actually catching it (each
' would either swallow it silently or produce a DIFFERENT, generic
' message, never the original text verbatim escaping uncaught). This
' reproduces that exact nesting shape with zero other project state on
' the call stack, to find out which layer - if any - is not doing what
' the code around it assumes.
'
'   1. Direct baseline: On Error Resume Next, Application.Run a Sub
'      that raises. Does the caller's Err.Number pick it up at all?
'   2. Same, but the target Sub does its own "On Error Resume Next /
'      On Error GoTo 0" before raising - VlaCheckSheetAbsent's exact
'      shape (checks something, clears its own handler, THEN raises).
'   3. The real nesting: outer On Error Resume Next (ExecStmtTrapped)
'      wraps a call to a middle Sub with NO handler of its own
'      (EvalDynamicHead), which calls an inner Sub with ITS OWN On
'      Error Resume Next around Application.Run (TryRuntimeHelper).
'   4. Same as 3, but the middle Sub is a Function returning a Variant
'      (EvalDynamicHead's real shape - AssignVar-style return), not a
'      plain Sub - in case a Function boundary changes anything.

Public Sub DiagRunAll()
    Debug.Print "===== VLA_Diag2: Application.Run + nested On Error ====="

    ' 1. Direct: does a raise inside Application.Run's target even
    ' respect a plain, single-level On Error Resume Next in the caller?
    Dim gotErr1 As Boolean, num1 As Long, desc1 As String, src1 As String
    On Error Resume Next
    Err.Clear
    Application.Run "VLA_Diag2.DiagRaiser"
    If Err.Number <> 0 Then
        gotErr1 = True: num1 = Err.Number: desc1 = Err.Description: src1 = Err.Source
    End If
    On Error GoTo 0
    If gotErr1 Then
        Debug.Print "1. direct Application.Run + On Error Resume Next: CAUGHT - num=" & num1 & " src=" & src1 & " desc=" & desc1
    Else
        Debug.Print "1. direct Application.Run + On Error Resume Next: NOT CAUGHT (Err.Number was 0 after return - either swallowed silently or never trapped)"
    End If

    ' 2. Target Sub does its own On Error Resume Next / On Error GoTo 0
    ' before raising - DiagRaiserTwoStep's exact shape, matching
    ' VlaCheckSheetAbsent (checks via O.E.R.N., clears with GoTo 0,
    ' THEN raises with no handler active).
    Dim gotErr2 As Boolean, num2 As Long, desc2 As String, src2 As String
    On Error Resume Next
    Err.Clear
    Application.Run "VLA_Diag2.DiagRaiserTwoStep"
    If Err.Number <> 0 Then
        gotErr2 = True: num2 = Err.Number: desc2 = Err.Description: src2 = Err.Source
    End If
    On Error GoTo 0
    If gotErr2 Then
        Debug.Print "2. Application.Run -> target has its own O.E.R.N./GoTo 0 before raising: CAUGHT - num=" & num2 & " src=" & src2 & " desc=" & desc2
    Else
        Debug.Print "2. Application.Run -> target has its own O.E.R.N./GoTo 0 before raising: NOT CAUGHT"
    End If

    ' 3. The real nesting, Sub-shaped middle layer.
    Debug.Print "3. real nesting (outer O.E.R.N. -> unprotected middle Sub -> inner O.E.R.N. + Application.Run):"
    DiagOuterCaller

    ' 4. The real nesting, Function-shaped middle layer (EvalDynamicHead
    ' is a Function, not a Sub - test whether that matters).
    Debug.Print "4. real nesting, middle layer is a Function returning Variant:"
    DiagOuterCallerFn

    Debug.Print "===== end ====="
End Sub

' Scenario 1/2 target: raises immediately, no handler of its own.
Public Sub DiagRaiser()
    Err.Raise 5, "VLA-Diag", "diag: scenario 1 raise"
End Sub

' Scenario 2 target: VlaCheckSheetAbsent's exact shape.
Public Sub DiagRaiserTwoStep()
    Dim ws As Object
    On Error Resume Next
    Set ws = Nothing ' stand-in for the ActiveWorkbook.Worksheets(nm) probe
    On Error GoTo 0
    Err.Raise 5, "VLA-Diag", "diag: scenario 2 raise, after its own On Error GoTo 0"
End Sub

' Scenario 3: outer On Error Resume Next (ExecStmtTrapped's shape).
Private Sub DiagOuterCaller()
    On Error Resume Next
    Err.Clear
    DiagMiddleUnprotected
    If Err.Number <> 0 Then
        Debug.Print "   CAUGHT at outer layer - num=" & Err.Number & " src=" & Err.Source & " desc=" & Err.Description
    Else
        Debug.Print "   NOT CAUGHT at outer layer (Err.Number was 0)"
    End If
    On Error GoTo 0
End Sub

' Middle layer: no error handling of its own (EvalDynamicHead's shape).
Private Sub DiagMiddleUnprotected()
    DiagInnerRuntimeHelper
End Sub

' Inner layer: TryRuntimeHelper's exact shape - its own On Error Resume
' Next around Application.Run, checks Err.Number, On Error GoTo 0.
Private Sub DiagInnerRuntimeHelper()
    On Error Resume Next
    Err.Clear
    Application.Run "VLA_Diag2.DiagRaiserTwoStep"
    Dim innerCaught As Boolean
    innerCaught = (Err.Number <> 0)
    On Error GoTo 0
    Debug.Print "   inner layer's own check: innerCaught=" & innerCaught
    ' Deliberately does NOT re-raise or clear beyond On Error GoTo 0 -
    ' matching TryRuntimeHelper exactly, to see what the outer layer
    ' actually observes in Err after this Sub returns normally.
End Sub

' Scenario 4: same nesting, middle layer is a Function (EvalDynamicHead
' really is "Private Function EvalDynamicHead(...) As Variant").
Private Sub DiagOuterCallerFn()
    On Error Resume Next
    Err.Clear
    Dim r As Variant
    r = DiagMiddleFn()
    If Err.Number <> 0 Then
        Debug.Print "   CAUGHT at outer layer - num=" & Err.Number & " src=" & Err.Source & " desc=" & Err.Description
    Else
        Debug.Print "   NOT CAUGHT at outer layer (Err.Number was 0)"
    End If
    On Error GoTo 0
End Sub

Private Function DiagMiddleFn() As Variant
    DiagInnerRuntimeHelper
    DiagMiddleFn = Empty
End Function
