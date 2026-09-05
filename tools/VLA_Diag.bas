Attribute VB_Name = "VLA_Diag"
Option Explicit

' Standalone diagnostic - NOT part of the VLA project, no dependency
' on VLA_Interpreter.bas, VLA_Tests.bas, or anything else in this
' workbook. Import this as its own module, run DiagRunAll directly
' (F5, or from the Immediate window: DiagRunAll), and read the
' Immediate window output. Nothing here uses CheckV/Report/any of
' this project's own comparison or dispatch machinery - the goal is
' to see raw CallByName + Collection.Add behavior with zero shared
' state, zero other Subs on the call stack before this one, and zero
' possible interference from anything else running in the same host
' test sweep.
'
' From the "twenty rounds" CallByName/array corruption investigation
' (TRENCHES.md, section IX) - kept here as a worked example of the
' project's own "build a standalone, zero-dependency repro" debugging
' discipline, not because the bug is still open (it isn't; the fix
' landed in VLA_Interpreter.bas and is described in full in TRENCHES.md).
'
' Eight independent scenarios, each a fresh Collection, each printed
' immediately after two Adds. 1-6 already ran once and gave a clean
' result: scalars survive any number of hops; arrays survive zero
' hops; only "array element, 3+ hops" breaks, matching the real
' interpreter (6) exactly. 7-8 narrow further - is it about WHERE the
' array gets indexed (inline in the CallByName call vs. a local var
' first), or WHEN (extracted late, at the last hop, vs. early, right
' after the array arrives)?
'   1. CallByName, scalar literal, same scope           (baseline)
'   2. CallByName, scalar, through ONE nested Sub call
'   3. CallByName, scalar, through THREE nested Sub calls
'   4. CallByName, array element, same scope             (baseline)
'   5. CallByName, array element, through THREE nested Sub calls
'   6. This project's OWN VLA_Interpreter.VlaInterpret, called here
'      and ONLY here in this entire run - isolates whether the real
'      bug reproduces even with nothing else sharing this call stack
'      or this module's own state at all.
'   7. Array element, 3 hops, but extracted into a plain local
'      variable right before the CallByName call (last hop), instead
'      of indexed inline as part of the CallByName argument list.
'   8. Array element, extracted into a local variable at the FIRST
'      hop (immediately after arriving), then passed as an ordinary
'      scalar through the remaining hops - the earliest possible
'      extraction point instead of the latest.

Public Sub DiagRunAll()
    Debug.Print "===== VLA_Diag: standalone, zero shared state ====="

    Dim c1 As Collection
    Set c1 = New Collection
    CallByName c1, "Add", VbMethod, 300
    CallByName c1, "Add", VbMethod, 400
    Debug.Print "1. scalar literal, same scope:            c1.Item(2) = " & c1.Item(2)

    Dim c2 As Collection
    Set c2 = New Collection
    DiagAddOneHop c2, 300
    DiagAddOneHop c2, 400
    Debug.Print "2. scalar, one nested Sub:                 c2.Item(2) = " & c2.Item(2)

    Dim c3 As Collection
    Set c3 = New Collection
    DiagAddThreeHopsA c3, 300
    DiagAddThreeHopsA c3, 400
    Debug.Print "3. scalar, three nested Subs:               c3.Item(2) = " & c3.Item(2)

    Dim c4 As Collection
    Set c4 = New Collection
    Dim vals4(0 To 0) As Variant
    vals4(0) = 300
    CallByName c4, "Add", VbMethod, vals4(0)
    vals4(0) = 400
    CallByName c4, "Add", VbMethod, vals4(0)
    Debug.Print "4. array element, same scope:               c4.Item(2) = " & c4.Item(2)

    Dim c5 As Collection
    Set c5 = New Collection
    Dim vals5(0 To 0) As Variant
    vals5(0) = 300
    DiagAddThreeHopsArrayA c5, vals5
    vals5(0) = 400
    DiagAddThreeHopsArrayA c5, vals5
    Debug.Print "5. array element, three nested Subs:        c5.Item(2) = " & c5.Item(2)

    On Error Resume Next
    Dim frame As Object
    Set frame = VLA_Interpreter.VlaInterpret( _
        "(begin (dim c) (obj-set! c (new Collection)) (. c add 300) (. c add 400) (dim x) (set! x (vlaitem c 2)))")
    If Err.Number <> 0 Then
        Debug.Print "6. real VLA_Interpreter.VlaInterpret:       ERROR - " & Err.Description
    Else
        Debug.Print "6. real VLA_Interpreter.VlaInterpret:       x = " & VLA_Runtime.VlaDictGet(frame, "x")
    End If
    On Error GoTo 0

    Dim c7 As Collection
    Set c7 = New Collection
    Dim vals7(0 To 0) As Variant
    vals7(0) = 300
    DiagAddThreeHopsArrayLateExtractA c7, vals7
    vals7(0) = 400
    DiagAddThreeHopsArrayLateExtractA c7, vals7
    Debug.Print "7. array elem, 3 hops, extracted to a LOCAL var right before CallByName (not indexed inline): c7.Item(2) = " & c7.Item(2)

    Dim c8 As Collection
    Set c8 = New Collection
    Dim vals8(0 To 0) As Variant
    vals8(0) = 300
    DiagAddThreeHopsArrayEarlyExtractA c8, vals8
    vals8(0) = 400
    DiagAddThreeHopsArrayEarlyExtractA c8, vals8
    Debug.Print "8. array elem extracted to a LOCAL var at the FIRST hop, then a scalar for the rest: c8.Item(2) = " & c8.Item(2)

    Debug.Print "===== end ====="
End Sub

Private Sub DiagAddThreeHopsArrayLateExtractA(ByVal c As Collection, ByVal argVals As Variant)
    DiagAddThreeHopsArrayLateExtractB c, argVals
End Sub
Private Sub DiagAddThreeHopsArrayLateExtractB(ByVal c As Collection, ByVal argVals As Variant)
    DiagAddThreeHopsArrayLateExtractC c, argVals
End Sub
Private Sub DiagAddThreeHopsArrayLateExtractC(ByVal c As Collection, ByVal argVals As Variant)
    Dim v As Variant
    v = argVals(0)
    CallByName c, "Add", VbMethod, v
End Sub

Private Sub DiagAddThreeHopsArrayEarlyExtractA(ByVal c As Collection, ByVal argVals As Variant)
    Dim v As Variant
    v = argVals(0)
    DiagAddThreeHopsArrayEarlyExtractB c, v
End Sub
Private Sub DiagAddThreeHopsArrayEarlyExtractB(ByVal c As Collection, ByVal v As Variant)
    DiagAddThreeHopsArrayEarlyExtractC c, v
End Sub
Private Sub DiagAddThreeHopsArrayEarlyExtractC(ByVal c As Collection, ByVal v As Variant)
    CallByName c, "Add", VbMethod, v
End Sub

Private Sub DiagAddOneHop(ByVal c As Collection, ByVal v As Variant)
    CallByName c, "Add", VbMethod, v
End Sub

Private Sub DiagAddThreeHopsA(ByVal c As Collection, ByVal v As Variant)
    DiagAddThreeHopsB c, v
End Sub
Private Sub DiagAddThreeHopsB(ByVal c As Collection, ByVal v As Variant)
    DiagAddThreeHopsC c, v
End Sub
Private Sub DiagAddThreeHopsC(ByVal c As Collection, ByVal v As Variant)
    CallByName c, "Add", VbMethod, v
End Sub

Private Sub DiagAddThreeHopsArrayA(ByVal c As Collection, ByVal argVals As Variant)
    DiagAddThreeHopsArrayB c, argVals
End Sub
Private Sub DiagAddThreeHopsArrayB(ByVal c As Collection, ByVal argVals As Variant)
    DiagAddThreeHopsArrayC c, argVals
End Sub
Private Sub DiagAddThreeHopsArrayC(ByVal c As Collection, ByVal argVals As Variant)
    CallByName c, "Add", VbMethod, argVals(0)
End Sub
