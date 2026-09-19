Attribute VB_Name = "VLA_Diag13"
Option Explicit

' Diagnostic for DATALOG.13 - NOT part of the VLA project. Unlike
' VLA_Diag12b it DOES call Frazaro, deliberately: it times the real
' VLA_Runtime wrappers (VlaDictGet, VlaDictHas, VlaDictSet) rather than a
' model of them, so the one number the fix rests on is measured on the code
' itself. Import it, run D13Wrappers from the Immediate window, read the
' output. Delete the module afterwards.
'
' WHAT IT IS FOR. DATALOG.12's pass 3 timed TypeName on a Dictionary at
' 0.148 ms (VLA_Diag12b section 6) and a MODEL of a row's dozen wrapper
' calls at 85% of a row. Neither timed the wrappers themselves. Run this
' once BEFORE the DATALOG.13 reload, on the old VLA_Runtime, and once
' AFTER: the Dictionary lines should fall from about 0.15 ms a call to a
' few thousandths, and the Collection lines, which never paid the guard,
' should not move. The version line says which runtime was measured.
'
' Nothing here writes to a sheet; the whole run is a few seconds.

' Calls per section. 10,000 keeps a 0.15 ms call at 1.5s a line.
Private Const D13_REPS As Long = 10000

Public Sub D13Wrappers()
    Dim d As Object, c As Object
    Debug.Print "=== VLA_Diag13: the VlaDict wrappers, timed on the real code ==="
    Debug.Print "    Excel " & Application.Version & ", " & Format$(Now, "yyyy-mm-dd hh:nn") & ", " & D13_REPS & " calls per line"
    Debug.Print "    VLA_RUNTIME_VERSION " & VLA_Runtime.VLA_RUNTIME_VERSION & _
                "  (PF4B.0 = before the fix, DATALOG13.0 = after)"
    Set d = VLA_Runtime.VlaDictNew()
    If TypeOf d Is Collection Then
        Debug.Print "    this host has no Scripting runtime - the Dictionary lines cannot be measured here"
    Else
        D13Time d, "Dictionary"
    End If
    Set c = New Collection
    D13Time c, "Collection"
    D13Controls
    Debug.Print "=== end ==="
End Sub

' Get, Has and Set on one representation, a single key, the shape a row's
' per-column calls take.
Private Sub D13Time(ByVal target As Object, ByVal kind As String)
    Dim i As Long
    Dim t As Double
    Dim v As Variant
    Dim found As Boolean
    VLA_Runtime.VlaDictSet target, "k-1", 1
    t = Timer
    For i = 1 To D13_REPS
        v = VLA_Runtime.VlaDictGet(target, "k-1")
    Next i
    D13Report "VlaDictGet on a " & kind, D13Elapsed(t)
    t = Timer
    For i = 1 To D13_REPS
        found = VLA_Runtime.VlaDictHas(target, "k-1")
    Next i
    D13Report "VlaDictHas on a " & kind, D13Elapsed(t)
    t = Timer
    For i = 1 To D13_REPS
        VLA_Runtime.VlaDictSet target, "k-1", i
    Next i
    D13Report "VlaDictSet on a " & kind & " (replacing one key)", D13Elapsed(t)
End Sub

' The two questions themselves, asked of a Dictionary with no wrapper
' around them, repeating VLA_Diag12b sections 6 and 9 as this run's
' control: if these move, the machine moved, not the code.
Private Sub D13Controls()
    Dim i As Long
    Dim t As Double
    Dim d As Object
    Dim nm As String
    Dim hits As Long
    On Error Resume Next
    Set d = CreateObject("Scripting.Dictionary")
    On Error GoTo 0
    If d Is Nothing Then Exit Sub
    t = Timer
    For i = 1 To D13_REPS
        nm = TypeName(d)
        If nm = "Dictionary" Then hits = hits + 1
    Next i
    D13Report "control: TypeName(d) on a Dictionary, no wrapper", D13Elapsed(t)
    t = Timer
    For i = 1 To D13_REPS
        If Not (TypeOf d Is Collection) Then hits = hits + 1
    Next i
    D13Report "control: TypeOf d Is Collection on a Dictionary, no wrapper", D13Elapsed(t)
End Sub

Private Function D13Elapsed(ByVal t0 As Double) As Double
    Dim dt As Double
    dt = Timer - t0
    If dt < 0 Then dt = dt + 86400#
    D13Elapsed = dt
End Function

Private Sub D13Report(ByVal what As String, ByVal dt As Double)
    Debug.Print "   " & what
    Debug.Print "      " & Format$(dt, "0.000") & "s for " & D13_REPS & "  =  " & _
                Format$(dt / D13_REPS * 1000#, "0.0000") & " ms each"
End Sub
