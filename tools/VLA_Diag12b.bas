Attribute VB_Name = "VLA_Diag12b"
Option Explicit

' Standalone diagnostic for DATALOG.12 - NOT part of the VLA project, and it
' calls nothing in it. No worksheet, no formula, no Frazaro code in the way:
' it times the VBA operations DATALOG performs per row, so the cost measured
' live can be attributed rather than guessed. Import it, run D12bCost from
' the Immediate window, read the output. Delete the module afterwards.
' VLA_Diag3.bas is the precedent - PROLOG.28's live pass found two costs no
' model could see, both by measuring them with no engine in the way.
'
' WHAT THIS ISOLATES. Step 1 of the live pass (2026-09-18, Excel 16.0 64-bit)
' measured a plain scan - one rule reading one Table, one comparison - at
' 0.273s over 100 rows, 2.254s over 1,000 and 25.547s over 10,000: linear, at
' about 2.5 MILLISECONDS per row. For scale, PROLOG.28 measured a unit of
' PROLOG's own work at about 1 microsecond on this machine, so a DATALOG row
' costs on the order of 2,500 of those. That gap is too large to be ordinary
' interpretation, so something the engine does PER ROW must be expensive.
'
' The suspect, from reading VLA_Datalog.bas: AtomMatches calls
' VLA_Runtime.VlaDictNew - CreateObject("Scripting.Dictionary") plus a
' CompareMode set - once for EVERY tuple it tests, on every body atom, on
' every rule pass. Step 1's program makes about two such calls per row. If
' creating that object costs about a millisecond here, it alone explains the
' measurement, and the fix is to stop creating one per row rather than to
' index anything.
'
' Sections 1 to 4 time each candidate on its own; section 5 puts them
' together in the shape one DATALOG row actually takes, and says how much of
' the 2.5 ms it accounts for. Nothing here writes to a sheet, so Excel stays
' responsive throughout: the whole run is seconds.

' Rows per section. 10,000 keeps every section well under a second even if
' each operation turns out to be slow.
Private Const REPS As Long = 10000

' What step 1 measured per row, live, on 2026-09-18: 25.547s / 10,000 rows.
Private Const MEASURED_MS_PER_ROW As Double = 2.5548

Public Sub D12bCost()
    Debug.Print "=== VLA_Diag12b: what one DATALOG row costs on this machine ==="
    Debug.Print "    Excel " & Application.Version & ", " & Format$(Now, "yyyy-mm-dd hh:nn") & ", " & REPS & " repetitions per section"
    Debug.Print "    live scan, step 1: " & Format$(MEASURED_MS_PER_ROW, "0.000") & " ms per row"
    D12bDictCreate
    D12bCollectionCreate
    D12bDictAddExists
    D12bKeyBuild
    D12bRowPipeline
    D12bTypeNameDict
    D12bTypeNameCollection
    D12bRowPipeline2
    D12bTypeOfTest
    D12bFlagTest
    Debug.Print "=== end ==="
End Sub

' 1. VlaDictNew's own shape: a late-bound Scripting.Dictionary, created and
'    dropped, with the CompareMode set the way VLA_Runtime.bas sets it. This
'    is what AtomMatches pays for every tuple it tests.
Private Sub D12bDictCreate()
    Dim i As Long
    Dim t As Double, dt As Double
    Dim d As Object
    t = Timer
    For i = 1 To REPS
        Set d = CreateObject("Scripting.Dictionary")
        d.CompareMode = 1
    Next i
    dt = Elapsed(t)
    Report "1. CreateObject(Scripting.Dictionary) + CompareMode", dt
End Sub

' 2. The same shape with a plain Collection, which needs no COM lookup - the
'    cheapest thing that could stand in for a per-row scratch dictionary.
Private Sub D12bCollectionCreate()
    Dim i As Long
    Dim t As Double, dt As Double
    Dim c As Collection
    t = Timer
    For i = 1 To REPS
        Set c = New Collection
    Next i
    dt = Elapsed(t)
    Report "2. Set c = New Collection", dt
End Sub

' 3. USING a dictionary, rather than creating one: one dictionary, one Add and
'    one Exists per row. This is the work RelTryAdd's membership index does.
Private Sub D12bDictAddExists()
    Dim i As Long
    Dim t As Double, dt As Double
    Dim d As Object
    Dim hits As Long
    Set d = CreateObject("Scripting.Dictionary")
    t = Timer
    For i = 1 To REPS
        If Not d.Exists("k" & i) Then d.Add "k" & i, True
        If d.Exists("k" & i) Then hits = hits + 1
    Next i
    dt = Elapsed(t)
    Report "3. One dictionary: Exists + Add + Exists", dt
End Sub

' 4. TupleKey's own shape: three cell values joined with Chr$(31), a CStr per
'    cell. A row is re-keyed five or six times on its way to the answer.
Private Sub D12bKeyBuild()
    Dim i As Long
    Dim t As Double, dt As Double
    Dim k As String
    Dim v1 As Variant, v2 As Variant, v3 As Variant
    t = Timer
    For i = 1 To REPS
        v1 = "B" & i
        v2 = "V" & (i Mod 50)
        v3 = CDbl((i Mod 100) * 200)
        k = CStr(v1) & Chr$(31) & CStr(v2) & Chr$(31) & CStr(v3) & Chr$(31)
    Next i
    dt = Elapsed(t)
    Report "4. Build one 3-column key string", dt
End Sub

' 5. One row's shape, put together: the two per-row dictionaries step 1's
'    program creates (one per body atom, per rule pass), plus six key builds
'    and six membership operations, which is what a row pays passing through
'    the filter, the join, the projection, the comparison, the full relation
'    and the delta. If this lands near the measured 2.5 ms, the per-row cost
'    is accounted for, and section 1's number says how much of it is the
'    CreateObject alone.
Private Sub D12bRowPipeline()
    Dim i As Long, j As Long
    Dim t As Double, dt As Double
    Dim scratch As Object, full As Object
    Dim k As String
    Dim share As Double
    Set full = CreateObject("Scripting.Dictionary")
    t = Timer
    For i = 1 To REPS
        For j = 1 To 2
            Set scratch = CreateObject("Scripting.Dictionary")
            scratch.CompareMode = 1
            scratch.Item("X") = "B" & i
        Next j
        For j = 1 To 6
            k = "B" & i & Chr$(31) & "V" & (i Mod 50) & Chr$(31) & CStr(CDbl((i Mod 100) * 200)) & Chr$(31) & CStr(j)
            If Not full.Exists(k) Then full.Add k, True
        Next j
    Next i
    dt = Elapsed(t)
    Report "5. One row's shape: 2 scratch dictionaries + 6 keys + 6 adds", dt
    share = (dt / REPS) * 1000# / MEASURED_MS_PER_ROW * 100#
    Debug.Print "   -> this accounts for " & Format$(share, "0") & "% of the " & Format$(MEASURED_MS_PER_ROW, "0.00") & " ms a row measured live."
    Debug.Print "   -> at this cost, 1,000 rows is " & Format$(dt / REPS * 1000#, "0.00") & "s and 10,000 rows is " & Format$(dt / REPS * 10000#, "0.0") & "s."
End Sub

' 6. TypeName on a late-bound Dictionary - added after the first run of this
'    module (2026-09-18) accounted for only 13% of a row. EVERY VlaDictGet,
'    VlaDictHas and VlaDictSet in VLA_Runtime.bas opens with
'    TypeName(d) = "Dictionary", and TypeName on a COM object has to ask the
'    object for its type information, which is not a cheap question. A row
'    goes through about a dozen of those calls: AtomMatches asks Has and then
'    Set per column, the comparison resolves two operands, the head resolves
'    one.
Private Sub D12bTypeNameDict()
    Dim i As Long
    Dim t As Double, dt As Double
    Dim d As Object
    Dim nm As String
    Dim hits As Long
    Set d = CreateObject("Scripting.Dictionary")
    t = Timer
    For i = 1 To REPS
        nm = TypeName(d)
        If nm = "Dictionary" Then hits = hits + 1
    Next i
    dt = Elapsed(t)
    Report "6. TypeName(d) on a Scripting.Dictionary", dt
End Sub

' 7. The same question asked of a Collection, which is VBA's own object
'    rather than a COM server: the contrast says whether TypeName itself is
'    slow, or only TypeName on a late-bound COM object.
Private Sub D12bTypeNameCollection()
    Dim i As Long
    Dim t As Double, dt As Double
    Dim c As Collection
    Dim nm As String
    Dim hits As Long
    Set c = New Collection
    t = Timer
    For i = 1 To REPS
        nm = TypeName(c)
        If nm = "Collection" Then hits = hits + 1
    Next i
    dt = Elapsed(t)
    Report "7. TypeName(c) on a Collection", dt
End Sub

' 8. One row's shape, v2: section 5 again, plus the dozen wrapper calls a row
'    really makes, each paying its own TypeName first. If this lands near the
'    measured 2.5 ms where section 5 reached an eighth of it, the wrappers are
'    the answer and the fix is to stop asking that question per call.
Private Sub D12bRowPipeline2()
    Dim i As Long, j As Long
    Dim t As Double, dt As Double
    Dim scratch As Object, full As Object
    Dim k As String, nm As String
    Dim share As Double
    Set full = CreateObject("Scripting.Dictionary")
    t = Timer
    For i = 1 To REPS
        For j = 1 To 2
            Set scratch = CreateObject("Scripting.Dictionary")
            scratch.CompareMode = 1
            scratch.Item("X") = "B" & i
        Next j
        For j = 1 To 12
            nm = TypeName(scratch)
            If nm = "Dictionary" Then
                If Not scratch.Exists("k" & j) Then scratch.Add "k" & j, j
            End If
        Next j
        For j = 1 To 6
            k = "B" & i & Chr$(31) & "V" & (i Mod 50) & Chr$(31) & CStr(CDbl((i Mod 100) * 200)) & Chr$(31) & CStr(j)
            If Not full.Exists(k) Then full.Add k, True
        Next j
    Next i
    dt = Elapsed(t)
    Report "8. One row's shape, v2: section 5 plus 12 TypeName-guarded calls", dt
    share = (dt / REPS) * 1000# / MEASURED_MS_PER_ROW * 100#
    Debug.Print "   -> this accounts for " & Format$(share, "0") & "% of the " & Format$(MEASURED_MS_PER_ROW, "0.00") & " ms a row measured live."
End Sub

' 9 and 10, added after live pass 4 (2026-09-18) confirmed the guard: what the
' two cheapest replacements for it cost, so the fix's own fork is decided on
' numbers rather than on taste. Section 9 asks the object whether it
' implements VBA's own Collection interface instead of asking COM for its
' name; section 10 reads a Boolean the code would decide once, when the
' dictionary was made.
Private Sub D12bTypeOfTest()
    Dim i As Long
    Dim t As Double, dt As Double
    Dim d As Object
    Dim hits As Long
    Set d = CreateObject("Scripting.Dictionary")
    t = Timer
    For i = 1 To REPS
        If Not (TypeOf d Is Collection) Then hits = hits + 1
    Next i
    dt = Elapsed(t)
    Report "9. TypeOf d Is Collection, asked of a Dictionary", dt
End Sub

Private Sub D12bFlagTest()
    Dim i As Long
    Dim t As Double, dt As Double
    Dim isDict As Boolean
    Dim hits As Long
    isDict = True
    t = Timer
    For i = 1 To REPS
        If isDict Then hits = hits + 1
    Next i
    dt = Elapsed(t)
    Report "10. A Boolean decided once and read per call", dt
End Sub

Private Function Elapsed(ByVal t0 As Double) As Double
    Dim dt As Double
    dt = Timer - t0
    If dt < 0 Then dt = dt + 86400#
    Elapsed = dt
End Function

Private Sub Report(ByVal what As String, ByVal dt As Double)
    Dim perItem As Double
    perItem = dt / REPS
    Debug.Print "   " & what
    Debug.Print "      " & Format$(dt, "0.000") & "s for " & REPS & "  =  " & Format$(perItem * 1000#, "0.0000") & " ms each" & _
                "   (" & Format$(perItem * 1000# / MEASURED_MS_PER_ROW * 100#, "0") & "% of a measured row)"
End Sub
