Attribute VB_Name = "VLA_Diag3"
Option Explicit

' Standalone diagnostic - NOT part of the VLA project, no dependency on
' VLA_Prolog.bas/VLA_Runtime.bas/anything else in this workbook. Import as
' its own module, run DiagRunAll (F5, or from the Immediate window), read
' the Immediate window output. Delete the module afterwards; nothing in
' Frazaro calls it.
'
' WHAT THIS ISOLATES. PROLOG.28 raised PROLOG's work budget from 120 units
' to 100,000, so a query may now hold tens of thousands of solutions and a
' list may hold tens of thousands of elements. The first live run of it
' HUNG EXCEL and crashed it, inside a test that asked for 99,999 solutions
' (owner's run, 2026-09-11). The suspected cause is not PROLOG at all: a
' VBA Collection is a linked list, so reading it by INDEX - the shape
' `For i = 1 To c.Count ... c.Item(i)` - may walk from the front every
' time, which is O(n^2) over the whole loop, where `For Each` is O(n).
' Under the old 120-unit ceiling no collection in the engine could hold
' more than about 120 items, so the difference was invisible.
'
' This module settles it by measurement rather than by reputation, with no
' Frazaro code in the way, and at the same time measures what a unit of
' PROLOG-shaped work actually costs in VBA on THIS machine - the number
' that decides whether 100,000 is the right budget or too high.
'
' Three questions, three answers in the Immediate window:
'   1. Does indexing a Collection cost more as it grows? (DiagCollectionCost)
'   2. How long does building and walking a 10,000-element cons chain take,
'      the shape a findall bag has? (DiagConsChainCost)
'   3. Roughly what does one unit of the engine's work cost - a clone of a
'      small environment plus a comparison, done 100,000 times?
'      (DiagWorkUnitCost)

Public Sub DiagRunAll()
    Debug.Print "=== VLA_Diag3: what a big Collection costs in VBA ==="
    DiagWorkUnitCost
    DiagCollectionCost
    DiagReleaseThreshold
    DiagDismantle
    Debug.Print "=== end ==="
End Sub

' 4. WHERE A CONS CHAIN BECOMES UNDROPPABLE. The run of 2026-09-11 found
'    this the hard way: building and walking a 10,000-cell chain is fast
'    (0.004s and 0.000s), and then RETURNING from the procedure raised "Out
'    of stack space" - releasing the head releases the tail, and VBA's own
'    release cascade is recursive however iterative our own walkers are.
'    This finds the length where that starts, which is the only number a
'    list budget can honestly be set from.
'
'    Each size is built inside its own procedure, which then returns: that
'    return IS the test. The error is trapped, the size reported, and the
'    run stops at the first failure - continuing after a stack overflow is
'    not something to trust.
Public Sub DiagReleaseThreshold()
    Debug.Print "4. Dropping a cons chain: where the release cascade breaks"
    Dim sizes As Variant
    sizes = Array(500, 1000, 2000, 3000, 5000, 8000, 10000)
    Dim si As Long
    For si = LBound(sizes) To UBound(sizes)
        Dim n As Long
        n = CLng(sizes(si))
        Err.Clear
        On Error Resume Next
        DiagBuildAndDrop n
        Dim code As Long, desc As String
        code = Err.Number
        desc = Err.Description
        On Error GoTo 0
        If code = 0 Then
            Debug.Print "  " & Format$(n, "@@@@@@") & "  dropped cleanly"
        Else
            Debug.Print "  " & Format$(n, "@@@@@@") & "  FAILED on release: " & code & " " & desc
            Debug.Print "  -> a list budget must sit well below " & n & "."
            Exit Sub
        End If
    Next si
    Debug.Print "  -> every size above dropped cleanly."
End Sub

Private Sub DiagBuildAndDrop(ByVal n As Long)
    Dim acc As Variant
    acc = "nil"
    Dim i As Long
    For i = n To 1 Step -1
        Dim cell As Collection
        Set cell = New Collection
        cell.Add "cons"
        cell.Add CStr(i)
        cell.Add acc
        Set acc = cell
    Next i
End Sub

' 5. CAN A CHAIN BE TAKEN APART SAFELY FIRST? Unlink each cell's tail from
'    the head down, so every cell loses its last reference while the walk
'    still holds the next one - one release at a time, no cascade. If this
'    survives a size section 4 could not drop, a long list is supportable
'    by dismantling it before it goes; if not, a budget is the only answer.
Public Sub DiagDismantle()
    Debug.Print "5. The same chain, unlinked cell by cell before it is dropped"
    Dim n As Long
    n = 10000
    Err.Clear
    On Error Resume Next
    DiagBuildDismantleAndDrop n
    Dim code As Long, desc As String
    code = Err.Number
    desc = Err.Description
    On Error GoTo 0
    If code = 0 Then
        Debug.Print "  " & n & " unlinked and dropped cleanly - dismantling works."
    Else
        Debug.Print "  " & n & " FAILED even when unlinked: " & code & " " & desc
    End If
End Sub

Private Sub DiagBuildDismantleAndDrop(ByVal n As Long)
    Dim acc As Variant
    acc = "nil"
    Dim i As Long
    For i = n To 1 Step -1
        Dim cell As Collection
        Set cell = New Collection
        cell.Add "cons"
        cell.Add CStr(i)
        cell.Add acc
        Set acc = cell
    Next i
    Dim w As Variant
    If IsObject(acc) Then Set w = acc Else w = acc
    Do While IsObject(w)
        Dim cur As Collection
        Set cur = w
        Dim nxt As Variant
        If IsObject(cur.Item(3)) Then Set nxt = cur.Item(3) Else nxt = cur.Item(3)
        cur.Remove 3
        cur.Add "nil"
        If IsObject(nxt) Then Set w = nxt Else w = nxt
    Loop
End Sub

' 1. INDEXED versus FOR EACH, at three sizes. If indexed access is linear,
'    the indexed column grows about four times when the size doubles, and
'    the For Each column about twice. If the two columns stay close, the
'    hang had another cause and PROLOG.28's loop rewrites bought nothing -
'    say so, rather than keeping them on a guess.
Public Sub DiagCollectionCost()
    Dim sizes As Variant
    sizes = Array(5000, 10000, 20000)
    Dim si As Long
    Debug.Print "1. Collection: indexed vs For Each (seconds)"
    Debug.Print "      n      indexed     For Each"
    For si = LBound(sizes) To UBound(sizes)
        Dim n As Long
        n = CLng(sizes(si))
        Dim c As Collection
        Set c = New Collection
        Dim i As Long
        For i = 1 To n
            c.Add i
        Next i
        Dim total As Double, t As Single, tIndexed As Single, tEach As Single
        total = 0
        t = Timer
        For i = 1 To n
            total = total + c.Item(i)
        Next i
        tIndexed = Timer - t
        total = 0
        Dim v As Variant
        t = Timer
        For Each v In c
            total = total + v
        Next v
        tEach = Timer - t
        Debug.Print "  " & Format$(n, "@@@@@@") & "  " & Format$(tIndexed, "0.000") & "s     " & Format$(tEach, "0.000") & "s"
    Next si
End Sub

' 2. A cons chain of 10,000, built and walked the way PROLOG's own term
'    walkers do since PROLOG.28 - a loop, never recursion.
'    NO LONGER CALLED BY DiagRunAll, and left here for the record: this is
'    the procedure whose RETURN raised "Out of stack space" on 2026-09-11.
'    Building took 0.004s and walking 0.000s - it was releasing the chain
'    that broke, which is what sections 4 and 5 now measure. Run it on its
'    own only if you want to see that again.
Public Sub DiagConsChainCost()
    Debug.Print "2. A 10,000-element cons chain"
    Dim items As Collection
    Set items = New Collection
    Dim i As Long
    For i = 1 To 10000
        items.Add CStr(i)
    Next i
    Dim t As Single
    t = Timer
    Dim acc As Variant
    acc = "nil"
    Dim arr() As Variant
    ReDim arr(1 To items.Count)
    Dim k As Long
    Dim v As Variant
    k = 0
    For Each v In items
        k = k + 1
        arr(k) = v
    Next v
    For i = UBound(arr) To 1 Step -1
        Dim cell As Collection
        Set cell = New Collection
        cell.Add "cons"
        cell.Add arr(i)
        cell.Add acc
        Set acc = cell
    Next i
    Debug.Print "  built in      " & Format$(Timer - t, "0.000") & "s"
    t = Timer
    Dim w As Variant
    Dim count As Long
    If IsObject(acc) Then Set w = acc Else w = acc
    Do While IsObject(w)
        Dim cur As Collection
        Set cur = w
        count = count + 1
        If IsObject(cur.Item(3)) Then Set w = cur.Item(3) Else w = cur.Item(3)
    Loop
    Debug.Print "  walked in     " & Format$(Timer - t, "0.000") & "s  (" & count & " cells)"
End Sub

' 3. What one unit of PROLOG's work costs. The engine clones its binding
'    environment for every candidate it tries and then compares a pair of
'    terms; this does the same thing 100,000 times, which is a whole work
'    budget. Whatever this prints is roughly the worst a single refused
'    query can cost a recalculating cell.
Public Sub DiagWorkUnitCost()
    Debug.Print "3. 100,000 work units (clone a small environment, compare a pair)"
    Dim envN As Collection, envT As Collection
    Set envN = New Collection
    Set envT = New Collection
    envN.Add "X": envT.Add "alice"
    envN.Add "Y": envT.Add "night"
    envN.Add "Z": envT.Add "42"
    Dim t As Single
    t = Timer
    Dim i As Long, hits As Long
    For i = 1 To 100000
        Dim cn As Collection, ct As Collection
        Set cn = New Collection
        Set ct = New Collection
        Dim v As Variant
        For Each v In envN
            cn.Add v
        Next v
        For Each v In envT
            ct.Add v
        Next v
        If CStr(ct.Item(1)) = "alice" Then hits = hits + 1
    Next i
    Debug.Print "  100,000 units " & Format$(Timer - t, "0.000") & "s  (" & hits & " matched)"
    Debug.Print "  -> if this is more than a few seconds, PROLOG_MAX_WORK (100000) is too high."
End Sub
