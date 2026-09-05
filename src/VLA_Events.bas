Attribute VB_Name = "VLA_Events"
Option Explicit
Public Const VLA_EVENTS_VERSION As String = "IN.7"
' IN.7 (button-click half): a second registry (mBtnWb/mBtnName/mBtnSrc/
' mBtnProc) added, keyed by (workbook, button name) - VlaRegisterButtonClickHandler/
' VlaDispatchButtonClick/VlaButtonClickDispatch/VlaHasButtonClickHandler,
' below. VlaUnregisterHandlers now clears both tables for a closing
' workbook. No new WithEvents sink: a Form Control button's .OnAction
' already IS "Excel calls a macro by name" - VlaButtonClickDispatch is
' that macro, reached the same way EnglishIdeCheck/EnglishIdeInterpret/
' etc. already are from a sheet button (VLA_IDE.bas's BuildWorkspace).
' =====================================================================
'  VLA_Events - the durable event -> source registry. IN.7's whole
'  point is running "When the sheet changes:" without injecting a
'  class into the user's workbook (L-TIER2's own injection surface);
'  the registry is what makes that possible - a plain add-in-resident
'  table, written when a program containing an on:sheet-change handler
'  is interpreted (VLA_IDE.bas's InterpretProgram), read by the add-
'  in's own WithEvents sink (VLA_EventSink.cls) when Excel actually
'  fires the event, later, in a call chain that has nothing else to do
'  with the run that registered it.
'
'  Deliberately the only module that knows "sheet-change" is a concept:
'  VLA_Interpreter stays event-agnostic (VlaInterpretEntry just runs a
'  named procedure by name; VlaHasProc just answers whether one was
'  declared) so a second event kind, later, does not have to touch the
'  interpreter's own dispatch at all.
'
'  Keyed by object identity (Is), not name - two different open
'  workbooks can share a display name only if one is a copy in a
'  different window, which Excel itself refuses, so identity is exact
'  and cheap, the same "Collection as a parallel-array table" idiom
'  VLA_English.bas's mActNames/mActParams/mActReq already use.
'
'  LAYER:     add-in-resident only - never exported or injected into a
'             user workbook (this table would be meaningless there).
'  MAY CALL:  VLA_Interpreter (VlaInterpretEntry).
'  SHIPS:     yes - VLA_Build.bas's mods array, alongside
'             VLA_EventSink.cls, the class that actually reads it.
' =====================================================================

Private mWb As Collection      ' registered hostWb Workbook objects
Private mSrc As Collection     ' parallel: the vla source text for each

' IN.7 (button-click half): a second table, same shape, keyed by
' (workbook, button name) instead of workbook alone - a workbook may
' register several buttons, each with its own source and its own
' entry proc within that source (VLA_English.bas's "on:click:N",
' handed back through EnglishClickHandlerNames/Procs). No WithEvents
' sink needed this time (unlike sheet-change's Application.SheetChange,
' a Form Control button's .OnAction is Excel calling a macro BY NAME,
' the interpreter-native seam VlaInterpretEntry already exists for) -
' VlaButtonClickDispatch below is that macro, and it is the only
' fixed OnAction string every button this project ever draws points
' at (VLA_Interpreter.bas's MakeVlaButton), Application.Caller telling
' it which one was actually clicked.
Private mBtnWb As Collection    ' registered hostWb Workbook objects, one per button
Private mBtnName As Collection  ' parallel: the button's name (== its caption, IN.7's design)
Private mBtnSrc As Collection   ' parallel: the vla source text for that program
Private mBtnProc As Collection  ' parallel: the entry proc name within that source

Private Sub EnsureInit()
    If mWb Is Nothing Then Set mWb = New Collection
    If mSrc Is Nothing Then Set mSrc = New Collection
End Sub

Private Sub EnsureBtnInit()
    If mBtnWb Is Nothing Then Set mBtnWb = New Collection
    If mBtnName Is Nothing Then Set mBtnName = New Collection
    If mBtnSrc Is Nothing Then Set mBtnSrc = New Collection
    If mBtnProc Is Nothing Then Set mBtnProc = New Collection
End Sub

' Called after a real interpret run whose program declared
' "on:sheet-change" (VLA_IDE.bas checks VLA_Interpreter.VlaHasProc
' right after its own VlaInterpret call). Re-running the same program
' replaces its earlier registration outright - the running program's
' CURRENT source is what should fire from now on, not whatever an
' earlier run left behind.
Public Sub VlaRegisterSheetChangeHandler(ByVal hostWb As Workbook, ByVal vlaSource As String)
    If hostWb Is Nothing Then Exit Sub
    EnsureInit
    VlaUnregisterHandlers hostWb
    mWb.Add hostWb
    mSrc.Add vlaSource
End Sub

' Called from VLA_EventSink's own WorkbookBeforeClose handler - once a
' registered workbook closes, its Workbook reference is no longer
' something a later event could legitimately match, and leaving it in
' the table would only grow it, unbounded, over a long Excel session.
' IN.7 (button-click half): also forgets every button handler this
' workbook registered - the same "closed means gone" reasoning, now
' covering both tables so a stray Workbook reference in either can
' never outlive the workbook it points to.
Public Sub VlaUnregisterHandlers(ByVal hostWb As Workbook)
    If Not mWb Is Nothing Then
        Dim i As Long
        For i = mWb.Count To 1 Step -1
            If mWb.Item(i) Is hostWb Then
                mWb.Remove i
                mSrc.Remove i
            End If
        Next
    End If
    If Not mBtnWb Is Nothing Then
        Dim j As Long
        For j = mBtnWb.Count To 1 Step -1
            If mBtnWb.Item(j) Is hostWb Then
                mBtnWb.Remove j
                mBtnName.Remove j
                mBtnSrc.Remove j
                mBtnProc.Remove j
            End If
        Next
    End If
End Sub

' Called from VLA_EventSink's Application_SheetChange handler. Must
' never let an error escape into Excel's own event dispatch - an
' unhandled error inside an event sink surfaces to the user as Excel
' itself misbehaving, not as a VLA error, so a broken handler is
' trapped and reported to the Immediate window instead of raised.
' EnableEvents is dropped for the call's duration and restored after:
' without it, a handler that writes so much as one cell re-fires
' SheetChange on its own write, and a handler runs itself into the
' ground the first time anyone gives it a realistic body (recompute a
' total when the sheet changes IS a write, by construction) - saved
' and restored rather than just set True, in case something else
' already had events off for its own reason.
Public Sub VlaDispatchSheetChange(ByVal changedWb As Workbook)
    If mWb Is Nothing Then Exit Sub
    Dim i As Long
    For i = 1 To mWb.Count
        If mWb.Item(i) Is changedWb Then
            Dim wasEnabled As Boolean
            wasEnabled = Application.EnableEvents
            Application.EnableEvents = False
            On Error Resume Next
            VLA_Interpreter.VlaInterpretEntry CStr(mSrc.Item(i)), "on:sheet-change", changedWb
            If Err.Number <> 0 Then
                Debug.Print "VLA-Events: on:sheet-change handler failed - " & Err.Description
                Err.Clear
            End If
            On Error GoTo 0
            Application.EnableEvents = wasEnabled
            Exit Sub
        End If
    Next
End Sub

' Dev-rig/test visibility only - lets a test assert a registration
' happened (or was cleared) without reaching into this module's
' Private state.
Public Function VlaHasSheetChangeHandler(ByVal hostWb As Workbook) As Boolean
    If mWb Is Nothing Then Exit Function
    Dim i As Long
    For i = 1 To mWb.Count
        If mWb.Item(i) Is hostWb Then
            VlaHasSheetChangeHandler = True
            Exit Function
        End If
    Next
End Function

' =====================================================================
'  IN.7 (button-click half): the button registry - VLA_IDE.bas's
'  InterpretProgram calls this once per 'When "<caption>" is
'  clicked:' handler a translation declared (EnglishClickHandlerNames/
'  Procs name them), right after the interpret run that declared it.
'  Re-running the same program replaces just that ONE caption's
'  registration - VlaRegisterSheetChangeHandler's "current run wins"
'  precedent, scoped per-button instead of per-workbook because one
'  program's re-run must not clear buttons a DIFFERENT program in the
'  same workbook already registered.
' =====================================================================
Public Sub VlaRegisterButtonClickHandler(ByVal hostWb As Workbook, ByVal buttonName As String, _
                                          ByVal vlaSource As String, ByVal entryProc As String)
    If hostWb Is Nothing Then Exit Sub
    EnsureBtnInit
    RemoveButtonHandler hostWb, buttonName
    mBtnWb.Add hostWb
    mBtnName.Add buttonName
    mBtnSrc.Add vlaSource
    mBtnProc.Add entryProc
End Sub

Private Sub RemoveButtonHandler(ByVal hostWb As Workbook, ByVal buttonName As String)
    If mBtnWb Is Nothing Then Exit Sub
    Dim i As Long
    For i = mBtnWb.Count To 1 Step -1
        If mBtnWb.Item(i) Is hostWb And VLA_Identity.Fold(CStr(mBtnName.Item(i))) = VLA_Identity.Fold(buttonName) Then
            mBtnWb.Remove i
            mBtnName.Remove i
            mBtnSrc.Remove i
            mBtnProc.Remove i
        End If
    Next
End Sub

' The pure/testable half: given a workbook and a button name, run
' its registered handler if one is registered, else do nothing - a
' click on an unregistered button (never wired, or from a workbook
' that closed) is a provable no-op, not an error. Same re-entrancy
' guard as VlaDispatchSheetChange, same reason: a handler body that
' writes a cell must not re-fire itself, and a broken handler must
' report to the Immediate window rather than surface as Excel itself
' misbehaving.
Public Sub VlaDispatchButtonClick(ByVal hostWb As Workbook, ByVal buttonName As String)
    If mBtnWb Is Nothing Then Exit Sub
    Dim i As Long
    For i = 1 To mBtnWb.Count
        If mBtnWb.Item(i) Is hostWb And VLA_Identity.Fold(CStr(mBtnName.Item(i))) = VLA_Identity.Fold(buttonName) Then
            Dim wasEnabled As Boolean
            wasEnabled = Application.EnableEvents
            Application.EnableEvents = False
            On Error Resume Next
            VLA_Interpreter.VlaInterpretEntry CStr(mBtnSrc.Item(i)), CStr(mBtnProc.Item(i)), hostWb
            If Err.Number <> 0 Then
                Debug.Print "VLA-Events: button '" & buttonName & "' handler failed - " & Err.Description
                Err.Clear
            End If
            On Error GoTo 0
            Application.EnableEvents = wasEnabled
            Exit Sub
        End If
    Next
End Sub

' The thin macro every button this project draws points its OnAction
' at (VLA_Interpreter.bas's MakeVlaButton - always the ADD-IN's own
' workbook name, never mHostWorkbook's, the IN.6/IN.7-sheet-change
' class of bug this module's own header warns about). Application.Caller
' is the clicked shape's Name, which IN.7's design makes identical to
' its caption; ActiveWorkbook is reliable here because a Form Control
' button can only be clicked while its own workbook's window is
' frontmost - the same trust VlaTimeIt's own "aim by activating"
' discipline already places in ActiveWorkbook elsewhere in this
' project. No WithEvents, no live event pump needed to exercise this -
' VlaDispatchButtonClick above is directly callable from a test with
' an explicit workbook and name; this macro only supplies those two
' from the real click.
Public Sub VlaButtonClickDispatch()
    On Error Resume Next
    Dim nm As String
    nm = CStr(Application.Caller)
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    On Error GoTo 0
    If Len(nm) = 0 Or wb Is Nothing Then Exit Sub
    VlaDispatchButtonClick wb, nm
End Sub

' Dev-rig/test visibility only - VlaHasSheetChangeHandler's own twin.
Public Function VlaHasButtonClickHandler(ByVal hostWb As Workbook, ByVal buttonName As String) As Boolean
    If mBtnWb Is Nothing Then Exit Function
    Dim i As Long
    For i = 1 To mBtnWb.Count
        If mBtnWb.Item(i) Is hostWb And VLA_Identity.Fold(CStr(mBtnName.Item(i))) = VLA_Identity.Fold(buttonName) Then
            VlaHasButtonClickHandler = True
            Exit Function
        End If
    Next
End Function
