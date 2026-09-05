VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VLA_EventSink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' IN.7: the add-in-resident event sink - the mechanism that makes
' "When the sheet changes:" run without injecting anything into the
' user's workbook (L-TIER2's own class-injection surface, the thing
' IN.7 exists to avoid). One instance lives for the add-in's whole
' session, owned by ThisWorkbook.Workbook_Open/BeforeClose
' (synthesized by VLA_Build.bas's VlaBuildAddin, the same hook point
' that already manages the ribbon menu) - WithEvents against the
' Application object itself, not any one workbook, so it keeps
' hearing SheetChange for whichever workbook VLA_Events has a handler
' registered against, across however many workbooks the user opens in
' the same Excel session. This is the first WithEvents usage anywhere
' in this codebase; VLA_Events.bas owns everything event-specific
' (what "sheet-change" means, the registry, the re-entrancy guard) so
' this class stays a thin, mechanical relay.
Private WithEvents mApp As Application

Private Sub Class_Initialize()
    Set mApp = Application
End Sub

Private Sub Class_Terminate()
    Set mApp = Nothing
End Sub

' Sh is a Worksheet for every real edit; SheetChange also fires for
' chart sheets, whose object model does not guarantee .Parent behaves
' the same way - trapped rather than assumed, since a bad guess here
' would surface as an Excel-level error, not a VLA one, the first time
' anyone edited a chart sheet in the same session as a registered
' handler.
Private Sub mApp_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    On Error Resume Next
    Dim wb As Workbook
    Set wb = Sh.Parent
    On Error GoTo 0
    If Not wb Is Nothing Then VLA_Events.VlaDispatchSheetChange wb
End Sub

Private Sub mApp_WorkbookBeforeClose(ByVal Wb As Workbook, Cancel As Boolean)
    VLA_Events.VlaUnregisterHandlers Wb
End Sub
