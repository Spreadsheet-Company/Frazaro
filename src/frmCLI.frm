VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCLI 
   Caption         =   "Frazaro CLI"
   ClientHeight    =   8440
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   9380.001
   OleObjectBlob   =   "frmCLI.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmCLI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' CLI.2 (owner-requested): resizable window. MSForms UserForm has no
' built-in sizable border at all (BorderStyle is only none/fixed-
' single) - the standard workaround is flipping WS_THICKFRAME onto the
' form's own window style via User32, after which Windows' own window
' manager handles the drag-resize natively. Deliberately NOT the
' AddressOf/subclassing pattern flagged as the real security concern
' when this whole feature started (BETA_ROADMAP.md's IN.13) - this is
' one style bit set on a window this form already owns, no custom
' WndProc, no callback pointer, nothing for AMSI/Defender heuristics to
' notice. SetWindowPos with SWP_FRAMECHANGED right after is required,
' not optional - Windows does not repaint the new frame style on its
' own until told to.
#If VBA7 Then
    Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hwnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Private Declare PtrSafe Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Private mHwnd As LongPtr
#Else
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Private Declare Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private mHwnd As Long
#End If

Private Const GWL_STYLE As Long = -16
Private Const WS_THICKFRAME As Long = &H40000
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOZORDER As Long = &H4
Private Const SWP_FRAMECHANGED As Long = &H20

' Layout constants captured from the original build
' (tools/build_cli_form.ps1) so UserForm_Resize computes every
' control's new position from the MARGINS and the bottom-reserved band,
' not from hardcoded absolutes.
Private Const CLI_MARGIN As Long = 12
Private Const CLI_BOTTOM_BAND As Long = 68        ' ClientHeight(272) - lblStatus.Top(204)
Private Const CLI_STATUS_TO_BUTTONS As Long = 22  ' cmdRun.Top(226) - lblStatus.Top(204)
Private Const CLI_BUTTON_GAP As Long = 6          ' cmdCancel.Left(346) - (cmdRun.Left+cmdRun.Width)

Private Sub UserForm_Initialize()
    cmdCancel.Cancel = True
    lblStatus.Caption = ""
    ' CLI.1: no font was ever set at build time - the original build
    ' script's own attempt to set txtCommand.Font.Name/.Size via
    ' PowerShell COM automation failed (a chained-property-access quirk
    ' specific to that automation path, not a VBA limitation) and was
    ' dropped, leaving MSForms' plain default font in place ever since.
    ' Font, like Caption, is a complex object property living in the
    ' binary .frx, not the plain-text .frm - set here for the same
    ' reason lblHint's Caption is: hand-editable without re-running
    ' tools/build_cli_form.ps1. Consolas, not the proportional default -
    ' owner-requested (paren-matching by eye is what a monospace font
    ' is FOR). Widely available since Vista/Office 2007, no fallback
    ' needed for this project's own baseline (customUI14 already
    ' assumes Office 2010+).
    txtCommand.Font.Name = "Consolas"
    txtCommand.Font.Size = 11
    ' CLI.1 (owner-caught live): raw VLA has no # comment - only the
    ' English/workspace-sheet dialect does. # has no special meaning to
    ' VLA.Tokenize at all (checked directly: only ";" is a recognized
    ' comment case), so a "#..." row reads as an ordinary bare symbol
    ' and gets evaluated as a variable reference - "nothing stored at
    ' key '#...'" is the normal unbound-variable error, not a hash-
    ' lookup feature. Deliberately not teaching the raw reader # too:
    ' # is genuinely unclaimed VLA syntax today, which is exactly the
    ' character a future reader macro (#t/#f, #x hex literals) would
    ' want - spending it on a second comment spelling forecloses that
    ' for a convenience this hint line covers for free. Set here, not
    ' baked into the .frm as a designer property, because a Label's
    ' Caption lives in the binary .frx, not the plain-text part of
    ' .frm - this is the hand-editable way to change it without
    ' re-running tools/build_cli_form.ps1.
    lblHint.Caption = "VLA or English - VLA comments use ; (not #) - Ctrl+Enter runs, Esc closes."
    ' CLI.1: boot-time example, not a placeholder - a real, runnable
    ' program showing both halves of the pipeline at once: a multi-line
    ' raw VLA form (L0.2's own payoff) defining a tail-recursive
    ' function, and msgbox as the exit point for actually SEEING a
    ' macro's result while testing one, the thing a new user has no
    ' other way to discover from the CLI alone. fib-iter mirrors L18's
    ' own proven test shape exactly (TestL18's "fact-iter", VLA_Tests_
    ' Grammar.bas) - accumulator-style, every parameter (byval), the
    ' recursive call in tail position via (return (fib-iter ...)).
    ' Honest caveat, not shown in the text itself but worth knowing:
    ' L18's actual rebind-and-jump transform is EmitProc's own trick -
    ' Compile only. Interpret (what the CLI always runs through) has no
    ' such optimization, so this genuinely recurses via CallUserProc
    ' rather than looping in the metal; correct either way for n=10,
    ' just not the O(1)-stack claim L18 makes under Compile.
    txtCommand.Text = _
        "(function" & vbCrLf & _
        "  fib-iter ((n) (byval a) (byval b Long)) Long " & vbCrLf & _
        "  " & Chr$(34) & "tail-recursive fibonacci, accumulator style" & Chr$(34) & vbCrLf & _
        "  (if (<= n 0)" & vbCrLf & _
        "    (then (return a))" & vbCrLf & _
        "    (else (return (fib-iter (- n 1) b (+ a b))))))" & vbCrLf & _
        vbCrLf & _
        "(defmacro" & vbCrLf & _
        "  (modal message)" & vbCrLf & _
        "  " & Chr$(34) & "alias for msgbox" & Chr$(34) & vbCrLf & _
        "  (msgbox message))" & vbCrLf & _
        vbCrLf & _
        "(modal (fib-iter 10 0 1))"
End Sub

Private Sub UserForm_Activate()
    txtCommand.SetFocus
    ' CLI.2: the window handle only reliably exists once the form is
    ' actually shown, not at Initialize - hence doing this here, not
    ' there. Best-effort: a failed resize-enable must never block
    ' opening the CLI itself. Me.hWnd is NOT a real member on UserForm
    ' in this Excel/VBA version (owner-caught live: "Method or data
    ' member not found") - FindWindow by class + caption instead, the
    ' same technique the very first hotkey snippet in this feature's
    ' own history used. ThunderDFrame is the standard UserForm window
    ' class; ThunderXFrame is the (rarer) alternate some forms get, so
    ' both are tried rather than assuming which one this form got.
    On Error Resume Next
    mHwnd = FindWindowA("ThunderDFrame", Me.Caption)
    If mHwnd = 0 Then mHwnd = FindWindowA("ThunderXFrame", Me.Caption)
    SetWindowLong mHwnd, GWL_STYLE, GetWindowLong(mHwnd, GWL_STYLE) Or WS_THICKFRAME
    SetWindowPos mHwnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or SWP_FRAMECHANGED
    ' Owner-caught live: the dialog renders oversized on first open,
    ' snapping to the correct scale the moment the user drags the
    ' border. MSForms renders through an internal layer that computed
    ' its scale once at initial paint and does not recompute just
    ' because the frame style changed out from under it afterward - it
    ' takes a REAL size change (WM_SIZE) to force that. Nudging Width by
    ' one point and back does that programmatically, so the correct
    ' scale shows up immediately instead of waiting on the user to touch
    ' the border themselves.
    Me.Width = Me.Width + 1
    Me.Width = Me.Width - 1
    On Error GoTo 0
End Sub

' CLI.2: MSForms controls have no built-in anchoring/docking at all -
' flipping WS_THICKFRAME alone would just reveal more empty dialog
' background on drag, with txtCommand and the buttons frozen at their
' original size and position. This is what actually makes the resize
' useful: txtCommand grows to fill the new space, the status line and
' both buttons stay pinned to the bottom-right corner. Every position is
' computed from the CLI_* margin/band constants captured from the
' original build, not hardcoded, so this generalizes to any new size.
Private Sub UserForm_Resize()
    On Error Resume Next   ' best-effort: a layout glitch must never crash the CLI mid-resize
    Dim w As Single, h As Single
    w = Me.InsideWidth
    h = Me.InsideHeight
    If w < 200 Or h < 150 Then Exit Sub   ' too small to lay out sanely - leave controls where they are

    lblHint.Width = w - 2 * CLI_MARGIN

    txtCommand.Width = w - 2 * CLI_MARGIN
    txtCommand.Height = h - CLI_BOTTOM_BAND - txtCommand.Top

    lblStatus.Top = h - CLI_BOTTOM_BAND
    cmdCancel.Top = lblStatus.Top + CLI_STATUS_TO_BUTTONS
    cmdCancel.Left = w - CLI_MARGIN - cmdCancel.Width
    cmdRun.Top = cmdCancel.Top
    cmdRun.Left = cmdCancel.Left - CLI_BUTTON_GAP - cmdRun.Width
    lblStatus.Width = cmdRun.Left - CLI_MARGIN - CLI_MARGIN
    On Error GoTo 0
End Sub

Private Sub cmdRun_Click()
    RunCurrentText
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub txtCommand_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn And (Shift And 2) = 2 Then   ' Ctrl+Enter
        KeyCode = 0
        RunCurrentText
    End If
End Sub

Private Sub RunCurrentText()
    Dim src As String
    src = txtCommand.Text
    If Len(Trim$(src)) = 0 Then Exit Sub
    lblStatus.Caption = "Running..."
    Me.Repaint
    lblStatus.Caption = VLA_IDE.VlaCliRun(src)
    txtCommand.SetFocus
End Sub
