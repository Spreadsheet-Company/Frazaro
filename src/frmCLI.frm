VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCLI 
   Caption         =   "Frazaro CLI"
   ClientHeight    =   9840.001
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   9580.001
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
' when this whole feature started (BETA_REARVIEW.md's IN.13) - this is
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
' control's new position from the MARGINS and the bands, not from
' hardcoded absolutes.
Private Const CLI_MARGIN As Long = 12
Private Const CLI_BUTTON_GAP As Long = 6          ' the least room kept between the two buttons
' CLI.5: the second band, and the owner's calls over four live looks.
' Top to bottom: the two buttons, Run at the left and Clear History at
' the right; the box where definitions are written, which takes all the
' height the window gives (the box is where the work is, so it is the
' one that grows); and the transcript, a fixed strip of about a dozen
' lines of Consolas along the very bottom of the window - where a
' terminal panel sits under an editor. No hint line: its words open the
' sample, as comments. No status line: the transcript says everything
' it said, and has its room. No Close button: Esc and the window's own
' X close it. Below CLI_MIN_HEIGHT the box would have no room, and
' narrower than the two buttons side by side they would overlap, so the
' layout stays where it is, as it already did below the old minimum.
Private Const CLI_TRANSCRIPT_BAND As Long = 156
Private Const CLI_PANE_GAP As Long = 8
Private Const CLI_MIN_HEIGHT As Long = 290

' CLI.3: where this box is in the history VLA_Console keeps. mStepsBack
' 0 is entry zero, the text being typed, and 1 is the newest command
' kept; mDraft holds entry zero while the box shows an older command.
' Here and not in VLA_Console: they describe this box, and the next
' opening starts with a fresh one. The commands themselves live there,
' because this form - and everything it holds - unloads on Esc.
Private mStepsBack As Long
Private mDraft As String
Private mRecalling As Boolean   ' True while ShowInBox itself changes the text
Private mJustOpened As Boolean  ' CLI.5: True until FocusBox first selects the sample

Private Sub UserForm_Initialize()
    ' CLI.1: no font was ever set at build time - the original build
    ' script's own attempt to set txtCommand.Font.Name/.Size via
    ' PowerShell COM automation failed (a chained-property-access quirk
    ' specific to that automation path, not a VBA limitation) and was
    ' dropped, leaving MSForms' plain default font in place ever since.
    ' Font, like Caption, is a complex object property living in the
    ' binary .frx, not the plain-text .frm - set here, where it stays
    ' hand-editable without re-running tools/build_cli_form.ps1.
    ' Consolas, not the proportional default - owner-requested (paren-
    ' matching by eye is what a monospace font is FOR). Widely available
    ' since Vista/Office 2007, no fallback needed for this project's own
    ' baseline (customUI14 already assumes Office 2010+).
    txtCommand.Font.Name = "Consolas"
    txtCommand.Font.Size = 11
    ' CLI.5: the transcript - read-only, the same Consolas a size smaller,
    ' flat and borderless on the dialog's own face colour, so it reads as
    ' part of the window around the box rather than as a second box (the
    ' owner's call, from the first live look: sunken, it looked like a box
    ' someone had switched off). Set here, like the font above, so it stays
    ' hand-editable without re-running tools/build_cli_form.ps1. It is
    ' filled in UserForm_Activate: scrolling it to its newest line needs
    ' focus, which Initialize - before the form is shown - cannot give.
    txtTranscript.Font.Name = "Consolas"
    txtTranscript.Font.Size = 10
    txtTranscript.Locked = True
    txtTranscript.BackColor = vbButtonFace
    txtTranscript.SpecialEffect = fmSpecialEffectFlat
    txtTranscript.BorderStyle = fmBorderStyleNone
    ' CLI.1's sample, a real program to run or to type over - since CLI.5
    ' under three ; comments carrying what the hint line above the box
    ' used to say (VLA_Console.VlaConsoleSample has why each part is
    ' there). FocusBox selects all of it on the way in.
    txtCommand.Text = VLA_Console.VlaConsoleSample()
    mJustOpened = True
End Sub

Private Sub UserForm_Activate()
    txtCommand.SetFocus
    ShowTranscript                      ' CLI.5: the session's transcript, as far as it goes
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

' CLI.5: where every opening leaves the focus - VLA_IDE.VlaOpenCli calls
' this last, once Show has returned. In the box; and the first time,
' with all of the sample selected, as if Ctrl+A had been pressed, so the
' first key typed replaces it (the owner's call). Here and not in
' Activate, which runs inside Show: SetFocus leaves the caret after the
' last character with nothing selected (MSForms' own rule), so the
' selection has to come after the opening's last SetFocus - this one.
' Opening a window already open leaves what is being typed as it is.
Public Sub FocusBox()
    txtCommand.SetFocus
    If mJustOpened Then
        mJustOpened = False
        txtCommand.SelStart = 0
        txtCommand.SelLength = Len(txtCommand.Text)
    End If
End Sub

' CLI.2: MSForms controls have no built-in anchoring/docking at all -
' flipping WS_THICKFRAME alone would just reveal more empty dialog
' background on drag, with txtCommand and the buttons frozen at their
' original size and position. This is what actually makes the resize
' useful: txtCommand grows to fill the new space (CLI.5: between the
' buttons, which stay along the top with Clear History at the right
' edge, and the transcript, which keeps the bottom edge). Every
' position is computed from the CLI_* margin/band constants captured
' from the original build, not hardcoded, so this generalizes to any
' new size.
Private Sub UserForm_Resize()
    On Error Resume Next   ' best-effort: a layout glitch must never crash the CLI mid-resize
    Dim w As Single, h As Single
    w = Me.InsideWidth
    h = Me.InsideHeight
    ' too small to lay out sanely - leave controls where they are
    If w < 2 * CLI_MARGIN + cmdRun.Width + CLI_BUTTON_GAP + cmdClearHistory.Width Then Exit Sub
    If h < CLI_MIN_HEIGHT Then Exit Sub

    ' CLI.5: the buttons along the top, the transcript keeping
    ' CLI_TRANSCRIPT_BAND along the bottom, and the box between them.
    cmdRun.Top = CLI_MARGIN
    cmdRun.Left = CLI_MARGIN
    cmdClearHistory.Top = CLI_MARGIN
    cmdClearHistory.Left = w - CLI_MARGIN - cmdClearHistory.Width

    txtTranscript.Top = h - CLI_MARGIN - CLI_TRANSCRIPT_BAND
    txtTranscript.Left = CLI_MARGIN
    txtTranscript.Width = w - 2 * CLI_MARGIN
    txtTranscript.Height = CLI_TRANSCRIPT_BAND

    txtCommand.Top = cmdRun.Top + cmdRun.Height + CLI_PANE_GAP
    txtCommand.Left = CLI_MARGIN
    txtCommand.Width = w - 2 * CLI_MARGIN
    txtCommand.Height = txtTranscript.Top - CLI_PANE_GAP - txtCommand.Top
    On Error GoTo 0
End Sub

Private Sub cmdRun_Click()
    RunCurrentText
End Sub

Private Sub cmdClearHistory_Click()
    AskAndClearHistory
End Sub

' CLI.5 (the owner's request, from the live test): forget the kept
' commands - and, by the button being on the form at all, show that they
' ARE kept. It asks first, the owner's call: forgetting is the one thing
' the CLI does that cannot be undone, so this is the one dialog of its
' own it still opens. The question names what goes, and No is the
' default, so a stray Enter keeps everything. With nothing kept there is
' nothing to ask about, and the transcript says so. The button and
' Ctrl+Shift+Delete (WindowKeys) both come here.
Private Sub AskAndClearHistory()
    Dim n As Long
    n = VLA_Console.VlaConsoleHistoryCount()
    If n = 0 Then
        VLA_Console.VlaConsoleClearHistory
        ShowTranscript
    ElseIf MsgBox(VLA_Console.VlaConsoleClearPrompt(n), vbYesNo + vbQuestion + vbDefaultButton2, "Frazaro") = vbYes Then
        VLA_Console.VlaConsoleClearHistory
        ShowTranscript
    End If
    txtCommand.SetFocus
End Sub

Private Sub txtCommand_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim older As Boolean
    If KeyCode = vbKeyReturn And (Shift And 2) = 2 Then   ' Ctrl+Enter
        KeyCode = 0
        RunCurrentText
    ElseIf Shift = 2 And (KeyCode = vbKeyUp Or KeyCode = vbKeyDown) Then
        ' CLI.3: Ctrl+Up / Ctrl+Down, with Ctrl alone - bare Up and Down
        ' move the caret between the lines of a many-line form, and
        ' Ctrl+Shift keeps whatever the box itself does with them.
        older = (KeyCode = vbKeyUp)
        KeyCode = 0
        RecallHistory older
    Else
        WindowKeys KeyCode, Shift
    End If
End Sub

Private Sub txtTranscript_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    WindowKeys KeyCode, Shift
End Sub

Private Sub cmdRun_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    WindowKeys KeyCode, Shift
End Sub

Private Sub cmdClearHistory_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    WindowKeys KeyCode, Shift
End Sub

' CLI.5: the keys that work wherever the focus is - Esc closes the
' window, as its X does, and Ctrl+Shift+Delete clears the history, as
' the Clear History button does, asking first. VLA_Console's
' VlaConsoleWindowKey decides which key is which, where the suite pins
' it; this does them. The Close button that used to answer Esc (its
' Cancel property) is gone - the owner's call: the X and Esc already say
' it - so every control that can hold the focus hands its keys here: the
' box, the transcript (a click puts the focus there, to select and copy)
' and the two buttons (Tab reaches them). Each key is spent before it
' acts, so the box never also deletes or cuts with it.
Private Sub WindowKeys(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case VLA_Console.VlaConsoleWindowKey(KeyCode.Value, Shift)
        Case "close"
            KeyCode = 0
            Unload Me
        Case "clear history"
            KeyCode = 0
            AskAndClearHistory
    End Select
End Sub

' CLI.3: any edit - a keystroke, a paste, a cut - makes the text the
' person's own again: the box goes back to entry zero, and what is in it
' now is what the next Ctrl+Up keeps. ShowInBox's own changes do not
' count.
Private Sub txtCommand_Change()
    If Not mRecalling Then mStepsBack = 0
End Sub

Private Sub RunCurrentText()
    Dim src As String, boxText As String
    src = txtCommand.Text
    If Len(Trim$(src)) = 0 Then Exit Sub
    ' CLI.3: the console's own words - history, history N, !N - are
    ' answered here, never run and never kept. CLI.5: their answers go into
    ' the transcript, history's list among them; only !N fills the box.
    If VLA_Console.VlaConsoleAnswer(src, boxText) Then
        If Len(boxText) > 0 Then ShowInBox boxText
        ShowTranscript
        Exit Sub
    End If
    ' CLI.3: kept BEFORE it runs, so a refused command is there to fix -
    ' and the box, still holding it, is entry zero again. CLI.4: kept on
    ' disk too. CLI.5: VlaConsoleRun does all of it - keeps the command,
    ' runs it with *, ** and *** bound, and writes it and everything that
    ' came back into the transcript, how it ended last.
    mStepsBack = 0
    VLA_Console.VlaConsoleRun src
    ShowTranscript
End Sub

' CLI.5: refill the transcript from VLA_Console and scroll it to its
' newest line. An MSForms textbox only scrolls to its caret while it has
' the focus, so the focus visits the pane and comes back to the box -
' with the box's own caret and selection put back as they were, since
' coming back into a textbox can otherwise select all of it.
Private Sub ShowTranscript()
    Dim keepStart As Long, keepLength As Long
    On Error Resume Next   ' best-effort, like the resize: a display glitch must never stop a run
    keepStart = txtCommand.SelStart
    keepLength = txtCommand.SelLength
    txtTranscript.Text = VLA_Console.VlaConsoleTranscript()
    txtTranscript.SetFocus
    txtTranscript.SelStart = Len(txtTranscript.Text)
    txtCommand.SetFocus
    txtCommand.SelStart = keepStart
    txtCommand.SelLength = keepLength
    On Error GoTo 0
End Sub

' CLI.3: Ctrl+Up (older) or Ctrl+Down, one step through the history.
' CLI.5: nothing more is said - the box shows the command, and at either
' end of the history it stays as it is, the way a terminal's prompt does.
Private Sub RecallHistory(ByVal older As Boolean)
    Dim newText As String
    If VLA_Console.VlaConsoleRecall(older, mStepsBack, mDraft, txtCommand.Text, newText) Then ShowInBox newText
End Sub

' CLI.3: put text in the box without it counting as an edit - the
' Change handler would otherwise send the history back to entry zero -
' and leave the caret at the end, where typing carries on.
Private Sub ShowInBox(ByVal shown As String)
    ' The parameter is not called "text": a declaration of text makes the
    ' VBE re-case every .Text in this module to .text, and the VBE's casing
    ' is what tools/build_cli_form.ps1's round trip writes back to src\.
    mRecalling = True
    On Error Resume Next   ' best-effort, like the resize: never leave mRecalling stuck True
    txtCommand.Text = shown
    txtCommand.SelStart = Len(txtCommand.Text)
    On Error GoTo 0
    mRecalling = False
End Sub
