Attribute VB_Name = "VLA_Console"
Option Explicit
Public Const VLA_CONSOLE_VERSION As String = "CLI.3"

' =====================================================================
'  VLA_Console - what the CLI remembers between commands. CLI.3: the
'  commands themselves, recalled with Ctrl+Up and Ctrl+Down, listed by
'  the word history, and brought back by number with !N.
'
'  A STANDARD MODULE, BECAUSE THE FORM FORGETS. frmCLI unloads on Esc,
'  and everything a UserForm holds dies with it. So the commands live
'  here, and the form keeps only what describes its own box - how far
'  back in the history the box is, and the half-typed text waiting at
'  entry zero - which SHOULD close with the window, since the next
'  opening starts with a fresh box.
'
'  THE SPLIT. Three pure functions make every decision, with the state
'  handed in, so VlaSelfTest pins them over a throwaway Collection and a
'  run in the middle of a CLI session leaves the session's history
'  alone:
'
'    VlaHistoryPush        keep one command: never a blank one, never a
'                          second copy of the newest, and the oldest
'                          let go past the cap
'    VlaHistoryMove        one Ctrl+Up or Ctrl+Down step
'    VlaConsoleWordAnswer  history, history N and !N - recognised,
'                          listed, fetched, or refused
'
'  Three thin session functions own the module's state, and they are
'  all frmCLI calls: VlaConsoleRemember, VlaConsoleRecall and
'  VlaConsoleAnswer. Only the last shows anything - a refusal, through
'  VlaShowError, the way VlaCliRun shows every other one - which is why
'  the pure half raises and never opens a dialog: a modal in the middle
'  of VlaSelfTest would stop the run.
'
'  THE CONSOLE'S WORDS go by their exact shape and nothing looser:
'  "history", "history N" and "!N", N a whole number, alone on one
'  line, in any case, with spaces around them ignored. Everything else
'  - an English sentence that begins "History", a rule someone wrote
'  into their own phrasebook, "History." with its full stop - reaches
'  the pipeline exactly as it did before CLI.3. The words are never
'  kept: they run nothing, so there is nothing to recall.
'
'  NUMBERS STAY PUT. When the cap lets the oldest command go, the rest
'  keep their numbers (mDropped counts what went), so a number read off
'  the history list still names the same command a few commands later.
'  The shell's own convention.
'
'  IN MEMORY ONLY, for as long as the add-in's VBA project keeps its
'  state: an Excel session, or until a reset - the VBE's Reset button,
'  VlaDevReload, or any scratch module injected into the dev workbook
'  (VlaTry, and VlaSelfTests' host half, both do that).
'  Never written into a workbook: a workbook travels, and what one
'  person typed at their console is not its business - CLI.4's entry
'  has the reasoning, and CLI.4 is what keeps it on disk instead.
'
'  LAYER:     add-in-resident only, beside VLA_IDE - the console is an
'             IDE surface, and nothing here is ever injected into a
'             user workbook.
'  MAY CALL:  VLA_Messages (the refusals); VLA_Runtime (VlaShowError,
'             from the session half only).
'  SHIPS:     yes - VLA_Build.bas's mods array, beside frmCLI, which
'             calls it.
' =====================================================================

' How many commands the session keeps: "a few hundred" (CLI.3), the
' shell's own long-standing figure. Past it, the oldest is let go.
Private Const HISTORY_CAP As Long = 500
' How many the word history lists when it is given no number.
Private Const HISTORY_LIST_DEFAULT As Long = 20
' How long one listed command may run before it is cut short with
' "...". The box is about seventy Consolas characters wide at the size
' it opens at, and a number and two spaces sit in front of each line.
Private Const HISTORY_LIST_WIDTH As Long = 60

Private mHistory As Collection   ' the session's commands, oldest first
Private mDropped As Long         ' how many the cap has let go, so numbers stay put
Private mLastList As String      ' the list the word history last put in the box

' ---------------------------------------------------------------------
'  The pure half. Everything arrives as an argument; nothing here reads
'  or writes the module's own state.
' ---------------------------------------------------------------------

' Keep one command, oldest first. Not kept: text that is nothing but
' spaces, tabs and line breaks, and text identical to the newest
' command already kept - running one command five times keeps it once.
' A repeat further back IS kept: it happened again, later. Answers how
' many of the oldest were let go to stay within cap, which the caller
' adds to its own count so that every number stays the same command.
Public Function VlaHistoryPush(ByVal entries As Collection, ByVal text As String, _
                               ByVal cap As Long) As Long
    If IsBlankText(text) Then Exit Function
    If entries.Count > 0 Then
        If CStr(entries.Item(entries.Count)) = text Then Exit Function
    End If
    entries.Add text
    Dim letGo As Long
    Do While entries.Count > cap And entries.Count > 0
        entries.Remove 1
        letGo = letGo + 1
    Loop
    VlaHistoryPush = letGo
End Function

' One Ctrl+Up (older) or Ctrl+Down step through entries. stepsBack says
' how far back the box is: 0 is entry zero - whatever is being typed -
' and 1 is the newest command kept. draft holds entry zero while the
' box shows an older command; it is taken from boxText at the moment
' the box leaves entry zero, which is what makes going away and coming
' back lose nothing. Answers False, and changes nothing, when there is
' nowhere to go. Otherwise newText is what the box should hold now, and
' entryNo is that command's number on the history list - base, the
' count the cap has let go, plus its place - or 0 for entry zero.
'
' One skip, the same in both directions, and it is what makes the first
' Ctrl+Up after a run do something: the command just run is still in
' the box AND is the newest kept, so stepping onto it would change
' nothing on the screen. Leaving entry zero steps over the newest when
' it is the text already in the box; coming back down steps over it
' again on the way to entry zero. No other step can land on its own
' text - VlaHistoryPush never keeps two identical commands in a row.
Public Function VlaHistoryMove(ByVal entries As Collection, ByVal base As Long, _
                               ByRef stepsBack As Long, ByRef draft As String, _
                               ByVal boxText As String, ByVal older As Boolean, _
                               ByRef newText As String, ByRef entryNo As Long) As Boolean
    Dim n As Long, toStep As Long
    n = entries.Count
    ' A position past the end of this history (nothing makes one today)
    ' starts again at entry zero rather than reading beyond it.
    If stepsBack < 0 Or stepsBack > n Then stepsBack = 0
    If older Then
        toStep = stepsBack + 1
        If stepsBack = 0 And n > 0 Then
            If CStr(entries.Item(n)) = boxText Then toStep = 2
        End If
        If toStep > n Then Exit Function
        If stepsBack = 0 Then draft = boxText
    Else
        If stepsBack = 0 Then Exit Function
        toStep = stepsBack - 1
        If toStep = 1 Then
            If CStr(entries.Item(n)) = draft Then toStep = 0
        End If
    End If
    stepsBack = toStep
    If toStep = 0 Then
        newText = draft
        entryNo = 0
    Else
        newText = CStr(entries.Item(n - toStep + 1))
        entryNo = base + n - toStep + 1
    End If
    VlaHistoryMove = True
End Function

' The console's own words, answered against the history handed in.
' False: the text is not one, and the caller runs it as a command,
' exactly as before CLI.3. True: boxText is what the box should hold
' now ("" leaves the box as it is), status is the status line, and
' listing comes back non-empty when boxText IS a history list, so the
' caller can hand it back as lastList. A refusal raises, through
' RaiseMsg; showing it is the caller's business.
'
' lastList is the list the word history last put in the box. Running
' that list unchanged is refused rather than handed to the English
' reader: a list of numbered commands is not a command, and it has no
' business being kept in the history as though it were one. Compared
' with line breaks and the spaces around it set aside, since what comes
' back out of a textbox is what the textbox made of it.
Public Function VlaConsoleWordAnswer(ByVal text As String, ByVal entries As Collection, _
                                     ByVal base As Long, ByVal lastList As String, _
                                     ByRef boxText As String, ByRef status As String, _
                                     ByRef listing As String) As Boolean
    Dim n As Long, kind As String, arg As String, howMany As Long, firstNo As Long
    boxText = ""
    status = ""
    listing = ""
    n = entries.Count
    If Len(lastList) > 0 Then
        If SameLines(text, lastList) Then
            VLA_Messages.RaiseMsg "cli-history-list-not-a-command", "last", CStr(base + n)
        End If
    End If
    kind = ConsoleWord(text, arg)
    If Len(kind) = 0 Then Exit Function
    VlaConsoleWordAnswer = True
    If kind = "recall" Then
        boxText = HistoryEntry(entries, base, arg)
        status = "Entry " & NoLeadingZeros(arg) & " is back in the box - Ctrl+Enter runs it."
        Exit Function
    End If
    If n = 0 Then
        status = "Nothing is kept yet - every command run here is kept, even a refused one."
        Exit Function
    End If
    If Len(arg) = 0 Then
        howMany = HISTORY_LIST_DEFAULT
    Else
        howMany = DigitsValue(arg)
        If howMany < 0 Then howMany = n     ' more digits than any history holds: all of it
    End If
    If howMany = 0 Then
        status = "history 0 lists nothing - history on its own lists the last " & HISTORY_LIST_DEFAULT & "."
        Exit Function
    End If
    If howMany > n Then howMany = n
    listing = HistoryListing(entries, base, howMany)
    boxText = listing
    firstNo = base + n - howMany + 1
    If howMany = 1 Then
        status = "Entry " & firstNo & " - type !" & firstNo & " to bring it back."
    Else
        status = "Entries " & firstNo & " to " & (base + n) & " - type !N, like !" & (base + n) & ", to bring one back."
    End If
End Function

' ---------------------------------------------------------------------
'  The pure half's own helpers.
' ---------------------------------------------------------------------

' "history" or "recall", with the number's digits in arg ("" when
' history was given none), or "" when the text is not one of the
' console's words - the module head has the exact shapes.
Private Function ConsoleWord(ByVal text As String, ByRef arg As String) As String
    Dim t As String, rest As String
    arg = ""
    t = TrimAll(text)
    If InStr(t, vbCr) > 0 Or InStr(t, vbLf) > 0 Then Exit Function
    If Left$(t, 1) = "!" Then
        If IsDigits(Mid$(t, 2)) Then
            arg = Mid$(t, 2)
            ConsoleWord = "recall"
        End If
        Exit Function
    End If
    t = Replace(t, vbTab, " ")
    If LCase$(t) = "history" Then
        ConsoleWord = "history"
    ElseIf LCase$(Left$(t, 8)) = "history " Then
        rest = Trim$(Mid$(t, 9))
        If IsDigits(rest) Then
            arg = rest
            ConsoleWord = "history"
        End If
    End If
End Function

' The command numbered digits, or a refusal that says what the history
' does hold.
Private Function HistoryEntry(ByVal entries As Collection, ByVal base As Long, _
                              ByVal digits As String) As String
    Dim n As Long, k As Long, asWritten As String
    n = entries.Count
    asWritten = NoLeadingZeros(digits)
    If n = 0 Then VLA_Messages.RaiseMsg "cli-history-empty", "n", asWritten
    k = DigitsValue(digits)          ' -1 when too long to be any command's number
    If k <= base Or k > base + n Then
        VLA_Messages.RaiseMsg "cli-history-no-such-entry", "n", asWritten, _
                              "first", CStr(base + 1), "last", CStr(base + n)
    End If
    HistoryEntry = CStr(entries.Item(k - base))
End Function

' The last howMany commands, one line each, numbered and lined up:
'
'      98  Put 5 in A1.
'      99  (defmacro (modal message) "alias for msgbox" (msgbox mess...
'     100  Put 6 in A2.
Private Function HistoryListing(ByVal entries As Collection, ByVal base As Long, _
                                ByVal howMany As Long) As String
    Dim n As Long, i As Long, w As Long, s As String
    n = entries.Count
    w = Len(CStr(base + n))
    For i = n - howMany + 1 To n
        If Len(s) > 0 Then s = s & vbCrLf
        s = s & Right$(Space$(w) & CStr(base + i), w) & "  " & _
                OneLine(CStr(entries.Item(i)), HISTORY_LIST_WIDTH)
    Next
    HistoryListing = s
End Function

' One command as one line of the list: every run of spaces, tabs and
' line breaks becomes one space, and past width it is cut short with
' "...", so a many-line defmacro lists as its opening words.
Private Function OneLine(ByVal text As String, ByVal width As Long) As String
    Dim s As String
    s = Replace(Replace(Replace(text, vbCr, " "), vbLf, " "), vbTab, " ")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    s = Trim$(s)
    If Len(s) > width Then s = Left$(s, width - 3) & "..."
    OneLine = s
End Function

' Two texts that are the same lines: every kind of line break counts as
' one kind, and spaces and line breaks at either end do not count.
Private Function SameLines(ByVal a As String, ByVal b As String) As Boolean
    SameLines = (OneKindOfBreak(TrimAll(a)) = OneKindOfBreak(TrimAll(b)))
End Function

Private Function OneKindOfBreak(ByVal s As String) As String
    OneKindOfBreak = Replace(Replace(s, vbCrLf, vbLf), vbCr, vbLf)
End Function

' The text with spaces, tabs and line breaks gone from both ends.
Private Function TrimAll(ByVal text As String) As String
    Dim a As Long, b As Long
    a = 1
    b = Len(text)
    Do While a <= b
        If InStr(" " & vbTab & vbCr & vbLf, Mid$(text, a, 1)) = 0 Then Exit Do
        a = a + 1
    Loop
    Do While b >= a
        If InStr(" " & vbTab & vbCr & vbLf, Mid$(text, b, 1)) = 0 Then Exit Do
        b = b - 1
    Loop
    TrimAll = Mid$(text, a, b - a + 1)
End Function

Private Function IsBlankText(ByVal text As String) As Boolean
    IsBlankText = (Len(TrimAll(text)) = 0)
End Function

' One or more of 0-9, and nothing else.
Private Function IsDigits(ByVal s As String) As Boolean
    If Len(s) = 0 Then Exit Function
    IsDigits = Not (s Like "*[!0-9]*")
End Function

' The digits as a person writes the number: leading zeros gone, "0" when
' nothing else is left.
Private Function NoLeadingZeros(ByVal digits As String) As String
    Dim i As Long
    i = 1
    Do While i < Len(digits) And Mid$(digits, i, 1) = "0"
        i = i + 1
    Loop
    NoLeadingZeros = Mid$(digits, i)
End Function

' The digits' value, or -1 when there are more of them than any
' command's number could need - nine is far past the cap - so a long run
' of digits is refused as no such command rather than overflowing.
Private Function DigitsValue(ByVal digits As String) As Long
    Dim s As String
    s = NoLeadingZeros(digits)
    If Len(s) > 9 Then
        DigitsValue = -1
    Else
        DigitsValue = CLng(s)
    End If
End Function

' ---------------------------------------------------------------------
'  The session half - everything frmCLI calls. The module's own state is
'  read and written here and nowhere else.
' ---------------------------------------------------------------------

Private Function SessionHistory() As Collection
    If mHistory Is Nothing Then Set mHistory = New Collection
    Set SessionHistory = mHistory
End Function

' Keep a command the console is about to run. Called BEFORE it runs,
' so a refused command is there to recall and fix.
Public Sub VlaConsoleRemember(ByVal text As String)
    mDropped = mDropped + VlaHistoryPush(SessionHistory(), text, HISTORY_CAP)
End Sub

' Ctrl+Up (older) or Ctrl+Down: one VlaHistoryMove step over the
' session's history, and the status line to leave. False leaves the box
' as it is, though status may still say why (nothing kept yet; the
' oldest already showing); "" leaves the status line alone too.
Public Function VlaConsoleRecall(ByVal older As Boolean, ByRef stepsBack As Long, _
                                 ByRef draft As String, ByVal boxText As String, _
                                 ByRef newText As String, ByRef status As String) As Boolean
    Dim entryNo As Long
    status = ""
    If VlaHistoryMove(SessionHistory(), mDropped, stepsBack, draft, boxText, older, newText, entryNo) Then
        If entryNo = 0 Then
            status = "Back to what you were typing."
        Else
            status = "History entry " & entryNo & " - Ctrl+Enter runs it."
        End If
        VlaConsoleRecall = True
    ElseIf older Then
        If SessionHistory().Count = 0 Then
            status = "Nothing is kept yet - every command run here is kept, even a refused one."
        Else
            status = "That is the oldest command kept."
        End If
    End If
End Function

' A console word, answered against the session's history. False: not
' one - run the text. True: boxText and status say what to show, ""
' leaving the box as it is. A refusal is shown the way VlaCliRun shows
' one, and the status line says so.
Public Function VlaConsoleAnswer(ByVal text As String, ByRef boxText As String, _
                                 ByRef status As String) As Boolean
    Dim listing As String
    On Error GoTo refused
    VlaConsoleAnswer = VlaConsoleWordAnswer(text, SessionHistory(), mDropped, mLastList, _
                                            boxText, status, listing)
    If Len(listing) > 0 Then mLastList = listing
    Exit Function
refused:
    Dim d As String
    d = Err.Description
    On Error Resume Next
    VLA_Runtime.VlaShowError d
    On Error GoTo 0
    boxText = ""
    status = "Failed - see message."
    VlaConsoleAnswer = True
End Function
