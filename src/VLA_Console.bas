Attribute VB_Name = "VLA_Console"
Option Explicit
Public Const VLA_CONSOLE_VERSION As String = "CLI.4"
' CLI.4: the history survives Excel - %APPDATA%\Frazaro\history.txt, one
' escaped line per command, read once a session and added to as commands
' run. A command naming a secret stays in memory only. ConsoleWord now
' folds through VLA_Identity.Fold rather than LCase$, which is locale-
' aware (LX.3): CLI.3 had written the one case R6 exists to prevent.

' =====================================================================
'  VLA_Console - what the CLI remembers between commands. CLI.3: the
'  commands themselves, recalled with Ctrl+Up and Ctrl+Down, listed by
'  the word history, and brought back by number with !N. CLI.4: the same
'  commands, kept on disk from one Excel session to the next.
'
'  A STANDARD MODULE, BECAUSE THE FORM FORGETS. frmCLI unloads on Esc,
'  and everything a UserForm holds dies with it. So the commands live
'  here, and the form keeps only what describes its own box - how far
'  back in the history the box is, and the half-typed text waiting at
'  entry zero - which SHOULD close with the window, since the next
'  opening starts with a fresh box.
'
'  THE SPLIT. Pure functions make every decision, with the state handed
'  in, so VlaSelfTest pins them over a throwaway Collection or a string
'  and a run in the middle of a CLI session leaves the session's history
'  - and the file - alone:
'
'    VlaHistoryPush        keep one command: never a blank one, never a
'                          second copy of the newest, and the oldest
'                          let go past the cap
'    VlaHistoryMove        one Ctrl+Up or Ctrl+Down step
'    VlaConsoleWordAnswer  history, history N and !N - recognised,
'                          listed, fetched, or refused
'    VlaHistoryEncode      CLI.4: one command as one line of the file,
'    VlaHistoryDecode      and one line back to its command
'    VlaHistoryFromFileText  a file's commands, and whether it is tidy
'    VlaHistoryFileText    the tidy file for a list of commands
'    VlaHistoryTailStart   where a tail read's first whole line begins
'    VlaHistoryTextFromBytes  the file's bytes as text, a bad line left out
'    VlaHistoryMayKeepOnDisk  whether a command may be written at all
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
'  ON DISK, IN THE PERSON'S OWN FOLDER (CLI.4): history.txt in
'  VLA_Loader.VlaProfileFolder, read the first time the console needs it
'  in a session - an Excel session, or the time after a reset: the VBE's
'  Reset button, VlaDevReload, or a scratch module injected into the dev
'  workbook (VlaTry and VlaSelfTests' host half both do that) - and added
'  to one line per command as commands run. UTF-8, through U.20's own
'  byte-honest helpers, never ANSI Print #. Read as data only: the text
'  goes into the list, and nothing loaded runs without the person's own
'  Ctrl+Enter. A read failure is silent - an empty history - and a write
'  failure is a note on the status line, never a dialog. A command that
'  names a secret is kept for the session and never written at all.
'  Never written into a workbook: a workbook travels, and what one person
'  typed at their console is not its business (CLI.4's roadmap entry:
'  the mild cousin of SEC.10).
'
'  LAYER:     add-in-resident only, beside VLA_IDE - the console is an
'             IDE surface, and nothing here is ever injected into a
'             user workbook.
'  MAY CALL:  VLA_Messages (the refusals); VLA_Identity (Fold);
'             VLA_Loader (the profile folder, the byte-honest file
'             helpers, the UTF-8 codec); VLA_Runtime (VlaShowError,
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
' CLI.4: the file, in VLA_Loader.VlaProfileFolder.
Private Const HISTORY_FILE_NAME As String = "history.txt"
' CLI.4: the most of the file one load reads - several times what 500
' ordinary commands take. A longer file (a runaway, or one something
' else wrote) is read from its last two megabytes, cut at a line break,
' and the tidy after the load brings it back under the cap.
Private Const HISTORY_READ_LIMIT As Long = 2097152

Private mHistory As Collection   ' the session's commands, oldest first; Nothing until the file is read
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
    If Not WouldKeep(entries, text) Then Exit Function
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
'  The pure half, continued - CLI.4's history.txt, as text and as bytes.
' ---------------------------------------------------------------------

' CLI.4: one command as one line of history.txt. Three escapes and no
' others: a backslash doubles, a carriage return is \r, a line feed is
' \n. So no record ever holds a line break, and every other character -
' an e-acute, a pound sign, a tab - is written as itself, UTF-8 on disk.
' The backslash goes first: the two escapes after it add backslashes
' that must not be doubled.
Public Function VlaHistoryEncode(ByVal text As String) As String
    VlaHistoryEncode = Replace(Replace(Replace(text, "\", "\\"), vbCr, "\r"), vbLf, "\n")
End Function

' CLI.4: one line of history.txt back to its command, in text. False,
' with text empty, for a line VlaHistoryEncode could not have written -
' a backslash before anything but \, r or n, or a lone one at the end -
' so a damaged record is skipped rather than recalled half-read. Built
' in a buffer of the record's own length (a decode is never longer), a
' run at a time between backslashes, so a long pasted program decodes
' in one pass rather than one concatenation per character.
Public Function VlaHistoryDecode(ByVal record As String, ByRef text As String) As Boolean
    Dim buf As String, i As Long, j As Long, k As Long, n As Long
    text = ""
    n = Len(record)
    buf = Space$(n)
    i = 1                                ' the next character of record to read
    Do
        k = InStr(i, record, "\")
        If k = 0 Then k = n + 1
        If k > i Then
            Mid$(buf, j + 1, k - i) = Mid$(record, i, k - i)
            j = j + k - i
        End If
        If k > n Then Exit Do
        If k = n Then Exit Function      ' a lone backslash at the very end
        j = j + 1
        Select Case Mid$(record, k + 1, 1)
            Case "\": Mid$(buf, j, 1) = "\"
            Case "r": Mid$(buf, j, 1) = vbCr
            Case "n": Mid$(buf, j, 1) = vbLf
            Case Else: Exit Function
        End Select
        i = k + 2
    Loop
    text = Left$(buf, j)
    VlaHistoryDecode = True
End Function

' CLI.4: the commands kept in a history file's text, oldest first,
' exactly as VlaHistoryPush would have kept them one run at a time:
' blank and damaged lines skipped, a repeat of the command before it
' collapsed, and only the newest cap. Any line ending reads - CRLF, a
' lone LF, a lone CR - and a byte-order mark at the front is set aside,
' so a file saved from Notepad still reads. tidy comes back True only
' when the text is already exactly VlaHistoryFileText of what came back,
' so the caller rewrites the file only when a rewrite would change it.
Public Function VlaHistoryFromFileText(ByVal fileText As String, ByVal cap As Long, _
                                       ByRef tidy As Boolean) As Collection
    Dim entries As Collection, lines As Variant, i As Long, body As String, text As String
    Set entries = New Collection
    body = fileText
    If Left$(body, 1) = ChrW$(&HFEFF&) Then body = Mid$(body, 2)
    lines = Split(Replace(Replace(body, vbCrLf, vbLf), vbCr, vbLf), vbLf)
    For i = LBound(lines) To UBound(lines)
        If Len(lines(i)) > 0 Then
            If VlaHistoryDecode(CStr(lines(i)), text) Then VlaHistoryPush entries, text, cap
        End If
    Next
    tidy = (VlaHistoryFileText(entries) = fileText)
    Set VlaHistoryFromFileText = entries
End Function

' CLI.4: a history file's whole text for these commands, oldest first -
' each as one VlaHistoryEncode line, each line ending CRLF. The tidy
' form VlaHistoryFromFileText compares against.
Public Function VlaHistoryFileText(ByVal entries As Collection) As String
    Dim parts() As String, i As Long
    If entries.Count = 0 Then Exit Function
    ReDim parts(1 To entries.Count)
    For i = 1 To entries.Count
        parts(i) = VlaHistoryEncode(CStr(entries.Item(i)))
    Next
    VlaHistoryFileText = Join(parts, vbCrLf) & vbCrLf
End Function

' CLI.4: where the first whole line begins in b(0) .. b(n - 1), the
' last n bytes of a file total bytes long: 0 when those n bytes ARE the
' whole file; otherwise just past the first line feed, since a tail
' read begins mid-line and perhaps mid-character; or n when there is no
' line feed at all, and nothing whole was read. A line feed is one byte
' in UTF-8 and never part of a longer sequence, so the cut always falls
' between characters.
Public Function VlaHistoryTailStart(ByRef b() As Byte, ByVal n As Long, _
                                    ByVal total As Long) As Long
    Dim i As Long
    If n >= total Then Exit Function
    For i = 0 To n - 1
        If b(i) = 10 Then
            VlaHistoryTailStart = i + 1
            Exit Function
        End If
    Next
    VlaHistoryTailStart = n
End Function

' CLI.4: the text of b(start) .. b(n - 1), read as UTF-8. All of it when
' it decodes, with clean True. Otherwise line by line, leaving out each
' line that does not decode, with clean False: one line saved in the
' wrong encoding - an e-acute written as a lone ANSI byte by some other
' editor - costs that line, not the whole history, and the caller's tidy
' rewrite then takes it out of the file, so it is paid for once. Empty
' lines are left out too; they mean nothing to a history.
Public Function VlaHistoryTextFromBytes(ByRef b() As Byte, ByVal start As Long, _
                                        ByVal n As Long, ByRef clean As Boolean) As String
    Dim s As String, badAt As Long, i As Long, lineStart As Long, lineEnd As Long
    Dim atEnd As Boolean, parts() As String, kept As Long
    clean = VLA_Loader.VlaUtf8Decode(b, start, n, s, badAt)
    If clean Then
        VlaHistoryTextFromBytes = s
        Exit Function
    End If
    ReDim parts(0 To 63)
    lineStart = start
    For i = start To n
        atEnd = (i = n)
        If Not atEnd Then atEnd = (b(i) = 10)
        If atEnd Then
            lineEnd = i
            If lineEnd > lineStart Then
                If b(lineEnd - 1) = 13 Then lineEnd = lineEnd - 1
            End If
            If lineEnd > lineStart Then
                If VLA_Loader.VlaUtf8Decode(b, lineStart, lineEnd, s, badAt) Then
                    If kept > UBound(parts) Then ReDim Preserve parts(0 To 2 * UBound(parts) + 1)
                    parts(kept) = s
                    kept = kept + 1
                End If
            End If
            lineStart = i + 1
        End If
    Next
    If kept > 0 Then
        ReDim Preserve parts(0 To kept - 1)
        VlaHistoryTextFromBytes = Join(parts, vbCrLf) & vbCrLf
    End If
End Function

' CLI.4: whether a command may be written to history.txt at all. One
' that names a password, a secret, a token, an API key or a credential -
' in any case, anywhere in it - is kept for this session and never
' written: "Protect this sheet with password X." is real grammar, and
' %APPDATA% is the roaming profile, which a company network may copy to
' a server. The owner's call, and the default PowerShell's own history
' keeps for the same reason. A word test, and wide on purpose: it can
' keep a harmless command off the disk, and it cannot catch a secret
' typed without a word that names it.
Public Function VlaHistoryMayKeepOnDisk(ByVal text As String) As Boolean
    Dim t As String, words As Variant, i As Long
    t = VLA_Identity.Fold(text)
    words = Array("password", "passwd", "pwd", "secret", "token", "apikey", "api key", _
                  "api-key", "api_key", "credential")
    For i = LBound(words) To UBound(words)
        If InStr(t, CStr(words(i))) > 0 Then Exit Function
    Next
    VlaHistoryMayKeepOnDisk = True
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
    If VLA_Identity.Fold(t) = "history" Then
        ConsoleWord = "history"
    ElseIf VLA_Identity.Fold(Left$(t, 8)) = "history " Then
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

' Whether VlaHistoryPush keeps text: never a blank one, never the newest
' again. CLI.4's writer asks the same question before it touches the
' file, so there is one answer to it.
Private Function WouldKeep(ByVal entries As Collection, ByVal text As String) As Boolean
    If IsBlankText(text) Then Exit Function
    If entries.Count > 0 Then
        If CStr(entries.Item(entries.Count)) = text Then Exit Function
    End If
    WouldKeep = True
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

' The session's commands - read from history.txt the first time they are
' needed (CLI.4), then kept here for the rest of the session.
Private Function SessionHistory() As Collection
    If mHistory Is Nothing Then Set mHistory = LoadHistoryFile()
    Set SessionHistory = mHistory
End Function

' Keep a command the console is about to run. Called BEFORE it runs,
' so a refused command is there to recall and fix. CLI.4: a command kept
' is also added to history.txt, unless it names a secret
' (VlaHistoryMayKeepOnDisk). Answers "" or, when that write failed, the
' note the status line should carry - never a dialog.
Public Function VlaConsoleRemember(ByVal text As String) As String
    Dim entries As Collection, isNew As Boolean
    Set entries = SessionHistory()
    isNew = WouldKeep(entries, text)
    mDropped = mDropped + VlaHistoryPush(entries, text, HISTORY_CAP)
    If isNew Then
        If VlaHistoryMayKeepOnDisk(text) Then VlaConsoleRemember = AppendHistoryFile(text)
    End If
End Function

' CLI.4: %APPDATA%\Frazaro\history.txt, or "" where there is no profile
' folder to keep it in.
Private Function HistoryFilePath() As String
    Dim folder As String
    folder = VLA_Loader.VlaProfileFolder()
    If Len(folder) > 0 Then HistoryFilePath = folder & "\" & HISTORY_FILE_NAME
End Function

' CLI.4: the file's commands, once a session - or an empty history when
' there is no profile folder, no file yet, or anything at all goes wrong
' reading it. A read failure is silent, by the entry's own terms. An
' untidy file - over the cap, a repeat, a damaged line, a line in the
' wrong encoding, or read only from its tail - is rewritten tidy,
' best-effort and just as silently: that is housekeeping, not the
' person's command. (A second Excel adding a line in the moment between
' this read and that rewrite loses it - rare enough to leave, and said
' here so nobody has to find it again.)
Private Function LoadHistoryFile() As Collection
    Dim path As String, b() As Byte, n As Long, total As Long, start As Long
    Dim fileText As String, clean As Boolean, tidy As Boolean, loaded As Collection
    Set LoadHistoryFile = New Collection
    path = HistoryFilePath()
    If Len(path) = 0 Then Exit Function
    On Error GoTo unreadable
    If Len(Dir$(path)) = 0 Then Exit Function
    n = VLA_Loader.VlaReadFileTailBytes(path, HISTORY_READ_LIMIT, b, total)
    If n <= 0 Then Exit Function
    start = VlaHistoryTailStart(b, n, total)
    fileText = VlaHistoryTextFromBytes(b, start, n, clean)
    Set loaded = VlaHistoryFromFileText(fileText, HISTORY_CAP, tidy)
    Set LoadHistoryFile = loaded
    If start > 0 Or Not clean Or Not tidy Then RewriteHistoryFile path, loaded
    Exit Function
unreadable:
    ' silent, by design: whatever went wrong, the history starts empty
End Function

' CLI.4: the file replaced by exactly these commands, tidy. Best-effort
' and silent - see LoadHistoryFile.
Private Sub RewriteHistoryFile(ByVal path As String, ByVal entries As Collection)
    Dim b() As Byte, n As Long
    On Error Resume Next
    n = VLA_Loader.VlaUtf8Encode(VlaHistoryFileText(entries), b)
    VLA_Loader.VlaWriteFileBytes path, b, n
    On Error GoTo 0
End Sub

' CLI.4: one command added to history.txt, making the profile folder
' when it is missing. Answers "" or the note for the status line. No
' profile folder at all (Mac Excel) is not a failure - there is nowhere
' to keep it - so it says nothing.
Private Function AppendHistoryFile(ByVal text As String) As String
    Dim folder As String, b() As Byte, n As Long, reason As String
    folder = VLA_Loader.VlaProfileFolder()
    If Len(folder) = 0 Then Exit Function
    On Error Resume Next
    If Len(Dir$(folder, vbDirectory)) = 0 Then MkDir folder
    If Err.Number <> 0 Then reason = Err.Description
    On Error GoTo 0
    If Len(reason) = 0 Then
        n = VLA_Loader.VlaUtf8Encode(VlaHistoryEncode(text) & vbCrLf, b)
        reason = VLA_Loader.VlaAppendFileBytes(folder & "\" & HISTORY_FILE_NAME, b, n)
    End If
    If Len(reason) > 0 Then AppendHistoryFile = "history.txt not saved: " & reason
End Function

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
