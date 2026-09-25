Attribute VB_Name = "VLA_Console"
Option Explicit
Public Const VLA_CONSOLE_VERSION As String = "CLI.5"
' CLI.5: the transcript and the last three results. Every command, and
' what came back - what it printed, what Frazaro said, the value it came
' to, the status - is kept for the session and shown under the box. The
' last three values are *, ** and *** in the next command. history's
' list goes to the transcript, so the box no longer ever holds a list,
' and CLI.3's guard against running one went with it. The window has no
' status line (the owner's call): every answer is a line in the
' transcript, a run's status carries its time, and Ctrl+Up says nothing
' but the command it brings back. The box opens with VlaConsoleSample,
' and Ctrl+Shift+Delete clears the history from the keyboard.
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
'    VlaTranscriptEntry    CLI.5: one command's lines in the transcript
'    VlaTranscriptPush     keep one entry, the oldest let go past the cap
'    VlaResultsPush        the newest result in front, three kept
'    VlaConsoleShowValue   a value as the transcript writes it
'    VlaConsoleClearPrompt  what the Clear History button asks first
'    VlaConsoleSample      what the box holds when the window opens
'    VlaConsoleWindowKey   what Esc and Ctrl+Shift+Delete do, anywhere
'
'  Thin session functions own the module's state, and they are all
'  frmCLI calls: VlaConsoleRun (CLI.5 - remember, run, record),
'  VlaConsoleRecall, VlaConsoleAnswer, VlaConsoleTranscript,
'  VlaConsoleHistoryCount and VlaConsoleClearHistory. None of them opens
'  a dialog: since CLI.5 a refusal is written into the transcript in its
'  own words. The pure half raises rather than showing anything, which
'  is also what lets VlaSelfTest run it - a modal in the middle of a
'  self-test run would stop it.
'
'  THE CONSOLE'S WORDS go by their exact shape and nothing looser:
'  "history", "history N" and "!N", N a whole number, and (CLI.5)
'  "clear", which empties the transcript and keeps the history - each
'  alone on one line, in any case, with spaces around them ignored.
'  Everything else - an English sentence that begins "History", a rule
'  someone wrote into their own phrasebook, "History." with its full
'  stop - reaches the pipeline exactly as it did before CLI.3. The words
'  are never kept: they run nothing, so there is nothing to recall.
'  Since CLI.5, history's list is written into the transcript; !N still
'  puts one command back in the box, unrun.
'
'  THE TRANSCRIPT (CLI.5): each command after a "~ ", then what its
'  debug-print statements printed, what Frazaro said (every dialog the
'  run would have raised, captured instead - VLA_IDE.VlaCliRun), the
'  value after "= ", and its status - OK, Failed or Translation failed,
'  and the time. A console word's answer is its own last line. Kept for
'  the Excel session, the newest TRANSCRIPT_KEEP commands, so closing
'  the window loses nothing; never written anywhere, since it holds
'  values out of the person's own workbook. The last three values ride
'  along as *, ** and *** (VLA_Interpreter binds them); only a command
'  that came to a value moves them along, so a debug-print between two
'  sums keeps * useful.
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
'  failure is a note on the command's status, never a dialog. A command
'  that names a secret is kept for the session and never written at all.
'  Never written into a workbook: a workbook travels, and what one person
'  typed at their console is not its business (CLI.4's roadmap entry:
'  the mild cousin of SEC.10).
'
'  LAYER:     add-in-resident only, beside VLA_IDE - the console is an
'             IDE surface, and nothing here is ever injected into a
'             user workbook.
'  MAY CALL:  VLA_Messages (the refusals); VLA_Identity (Fold);
'             VLA_Loader (the profile folder, the byte-honest file
'             helpers, the UTF-8 codec); VLA (VlaWriteForm, to show a
'             list); VLA_IDE (VlaCliRun, from the session half only).
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
' CLI.5: how many commands the transcript keeps - enough to scroll back
' through a long sitting, and never so much that refilling the pane each
' command is felt.
Private Const TRANSCRIPT_KEEP As Long = 200
' CLI.5: how long one shown value may run before it is cut with "...".
Private Const VALUE_SHOW_WIDTH As Long = 400
' CLI.5: *, ** and *** - the last three results.
Private Const RESULTS_KEEP As Long = 3

Private mHistory As Collection   ' the session's commands, oldest first; Nothing until the file is read
Private mDropped As Long         ' how many the cap has let go, so numbers stay put
Private mResults As Collection   ' CLI.5: the last three values, newest first
Private mTranscript As Collection ' CLI.5: the session's transcript, one entry per command

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
' now - only !N ever puts anything there; "" leaves the box as it is -
' listing is history's list (CLI.5: it goes to the transcript), and
' status is the answer's own line there. clearsTranscript comes back
' True for the word clear - emptying the transcript is the caller's to
' do. A refusal raises, through RaiseMsg; what becomes of it is the
' caller's business.
Public Function VlaConsoleWordAnswer(ByVal text As String, ByVal entries As Collection, _
                                     ByVal base As Long, ByRef boxText As String, _
                                     ByRef status As String, ByRef listing As String, _
                                     Optional ByRef clearsTranscript As Boolean) As Boolean
    Dim n As Long, kind As String, arg As String, howMany As Long, firstNo As Long
    boxText = ""
    status = ""
    listing = ""
    clearsTranscript = False
    n = entries.Count
    kind = ConsoleWord(text, arg)
    If Len(kind) = 0 Then Exit Function
    VlaConsoleWordAnswer = True
    If kind = "clear" Then
        clearsTranscript = True
        status = "Transcript cleared - the history, *, ** and *** are kept."
        Exit Function
    End If
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
'  The pure half, continued - CLI.5's transcript, and the last three
'  results.
' ---------------------------------------------------------------------

' CLI.5: one command's lines in the transcript. The command after "~ "
' (the owner's call; it was "> " at first), its later lines indented two
' spaces to sit under it; then what its debug-print statements printed,
' line by line; then what Frazaro said - each captured dialog's own
' words, without the "Frazaro: " title the capture writes in front; then
' "= " and the value, when it came to one; then the status, when there
' is one.
'
'    ~ (* 6 7)
'    = 42
'    OK - 18:02:11
Public Function VlaTranscriptEntry(ByVal cmdText As String, ByVal printed As Collection, _
                                   ByVal said As String, ByVal hasValue As Boolean, _
                                   ByVal valueText As String, ByVal status As String) As String
    Dim cmdLines As Variant, i As Long, out As String, p As Variant
    cmdLines = Split(Replace(Replace(TrimAll(cmdText), vbCrLf, vbLf), vbCr, vbLf), vbLf)
    out = "~"
    If UBound(cmdLines) >= 0 Then out = "~ " & cmdLines(0)
    For i = 1 To UBound(cmdLines)
        out = out & vbCrLf & "  " & cmdLines(i)
    Next
    If Not printed Is Nothing Then
        For Each p In printed
            out = out & vbCrLf & CStr(p)
        Next
    End If
    If Len(said) > 0 Then out = out & vbCrLf & SaidText(said)
    If hasValue Then out = out & vbCrLf & "= " & valueText
    If Len(status) > 0 Then out = out & vbCrLf & status
    VlaTranscriptEntry = out
End Function

' CLI.5 (the owner's call): the question frmCLI's Clear History button
' asks before anything is forgotten - naming how many commands go, and
' the file, because it cannot be undone.
Public Function VlaConsoleClearPrompt(ByVal n As Long) As String
    VlaConsoleClearPrompt = "Clear the history? This forgets the " & n & IIf(n = 1, " command", " commands") & _
        " kept here - Ctrl+Up, history and !N will start again from nothing - and deletes history.txt" & _
        " from your Frazaro folder. It cannot be undone."
End Function

' CLI.5: what the box holds when the window opens - CLI.1's sample,
' under four ; comments in the owner's own words: what the hint line
' above the box used to say (the owner's call: the line went, and its
' words moved in here, where they are read once and typed over), then
' the history's keys. "history" lists 20, CLI.3's default, not all of
' it, so the third line says 20. frmCLI selects all of it on the way
' in, so the first key typed replaces it.
'
' CLI.1's reasons, moved here with the text. ; and not #: raw VLA has no
' # comment - only the English/workspace-sheet dialect does. # has no
' special meaning to VLA.Tokenize at all (only ";" is a comment case),
' so a "#..." row reads as an ordinary bare symbol and gets evaluated as
' a variable reference - "nothing stored at key '#...'" is the normal
' unbound-variable error, not a hash-lookup feature. Deliberately not
' teaching the raw reader # too: # is genuinely unclaimed VLA syntax
' today, which is exactly the character a future reader macro (#t/#f,
' #x hex literals) would want - spending it on a second comment spelling
' forecloses that for a convenience one comment line covers for free.
' And a real, runnable program, not a placeholder, showing both halves
' of the pipeline at once: a multi-line raw VLA form (L0.2's own payoff)
' defining a tail-recursive function, and msgbox as the exit point for
' actually SEEING a macro's result while testing one, the thing a new
' user has no other way to discover from the CLI alone. fib-iter mirrors
' L18's own proven test shape exactly (TestL18's "fact-iter",
' VLA_Tests_Grammar.bas) - accumulator-style, every parameter (byval),
' the recursive call in tail position via (return (fib-iter ...)).
' Honest caveat, not shown in the text itself but worth knowing: L18's
' actual rebind-and-jump transform is EmitProc's own trick - Compile
' only. Interpret (what the CLI always runs through) has no such
' optimization, so this genuinely recurses via CallUserProc rather than
' looping in the metal; correct either way for n=10, just not the
' O(1)-stack claim L18 makes under Compile.
Public Function VlaConsoleSample() As String
    VlaConsoleSample = _
        "; VLA or English - [Control+Enter] runs it, [Escape] closes this window." & vbCrLf & _
        "; VLA comments start with ; (not #), like these four lines." & vbCrLf & _
        "; [Control+UpArrow] brings back what you ran last; " & Chr$(34) & "history" & Chr$(34) & " lists 20." & vbCrLf & _
        "; [Control+Shift+Delete] clears the history - it asks first." & vbCrLf & _
        vbCrLf & _
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
End Function

' CLI.5: what a key does wherever the focus is in frmCLI - in the box,
' in the transcript or on a button. "close" for Esc, as the window's X
' (the Close button went); "clear history" for Ctrl+Shift+Delete, which
' asks first, as the Clear History button does; "" for every other key,
' which is left to the control that has it. The chord is the owner's
' call - a CLI is the keyboard's, so nothing it does may need the mouse -
' and it is the one every browser uses to clear its history, and one
' nothing in the box uses. Ctrl+Backspace, the first thought, is every
' text box's delete-the-word-before, and stays that. mask is MSForms'
' Shift argument: 1 Shift, 2 Ctrl, 4 Alt - exactly Ctrl and Shift here.
Public Function VlaConsoleWindowKey(ByVal pressed As Long, ByVal mask As Long) As String
    If pressed = vbKeyEscape Then
        VlaConsoleWindowKey = "close"
    ElseIf pressed = vbKeyDelete And mask = 3 Then
        VlaConsoleWindowKey = "clear history"
    End If
End Function

' CLI.5: keep one transcript entry, and only the newest cap of them.
Public Sub VlaTranscriptPush(ByVal entries As Collection, ByVal entry As String, _
                             ByVal cap As Long)
    entries.Add entry
    Do While entries.Count > cap And entries.Count > 0
        entries.Remove 1
    Loop
End Sub

' CLI.5: a new result goes in front - it is * now, the old * is **, the
' old ** is *** - and only the newest RESULTS_KEEP stay.
Public Sub VlaResultsPush(ByVal results As Collection, ByVal v As Variant)
    If results.Count = 0 Then
        results.Add v
    Else
        results.Add v, Before:=1
    End If
    Do While results.Count > RESULTS_KEEP
        results.Remove results.Count
    Loop
End Sub

' CLI.5: a value the way the transcript writes it after "= ". Text is
' quoted and escaped the way VLA writes a string, so what is shown can be
' typed back; a number is written with a full stop whatever the
' machine's locale; true and false as VLA spells them; a date as
' yyyy-mm-dd, with the time when it has one; a list the way VLA writes
' it (VLA.VlaWriteForm); a one-dimensional array as (array ...); a grid
' - a range's values - by its size; a range by its sheet and address;
' any other object by its kind. Cut at VALUE_SHOW_WIDTH.
Public Function VlaConsoleShowValue(ByVal v As Variant) As String
    Dim s As String
    s = ShowValue(v)
    If Len(s) > VALUE_SHOW_WIDTH Then s = Left$(s, VALUE_SHOW_WIDTH - 3) & "..."
    VlaConsoleShowValue = s
End Function

' The rules VlaConsoleShowValue names, one value at a time.
Private Function ShowValue(ByRef v As Variant) As String
    If IsObject(v) Then
        ShowValue = ShowObject(v)
        Exit Function
    End If
    If IsArray(v) Then
        ShowValue = ShowArray(v)
        Exit Function
    End If
    Select Case VarType(v)
        Case vbString
            ShowValue = """" & Replace(Replace(v, "\", "\\"), """", "\""") & """"
        Case vbBoolean
            If v Then ShowValue = "true" Else ShowValue = "false"
        Case vbDate
            If v = Int(v) Then
                ShowValue = Format$(v, "yyyy-mm-dd")
            Else
                ShowValue = Format$(v, "yyyy-mm-dd hh:nn:ss")
            End If
        Case vbError
            ShowValue = ShowErrorValue(v)
        Case vbNull
            ShowValue = "null"
        Case vbEmpty
            ShowValue = "nothing"
        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
            ShowValue = Trim$(Str$(v))
        Case Else
            ShowValue = CStr(v)
    End Select
End Function

' An object: a list as VLA writes it, a range by its sheet and address,
' a sheet or a workbook by name, and anything else - or anything that
' cannot say (a range on a sheet since deleted) - by its kind.
Private Function ShowObject(ByRef v As Variant) As String
    Dim s As String
    If v Is Nothing Then
        ShowObject = "nothing"
        Exit Function
    End If
    On Error Resume Next
    Select Case TypeName(v)
        Case "Collection": s = VLA.VlaWriteForm(v)
        Case "Range": s = "the range " & v.Parent.Name & "!" & v.Address(False, False)
        Case "Worksheet": s = "the sheet " & v.Name
        Case "Workbook": s = "the workbook " & v.Name
    End Select
    On Error GoTo 0
    If Len(s) = 0 Then s = "a " & TypeName(v)
    ShowObject = s
End Function

' An array: one dimension as (array ...), its items shown by the same
' rules and the first twenty only; two dimensions - what a range's
' values come back as - by their size, since a grid is not a line.
Private Function ShowArray(ByRef v As Variant) As String
    Dim dims As Long, i As Long, s As String, shown As Long
    dims = ArrayDims(v)
    If dims = 0 Then
        ShowArray = "(array)"
    ElseIf dims = 1 Then
        s = "(array"
        For i = LBound(v) To UBound(v)
            If shown = 20 Then
                s = s & " ..."
                Exit For
            End If
            s = s & " " & ShowValue(v(i))
            shown = shown + 1
        Next
        ShowArray = s & ")"
    ElseIf dims = 2 Then
        ShowArray = "an array of " & (UBound(v, 1) - LBound(v, 1) + 1) & " rows and " & _
                    (UBound(v, 2) - LBound(v, 2) + 1) & " columns"
    Else
        ShowArray = "an array of " & dims & " dimensions"
    End If
End Function

' How many dimensions an array has - 0 for one not yet given any.
Private Function ArrayDims(ByRef v As Variant) As Long
    Dim d As Long, probe As Long
    On Error GoTo done
    Do
        probe = UBound(v, d + 1)
        d = d + 1
    Loop
done:
    ArrayDims = d
End Function

' A worksheet error value by the name Excel shows for it.
Private Function ShowErrorValue(ByRef v As Variant) As String
    Select Case CStr(v)
        Case "Error 2000": ShowErrorValue = "#NULL!"
        Case "Error 2007": ShowErrorValue = "#DIV/0!"
        Case "Error 2015": ShowErrorValue = "#VALUE!"
        Case "Error 2023": ShowErrorValue = "#REF!"
        Case "Error 2029": ShowErrorValue = "#NAME?"
        Case "Error 2036": ShowErrorValue = "#NUM!"
        Case "Error 2042": ShowErrorValue = "#N/A"
        Case Else: ShowErrorValue = CStr(v)
    End Select
End Function

' What the captured dialogs said, without the "Frazaro: " title that
' VlaShowError's capture writes in front of each - the transcript is
' Frazaro's own page already.
Private Function SaidText(ByVal said As String) As String
    Dim s As String
    s = said
    If Left$(s, 9) = "Frazaro: " Then s = Mid$(s, 10)
    SaidText = Replace(s, vbCrLf & "Frazaro: ", vbCrLf)
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
    ' CLI.5: clear - what it means in every terminal, the screen: here, the
    ' transcript. Never the history, so the word typed from habit loses
    ' nothing; forgetting commands is the Clear History button's, which
    ' asks first. "clear." and "Clear range A1:C10." stay English's.
    If VLA_Identity.Fold(t) = "clear" Then
        ConsoleWord = "clear"
        Exit Function
    End If
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
' note the command's status should carry - never a dialog.
Public Function VlaConsoleRemember(ByVal text As String) As String
    Dim entries As Collection, isNew As Boolean
    Set entries = SessionHistory()
    isNew = WouldKeep(entries, text)
    mDropped = mDropped + VlaHistoryPush(entries, text, HISTORY_CAP)
    If isNew Then
        If VlaHistoryMayKeepOnDisk(text) Then VlaConsoleRemember = AppendHistoryFile(text)
    End If
End Function

' CLI.5: run one command the way the CLI runs it - kept first (CLI.3,
' and on disk, CLI.4), then run with the last three results bound as *,
' ** and ***, then written into the transcript with everything that came
' back. A value moves the three along: it becomes *, and the oldest of
' them goes. The status - how the run ended, and when - is the entry's
' last line, with a note from a failed history write on the end of it;
' the window has no status line of its own.
Public Sub VlaConsoleRun(ByVal text As String)
    Dim note As String, status As String, said As String, printed As Collection
    Dim runValue As Variant, hasValue As Boolean, valueText As String
    note = VlaConsoleRemember(text)
    status = VLA_IDE.VlaCliRun(text, SessionResults(), said, printed, runValue, hasValue)
    If hasValue Then
        valueText = VlaConsoleShowValue(runValue)
        VlaResultsPush SessionResults(), runValue
    End If
    If Len(note) > 0 Then status = status & " - " & note
    VlaTranscriptPush SessionTranscript(), _
                      VlaTranscriptEntry(text, printed, said, hasValue, valueText, status), TRANSCRIPT_KEEP
End Sub

' CLI.5: the whole transcript so far, oldest first - what the strip
' under the box shows, refilled after every command and when the window
' opens again.
Public Function VlaConsoleTranscript() As String
    Dim entries As Collection, parts() As String, i As Long
    Set entries = SessionTranscript()
    If entries.Count = 0 Then Exit Function
    ReDim parts(1 To entries.Count)
    For i = 1 To entries.Count
        parts(i) = CStr(entries.Item(i))
    Next
    VlaConsoleTranscript = Join(parts, vbCrLf)
End Function

' CLI.5 (the owner's request, from the live test): forget every kept
' command - the session's list and history.txt both - and the numbers
' start again at 1. The caller has already asked (frmCLI's Clear History
' button, or Ctrl+Shift+Delete), whenever there was anything to forget;
' this only does it, and writes one line into the transcript saying so -
' or, with nothing kept, saying that. The transcript and *, ** and ***
' are left as they are. All or nothing: when history.txt cannot be
' deleted (read-only, say), nothing is forgotten, and the line says why
' - a list forgotten here but still on disk would come back with the
' next Excel session, while this one insisted there was nothing left to
' clear.
Public Sub VlaConsoleClearHistory()
    Dim n As Long, status As String, reason As String
    n = SessionHistory().Count
    If n = 0 Then
        status = "Nothing is kept yet - there is no history to clear."
    Else
        reason = DeleteHistoryFile()
        If Len(reason) > 0 Then
            status = "History not cleared - history.txt could not be deleted (" & reason & "), so nothing was forgotten."
        Else
            Set mHistory = New Collection   ' not Nothing: nothing is to be read back this session
            mDropped = 0
            status = "History cleared - " & n & IIf(n = 1, " command", " commands") & " forgotten."
        End If
    End If
    VlaTranscriptPush SessionTranscript(), status, TRANSCRIPT_KEEP
End Sub

' CLI.5: how many commands are kept - what the Clear History question
' names, and nothing when there is nothing to clear.
Public Function VlaConsoleHistoryCount() As Long
    VlaConsoleHistoryCount = SessionHistory().Count
End Function

' CLI.5: history.txt deleted - "" when it went, or when there was none
' to delete; VBA's own reason when it would not go.
Private Function DeleteHistoryFile() As String
    Dim path As String
    path = HistoryFilePath()
    If Len(path) = 0 Then Exit Function
    On Error Resume Next
    If Len(Dir$(path)) > 0 Then Kill path
    If Err.Number <> 0 Then DeleteHistoryFile = Err.Description
    On Error GoTo 0
End Function

Private Function SessionResults() As Collection
    If mResults Is Nothing Then Set mResults = New Collection
    Set SessionResults = mResults
End Function

Private Function SessionTranscript() As Collection
    If mTranscript Is Nothing Then Set mTranscript = New Collection
    Set SessionTranscript = mTranscript
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
' when it is missing. Answers "" or the note for the command's status. No
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
' session's history. False leaves the box as it is - nothing kept yet,
' or the oldest already showing. CLI.5: nothing is said either way, now
' the window has no status line: the box shows the command, and at
' either end of the history it stays as it is, the way a terminal's
' prompt does.
Public Function VlaConsoleRecall(ByVal older As Boolean, ByRef stepsBack As Long, _
                                 ByRef draft As String, ByVal boxText As String, _
                                 ByRef newText As String) As Boolean
    Dim entryNo As Long
    VlaConsoleRecall = VlaHistoryMove(SessionHistory(), mDropped, stepsBack, draft, boxText, older, newText, entryNo)
End Function

' A console word, answered against the session's history. False: not
' one - run the text. True: boxText is what the box should hold now, ""
' leaving it as it is. CLI.5: the word and its answer - history's list,
' or a refusal in its own words - go into the transcript, and nothing
' opens a dialog. clear empties the transcript and leaves one line in
' it saying what was kept, the way a browser's console says it was
' cleared.
Public Function VlaConsoleAnswer(ByVal text As String, ByRef boxText As String) As Boolean
    Dim listing As String, d As String, wipes As Boolean, status As String
    On Error GoTo refused
    VlaConsoleAnswer = VlaConsoleWordAnswer(text, SessionHistory(), mDropped, boxText, status, listing, wipes)
    If wipes Then
        Set mTranscript = New Collection    ' clear: the pane empties, as a terminal's does
        VlaTranscriptPush SessionTranscript(), status, TRANSCRIPT_KEEP
    ElseIf VlaConsoleAnswer Then
        VlaTranscriptPush SessionTranscript(), _
                          VlaTranscriptEntry(text, Nothing, listing, False, "", status), TRANSCRIPT_KEEP
    End If
    Exit Function
refused:
    d = Err.Description
    boxText = ""
    ' the refusal is the entry's last line: a console word has no status
    VlaTranscriptPush SessionTranscript(), _
                      VlaTranscriptEntry(text, Nothing, d, False, "", ""), TRANSCRIPT_KEEP
    VlaConsoleAnswer = True
End Function
