Attribute VB_Name = "VLA_Provenance"
Option Explicit
Public Const VLA_PROVENANCE_VERSION As String = "SEC8.0"

' =====================================================================
'  VLA_Provenance - SEC.8: does the workbook carrying this program come
'  from outside, and may it therefore have external effect?
'
'  WHY THIS EXISTS. Office's 2022 default blocks macros in a file that
'  arrived from the internet. That block reads the file's Mark-of-the-
'  Web and applies to VBA projects. A Frazaro program is not a VBA
'  project - it lives in CELLS - so nothing in that default has ever
'  applied to it. A plain .xlsx mailed from outside carries its program
'  in cells, the person clicks the ordinary "Enable Editing", and the
'  interpreter then runs that program with the ADD-IN's privilege,
'  which includes real external effect. This module is the missing
'  half: Office's own provenance model, applied at Frazaro's own
'  dispatch.
'
'  THE SHAPE, and why it is split this way. Three PURE functions decide
'  the policy (VlaParseZoneIdentifier / VlaPathIsDemonstrablyLocal /
'  VlaProvenanceRefuses) and one impure function reads the byte source
'  (ReadZoneStream). The split is not decoration: the pure three carry
'  real VlaSelfTest coverage with no COM and no file at all, which is
'  the only automated coverage this item can have - see SEC.8's own
'  roadmap entry on what stays uncovered and why.
'
'  THE SIGNAL IS POSITIVE-ONLY, and that is forced, not chosen. The
'  Zone.Identifier alternate data stream cannot be read on a non-NTFS
'  volume, and it is simply ABSENT on any ordinary locally-authored
'  file. Measured at scoping time (2026-09-08): both cases fail with
'  the SAME error, 0x800A0BBA "File could not be opened." There is no
'  way to tell "this file is not from the internet" apart from "I
'  cannot tell where this file is from" by the error. So a refusal
'  rests entirely on a zone id we positively read and parsed. Anything
'  else is not evidence of safety - it is absence of evidence, and this
'  module says so by name rather than pretending otherwise.
'
'  ...which is why VlaPathIsDemonstrablyLocal exists. Absence of a mark
'  is only reassuring when the file sits somewhere a mark COULD have
'  been stored and simply was not. A local fixed-drive path is such a
'  place. A UNC share or an http(s):// (WebDAV/SharePoint) path is not:
'  those can silently carry no stream at all, so an unreadable mark
'  there is genuinely unknown provenance and is refused. Known
'  false-positive class, recorded deliberately in the roadmap entry: a
'  workbook opened straight from an https:// SharePoint URL refuses
'  external effect. A OneDrive-SYNCED folder is an ordinary C:\ path
'  and is unaffected.
'
'  THERE IS NO TRUST STORE HERE, AND THAT IS THE POINT. SEC.10 exists
'  because a consent record kept INSIDE a workbook arrives already
'  granted from a workbook someone sends you; SEC.11 exists because the
'  hash such a record is keyed to is a 32-bit polynomial and forgeable.
'  A store is attack surface. This module writes nothing, anywhere: the
'  only way to grant is Windows own out-of-band Unblock (file
'  Properties -> Unblock, or a Trusted Location), a channel no workbook
'  content can reach, pre-fill, or collide with. SEC.8 therefore does
'  NOT wait on SEC.10/SEC.11, and must never grow a store of its own -
'  doing so would re-import both bugs at once.
' =====================================================================

' Windows URL security zones (URLMON's own numbering). Public so the
' pure tests can name them, and so a reader never has to decode a bare
' 3 at a comparison site.
Public Const VlaZoneUnreadable As Long = -1     ' no mark we could read and parse
Public Const VlaZoneLocalMachine As Long = 0
Public Const VlaZoneIntranet As Long = 1
Public Const VlaZoneTrustedSites As Long = 2
Public Const VlaZoneInternet As Long = 3
Public Const VlaZoneRestricted As Long = 4

' The per-command memo. Keyed by the workbook's own full path rather
' than held as a "current command" flag: a stale flag is a silent
' security hole (it would answer for the wrong workbook), whereas a
' stale PATH simply misses the cache and re-reads. The failure mode of
' this cache is a redundant ADS read, never a wrong verdict.
Private mMemoPath As String
Private mMemoValid As Boolean
Private mMemoZone As Long

' What an external-effect call means when NO workbook was ever captured
' for this command (see VlaProvenanceGuardCaptured for the full note).
' False = allow, relying on VLA_IDE.CaptureHost having captured on every
' path a person can reach, which check_sec8_provenance_gate.ps1 pins.
' True = refuse. True is the safer end state and the intended one; it is
' held False only until a live host pass confirms it breaks nothing.
Private Const UncapturedRefuses As Boolean = False

' ---------------------------------------------------------------
'  PURE: parse a Zone.Identifier stream's text.
' ---------------------------------------------------------------
' The stream is a small INI: a [ZoneTransfer] section with a ZoneId=N
' line, usually alongside ReferrerUrl / HostUrl lines we deliberately
' ignore (a referrer is attacker-supplied text, and nothing here should
' ever make a decision from it). Returns VlaZoneUnreadable for anything
' it cannot positively parse, per this module's positive-only rule.
Public Function VlaParseZoneIdentifier(ByVal streamText As String) As Long
    VlaParseZoneIdentifier = VlaZoneUnreadable
    If Len(streamText) = 0 Then Exit Function

    ' Accept CRLF, bare LF, or bare CR: the stream is written by
    ' whatever put the mark there, not by us, so its line endings are
    ' not ours to assume.
    Dim normalised As String
    normalised = Replace(Replace(streamText, vbCrLf, vbLf), vbCr, vbLf)

    Dim lines As Variant
    lines = Split(normalised, vbLf)

    Dim i As Long
    Dim ln As String
    Dim eqPos As Long
    Dim key As String
    Dim zoneText As String
    For i = LBound(lines) To UBound(lines)
        ln = Trim$(CStr(lines(i)))
        eqPos = InStr(ln, "=")
        If eqPos > 1 Then
            key = VLA_Identity.Fold(Trim$(Left$(ln, eqPos - 1)))
            If key = "zoneid" Then
                zoneText = Trim$(Mid$(ln, eqPos + 1))
                ' Digits only. IsNumeric alone would accept "3.7" and
                ' CLng would then ROUND it to 4 - a decimal must not be
                ' able to land on a zone id it does not name.
                '
                ' Length-capped before CLng, and the cap is not cosmetic:
                ' the bytes of this stream travel WITH the file, so an
                ' attacker who authored the workbook authored them too. A
                ' crafted "ZoneId=99999999999999" is all digits and would
                ' overflow CLng, raising error 6 out of a function this
                ' module documents as pure and total. Ten digits cannot
                ' overflow a Long's ten-digit range... except at the very
                ' top of it, so the parsed value is range-checked against
                ' the zones that actually exist rather than trusted for
                ' being short. Anything else is simply not a zone id, and
                ' "not a zone id" is already this function's answer for
                ' everything it cannot positively parse.
                If AllDigits(zoneText) Then
                    If Len(zoneText) <= 9 Then
                        Dim parsed As Long
                        parsed = CLng(zoneText)
                        If parsed >= VlaZoneLocalMachine And parsed <= VlaZoneRestricted Then
                            VlaParseZoneIdentifier = parsed
                        End If
                    End If
                    Exit Function
                End If
            End If
        End If
    Next
End Function

Private Function AllDigits(ByVal s As String) As Boolean
    Dim i As Long
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        Select Case Mid$(s, i, 1)
            Case "0" To "9"
            Case Else: Exit Function
        End Select
    Next
    AllDigits = True
End Function

' ---------------------------------------------------------------
'  PURE: is this path somewhere a Mark-of-the-Web could have lived?
' ---------------------------------------------------------------
' "Demonstrably local" means a drive-letter path (C:\...). Not local: a
' UNC share (\\server\...), an http(s):// WebDAV/SharePoint path, and
' anything else we do not recognise. An UNSAVED workbook - empty path -
' counts as local: it was created in this Excel session and there is no
' file that could have been mailed to anyone.
Public Function VlaPathIsDemonstrablyLocal(ByVal fullPath As String) As Boolean
    Dim p As String
    p = Trim$(fullPath)
    If Len(p) = 0 Then
        VlaPathIsDemonstrablyLocal = True    ' never saved; no file to have arrived
        Exit Function
    End If
    If Left$(p, 2) = "\\" Then Exit Function ' UNC

    Dim folded As String
    folded = VLA_Identity.Fold(p)
    If Left$(folded, 7) = "http://" Then Exit Function
    If Left$(folded, 8) = "https://" Then Exit Function

    ' Drive-letter form: exactly "X:\" at the head.
    If Len(p) >= 3 Then
        Select Case Mid$(folded, 1, 1)
            Case "a" To "z"
                If Mid$(p, 2, 2) = ":\" Then VlaPathIsDemonstrablyLocal = True
        End Select
    End If
End Function

' ---------------------------------------------------------------
'  PURE: the policy itself, in one place.
' ---------------------------------------------------------------
' True means "refuse external effect for a program carried by this
' workbook". Kept separate from both the stream read and the dispatch
' sites so that the whole of SEC.8's decision is one reviewable
' function with full test coverage.
Public Function VlaProvenanceRefuses(ByVal zoneId As Long, _
                                     ByVal pathIsLocal As Boolean) As Boolean
    Select Case zoneId
        Case VlaZoneInternet, VlaZoneRestricted
            VlaProvenanceRefuses = True      ' positively marked as outside
        Case VlaZoneLocalMachine, VlaZoneIntranet, VlaZoneTrustedSites
            VlaProvenanceRefuses = False     ' positively marked as inside
        Case Else
            ' Unreadable. Safe ONLY where a mark could have been stored
            ' and was not; unknown provenance anywhere else.
            VlaProvenanceRefuses = Not pathIsLocal
    End Select
End Function

' ---------------------------------------------------------------
'  IMPURE: read the Zone.Identifier alternate data stream.
' ---------------------------------------------------------------
' Same ADODB.Stream idiom VLA_Loader.ReadTextFile already uses to read
' bytes, pointed at "<path>:Zone.Identifier". Verified at scoping time
' that this is not merely convenient but necessary: .NET's FileStream
' and PowerShell's Test-Path both REJECT an ADS path outright ("The
' given path's format is not supported"), while ADODB opens it. VBA's
' own Open ... For Binary would also work (it reaches CreateFileW), but
' reusing the house idiom keeps one file-reading mechanism in the code.
'
' Returns "" for every failure. Distinguishing them is not possible
' here (see this module's header) and not useful: "" means unreadable,
' and the policy function decides what unreadable is worth.
Private Function ReadZoneStream(ByVal fullPath As String) As String
    On Error GoTo unreadable
    If Len(Trim$(fullPath)) = 0 Then Exit Function
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2                              ' adTypeText
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile fullPath & ":Zone.Identifier"
    ReadZoneStream = st.ReadText(-1)         ' adReadAll
    st.Close
    Exit Function
unreadable:
    ' No Zone.Identifier, a non-NTFS volume, a path we cannot open, or
    ' no ADODB at all - one answer for all of them, by design.
    ReadZoneStream = ""
End Function

' ---------------------------------------------------------------
'  The verdict, memoised per workbook path.
' ---------------------------------------------------------------
' VLA_IDE calls VlaProvenanceCapture at CaptureHost, so the ordinary
' path reads the stream exactly ONCE per command - matching D1's own
' read-the-ambient-world-once design. Every dispatch-site guard then
' hits the memo. A site reached WITHOUT a prior capture is not a hole:
' the guard resolves the path itself and reads, so the gate can never
' be bypassed merely by finding an entry point that forgot to capture.
Public Sub VlaProvenanceCapture(ByVal fullPath As String)
    mMemoPath = fullPath
    mMemoZone = VlaParseZoneIdentifier(ReadZoneStream(fullPath))
    mMemoValid = True
End Sub

Public Function VlaProvenanceZoneFor(ByVal fullPath As String) As Long
    If mMemoValid Then
        If mMemoPath = fullPath Then
            VlaProvenanceZoneFor = mMemoZone
            Exit Function
        End If
    End If
    VlaProvenanceCapture fullPath
    VlaProvenanceZoneFor = mMemoZone
End Function

' Test seam, and a correctness guard: a memo held across two different
' workbooks would answer for the wrong one. VlaProvenanceZoneFor clears
' on a different path automatically (the key check above), but an
' explicit reset keeps the pure tests independent of each other.
Public Sub VlaProvenanceResetMemo()
    mMemoPath = ""
    mMemoValid = False
    mMemoZone = VlaZoneUnreadable
End Sub

' Pure-test seam: seed the memo with a zone WITHOUT touching the disk,
' so the guard's own refusal path can be exercised by VlaSelfTest with
' no COM, no file, and no Mark-of-the-Web to fabricate. Deliberately
' NOT a way to grant trust - it can only be called from inside this
' VBA project, and every value it can set is one ReadZoneStream could
' have returned anyway.
Public Sub VlaProvenanceSeedMemoForTest(ByVal fullPath As String, ByVal zoneId As Long)
    mMemoPath = fullPath
    mMemoZone = zoneId
    mMemoValid = True
End Sub

' ---------------------------------------------------------------
'  The guard every external-effect dispatch site calls.
' ---------------------------------------------------------------
' verbLabel is the English-facing name of the effect being refused, so
' the refusal names what the program actually tried to do rather than
' failing generically. Refuses through VLA_Messages.RaiseMsg with a
' stable id - never a raw Err.Raise (SD-2/LX.2's discipline, and
' check_raise_ratchet.ps1 holds this module at an implicit ceiling
' of 0).
Public Sub VlaProvenanceGuard(ByVal verbLabel As String, ByVal hostPath As String)
    Dim z As Long
    z = VlaProvenanceZoneFor(hostPath)
    If Not VlaProvenanceRefuses(z, VlaPathIsDemonstrablyLocal(hostPath)) Then Exit Sub

    Dim shownPath As String
    shownPath = hostPath
    If Len(Trim$(shownPath)) = 0 Then shownPath = "(this unsaved workbook)"
    VLA_Messages.RaiseMsg "sec8-untrusted-workbook-effect", _
                          "verb", verbLabel, _
                          "path", shownPath, _
                          "why", ZoneReason(z)
End Sub

' WHY THE INTERPRETER MUST NOT RE-RESOLVE THE WORKBOOK ITSELF.
' The obvious shape - each dispatch site reads ActiveWorkbook.FullName
' and gates on that - is UNSOUND, and provably so with code that already
' ships: "activate" is a dispatchable member (VLA_Interpreter.bas, the
' one-positional-argument census, ~line 2504, reached by "Go to sheet
' ..."), so a running program can change which workbook is active. A
' program carried by an internet-marked workbook could then activate an
' innocent local workbook and have its NEXT external-effect call read
' the innocent one's provenance. That is provenance laundering, and it
' would defeat the whole item.
'
' So the gate answers for the workbook that CARRIED the program - the
' one VLA_IDE.CaptureHost bound at command entry - and never for
' whatever happens to be active when a given line runs. That is also
' exactly SEC.8's scope answer: this gate is about the provenance of the
' program's own host, not of every workbook it touches (a locally
' authored program opening an internet-sourced DATA file is a different
' axis, and belongs to SEC.15, not here).
Public Sub VlaProvenanceGuardCaptured(ByVal verbLabel As String)
    If Not mMemoValid Then
        ' Reached without VLA_IDE.CaptureHost having run: the Immediate
        ' window, the host test rig, or a future entry point that forgets
        ' to capture.
        '
        ' THIS IS THE ONE PLACE THIS ITEM CHOSE COMPATIBILITY OVER
        ' STRICTNESS, and it is a one-line, reviewable choice rather than
        ' an accident - see the Const and its note below. Flipping it to
        ' True is strictly safer and is the intended end state once a
        ' live host pass has confirmed that nothing in VLA_Tests_Host
        ' routes an external-effect member through dispatch without a
        ' capture first (as of SEC.8's own scoping read, nothing does:
        ' close/printout/exportasfixedformat are deliberately unexercised
        ' there, and the suite's own SaveCopyAs/Close calls are direct
        ' VBA in the harness, not interpreter dispatch).
        '
        ' It is not a silent hole either way: check_sec8_provenance_gate.ps1
        ' fails if CaptureHost stops capturing, which is what makes the
        ' captured memo dependable on every path a person can actually
        ' reach.
        If Not UncapturedRefuses Then Exit Sub
        VLA_Messages.RaiseMsg "sec8-untrusted-workbook-effect", _
                              "verb", verbLabel, _
                              "path", "(unknown - no workbook was captured for this command)", _
                              "why", "Frazaro could not establish which workbook this program came from"
        Exit Sub
    End If
    VlaProvenanceGuard verbLabel, mMemoPath
End Sub

' The human half of the refusal: which of the two distinct reasons
' applies, since the remedy differs between them.
Private Function ZoneReason(ByVal zoneId As Long) As String
    Select Case zoneId
        Case VlaZoneInternet
            ZoneReason = "this workbook is marked as having come from the internet"
        Case VlaZoneRestricted
            ZoneReason = "this workbook is marked as having come from a restricted site"
        Case Else
            ZoneReason = "this workbook sits where Windows cannot record where a file came from (a network share or a web address), so its origin is unknown"
    End Select
End Function
