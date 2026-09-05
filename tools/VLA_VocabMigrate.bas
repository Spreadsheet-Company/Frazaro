Attribute VB_Name = "VLA_VocabMigrate"
Option Explicit

' F.13 (docs/BETA_ROADMAP.md) - throwaway, one-time migration tool.
' Converts english.vla's OLD directive DSL (pattern => template,
' test:/fail:/macro:/function:, #/' comments) into the NEW grammar
' (english-vla/english-vla-override/test-success/test-fail/bare
' defmacro/english-function forms, ; comments) - see F.13's own
' grammar table for the full mapping and its reasoning.
'
' Deliberately does NOT depend on VLA_English.bas's own Private
' internals (VocabParensClosed, VocabReadFile, etc.) - a throwaway
' tool has no business reaching into another module's private state,
' so the (small) paren-gathering and file-I/O logic below is its own
' fresh copy, not a shared one. Does depend on this project's
' VLA_Identity.Fold (a tiny, stable, foundational utility - reusing it
' is safer than reimplementing case-folding a second time).
'
' Usage: import this module into the dev workbook, then from the
' Immediate window:
'   VlaMigrateVocab "C:\...\scripts\english.vla", "C:\...\scripts\english_new.vla"
' Read the printed per-directive counts against english.vla's own
' known counts (105 rules, 210 test:, 5 fail:, 77 macro:, 5 function:,
' cross-checked three independent ways during the F.13 build - this
' tool's own count, a raw line recount, and EnglishLoadVocabulary's own
' return value against the real, already-trusted loader, all agreeing;
' an earlier scoping-time estimate of "111" was a casual grep artifact,
' not a real count, and does not apply) before trusting the output at
' scale, and hand-review a sample of the written file. Delete this
' module once the migration is complete and committed - it has no
' reason to exist once english.vla is converted.
'
' F.13 follow-up: also carries VlaReformatVocab, a second one-time
' tool for the same throwaway reason - reflows every top-level form so
' each of its own arguments sits on its own line (head alone on the
' opening line, closing paren trailing the last argument; nested
' sub-forms stay flat - only the OUTERMOST directive's own arguments
' split out). Comments and blank lines pass through verbatim. Reuses
' this module's own ParensClosed/file-I/O helpers rather than
' duplicating them a second time. Delete the whole module once both
' jobs are done and committed.

Public Sub VlaMigrateVocab(ByVal inPath As String, ByVal outPath As String)
    Dim srcText As String
    srcText = ReadFileUtf8(inPath)
    Dim lines() As String
    lines = Split(Replace(srcText, vbCrLf, vbLf), vbLf)

    Dim outBuf As String, outUsed As Long
    Dim i As Long, buf As String, line As String, p As Long, startLine As Long
    Dim cRule As Long, cOverride As Long, cTest As Long, cFail As Long
    Dim cMacro As Long, cFunc As Long, cComment As Long, cBlank As Long

    i = LBound(lines)
    Do While i <= UBound(lines)
        buf = lines(i)
        startLine = i + 1

        ' Same two multi-line-gathering rules the old loader uses
        ' (VLA_English.bas's EnglishLoadVocabularyText), reproduced
        ' here unchanged - paren-balance is quote-aware, matching
        ' ParensClosed below exactly.
        If VLA_Identity.Fold(Left$(Trim$(buf), 6)) = "macro:" And _
           Not ParensClosed(Mid$(Trim$(buf), 7)) Then
            Do While Not ParensClosed(Mid$(Trim$(buf), 7))
                i = i + 1
                If i > UBound(lines) Then
                    Err.Raise 5, "VlaMigrateVocab", "line " & startLine & _
                        ": the macro: form starting here never closes its parentheses"
                End If
                buf = buf & vbLf & lines(i)
            Loop
        End If
        If Right$(RTrim$(buf), 2) = "=>" And VLA_Identity.Fold(Left$(Trim$(buf), 5)) <> "fail:" Then
            buf = RTrim$(buf)
            Do While Not ParensClosed(Mid$(buf, InStr(buf, "=>") + 2))
                i = i + 1
                If i > UBound(lines) Then
                    Err.Raise 5, "VlaMigrateVocab", "line " & startLine & _
                        ": the template starting here never closes its parentheses"
                End If
                buf = buf & " " & Trim$(lines(i))
            Loop
        End If
        i = i + 1

        line = Trim$(buf)
        If Len(line) = 0 Then
            cBlank = cBlank + 1
            SbAdd outBuf, outUsed, vbLf
            GoTo nextLine
        End If
        If Left$(line, 1) = "#" Or Left$(line, 1) = "'" Then
            cComment = cComment + 1
            SbAdd outBuf, outUsed, ";" & Mid$(line, 2) & vbLf
            GoTo nextLine
        End If

        If VLA_Identity.Fold(Left$(line, 9)) = "function:" Then
            cFunc = cFunc + 1
            Dim fLine As String
            fLine = Trim$(Mid$(line, 10))
            p = InStr(fLine, "=>")
            ' The matched phrase is English prose (quoted, same as every
            ' other pattern/sentence argument in the grammar); the
            ' target is a reference to a VLA/VBA function, the exact
            ' same role english-vla's own template argument plays - a
            ' bare atom, never a string. The owner's own catch.
            SbAdd outBuf, outUsed, "(english-function " & _
                QuoteStr(Trim$(Left$(fLine, p - 1))) & " " & _
                Trim$(Mid$(fLine, p + 2)) & ")" & vbLf
            GoTo nextLine
        End If

        If VLA_Identity.Fold(Left$(line, 5)) = "test:" Then
            cTest = cTest + 1
            Dim tLine As String
            tLine = Trim$(Mid$(line, 6))
            p = InStr(tLine, "=>")
            SbAdd outBuf, outUsed, "(test-success " & _
                QuoteStr(Trim$(Left$(tLine, p - 1))) & " " & _
                Trim$(Mid$(tLine, p + 2)) & ")" & vbLf
            GoTo nextLine
        End If

        If VLA_Identity.Fold(Left$(line, 5)) = "fail:" Then
            cFail = cFail + 1
            Dim faLine As String
            faLine = Trim$(Mid$(line, 6))
            p = InStr(faLine, "=>")
            SbAdd outBuf, outUsed, "(test-fail " & _
                QuoteStr(Trim$(Left$(faLine, p - 1))) & " " & _
                QuoteStr(Trim$(Mid$(faLine, p + 2))) & ")" & vbLf
            GoTo nextLine
        End If

        If VLA_Identity.Fold(Left$(line, 6)) = "macro:" Then
            cMacro = cMacro + 1
            SbAdd outBuf, outUsed, Trim$(Mid$(line, 7)) & vbLf
            GoTo nextLine
        End If

        p = InStr(line, "=>")
        If p = 0 Then
            Err.Raise 5, "VlaMigrateVocab", "line " & startLine & _
                ": expected a directive or 'pattern => template', got: " & line
        End If
        If VLA_Identity.Fold(Left$(Trim$(Left$(line, p - 1)), 9)) = "override:" Then
            cOverride = cOverride + 1
            SbAdd outBuf, outUsed, "(english-vla-override " & _
                QuoteStr(Trim$(Mid$(Trim$(Left$(line, p - 1)), 10))) & " " & _
                Trim$(Mid$(line, p + 2)) & ")" & vbLf
        Else
            cRule = cRule + 1
            SbAdd outBuf, outUsed, "(english-vla " & _
                QuoteStr(Trim$(Left$(line, p - 1))) & " " & _
                Trim$(Mid$(line, p + 2)) & ")" & vbLf
        End If
nextLine:
    Loop

    WriteFileUtf8NoBom outPath, SbText(outBuf, outUsed)

    Debug.Print "=== VlaMigrateVocab"
    Debug.Print "  english-vla (rules):    " & cRule
    Debug.Print "  english-vla-override:   " & cOverride
    Debug.Print "  test-success:           " & cTest
    Debug.Print "  test-fail:              " & cFail
    Debug.Print "  defmacro (was macro:):  " & cMacro
    Debug.Print "  english-function:       " & cFunc
    Debug.Print "  comment lines:          " & cComment
    Debug.Print "  blank lines:            " & cBlank
    Debug.Print "  written to: " & outPath
End Sub

' F.13 follow-up: reflows every top-level form in a new-grammar vocab
' file so each of its own arguments sits on its own line, for
' readability. Comments (;) and blank lines pass through byte-for-byte
' (VLA.VlaReadForms strips comments entirely - they are not part of
' any form's own structure - so this walks the file line-by-line the
' same way VlaMigrateVocab does, gathering only the lines that make up
' one top-level form and leaving everything else untouched). Purely
' cosmetic - every gathered form is read back through VLA.VlaReadForms
' and re-emitted from ITS OWN structure, so the output is guaranteed
' to parse identically to the input; nothing about VLA's own
' whitespace-agnostic reading is at stake here.
Public Sub VlaReformatVocab(ByVal inPath As String, ByVal outPath As String)
    Dim srcText As String
    srcText = ReadFileUtf8(inPath)
    Dim lines() As String
    lines = Split(Replace(srcText, vbCrLf, vbLf), vbLf)

    Dim outBuf As String, outUsed As Long
    Dim i As Long, buf As String, line As String
    Dim cForm As Long

    i = LBound(lines)
    Do While i <= UBound(lines)
        buf = lines(i)
        line = Trim$(buf)
        If Len(line) = 0 Then
            SbAdd outBuf, outUsed, vbLf
            i = i + 1
        ElseIf Left$(line, 1) = ";" Then
            SbAdd outBuf, outUsed, buf & vbLf
            i = i + 1
        ElseIf Left$(line, 1) = "(" Then
            Do While Not ParensClosed(buf)
                i = i + 1
                If i > UBound(lines) Then
                    Err.Raise 5, "VlaReformatVocab", "a form starting at """ & _
                        Left$(line, 60) & "..."" never closes its parentheses"
                End If
                buf = buf & vbLf & lines(i)
            Loop
            i = i + 1
            Dim forms As Collection
            Set forms = VLA.VlaReadForms(buf)
            Dim f As Variant
            For Each f In forms
                cForm = cForm + 1
                SbAdd outBuf, outUsed, ReformatForm(f) & vbLf
            Next
        Else
            ' Not expected in a well-formed vocab file - pass through
            ' rather than guess.
            SbAdd outBuf, outUsed, buf & vbLf
            i = i + 1
        End If
    Loop

    WriteFileUtf8NoBom outPath, SbText(outBuf, outUsed)
    Debug.Print "=== VlaReformatVocab"
    Debug.Print "  forms reformatted: " & cForm
    Debug.Print "  written to: " & outPath
End Sub

' Head alone on the opening line; every remaining top-level argument
' on its own 4-space-indented line, flat (VLA.VlaWriteForm - nested
' sub-forms are never further split); the closing paren trails the
' last argument, no line of its own. A form with zero or one argument
' has nothing worth splitting and stays flat.
Private Function ReformatForm(ByVal f As Variant) As String
    If Not IsObject(f) Then
        ReformatForm = VLA.VlaWriteForm(f)
        Exit Function
    End If
    Dim fl As Collection
    Set fl = f
    If fl.Count <= 1 Then
        ReformatForm = VLA.VlaWriteForm(f)
        Exit Function
    End If
    Dim r As String
    r = "(" & VLA.VlaWriteForm(fl.Item(1)) & vbLf
    Dim k As Long
    For k = 2 To fl.Count
        r = r & "    " & VLA.VlaWriteForm(fl.Item(k))
        If k < fl.Count Then r = r & vbLf
    Next
    r = r & ")"
    ReformatForm = r
End Function

' Quote-aware paren-depth check - VocabParensClosed's own algorithm
' (VLA_English.bas), reproduced fresh rather than reaching into
' another module's Private internals from a throwaway tool.
Private Function ParensClosed(ByVal t As String) As Boolean
    Dim i As Long, c As String, depth As Long, seen As Boolean, inQuote As Boolean
    For i = 1 To Len(t)
        c = Mid$(t, i, 1)
        If c = """" Then
            inQuote = Not inQuote
        ElseIf Not inQuote Then
            If c = "(" Then
                depth = depth + 1
                seen = True
            ElseIf c = ")" Then
                depth = depth - 1
            End If
        End If
    Next
    ParensClosed = seen And depth = 0
End Function

' Wrap s as a VLA string literal - backslashes escaped first, then
' quotes, matching this project's own established escaping convention
' (VLA_English.bas's VlaStringLit). Old-format pattern/sentence/
' fragment/function-word text is bare, unquoted prose today (and
' commonly contains its OWN literal quotes already - "Put total plus 1
' into cell "B2"." is a real corpus sentence), and becomes a quoted
' argument in the new grammar - this is the one place a naive
' conversion would silently corrupt real content.
Private Function QuoteStr(ByVal s As String) As String
    QuoteStr = """" & Replace(Replace(s, "\", "\\"), """", "\""") & """"
End Function

Private Sub SbAdd(ByRef buf As String, ByRef used As Long, ByVal s As String)
    Dim n As Long
    n = Len(s)
    If n = 0 Then Exit Sub
    If used + n > Len(buf) Then
        Dim cap As Long
        cap = Len(buf)
        If cap < 4096 Then cap = 4096
        Do While used + n > cap
            cap = cap * 2
        Loop
        buf = buf & Space$(cap - Len(buf))
    End If
    Mid$(buf, used + 1, n) = s
    used = used + n
End Sub

Private Function SbText(ByRef buf As String, ByVal used As Long) As String
    SbText = Left$(buf, used)
End Function

Private Function ReadFileUtf8(ByVal filePath As String) As String
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2 ' adTypeText
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile filePath
    ReadFileUtf8 = st.ReadText(-1) ' adReadAll
    st.Close
End Function

' Writes UTF-8 WITHOUT a BOM, matching english.vla's own current
' convention (no BOM today) - ADODB always stamps a BOM on a text-mode
' UTF-8 save, so the standard workaround applies: write as text, flip
' the same stream to binary, skip its first 3 bytes (the BOM), copy
' the rest to a fresh binary stream, save that.
Private Sub WriteFileUtf8NoBom(ByVal filePath As String, ByVal content As String)
    Dim tStream As Object, bStream As Object
    Set tStream = CreateObject("ADODB.Stream")
    tStream.Type = 2 ' adTypeText
    tStream.Charset = "utf-8"
    tStream.Open
    tStream.WriteText content
    tStream.Position = 0
    tStream.Type = 1 ' adTypeBinary
    tStream.Position = 3 ' skip the UTF-8 BOM (EF BB BF)

    Set bStream = CreateObject("ADODB.Stream")
    bStream.Type = 1 ' adTypeBinary
    bStream.Open
    tStream.CopyTo bStream
    bStream.SaveToFile filePath, 2 ' adSaveCreateOverWrite
    bStream.Close
    tStream.Close
End Sub
