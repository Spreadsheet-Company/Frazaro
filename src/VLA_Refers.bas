Attribute VB_Name = "VLA_Refers"
Option Explicit
Public Const VLA_REFERS_VERSION As String = "AXM.7"
' AXM.7: THE FORMULA-REFERENCE READER (Stage 0.2 of docs/SINGULARITY.md).
' Pure string work: the references inside a formula, read from the
' formula's text alone, which is what refers(from, to) is built from. No
' Excel object, no Range, no Application, no file - a function from text
' to records that the pure suite calls with no workbook (TestFormulaRefs,
' VLA_Tests.bas), and that VlaWriteRefersGolden runs over scripts/refers.txt
' to write scripts/refers_golden.txt, the golden the core's refers.rs
' (PORT.8, slice 8b) is held to. A parser for references, not for the
' formula language: operators and functions are stepped over, and the
' binders of LET and LAMBDA read as names (a recorded limit).
'
' WHAT IT READS, each a kind of the record below. A cell (A1, $A$1, a$1),
' a range (A1:B2), a whole column (A:A, $A:C) and a whole row (1:3), each
' bare or behind a sheet qualifier (Data!A1, 'Q1 Data'!A1, 'It''s'!A1). A
' name (Rate, Tax_Rate.2024, Model!Local). A structured reference as
' written (Sales[Amount], Sales[[#This Row],[Amount]], [@Amount]), its
' brackets counted so a nested one is taken whole. An external reference
' as written: [1]Sheet1!A1 as a file holds it, [Book.xlsx]Sheet1!A1 as the
' formula bar shows it, '[Book.xlsx]Q1 Data'!A1, 'C:\dir\[Book.xlsx]Sheet1'!A1,
' and an external name [1]!Rate. A 3D span as written (Jan:Dec!A1,
' 'Q1 Data:Q4 Data'!A1). A spill (A1#, Data!A1#, Rate#). INDIRECT and OFFSET,
' unreadable by name, one record per call with the read going on inside
' the call (OFFSET(A1,1,0) still yields A1, a true precedent). And #REF!,
' a broken reference (#REF!, Data!#REF!, and #REF!A1 after a deleted
' sheet). Nothing else is a reference: text inside a string literal
' ("A1"), a function spelled like a cell (LOG10( is a call because of the
' parenthesis), TRUE and FALSE, the other error literals, numbers (1.5,
' 2E+3, .5), the @ of implicit intersection (stepped over; the reference
' behind it is read as written, the catch the roadmap names) and the
' _xlfn. prefixes a file holds (_xlfn.STDEV.S( is a call like any other).
' A token spelled like a cell past the last column or row (XFE1,
' A1048577) is a name, as Excel reads it (#NAME?); XFD1048576 is the last
' cell. References come back in formula order with duplicates kept.
'
' THREE READINGS OF ONE SCAN. RefersScan tokenizes. RefersSpell renders a
' record as the second field of a refers row in the treaty's spelling
' (conformance/README.md, oracle 8): a cell, range, column or row
' qualified by its sheet, or by the home sheet when written bare, with
' its $ marks dropped and its letters upper-cased (Model!B1, Data!A:A,
' 'Q1 Data'!A1:B2); a name, a structured reference, an external reference
' and a 3D span as written; a spill as its cell with # (Model!A1#);
' (unreadable "INDIRECT"); a broken reference as Model!#REF! or as
' written. RefersR1C1 renders the whole formula in R1C1 relative to a home
' cell as Excel's FormulaR1C1 spells it: R[-2]C[-1], R1C1, RC, a whole
' column A:A collapsed to C[-1] when both ends render alike, every
' qualifier kept as written. The sheet quoting rule (RefersQuoteSheet):
' quotes when the name holds a character outside letters, digits and the
' underscore, starts with a digit, or is itself cell-shaped or
' R1C1-shaped; an apostrophe inside is doubled.
'
' Fold (VLA_Identity) for every comparison without case; no LCase$.
' Nothing here raises: a text the reader cannot read yields the
' references it could read, and the home cell's parse answers False.

' One reference, as RefersScan found it.
Public Type FormulaRef
    Kind As String        ' cell range column row name structured external 3d spill unreadable broken
    PartShape As String   ' the reference part's own shape where it has numbers: cell range column row; "" otherwise
    Book As String        ' external: the text inside the [ ] ("1", "Book.xlsx"); "" otherwise
    SheetName As String   ' the sheet as named: unquoted, '' undoubled; a span as "Jan:Dec"; "" where none is written
    Written As String     ' the whole token as it stands in the formula
    Part As String        ' the part after any qualifier or #REF!: $A$1, A1:B2, A1#, Rate, Sales[Amount], INDIRECT, #REF!
    Row1 As Long          ' the numbers PartShape names; 0 where none
    Col1 As Long
    Row2 As Long
    Col2 As Long
    RowAbs1 As Boolean
    ColAbs1 As Boolean
    RowAbs2 As Boolean
    ColAbs2 As Boolean
    Pos As Long           ' 1-based position of the token in the text the scan was given
    Span As Long          ' its length
    PartPos As Long       ' where Part begins, for a renderer
End Type

Private Const MAX_COL As Long = 16384          ' XFD
Private Const MAX_ROW As Long = 1048576

' ---------------------------------------------------------------------
'  The scan
' ---------------------------------------------------------------------

' Every reference in a formula's text, in formula order, duplicates kept.
' A leading = is stepped over. The count comes back; refs() holds 1 To
' count, and is left unallocated when the count is 0.
Public Function RefersScan(ByVal formulaText As String, ByRef refs() As FormulaRef) As Long
    Dim cnt As Long, i As Long, n As Long
    Dim ch As String
    Erase refs
    n = Len(formulaText)
    i = 1
    If n >= 1 Then
        If Mid$(formulaText, 1, 1) = "=" Then i = 2
    End If
    Do While i <= n
        ch = Mid$(formulaText, i, 1)
        If ch = """" Then
            i = SkipDelimited(formulaText, i, """")
        ElseIf Not ReadReference(formulaText, i, refs, cnt) Then
            i = i + 1
        End If
    Loop
    RefersScan = cnt
End Function

' What can begin at i: a qualifier and its part, a bare part, a bare
' structured reference, an error literal, a function name, a number, or a
' quoted text that qualifies nothing. Moves i past what it read and
' answers True; False where nothing of the kind begins (an operator, a
' space, a parenthesis).
Private Function ReadReference(ByRef s As String, ByRef i As Long, ByRef refs() As FormulaRef, _
                               ByRef cnt As Long) As Boolean
    Dim r As FormulaRef
    Dim j As Long, ch As String
    ch = Mid$(s, i, 1)
    r.Pos = i
    j = i
    If ReadQualifier(s, j, r) Then
        r.PartPos = j
        If ReadPart(s, j, r) Then FinishRef s, j, r, refs, cnt
        i = j                                    ' past the qualifier, and past the part when it read one
        ReadReference = True
        Exit Function
    End If
    If ch = "'" Then
        i = SkipDelimited(s, i, "'")             ' a quote that qualifies nothing
        ReadReference = True
    ElseIf ch = "[" Then
        j = BracketGroupEnd(s, i)                ' a bare structured reference: [@Amount], [Amount]
        r.Kind = "structured"
        r.PartPos = i
        r.Part = Mid$(s, i, j - i)
        FinishRef s, j, r, refs, cnt
        i = j
        ReadReference = True
    ElseIf ch = "#" Or ch = "$" Or IsIdentChar(ch) Then
        r.PartPos = i
        If ReadPart(s, j, r) Then FinishRef s, j, r, refs, cnt
        If j = i Then j = i + 1                  ' nothing read at all: never stand still
        i = j
        ReadReference = True
    End If
End Function

' A sheet qualifier at j: 'quoted'!, [book]sheet!, [book]!, sheet!, or a
' span first:last!. Fills Book and SheetName, moves j past the ! and
' answers True; otherwise leaves j where it was and answers False.
Private Function ReadQualifier(ByRef s As String, ByRef j As Long, ByRef r As FormulaRef) As Boolean
    Dim n As Long, k As Long, p As Long, p2 As Long, b As Long, e As Long
    Dim q As String, ch As String
    n = Len(s)
    ch = Mid$(s, j, 1)
    If ch = "'" Then
        k = DelimitedEnd(s, j, "'")              ' the position after the closing quote, or 0
        If k = 0 Then Exit Function
        If Mid$(s, k, 1) <> "!" Then Exit Function
        q = Replace(Mid$(s, j + 1, k - j - 2), "''", "'")
        b = InStr(1, q, "[", vbBinaryCompare)
        e = InStr(1, q, "]", vbBinaryCompare)
        If b > 0 And e > b Then                  ' '[Book.xlsx]Q1 Data' or 'C:\dir\[Book.xlsx]Sheet1'
            r.Book = Mid$(q, b + 1, e - b - 1)
            r.SheetName = Mid$(q, e + 1)
        Else
            r.SheetName = q
        End If
        j = k + 1
        ReadQualifier = True
        Exit Function
    End If
    k = j
    If ch = "[" Then
        e = InStr(j + 1, s, "]", vbBinaryCompare)
        If e = 0 Then Exit Function
        b = InStr(j + 1, s, "[", vbBinaryCompare)
        If b > 0 Then
            If b < e Then Exit Function          ' a nested [: a structured reference, not a book
        End If
        k = e + 1
        If Mid$(s, k, 1) = "!" Then              ' [1]!Rate, an external name
            r.Book = Mid$(s, j + 1, e - j - 1)
            r.SheetName = ""
            j = k + 1
            ReadQualifier = True
            Exit Function
        End If
    End If
    p = k
    Do While p <= n
        If Not IsIdentChar(Mid$(s, p, 1)) Then Exit Do
        p = p + 1
    Loop
    If p = k Then Exit Function                  ' no sheet name here
    q = Mid$(s, k, p - k)
    If Mid$(s, p, 1) = ":" Then                  ' first:last!, a 3D span
        p2 = p + 1
        Do While p2 <= n
            If Not IsIdentChar(Mid$(s, p2, 1)) Then Exit Do
            p2 = p2 + 1
        Loop
        If p2 > p + 1 Then
            If Mid$(s, p2, 1) = "!" Then
                q = q & ":" & Mid$(s, p + 1, p2 - p - 1)
                p = p2
            End If
        End If
    End If
    If Mid$(s, p, 1) <> "!" Then Exit Function
    If ch = "[" Then r.Book = Mid$(s, j + 1, e - j - 1)
    r.SheetName = q
    j = p + 1
    ReadQualifier = True
End Function

' The reference part at j, after a qualifier or bare: #REF! (with the cell
' or range a deleted sheet leaves after it), a cell, a range, a column, a
' row, a structured reference with its table, a name, either's spill #, or
' an unreadable call's name. Sets Kind, PartShape, Part and the numbers,
' moves j past the part and answers True. A function name, TRUE, FALSE, a
' number or another error literal is stepped over with False; j stays
' only where nothing begins.
Private Function ReadPart(ByRef s As String, ByRef j As Long, ByRef r As FormulaRef) As Boolean
    Dim k As Long
    Dim tok As String, nxt As String, first As String
    If Mid$(s, j, 1) = "#" Then
        k = ErrorLiteralEnd(s, j)
        tok = Mid$(s, j, k - j)
        j = k
        If VLA_Identity.Fold(tok) <> "#ref!" Then Exit Function
        r.Kind = "broken"
        r.Part = tok
        If ReadCellLike(s, j, r) Then r.PartPos = j - Len(r.Part)   ' #REF!A1: the cell keeps its numbers
        ReadPart = True
        Exit Function
    End If
    k = TokenEnd(s, j)
    If k = j Then Exit Function                  ' nothing begins here
    tok = Mid$(s, j, k - j)
    nxt = Mid$(s, k, 1)
    If nxt = "(" Then                            ' a function call: LOG10(, _xlfn.STDEV.S(, TRUE()
        j = k
        If VLA_Identity.Fold(tok) = "indirect" Or VLA_Identity.Fold(tok) = "offset" Then
            r.Kind = "unreadable"
            r.Part = tok
            ReadPart = True
        End If
        Exit Function
    End If
    If nxt = "[" Then                            ' a structured reference with its table: Sales[Amount]
        k = BracketGroupEnd(s, k)
        r.Kind = "structured"
        r.Part = Mid$(s, j, k - j)
        j = k
        ReadPart = True
        Exit Function
    End If
    If ReadCellLike(s, j, r) Then
        If r.PartShape = "cell" Then
            If Mid$(s, j, 1) = "#" Then          ' a spill: A1#
                r.Kind = "spill"
                r.Part = r.Part & "#"
                j = j + 1
            Else
                r.Kind = "cell"
            End If
        Else
            r.Kind = r.PartShape                 ' range, column, row
        End If
        ReadPart = True
        Exit Function
    End If
    first = Mid$(tok, 1, 1)
    If first = "$" Then first = Mid$(tok, 2, 1)
    If first = "." Or IsDigitChar(first) Then    ' a number: 1.5, 2E (its +3 follows), .5
        j = k
        Exit Function
    End If
    If VLA_Identity.Fold(tok) = "true" Or VLA_Identity.Fold(tok) = "false" Then
        j = k
        Exit Function
    End If
    r.Kind = "name"
    r.Part = tok
    j = k
    If Mid$(s, j, 1) = "#" Then                  ' a name's spill: Rate#
        r.Kind = "spill"
        r.Part = r.Part & "#"
        j = j + 1
    End If
    ReadPart = True
End Function

' A cell, range, whole column or whole row at j, its second half read
' after a colon when the two halves are of one shape. Sets PartShape, Part
' and the numbers, moves j and answers True; otherwise leaves everything
' as it was.
Private Function ReadCellLike(ByRef s As String, ByRef j As Long, ByRef r As FormulaRef) As Boolean
    Dim k As Long, k2 As Long
    Dim tok As String, tok2 As String, shp As String
    Dim r1 As Long, c1 As Long, r2 As Long, c2 As Long
    Dim ra1 As Boolean, ca1 As Boolean, ra2 As Boolean, ca2 As Boolean
    k = TokenEnd(s, j)
    If k = j Then Exit Function
    tok = Mid$(s, j, k - j)
    If Mid$(s, k, 1) = ":" Then
        k2 = TokenEnd(s, k + 1)
        If k2 > k + 1 Then tok2 = Mid$(s, k + 1, k2 - k - 1)
    End If
    If ParseCellPart(tok, r1, c1, ra1, ca1) Then
        shp = "cell"
        If Len(tok2) > 0 Then
            If ParseCellPart(tok2, r2, c2, ra2, ca2) Then shp = "range"
        End If
    ElseIf Len(tok2) > 0 Then
        If ParseColPart(tok, c1, ca1) Then
            If ParseColPart(tok2, c2, ca2) Then shp = "column"
        End If
        If Len(shp) = 0 Then
            If ParseRowPart(tok, r1, ra1) Then
                If ParseRowPart(tok2, r2, ra2) Then shp = "row"
            End If
        End If
    End If
    If Len(shp) = 0 Then Exit Function
    r.PartShape = shp
    r.Row1 = r1
    r.Col1 = c1
    r.RowAbs1 = ra1
    r.ColAbs1 = ca1
    If shp = "cell" Then
        r.Part = tok
        j = k
    Else
        r.Row2 = r2
        r.Col2 = c2
        r.RowAbs2 = ra2
        r.ColAbs2 = ca2
        r.Part = Mid$(s, j, k2 - j)
        j = k2
    End If
    ReadCellLike = True
End Function

' The record is complete: its text, and the kind a qualifier decides.
Private Sub FinishRef(ByRef s As String, ByVal j As Long, ByRef r As FormulaRef, _
                      ByRef refs() As FormulaRef, ByRef cnt As Long)
    r.Span = j - r.Pos
    r.Written = Mid$(s, r.Pos, r.Span)
    If Len(r.Book) > 0 Then
        r.Kind = "external"
    ElseIf InStr(1, r.SheetName, ":", vbBinaryCompare) > 0 Then
        r.Kind = "3d"
    End If
    If cnt = 0 Then
        ReDim refs(1 To 8)
    ElseIf cnt = UBound(refs) Then
        ReDim Preserve refs(1 To cnt * 2)
    End If
    cnt = cnt + 1
    refs(cnt) = r
End Sub

' ---------------------------------------------------------------------
'  The pieces of a part
' ---------------------------------------------------------------------

' $?LETTERS$?DIGITS, the whole token, within the sheet: the numbers and
' marks are written only when it is one.
Private Function ParseCellPart(ByVal tok As String, ByRef rowNum As Long, ByRef colNum As Long, _
                               ByRef rowAbs As Boolean, ByRef colAbs As Boolean) As Boolean
    Dim p As Long, k As Long, n As Long
    Dim letters As String, digits As String
    Dim lRow As Long, lCol As Long, lRowAbs As Boolean, lColAbs As Boolean
    n = Len(tok)
    p = 1
    If Mid$(tok, 1, 1) = "$" Then
        lColAbs = True
        p = 2
    End If
    k = p
    Do While k <= n
        If Not IsAsciiLetter(Mid$(tok, k, 1)) Then Exit Do
        k = k + 1
    Loop
    letters = Mid$(tok, p, k - p)
    If Len(letters) = 0 Or Len(letters) > 3 Then Exit Function
    If Mid$(tok, k, 1) = "$" Then
        lRowAbs = True
        k = k + 1
    End If
    p = k
    Do While k <= n
        If Not IsDigitChar(Mid$(tok, k, 1)) Then Exit Do
        k = k + 1
    Loop
    digits = Mid$(tok, p, k - p)
    If Len(digits) = 0 Or Len(digits) > 7 Then Exit Function
    If k <= n Then Exit Function                 ' something after the digits: a name (ABC123X)
    lCol = LettersToCol(letters)
    If lCol = 0 Then Exit Function
    lRow = CLng(digits)
    If lRow = 0 Or lRow > MAX_ROW Then Exit Function
    rowNum = lRow
    colNum = lCol
    rowAbs = lRowAbs
    colAbs = lColAbs
    ParseCellPart = True
End Function

' $?LETTERS, the whole token, for one end of a whole-column reference.
Private Function ParseColPart(ByVal tok As String, ByRef colNum As Long, ByRef colAbs As Boolean) As Boolean
    Dim p As Long, i As Long, lCol As Long
    Dim letters As String
    p = 1
    If Mid$(tok, 1, 1) = "$" Then p = 2
    letters = Mid$(tok, p)
    If Len(letters) = 0 Or Len(letters) > 3 Then Exit Function
    For i = 1 To Len(letters)
        If Not IsAsciiLetter(Mid$(letters, i, 1)) Then Exit Function
    Next i
    lCol = LettersToCol(letters)
    If lCol = 0 Then Exit Function
    colNum = lCol
    colAbs = (p = 2)
    ParseColPart = True
End Function

' $?DIGITS, the whole token, for one end of a whole-row reference.
Private Function ParseRowPart(ByVal tok As String, ByRef rowNum As Long, ByRef rowAbs As Boolean) As Boolean
    Dim p As Long, i As Long, lRow As Long
    Dim digits As String
    p = 1
    If Mid$(tok, 1, 1) = "$" Then p = 2
    digits = Mid$(tok, p)
    If Len(digits) = 0 Or Len(digits) > 7 Then Exit Function
    For i = 1 To Len(digits)
        If Not IsDigitChar(Mid$(digits, i, 1)) Then Exit Function
    Next i
    lRow = CLng(digits)
    If lRow = 0 Or lRow > MAX_ROW Then Exit Function
    rowNum = lRow
    rowAbs = (p = 2)
    ParseRowPart = True
End Function

' A = 1 ... XFD = 16384; 0 for letters past the last column.
Private Function LettersToCol(ByVal letters As String) As Long
    Dim i As Long, c As Long, v As Long
    For i = 1 To Len(letters)
        c = AscW(VLA_Identity.Fold(Mid$(letters, i, 1))) - 96
        If c < 1 Or c > 26 Then Exit Function
        v = v * 26 + c
    Next i
    If v > MAX_COL Then Exit Function
    LettersToCol = v
End Function

' 1 = A ... 16384 = XFD.
Public Function RefersColumnLetters(ByVal colNum As Long) As String
    Dim v As Long, rm As Long
    Dim letters As String
    v = colNum
    Do While v > 0
        rm = (v - 1) Mod 26
        letters = ChrW$(65 + rm) & letters
        v = (v - 1) \ 26
    Loop
    RefersColumnLetters = letters
End Function

' ---------------------------------------------------------------------
'  Walking the text
' ---------------------------------------------------------------------

' The position after the run of identifier characters and $ signs at j.
Private Function TokenEnd(ByRef s As String, ByVal j As Long) As Long
    Dim k As Long, n As Long
    Dim ch As String
    n = Len(s)
    k = j
    Do While k <= n
        ch = Mid$(s, k, 1)
        If ch <> "$" Then
            If Not IsIdentChar(ch) Then Exit Do
        End If
        k = k + 1
    Loop
    TokenEnd = k
End Function

' The position after the error literal at j (its #): letters, digits, /
' and _, then a closing ! or ? when there is one. #REF!, #N/A, #DIV/0!,
' #NAME?, #GETTING_DATA.
Private Function ErrorLiteralEnd(ByRef s As String, ByVal j As Long) As Long
    Dim k As Long, n As Long
    Dim ch As String
    n = Len(s)
    k = j + 1
    Do While k <= n
        ch = Mid$(s, k, 1)
        If Not (IsAsciiLetter(ch) Or IsDigitChar(ch) Or ch = "/" Or ch = "_") Then Exit Do
        k = k + 1
    Loop
    ch = Mid$(s, k, 1)
    If ch = "!" Or ch = "?" Then k = k + 1
    ErrorLiteralEnd = k
End Function

' The position after the closing delimiter of the delimited text at j (a
' string literal's ", a quoted name's '); a doubled delimiter stays
' inside. 0 when it never closes.
Private Function DelimitedEnd(ByRef s As String, ByVal j As Long, ByVal d As String) As Long
    Dim k As Long, n As Long
    n = Len(s)
    k = j + 1
    Do While k <= n
        If Mid$(s, k, 1) <> d Then
            k = k + 1
        ElseIf Mid$(s, k + 1, 1) = d Then
            k = k + 2
        Else
            DelimitedEnd = k + 1
            Exit Function
        End If
    Loop
End Function

' DelimitedEnd, or the end of the text when the delimiter never closes.
Private Function SkipDelimited(ByRef s As String, ByVal j As Long, ByVal d As String) As Long
    Dim k As Long
    k = DelimitedEnd(s, j, d)
    If k = 0 Then k = Len(s) + 1
    SkipDelimited = k
End Function

' The position after the ] that closes the [ at i, nested brackets
' counted and a ' inside taken as the escape it is there; the end of the
' text when it never closes.
Private Function BracketGroupEnd(ByRef s As String, ByVal i As Long) As Long
    Dim k As Long, n As Long, depth As Long
    Dim ch As String
    n = Len(s)
    k = i
    Do While k <= n
        ch = Mid$(s, k, 1)
        If ch = "'" Then
            k = k + 2
        Else
            If ch = "[" Then depth = depth + 1
            If ch = "]" Then depth = depth - 1
            k = k + 1
            If depth = 0 Then
                BracketGroupEnd = k
                Exit Function
            End If
        End If
    Loop
    BracketGroupEnd = n + 1
End Function

Private Function IsAsciiLetter(ByVal ch As String) As Boolean
    Dim a As Long
    If Len(ch) <> 1 Then Exit Function
    a = AscW(ch)
    IsAsciiLetter = (a >= 65 And a <= 90) Or (a >= 97 And a <= 122)
End Function

Private Function IsDigitChar(ByVal ch As String) As Boolean
    Dim a As Long
    If Len(ch) <> 1 Then Exit Function
    a = AscW(ch)
    IsDigitChar = (a >= 48 And a <= 57)
End Function

' A letter, a digit, _ or ., or any character past ASCII (a letter in
' another alphabet, as a name or a sheet may hold).
Private Function IsIdentChar(ByVal ch As String) As Boolean
    Dim a As Long
    If Len(ch) <> 1 Then Exit Function
    If IsAsciiLetter(ch) Or IsDigitChar(ch) Then
        IsIdentChar = True
    ElseIf ch = "_" Or ch = "." Then
        IsIdentChar = True
    Else
        a = AscW(ch)
        IsIdentChar = (a < 0 Or a > 127)
    End If
End Function

' ---------------------------------------------------------------------
'  The spellings
' ---------------------------------------------------------------------

' The second field of a refers row for one record, the home sheet
' qualifying a reference written bare.
Public Function RefersSpell(ByRef r As FormulaRef, ByVal homeSheet As String) As String
    Dim q As String
    If Len(r.SheetName) > 0 Then
        q = RefersQuoteSheet(r.SheetName) & "!"
    Else
        q = RefersQuoteSheet(homeSheet) & "!"
    End If
    Select Case r.Kind
    Case "cell", "range", "column", "row"
        RefersSpell = q & SpelledPart(r)
    Case "spill"
        If r.PartShape = "cell" Then
            RefersSpell = q & SpelledPart(r) & "#"
        Else
            RefersSpell = r.Written
        End If
    Case "broken"
        If Len(r.PartShape) > 0 Then
            RefersSpell = r.Written
        Else
            RefersSpell = q & "#REF!"
        End If
    Case "unreadable"
        If VLA_Identity.Fold(r.Part) = "indirect" Then
            RefersSpell = "(unreadable ""INDIRECT"")"
        Else
            RefersSpell = "(unreadable ""OFFSET"")"
        End If
    Case Else                                    ' name, structured, external, 3d: as written
        RefersSpell = r.Written
    End Select
End Function

' The part in A1 with its marks dropped and its letters upper-cased.
Private Function SpelledPart(ByRef r As FormulaRef) As String
    Select Case r.PartShape
    Case "cell"
        SpelledPart = RefersColumnLetters(r.Col1) & CStr(r.Row1)
    Case "range"
        SpelledPart = RefersColumnLetters(r.Col1) & CStr(r.Row1) & ":" & RefersColumnLetters(r.Col2) & CStr(r.Row2)
    Case "column"
        SpelledPart = RefersColumnLetters(r.Col1) & ":" & RefersColumnLetters(r.Col2)
    Case "row"
        SpelledPart = CStr(r.Row1) & ":" & CStr(r.Row2)
    End Select
End Function

' A sheet name as a reference spells it: quoted when it must be, an
' apostrophe inside doubled.
Public Function RefersQuoteSheet(ByVal sheetName As String) As String
    If SheetNeedsQuotes(sheetName) Then
        RefersQuoteSheet = "'" & Replace(sheetName, "'", "''") & "'"
    Else
        RefersQuoteSheet = sheetName
    End If
End Function

' Excel's rule as this reader predicts it (the live pass checks it where
' Excel has one): a character outside letters, digits and _, a leading
' digit, or a name that is itself a cell or an R1C1 reference.
Private Function SheetNeedsQuotes(ByVal sheetName As String) As Boolean
    Dim n As Long, i As Long
    Dim ch As String
    Dim rw As Long, cl As Long, ra As Boolean, ca As Boolean
    n = Len(sheetName)
    If n = 0 Then
        SheetNeedsQuotes = True
        Exit Function
    End If
    If IsDigitChar(Mid$(sheetName, 1, 1)) Then
        SheetNeedsQuotes = True
        Exit Function
    End If
    For i = 1 To n
        ch = Mid$(sheetName, i, 1)
        If Not IsIdentChar(ch) Or ch = "." Then
            SheetNeedsQuotes = True
            Exit Function
        End If
    Next i
    If ParseCellPart(sheetName, rw, cl, ra, ca) Then
        SheetNeedsQuotes = True
        Exit Function
    End If
    SheetNeedsQuotes = IsR1C1Shaped(sheetName)
End Function

' R, C, R1, C1, RC, R1C1, R12C34: a sheet so named is quoted.
Private Function IsR1C1Shaped(ByVal sheetName As String) As Boolean
    Dim f As String, ch As String
    Dim i As Long, n As Long
    Dim seenC As Boolean
    f = VLA_Identity.Fold(sheetName)
    n = Len(f)
    If n = 0 Then Exit Function
    ch = Mid$(f, 1, 1)
    If ch = "c" Then
        seenC = True
    ElseIf ch <> "r" Then
        Exit Function
    End If
    For i = 2 To n
        ch = Mid$(f, i, 1)
        If ch = "c" And Not seenC Then
            seenC = True
        ElseIf Not IsDigitChar(ch) Then
            Exit Function
        End If
    Next i
    IsR1C1Shaped = True
End Function

' ---------------------------------------------------------------------
'  R1C1, and the home cell
' ---------------------------------------------------------------------

' The formula rendered in R1C1 relative to the cell at (homeRow, homeCol):
' every part with numbers rewritten, everything else as written.
Public Function RefersR1C1(ByVal formulaText As String, ByVal homeRow As Long, ByVal homeCol As Long) As String
    Dim refs() As FormulaRef
    Dim cnt As Long, i As Long, p As Long
    Dim outText As String
    cnt = RefersScan(formulaText, refs)
    p = 1
    For i = 1 To cnt
        If Len(refs(i).PartShape) > 0 Then
            outText = outText & Mid$(formulaText, p, refs(i).PartPos - p) & R1C1Part(refs(i), homeRow, homeCol)
            If refs(i).Kind = "spill" Then outText = outText & "#"
            p = refs(i).Pos + refs(i).Span
        End If
    Next i
    RefersR1C1 = outText & Mid$(formulaText, p)
End Function

' One part in R1C1: a whole column or row whose two ends render alike is
' written once, as Excel writes it.
Private Function R1C1Part(ByRef r As FormulaRef, ByVal homeRow As Long, ByVal homeCol As Long) As String
    Dim a As String, b As String
    Select Case r.PartShape
    Case "cell"
        R1C1Part = "R" & RelSpell(r.Row1, r.RowAbs1, homeRow) & "C" & RelSpell(r.Col1, r.ColAbs1, homeCol)
    Case "range"
        R1C1Part = "R" & RelSpell(r.Row1, r.RowAbs1, homeRow) & "C" & RelSpell(r.Col1, r.ColAbs1, homeCol) & ":" & _
                   "R" & RelSpell(r.Row2, r.RowAbs2, homeRow) & "C" & RelSpell(r.Col2, r.ColAbs2, homeCol)
    Case "column"
        a = "C" & RelSpell(r.Col1, r.ColAbs1, homeCol)
        b = "C" & RelSpell(r.Col2, r.ColAbs2, homeCol)
        If a = b Then R1C1Part = a Else R1C1Part = a & ":" & b
    Case "row"
        a = "R" & RelSpell(r.Row1, r.RowAbs1, homeRow)
        b = "R" & RelSpell(r.Row2, r.RowAbs2, homeRow)
        If a = b Then R1C1Part = a Else R1C1Part = a & ":" & b
    End Select
End Function

' 5 for an absolute 5; [2], [-3] or nothing for a relative one.
Private Function RelSpell(ByVal n As Long, ByVal isAbs As Boolean, ByVal home As Long) As String
    If isAbs Then
        RelSpell = CStr(n)
    ElseIf n = home Then
        RelSpell = ""
    Else
        RelSpell = "[" & CStr(n - home) & "]"
    End If
End Function

' A formula's own cell as a refers row spells its first field, Model!B3
' or 'Q1 Data'!C5, read with the same scan: its sheet, row and column.
' False for anything else - a bare B3, a range, a book, more text.
Public Function RefersParseHome(ByVal homeText As String, ByRef sheetName As String, _
                                ByRef rowNum As Long, ByRef colNum As Long) As Boolean
    Dim refs() As FormulaRef
    Dim cnt As Long
    cnt = RefersScan(homeText, refs)
    If cnt <> 1 Then Exit Function
    If refs(1).Kind <> "cell" Then Exit Function
    If Len(refs(1).SheetName) = 0 Then Exit Function
    If refs(1).Pos <> 1 Or refs(1).Span <> Len(homeText) Then Exit Function
    sheetName = refs(1).SheetName
    rowNum = refs(1).Row1
    colNum = refs(1).Col1
    RefersParseHome = True
End Function
