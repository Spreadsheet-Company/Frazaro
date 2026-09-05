Attribute VB_Name = "VLA_English"
Option Explicit
Public Const VLA_ENGLISH_VERSION As String = "LX5.2"
' LX5.2: BETA_ROADMAP2.md's LX.5, phase 2 - the physical split. Every
' procedure that touches the shared rule-store state (grammar
' registration, the DCG matcher, G-RENDER, statement/condition/
' expression parsing, vocabulary-file loading, sub-assembly bookkeeping,
' diagnostics, and the Public front door itself - EnglishToVla drives
' the parse loop directly, so it was never separable from the engine it
' orchestrates) moved to the new VLA_SentenceEngine.bas. A direct
' variable-usage audit - every one of this file's ~65 module-level
' Private Collections, checked against every one of its 216 procedures -
' found the two sides were NOT separable along the "grammar sections"
' this file used to be organized into: mPatItems alone was touched
' directly from 13 different regions spanning the whole file. So the
' real, honest boundary is smaller than first estimated: only the
' twelve procedures below never touch shared state at all (pure
' functions of their own parameters), plus this file's own LX5.1
' keyword-alias mechanism. Everything else genuinely IS the language-
' neutral engine LX.5 set out to isolate; this file is what's left once
' it's isolated.
' Public API today: NumberWord/OrdinalWord/ExprOpWord/IsExprOpWord/
' IsNoiseWord/IsDroppedWord/SkipArticles/IsColorWord/SlotDesc/
' StrayCharHint (this language's own vocabulary tables and teaching
' prose - the "data" LX.1's audit already named) and
' RegisterKeywordAlias/CanonicalizeStructuralWords (LX5.1's own seam - a
' phrasebook's (keyword-alias "si" "if") directive resolves here).
' Every one of the ~200 moved procedures kept its exact Public/Private
' visibility and its exact name; VBA resolves an unqualified Public
' call project-wide regardless of which module hosts it, so no caller
' anywhere in the project (VLA_Browser.bas, VLA_IDE.bas, the test
' suite) needed to change. Comments elsewhere in this project citing
' "VLA_English.bas" for a now-relocated mechanism (CandidateShapeOk,
' FindRenderRule, ClickHandlerSlug, MatchRefToken, and others) are
' historically accurate about where that code was WRITTEN, not about
' where it lives today - not swept in this pass; a future static check
' could catch the drift the way F.15 already catches the module-
' manifest one.

Public mKeywordAlias As Collection ' LX5.1: folded surface word -> canonical
                                   ' English spelling, populated only by a
                                   ' phrasebook's (keyword-alias ...)
                                   ' directive; see CanonicalizeStructuralWords


Public Function ExprOpWord(ByVal opSym As String) As String
    Select Case opSym
        Case "+": ExprOpWord = "plus"
        Case "-": ExprOpWord = "minus"
        Case "*": ExprOpWord = "times"
        Case "/": ExprOpWord = "divided by"
        Case "&": ExprOpWord = "joined with"
        Case "=": ExprOpWord = "equals"
        Case "<>": ExprOpWord = "does not equal"
        Case "<": ExprOpWord = "is less than"
        Case ">": ExprOpWord = "is greater than"
        Case "<=": ExprOpWord = "is at most"
        Case ">=": ExprOpWord = "is at least"
    End Select
End Function

' Words the expression parser treats as operators: a literal from this
' set directly after an {x:expr} or {x:cond} slot is unreachable, because
' the expression will consume it. (The column-A class of bug: a rule
' that can never match, discoverable only by accident.)
Public Function IsExprOpWord(ByVal w As String) As Boolean
    Select Case w
        Case "plus", "minus", "times", "divided", "multiplied", "joined", "followed"
            IsExprOpWord = True
    End Select
End Function

' Words stripped from PATTERNS at registration (rule authors may
' write any of these; they never match as literals).
Public Function IsNoiseWord(ByVal w As String) As Boolean
    IsNoiseWord = (w = "the" Or w = "a" Or w = "an" Or w = "please")
End Function

' Words dropped from INPUT at tokenize time. Deliberately narrower:
' "the" and "please" can never be references, but "a"/"an" can be
' (column A!), so those survive as tokens and are skipped by the
' matcher only where grammar words are expected.
Public Function IsDroppedWord(ByVal w As String) As Boolean
    IsDroppedWord = (w = "the" Or w = "please")
End Function

' Advance past articles where a grammar word (not a reference) is
' expected: before literals and structural keywords.
Public Sub SkipArticles(toks() As String, ByRef p As Long)
    Do While TokAt(toks, p) = "a" Or TokAt(toks, p) = "an"
        p = p + 1
    Loop
End Sub

' LX5.1: the language-neutrality seam. Rewrites any bare-word token a
' phrasebook has aliased (RegisterKeywordAlias, below) to its canonical
' English spelling, so ParseStmt/ParseCond/ParseExpr/ParsePrim/the
' Oxford-comma list logic - all of which only ever compare tokens
' against English spellings - never have to change to read a second
' language's control-flow words. Only IsWordTok tokens are candidates:
' a number token starts with a digit and a quoted-text token starts
' with Chr$(34), so neither is ever a key in mKeywordAlias and both
' pass through untouched, the same protection SkipArticles/IsNoiseWord
' already rely on. A no-op whenever no phrasebook has registered an
' alias - English's own load path never does - so this changes nothing
' for the shipped product until a language file asks for it.
Public Function CanonicalizeStructuralWords(toks() As String) As String()
    ' P-TOK: toks is now an array, assigned by VALUE (not by reference
    ' the way "Set x = coll" aliased the same Collection). Mutating
    ' toks in place and only THEN copying it out as the return value
    ' (rather than this function's own old order: alias first, mutate
    ' after) is required now - an early "CanonicalizeStructuralWords =
    ' toks" would snapshot a copy before any rewrite below applied.
    If mKeywordAlias.Count > 0 Then
        Dim i As Long, t As String, canon As Variant
        For i = 1 To UBound(toks)
            t = toks(i)
            If IsWordTok(t) Then
                canon = Empty
                On Error Resume Next
                canon = mKeywordAlias.Item(t)
                On Error GoTo 0
                If Not IsEmpty(canon) Then toks(i) = CStr(canon)   ' direct
                    ' write, was ReplaceAt's remove+re-add dance - only
                    ' needed for a Collection, which has no in-place Item set
            End If
        Next i
    End If
    CanonicalizeStructuralWords = toks
End Function

' The teaching half of the stray-character refusal: the symbols
' spreadsheet users reach for, mapped to the words the grammar knows.
Public Function StrayCharHint(ByVal c As String) As String
    Dim h As String
    Select Case c
        Case "+": h = "write 'plus' for addition"
        Case "*": h = "write 'times' for multiplication"
        Case "/": h = "write 'divided by' for division"
        ' G-PATH: EnTokenize is a context-free pass over raw
        ' characters - it has no way to know a {:path} slot was
        ' expected here, so this is the message anyone typing a real
        ' Windows path unquoted actually hits (MatchPathToken's own
        ' colon-lookahead refusal only catches the rarer backslash-
        ' free shape, since tokenization dies here first on any real
        ' one). Was falling through to Case Else's generic "put it in
        ' quotes" before this - true, but a bare backslash is common
        ' and specific enough now to earn its own, more direct hint.
        Case "\": h = "if this is a file path, put the WHOLE path in ""quotes"" (a period ends a sentence otherwise)"
        Case "=": h = "write 'is' (or 'equals') to compare, 'Set ... to ...' to assign"
        Case "<": h = "write 'is less than'"
        Case ">": h = "write 'is greater than'"
        Case "&": h = "write 'joined with' to combine text"
        Case "$": h = "write the plain number (or 'Format ... as currency.' for display)"
        Case ";": h = "end the sentence with '.' and start a new one"
        Case "(": h = "a VLA form must begin its row - start the cell with '(' to write VLA, or say it in words"
        Case ")": h = "a ')' without its form - VLA rows start with '(' at the first character"
        Case "'", ChrW$(8217): h = "for a spaced sheet name, attach the reference: 'Q1 Data'!A1 - otherwise use straight double quotes "" for text, and names without apostrophes"
        Case ChrW$(8211), ChrW$(8212): h = "use a plain hyphen - (Word may have autocorrected it)"
        Case Else: h = "remove it, or put it inside ""quotes"" if it is part of a text value"
    End Select
    StrayCharHint = " - " & h
End Function

Public Function NumberWord(ByVal w As String) As String
    Select Case w
        Case "zero": NumberWord = "0"
        Case "one": NumberWord = "1"
        Case "two": NumberWord = "2"
        Case "three": NumberWord = "3"
        Case "four": NumberWord = "4"
        Case "five": NumberWord = "5"
        Case "six": NumberWord = "6"
        Case "seven": NumberWord = "7"
        Case "eight": NumberWord = "8"
        Case "nine": NumberWord = "9"
        Case "ten": NumberWord = "10"
        Case "eleven": NumberWord = "11"
        Case "twelve": NumberWord = "12"
        ' G8: thirteen..twenty - the rest of the bounded cardinal range.
        ' Safe at the tokenizer, unconditionally, unlike ordinals below:
        ' none of these thirteen words is claimed anywhere else in the
        ' vocab or the engine (checked before adding them), so there is
        ' no shipped idiom or pattern-literal for a blanket rewrite to
        ' collide with. Beyond twenty stays digits by doctrine.
        Case "thirteen": NumberWord = "13"
        Case "fourteen": NumberWord = "14"
        Case "fifteen": NumberWord = "15"
        Case "sixteen": NumberWord = "16"
        Case "seventeen": NumberWord = "17"
        Case "eighteen": NumberWord = "18"
        Case "nineteen": NumberWord = "19"
        Case "twenty": NumberWord = "20"
        Case Else: NumberWord = w
    End Select
End Function

' G8: ordinal words (first..twentieth), recognized only inside the
' expr grammar (ParsePrimCore), never at the tokenizer. "first" is
' safe here despite being reserved by mFnOf's "first of X" idiom: the
' call site only reaches this table after that check has already
' failed (either the word isn't "first"/"last", or it wasn't followed
' by "of"), so "first of found-items" is never at risk. Returns "" for
' anything unrecognized so the call site's Len() check doubles as the
' presence test.
Public Function OrdinalWord(ByVal w As String) As String
    Select Case w
        Case "first": OrdinalWord = "1"
        Case "second": OrdinalWord = "2"
        Case "third": OrdinalWord = "3"
        Case "fourth": OrdinalWord = "4"
        Case "fifth": OrdinalWord = "5"
        Case "sixth": OrdinalWord = "6"
        Case "seventh": OrdinalWord = "7"
        Case "eighth": OrdinalWord = "8"
        Case "ninth": OrdinalWord = "9"
        Case "tenth": OrdinalWord = "10"
        Case "eleventh": OrdinalWord = "11"
        Case "twelfth": OrdinalWord = "12"
        Case "thirteenth": OrdinalWord = "13"
        Case "fourteenth": OrdinalWord = "14"
        Case "fifteenth": OrdinalWord = "15"
        Case "sixteenth": OrdinalWord = "16"
        Case "seventeenth": OrdinalWord = "17"
        Case "eighteenth": OrdinalWord = "18"
        Case "nineteenth": OrdinalWord = "19"
        Case "twentieth": OrdinalWord = "20"
        Case Else: OrdinalWord = ""
    End Select
End Function

' METAVOCAB: the six recognized directive shapes, unchanged from the
' original loader - EXCEPT the final Else, which used to raise
' immediately and now tries one thing first (see ExpandVocabMacroCall).
' allowExpansion is False for anything already reached BY that
' expansion (a generator's own spliced output is fully expanded
' already - trying to expand it again would be pointless at best,
' and Substitute/ExpandMacros give no guarantee a second attempt is
' even a no-op) - only a form read straight from the file gets the
' one real attempt. depth guards nested (begin ...) splicing only (a
' generator whose own body contains another (begin ...), or one
' generator calling another) - it is NOT a macro-expansion budget;
' VLA.VlaExpandText's own toFixpoint:=True already expands to a
' stable form in ONE call, so nothing here ever re-expands anything.
' rowTag rides alongside startLine, inherited unchanged through every
' splice exactly the same way startLine already is, EXCEPT at-row
' itself (below), which is the only place it ever changes.
' LX5.1: registers one (keyword-alias "surface" "canonical") entry -
' surface is the language's own spelling, canonical must be one of the
' closed set of English structural words CanonicalizeStructuralWords'
' consumers already hardcode (if/repeat/until/times/while/for/each/in/
' stop/loop/increase/decrease/add/get/give/try/that/fails/when/clicked/
' create/define/done/otherwise/and/or/is/equals/greater/less/at/least/
' most/does/not/plus/minus/divided/joined/with/percent/of). Re-declaring
' the same surface word overwrites its earlier mapping rather than
' erroring, the same last-write-wins shape AddKeyed already uses
' elsewhere in this file, so a later phrasebook may correct an earlier
' one's choice without ceremony.
Public Sub RegisterKeywordAlias(ByVal surface As String, ByVal canonical As String)
    Dim k As String
    k = VLA_Identity.Fold(surface)
    On Error Resume Next
    mKeywordAlias.Remove k
    On Error GoTo 0
    mKeywordAlias.Add VLA_Identity.Fold(canonical), k
End Sub

Public Function IsColorWord(ByVal s As String) As Boolean
    Select Case s
        Case "black", "white", "red", "green", "blue", "yellow", "magenta", "cyan"
            IsColorWord = True
    End Select
End Function

Public Function SlotDesc(ByVal cat As String) As String
    Select Case cat
        Case "name", "var": SlotDesc = "a name (one word, like total)"
        Case "text": SlotDesc = "a reference (like B2 or ""Sheet1"")"
        Case "expr": SlotDesc = "a value (like 5, ""text"", or total plus 1)"
        Case "cond": SlotDesc = "a condition (like total is greater than 5)"
        ' G2: the typed reference slots teach their shapes; quotes are
        ' always the named-thing escape hatch, so each says so.
        Case "range": SlotDesc = "a range (like A1:C50, B2, or Data!A1:B10 - quotes for a named range)"
        Case "cell": SlotDesc = "a cell (like B2 or Data!B2 - quotes for a named cell)"
        Case "column": SlotDesc = "a column letter (like C or AA)"
        Case "sheet": SlotDesc = "a sheet name (like Data)"
        Case "color": SlotDesc = "a color (like red or yellow - quotes for a code like ""#FF69B4"")"
        ' G-PATH: quotes aren't an escape hatch here, they're the
        ' normal case - a bare word is only accepted as a variable
        ' name (report-path), never as literal path text.
        Case "path": SlotDesc = "a quoted path (like ""C:\Reports\file.xlsx"" - always quoted, a bare period would end the sentence) or a variable name"
        ' G6: list-valued slots - one comma-separated description per
        ' item shape, Oxford comma required (the example text teaches
        ' it by using it).
        Case "text-list": SlotDesc = "a list of names, comma-separated (like Region, Product, and Date)"
        Case "range-list": SlotDesc = "a list of ranges, comma-separated (like A1:C50, B2, and Data!A1:B10)"
        Case "cell-list": SlotDesc = "a list of cells, comma-separated (like B2, C3, and D4)"
        Case "column-list": SlotDesc = "a list of column letters, comma-separated (like B, C, and F)"
        Case "sheet-list": SlotDesc = "a list of sheet names, comma-separated (like Data, Summary, and Notes)"
        Case "color-list": SlotDesc = "a list of colors, comma-separated (like red, green, and blue)"
        Case Else: SlotDesc = cat
    End Select
End Function
