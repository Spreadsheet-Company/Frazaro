Attribute VB_Name = "VLA_Browser"
Option Explicit
Public Const VLA_BROWSER_VERSION As String = "PORT.1"
' PORT.1: the pure, host-free translate contract - the API a non-VBA
' port (a browser page, a Tauri/Electron shell, any future substrate
' SUBSTRATE.md's watch-list ever needs) targets. Not a new engine: this
' module composes primitives VLA_English.bas and VLA.bas already own
' (EnglishResetGrammar, EnglishLoadVocabularyText, EnglishToVla,
' EnglishToVba, VlaSetPreludeOverride) and adds exactly one thing none
' of them offer on their own - a caller that never touches a file, never
' shows a MsgBox, and gets a refusal back as the same teaching text the
' product already generates, as a return value instead of a dialog.
'
' Host-free, confirmed rather than assumed: EnglishToVla/EnglishToVba
' were already pure string-in/string-out before this module existed (a
' `grep -c "ActiveSheet\|ActiveWorkbook\|Application\.\|Worksheets(\|
' ThisWorkbook\|\.Cells\|\.Range"` over VLA_English.bas returns 7 hits,
' every one a comment or Application.PathSeparator on the FILE-based
' EnglishTranslateToVla/ToVba's own path, none inside EnglishToVla/
' EnglishToVba/EnglishResetGrammar/EnglishLoadVocabularyText/EnsureInit
' themselves). The one real touch was VlaTranspile's own prelude load
' (PreludeMacros, VLA.bas) - fixed at its source by
' VlaSetPreludeOverride/VlaClearPreludeOverride, not routed around here.
' tools/check_translate_purity.ps1 (PORT.2) pins this so it stays true.
'
' Deliberately NOT a wrapper around EnglishTranslateToVla/ToVba (the
' existing file-based pair, VLA_English.bas ~L7540) - those own Dir$
' checks and a VlaShowError MsgBox in their failed: handler are exactly
' what a host-free caller must never reach, even transitively. This
' module is built from the same primitives, independently, so the
' existing pair's behavior is untouched by this file's existence.
'
' Signature note: ByRef out-parameters before a trailing ParamArray is
' the one ordering VBA requires here - Optional is what cannot precede
' a ParamArray (EnglishTranslateToVla's own header comment already
' found this the hard way); a required ByRef parameter has no such
' restriction and is the closer match to this module's contract anyway
' (every parameter here is required - a caller with nothing to translate
' has nothing to call this for).

' English -> VLA only. No prelude text needed: VlaTranspile (the VLA ->
' VBA stage) is the only consumer of the prelude, so this stage was
' already 100% host-free before this module existed - the only thing
' added is a refusal returned as text rather than shown as a dialog.
' vocabTexts: the text content of each phrasebook to load (english.vla,
' any others), in the order they should apply - the text equivalent of
' EnglishTranslateToVla's own vocabPaths ParamArray.
Public Function EnglishTranslateTextToVla(ByVal programText As String, _
                                           ByRef outVla As String, _
                                           ByRef refusal As String, _
                                           ParamArray vocabTexts() As Variant) As Boolean
    On Error GoTo failed
    outVla = ""
    refusal = ""

    EnglishResetGrammar
    Dim i As Long
    For i = LBound(vocabTexts) To UBound(vocabTexts)
        EnglishLoadVocabularyText CStr(vocabTexts(i)), "vocab-" & CStr(i - LBound(vocabTexts) + 1)
    Next

    outVla = EnglishToVla(programText)
    EnglishTranslateTextToVla = True
    Exit Function

failed:
    outVla = ""
    refusal = Err.Description
    EnglishTranslateTextToVla = False
End Function

' English -> VBA. preludeText is required (not Optional - VlaTranspile's
' prelude is core language infrastructure, never silently defaulted for
' a host-free caller who has no ThisWorkbook/embedded-sheet fallback to
' fall back to) and is threaded through VlaSetPreludeOverride/
' VlaClearPreludeOverride (VLA.bas, PORT.1) - cleared on every exit path,
' success or failure, so this call can never leave stale override state
' behind for an unrelated later caller (a live Excel session running
' both a normal host-based Compile and a pure test in the same session
' is exactly the case a leaked override would corrupt silently).
Public Function EnglishTranslateTextToVba(ByVal programText As String, _
                                           ByVal preludeText As String, _
                                           ByRef outVba As String, _
                                           ByRef refusal As String, _
                                           ParamArray vocabTexts() As Variant) As Boolean
    On Error GoTo failed
    outVba = ""
    refusal = ""

    EnglishResetGrammar
    Dim i As Long
    For i = LBound(vocabTexts) To UBound(vocabTexts)
        EnglishLoadVocabularyText CStr(vocabTexts(i)), "vocab-" & CStr(i - LBound(vocabTexts) + 1)
    Next

    VLA.VlaSetPreludeOverride preludeText
    outVba = EnglishToVba(programText)
    VLA.VlaClearPreludeOverride

    EnglishTranslateTextToVba = True
    Exit Function

failed:
    VLA.VlaClearPreludeOverride
    outVba = ""
    refusal = Err.Description
    EnglishTranslateTextToVba = False
End Function
