Attribute VB_Name = "VLA_Identity"
Option Explicit
Public Const VLA_IDENTITY_VERSION As String = "LX3.0"

' =====================================================================
'  VLA_Identity - decides SAMENESS. Exactly two comparison verbs, and
'  nothing else lives here; a junk drawer defeats the point the day
'  it is created (REBUILD.md SS4, "no VLA_Utilities").
'
'    Fold(s)         invariant ASCII case fold - decides IDENTITY.
'                     Dispatch keys, Collection keys, macro/rule
'                     names, DSL keywords ("macro:", "test:", "if",
'                     "otherwise" ...). Never varies with the host's
'                     Windows locale.
'    SameText(a, b)   locale text comparison - for things a USER
'                     TYPED and meant as words. Not currently called
'                     anywhere in the corpus; kept as the documented
'                     alternative so a future author reaches for this
'                     instead of writing a new LCase$ site.
'
'  LAYER:     0
'  MAY CALL:  (nothing - Layer 0 has no dependencies)
'  SHIPS:     add-in (VLA_ prefix; not injected into user workbooks)
'  PAYS INTO: LX.3 (this module IS the fix), SD-8 (identifiers folded
'             invariantly, identity never depends on locale), R6 (no
'             LCase$ outside this module - a lint can now grep for
'             zero hits), P-TOK (the fold is already single-scan),
'             EN.4, LX.6 (the non-ASCII identifier policy this module
'             makes answerable, since Fold only ever touches A-Z)
'  REASON:    VBA's LCase$/UCase$ are LOCALE-AWARE. Under a Turkish
'             Windows locale, LCase$("I") folds to a dotless "i", not
'             the ASCII "i" - so identifier identity, macro-name
'             dispatch, and vocabulary-keyword matching could silently
'             depend on which Windows locale happened to be running on
'             the user's machine. ~110 call sites decided identity
'             through LCase$ before this pass (VLA.bas, VLA_English.bas,
'             VLA_Runtime.bas, VLA_IDE.bas, VLA_Loader.bas,
'             VLA_DevRig.bas); every one now folds through here.
' =====================================================================

' Invariant ASCII case fold - decides IDENTITY. Only the 26 ASCII
' letters move (A-Z -> a-z); accented letters, punctuation, digits,
' and every other code point pass through untouched, so a
' non-ASCII identifier (LX.6) is never mangled by a fold meant for
' the language's own ASCII keyword set.
Public Function Fold(ByVal s As String) As String
    Dim n As Long, i As Long, c As Integer
    n = Len(s)
    If n = 0 Then Exit Function
    Dim buf As String
    buf = s
    For i = 1 To n
        c = AscW(Mid$(buf, i, 1))
        If c >= 65 And c <= 90 Then Mid$(buf, i, 1) = ChrW$(c + 32)
    Next i
    Fold = buf
End Function

' Locale text comparison - for user-facing text, never identity.
Public Function SameText(ByVal a As String, ByVal b As String) As Boolean
    SameText = (StrComp(a, b, vbTextCompare) = 0)
End Function
