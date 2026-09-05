Attribute VB_Name = "VLA_Lint"
Option Explicit
Public Const VLA_LINT_VERSION As String = "LX2.0"
' LX2.0: this module's 2 raw Err.Raise refusal sites now route through
' VLA_Messages.RaiseMsg with a stable id - SD-2/LX.2's first migrated
' batch. Rendered text and Err.Number/Err.Source are unchanged.
' VLALINT.0: a VBA port of tools/vla_lint.pl - the house-style pretty-
' printer for .vla (S-expression) source text, promoted from a Perl
' script into native VBA so linting can be a standing, built-in default
' (every future .vla-writing code path can call VlaLintFormat directly)
' rather than a manual external step nothing in the runtime can invoke.
' House rule, seeded from the Perl tool but since revised (owner-
' requested, across three passes this session):
'   - a form's SIGNATURE ((defmacro (name params) ...)) always stays on
'     one line, regardless of arity - untouched by this pass.
'   - any form renders FLAT (one line) if the whole thing fits within
'     MAX_WIDTH columns counting from its own starting column - argument
'     COUNT never decides this on its own, only measured width does.
'     Two exceptions force vertical regardless of width: the node
'     ITSELF is one of the six real-directive shapes (english-vla,
'     english-vla-override, test-success, test-fail, defmacro,
'     english-function - the same six VLA_English.bas's own
'     DispatchVocabForm recognizes) - "the pairing should always be
'     visibly separate," English paired with VLA/expected-result; or a
'     NON-FIRST argument is one of those six - it must start its own
'     line, never get inlined after a sibling (the FIRST argument is
'     exempt from this second rule; it can still render flat with its
'     parent if the parent itself fits and isn't otherwise forced).
'   - a form that doesn't fit flat (or is forced) verticalizes in one of
'     two shapes, chosen by the functor - see PpForm's own call site for
'     the exact decision:
'       - a directive-shaped node, or an ordinary form whose functor is
'         longer than MAX_INLINE_FUNCTOR (a generator call like
'         table-property-family, tablespec-row, ...): PpVerticalize's
'         layout, the same shape PpDefmacro already used - the functor
'         stands alone on its own line, EVERY argument (including the
'         first) gets its own line below it, indented 4 past the form's
'         own starting column. A long functor with its first argument
'         fused to it wastes "a giant square of space" (owner's own
'         framing); this is the fix for that, specifically.
'       - anything else (a short functor - VLA's own control-flow and
'         definition keywords: if, begin, while, quote-if, deflambda,
'         sub, on-error, ...): PpVerticalizeInline's layout, PpForm's
'         original shape - the functor and its first argument share the
'         opening line, every remaining argument aligns beneath the
'         first argument's own column. Reads as ordinary Lisp at this
'         length; a defmacro-style full split here would waste vertical
'         space on nothing, not fix anything.
' Applies recursively - a nested argument that doesn't itself fit flat
' verticalizes in turn, wherever it lands.
'
' Two owner-requested house rules, VLA_Lint.bas's own addition (no
' Perl-tool ancestor - archive/vla_lint.pl is retired and was never
' updated to match):
'   - a bare TOP-LEVEL (defmacro ...) form is always followed by a
'     blank line before whatever comes next (another form, a comment,
'     ...) - the closing paren must never sit directly above the next
'     line. Only checked at top level, one segment lookahead; a
'     defmacro already followed by a blank verbatim line is left alone
'     (not duplicated).
'   - the formatted text always ends with exactly two blank lines (three
'     trailing line breaks) - opening the file and going to its very
'     end should never require pressing Enter twice before typing a
'     new form.
'
' Reuse, not reinvention: VLA.VlaReadForms parses an extracted form's
' raw text into the same Collection-of-Collections/Variants tree this
' whole codebase already uses (atom = Not IsObject, list = IsObject);
' VLA.VlaWriteForm IS the "flat candidate text" the width check needs,
' already correct (string quote-escaping included) - no new tokenizer,
' no new flat-renderer. What VLA.bas's reader CANNOT supply - and what
' this module must own itself - is the comment/blank-line-preserving
' segmenter: VLA.Tokenize discards comment text entirely, and
' VlaReadFormsWithLines exposes only a start line per top-level form,
' never an end line or the gaps between forms. SplitSegments (below) is
' a fresh port of the Perl tool's own Pass 1 (paren-depth + string-
' escape + ";"-to-end-of-line scan over raw lines) for exactly that
' reason. VLA.bas already has a different, simpler public pretty-
' printer (VlaFormat/WritePretty, 90-col, no defmacro/directive
' awareness, drops all comments on round-trip) - deliberately left
' untouched: it has no real callers today, and comment-dropping makes
' it unsuitable for this module's actual job regardless.
'
' Known inherited behavior, not a bug to fix here: a ";" comment
' sitting INSIDE a multi-line form's own parens (not at top level,
' between forms) is swallowed into that form's raw-text blob by
' SplitSegments and then silently dropped by VLA.Tokenize's own
' comment handling - it never reaches the pretty-printer and is never
' re-emitted. Only comments/blanks that are their own TOP-LEVEL line
' (before or after a form, at paren-depth 0) survive. This is
' tools/vla_lint.pl's own already-installed, already-verified behavior;
' this port reproduces it exactly, on purpose - see TestVlaLint's own
' pin documenting it (VLA_Tests.bas).
'
' Documented latent divergence from the Perl tool, safe today, flagged
' so a future regression is discoverable: Perl's string tokens are
' opaque raw substrings, reproduced byte-for-byte on output. VLA.Tokenize/
' WriteDatum DECODE \"/\\ on read and RE-ENCODE on write - not
' equivalent for a string literal containing a lone, unpaired backslash
' (e.g. "C:\Users\foo") - Perl reproduces it unchanged; this reuse
' doubles the lone backslash on re-emit. Verified empirically against
' today's corpus (scripts/english.vla, scripts/prelude.vla): zero lone
' backslashes inside any string literal in either file (english.vla's
' own one lone backslash sits inside a ";" comment, never touched by
' either pipeline). A future string literal with an unpaired backslash
' would format differently than tools/vla_lint.pl would have - nothing
' here detects that case; this note is the only guard.

Private Const MAX_WIDTH As Long = 100
' Owner-requested boundary (this session, third pass): a form that
' verticalizes uses PpVerticalizeInline's shape (functor + first
' argument share the opening line) when the functor is no longer than
' this; a functor longer than it falls back to PpVerticalize's own full
' split instead. Chosen from the real corpus, not guessed: every
' control-flow/definition keyword that must stay inline for this to
' read as ordinary Lisp - if(2), dim(3), sub(3), try(3), for(3),
' goto(4), set!(4), begin(5), while(5), label(5), quote(5), lambda(6),
' include(7), quote-if(8), for-each(8), obj-set!(8), deflambda(9),
' on-error(9), debug-print(11) - sits at or under 11; every generator-
' style call this rule exists FOR - tablespec-row(13),
' table-property-family(22) - sits at 13 or above. 12 sits cleanly in
' the gap between the two groups.
Private Const MAX_INLINE_FUNCTOR As Long = 12

' =====================================================================
'  V8-style string builder - this module's own copy of VLA.bas's
'  SbAdd/SbText (Rule 12: Privates are invisible cross-module, and a
'  hot helper should not pay a cross-module call anyway - VLA_English.bas
'  already carries an identical twin for the same reason).
' =====================================================================
Private Sub SbAdd(ByRef buf As String, ByRef used As Long, ByVal s As String)
    Dim n As Long
    n = Len(s)
    If n = 0 Then Exit Sub
    If used + n > Len(buf) Then
        Dim cap As Long
        cap = Len(buf)
        If cap < 64 Then cap = 64
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

' =====================================================================
'  Pass 1: segmenter - splits raw text into "form" segments (a run of
'  lines starting at paren-depth 0 with "(" as the first non-blank
'  char, continuing until depth returns to 0, tracking string-literal
'  escapes and ";" end-of-line comments while scanning) and "verbatim"
'  segments (everything else - comments, blank lines - one line each,
'  passed through byte-for-byte). Ported line-for-line from
'  tools/vla_lint.pl's own Pass 1 - see this module's own header note
'  on why VLA.bas's reader cannot supply this.
' =====================================================================
Private Function FirstNonBlankIsParen(ByVal ln As String) As Boolean
    Dim i As Long
    For i = 1 To Len(ln)
        Dim c As String
        c = Mid$(ln, i, 1)
        If c <> " " And c <> vbTab Then
            FirstNonBlankIsParen = (c = "(")
            Exit Function
        End If
    Next
End Function

Private Sub SplitSegments(ByVal text As String, _
                           ByRef segTypes() As String, ByRef segTexts() As String, ByRef segCount As Long)
    segCount = 0
    If Len(text) = 0 Then Exit Sub

    Dim lines() As String
    lines = Split(Replace(text, vbCrLf, vbLf), vbLf)
    Dim lastIdx As Long
    lastIdx = UBound(lines)

    Dim i As Long
    i = LBound(lines)
    Do While i <= lastIdx
        Dim ln As String
        ln = lines(i)
        If FirstNonBlankIsParen(ln) Then
            Dim depth As Long, started As Boolean
            depth = 0
            started = False
            Dim buf As String
            buf = ""
            Dim firstLineInForm As Boolean
            firstLineInForm = True

            Do While i <= lastIdx
                Dim l As String
                l = lines(i)
                If firstLineInForm Then
                    buf = l
                    firstLineInForm = False
                Else
                    buf = buf & vbLf & l
                End If

                Dim inQuote As Boolean, esc As Boolean
                inQuote = False
                esc = False
                Dim ci As Long
                For ci = 1 To Len(l)
                    Dim c As String
                    c = Mid$(l, ci, 1)
                    If inQuote Then
                        If esc Then
                            esc = False
                        ElseIf c = "\" Then
                            esc = True
                        ElseIf c = """" Then
                            inQuote = False
                        End If
                    Else
                        If c = ";" Then
                            Exit For
                        ElseIf c = """" Then
                            inQuote = True
                        ElseIf c = "(" Then
                            depth = depth + 1
                            started = True
                        ElseIf c = ")" Then
                            depth = depth - 1
                        End If
                    End If
                Next
                i = i + 1
                If started And depth = 0 Then Exit Do
            Loop

            segCount = segCount + 1
            ReDim Preserve segTypes(1 To segCount)
            ReDim Preserve segTexts(1 To segCount)
            segTypes(segCount) = "form"
            segTexts(segCount) = buf
        Else
            segCount = segCount + 1
            ReDim Preserve segTypes(1 To segCount)
            ReDim Preserve segTexts(1 To segCount)
            segTypes(segCount) = "verbatim"
            segTexts(segCount) = ln
            i = i + 1
        End If
    Loop
End Sub

' =====================================================================
'  Shape predicates
' =====================================================================

' A (defmacro (name params) ...) shaped list - recognized wherever it
' occurs, nested or top-level (a generator's own body may define new
' defmacros INSIDE a (begin ...)). elems(2) (the docstring slot) is
' deliberately NOT checked for shape - a generator's own unexpanded
' template may carry a bare placeholder symbol there instead of a real
' string, and this pass runs on source text, never expanded output.
Private Function IsDefmacroShaped(ByVal node As Variant) As Boolean
    If Not IsObject(node) Then Exit Function
    Dim flc As Collection
    Set flc = node
    If flc.Count < 3 Then Exit Function
    If IsObject(flc.Item(1)) Then Exit Function
    If StrComp(CStr(flc.Item(1)), "defmacro", vbBinaryCompare) <> 0 Then Exit Function
    If Not IsObject(flc.Item(2)) Then Exit Function
    IsDefmacroShaped = True
End Function

' Owner-requested: a defmacro OR deflambda sitting as one BODY ELEMENT
' of an enclosing form (top-level defmacro already got this treatment
' from VlaLintFormat's own segment-level rule; this is the same rule
' applied WITHIN a rendered body list too - alonzo.vla's own
' register-bricks, a sub whose body is five back-to-back deflambdas,
' is the case that found the gap). Head-only check, deliberately not as
' strict as IsDefmacroShaped's own arity/shape validation - this
' decides BLANK-LINE SPACING only, never which layout renders the form
' itself (deflambda has no dedicated PpForm branch of its own; it
' renders via the ordinary PpVerticalize/PpVerticalizeInline path like
' any other call, its own shape not being defmacro's - name and params
' are two separate leading items, not one combined signature list).
Private Function IntroducesDefinition(ByVal node As Variant) As Boolean
    If Not IsObject(node) Then Exit Function
    Dim flc As Collection
    Set flc = node
    If flc.Count < 1 Then Exit Function
    If IsObject(flc.Item(1)) Then Exit Function
    Dim head As String
    head = CStr(flc.Item(1))
    IntroducesDefinition = (head = "defmacro" Or head = "deflambda")
End Function

' Directive-shaped: the real-directive head shapes VLA_English.bas's own
' DispatchVocabForm recognizes - defmacro, english-vla (any "-vla"
' suffix - the source language rides in the head, F.13), english-vla-
' override (any "-vla-override" suffix, checked separately since its
' own tail does NOT end in plain "-vla"), test-success, test-fail,
' english-function (any "-function" suffix), keyword-alias (LX5.1 - a
' language file's own control-flow vocabulary). Used to decide (a)
' whether a NON-FIRST argument of some enclosing form must start its
' own line, and (b) whether a form is itself one that always splits its
' first two arguments - not to control how the directive's own
' internals render, which stays ordinary PpForm recursion for every one
' of these except defmacro (which alone gets PpDefmacro's own dedicated
' layout, checked separately, earlier).
Private Function IsDirectiveShaped(ByVal node As Variant) As Boolean
    If Not IsObject(node) Then Exit Function
    Dim flc As Collection
    Set flc = node
    If flc.Count < 1 Then Exit Function
    If IsObject(flc.Item(1)) Then Exit Function
    Dim head As String
    head = CStr(flc.Item(1))
    If head = "defmacro" Then IsDirectiveShaped = True: Exit Function
    If head = "test-success" Then IsDirectiveShaped = True: Exit Function
    If head = "test-fail" Then IsDirectiveShaped = True: Exit Function
    If head = "keyword-alias" Then IsDirectiveShaped = True: Exit Function
    If Len(head) > 13 And Right$(head, 13) = "-vla-override" Then IsDirectiveShaped = True: Exit Function
    If Len(head) > 4 And Right$(head, 4) = "-vla" Then IsDirectiveShaped = True: Exit Function
    If Len(head) > 9 And Right$(head, 9) = "-function" Then IsDirectiveShaped = True: Exit Function
End Function

' True if node, or any descendant at any depth, is defmacro-shaped -
' gates trusting VLA.VlaWriteForm's flat rendering as a legitimate
' width-check candidate. VlaWriteForm is a generic, unconditional flat
' renderer with no defmacro awareness at all; PpForm ALWAYS forces
' defmacro multi-line via PpDefmacro (see there), so a flat_text that
' contains one anywhere describes a rendering PpForm could never
' actually produce - must be rejected before it's ever compared against
' MAX_WIDTH, not just when the node itself is the offender.
Private Function ContainsDefmacro(ByVal node As Variant) As Boolean
    If Not IsObject(node) Then Exit Function
    If IsDefmacroShaped(node) Then
        ContainsDefmacro = True
        Exit Function
    End If
    Dim flc As Collection
    Set flc = node
    Dim child As Variant
    For Each child In flc
        If ContainsDefmacro(child) Then
            ContainsDefmacro = True
            Exit Function
        End If
    Next
End Function

' =====================================================================
'  Line-array helpers - VBA Collection has no replace-by-index, and
'  Perl's own pp() mutates the LAST element of its own rest-lines
'  return value in place (e.g. appending " " & $x to continue a
'  sibling onto the same line) - so rest-lines here is a resizable
'  1-based String() array instead of a Collection, specifically so
'  restLines(restCount) = restLines(restCount) & ... works. The VALUES
'  produced are identical to the Perl tool's own; only this plumbing
'  shape differs, forced by the language.
' =====================================================================
Private Sub AppendLine(ByRef arr() As String, ByRef cnt As Long, ByVal s As String)
    cnt = cnt + 1
    ReDim Preserve arr(1 To cnt)
    arr(cnt) = s
End Sub

Private Sub AppendLines(ByRef dest() As String, ByRef destCount As Long, ByRef src() As String, ByVal srcCount As Long)
    Dim i As Long
    For i = 1 To srcCount
        AppendLine dest, destCount, src(i)
    Next
End Sub

' =====================================================================
'  The pretty-printer - a 1:1 port of tools/vla_lint.pl's own pp(),
'  1-based flc.Item(k) standing in for Perl's 0-based elems[k-1]:
'  Item(1) = functor, Item(2) = first argument, Item(3.. ) = every
'  argument after the first ("non-first", in every comment above).
' =====================================================================
Private Sub PpForm(ByVal node As Variant, ByVal col As Long, _
                    ByRef firstLine As String, ByRef restLines() As String, ByRef restCount As Long, _
                    ByRef endCol As Long)
    restCount = 0

    If Not IsObject(node) Then
        ' VLA.VlaWriteForm, not bare CStr - a string-literal atom's
        ' STORED form is only a leading quote MARKER plus its decoded
        ' content (no trailing quote at all - WriteDatum reconstructs
        ' both quotes from that marker on output). CStr(node) would
        ' emit the raw internal storage form verbatim, silently
        ' dropping the closing quote on every quoted-string argument
        ' rendered through this path - PpDefmacro's own docstring line
        ' already got this right (see its own comment); this is the
        ' same fix, generalized to every atom this function renders.
        firstLine = VLA.VlaWriteForm(node)
        endCol = col + Len(firstLine)
        Exit Sub
    End If

    Dim flc As Collection
    Set flc = node
    If flc.Count = 0 Then
        firstLine = "()"
        endCol = col + 2
        Exit Sub
    End If

    If IsDefmacroShaped(flc) Then
        PpDefmacro flc, col, firstLine, restLines, restCount, endCol
        Exit Sub
    End If

    Dim nargs As Long
    nargs = flc.Count - 1

    If nargs = 0 Then
        ' A bare functor call, e.g. (stop) - always flat; still goes
        ' through PpForm on the functor itself so a (hypothetically)
        ' multi-line compound functor still renders correctly.
        Dim ff As String, fr() As String, frc As Long, fec As Long
        PpForm flc.Item(1), col + 1, ff, fr, frc, fec
        If frc > 0 Then
            AppendLines restLines, restCount, fr, frc
            restLines(restCount) = restLines(restCount) & ")"
            firstLine = "(" & ff
        Else
            firstLine = "(" & ff & ")"
        End If
        endCol = fec + 1
        Exit Sub
    End If

    ' Should this form verticalize even if it WOULD fit flat? Two cases:
    '   - the node itself is one of the non-defmacro directive shapes
    '     (english-vla/-vla-override/test-success/test-fail/
    '     english-function/keyword-alias [LX5.1]) - "the pairing should
    '     always be visibly separate," regardless of width.
    '   - a non-first argument (Item(3) onward - Item(2), the first
    '     argument, is exempt) is directive-shaped (any of them) -
    '     it must start its own line, never get inlined after a sibling.
    Dim mustVerticalize As Boolean
    mustVerticalize = IsDirectiveShaped(flc)
    If Not mustVerticalize Then
        Dim k As Long
        For k = 3 To flc.Count
            If IsDirectiveShaped(flc.Item(k)) Then
                mustVerticalize = True
                Exit For
            End If
        Next
    End If

    ' MAX_WIDTH gate: try the WHOLE form flat first, regardless of raw
    ' argument count - skipped entirely when forced vertical above, or
    ' when a defmacro sits anywhere in the subtree (PpDefmacro's own
    ' layout is unconditionally multi-line, so no flat rendering of it
    ' could ever be one PpForm would actually produce).
    If Not mustVerticalize And Not ContainsDefmacro(flc) Then
        Dim flatTxt As String
        flatTxt = VLA.VlaWriteForm(flc)
        If col + Len(flatTxt) <= MAX_WIDTH Then
            firstLine = flatTxt
            endCol = col + Len(flatTxt)
            Exit Sub
        End If
    End If

    ' Doesn't fit flat (or forced) - which vertical shape? Owner-
    ' requested boundary, third pass on this same question: a single
    ' "everything splits" rule for every verticalized form (the second
    ' pass) turned out to overreach - right for a generator call like
    ' table-property-family (a long functor name, first argument fused
    ' to it, "a giant square of wasted space"), wrong for VLA's own
    ' control-flow/definition keywords (if, begin, quote-if, while,
    ' deflambda, sub, ...), which read as alien Lisp once their (short)
    ' functor stands alone with nothing on its own line. The node
    ' itself being directive-shaped ALWAYS gets PpVerticalize's full
    ' split, unconditional on length - the six real-directive shapes'
    ' own standing rule, unaffected by this decision. Otherwise the
    ' functor's own length decides: at or under MAX_INLINE_FUNCTOR,
    ' PpVerticalizeInline's shape (functor + first argument share the
    ' opening line, ordinary Lisp style); longer than it,
    ' PpVerticalize's full split, the same shape defmacro's own
    ' signature/docstring/body layout already uses in spirit.
    If IsDirectiveShaped(flc) Or Len(VLA.VlaWriteForm(flc.Item(1))) > MAX_INLINE_FUNCTOR Then
        PpVerticalize flc, col, firstLine, restLines, restCount, endCol
    Else
        PpVerticalizeInline flc, col, firstLine, restLines, restCount, endCol
    End If
End Sub

' defmacro's own layout: "(defmacro" alone on its own line; signature
' stays on ONE line always (JoinSig, never run through PpForm's own
' recursive rules - english.vla's own convention); docstring on its own
' line; each body form PpForm'd in turn, indented 4 past wherever this
' defmacro form itself starts (so a NESTED defmacro, e.g. inside a
' generator's own (begin ...), indents 4 past ITS OWN column, not a
' fixed 4 from the file's own left margin).
Private Sub PpDefmacro(ByVal node As Variant, ByVal col As Long, _
                        ByRef firstLine As String, ByRef restLines() As String, ByRef restCount As Long, _
                        ByRef endCol As Long)
    Dim flc As Collection
    Set flc = node
    Dim bodyCol As Long
    bodyCol = col + 4
    restCount = 0

    AppendLine restLines, restCount, Space$(bodyCol) & JoinSig(flc.Item(2))
    ' The docstring slot: flc.Item(3) is already a DECODED atom (VLA.
    ' Tokenize strips quotes/un-escapes on read) - CStr() alone would
    ' emit it WITHOUT its surrounding quotes. VLA.VlaWriteForm re-
    ' encodes exactly like every other atom, and returns an unexpanded
    ' generator placeholder symbol (a bare word, not a string) here
    ' unchanged - correct for both cases.
    AppendLine restLines, restCount, Space$(bodyCol) & VLA.VlaWriteForm(flc.Item(3))

    Dim k As Long
    For k = 4 To flc.Count
        Dim bf As String, br() As String, brc As Long, bec As Long
        PpForm flc.Item(k), bodyCol, bf, br, brc, bec
        AppendLine restLines, restCount, Space$(bodyCol) & bf
        AppendLines restLines, restCount, br, brc
        ' A nested defmacro/deflambda body form gets a blank line after
        ' it too, same as a top-level one (VlaLintFormat's own rule) -
        ' IntroducesDefinition's own header has the full reasoning.
        If IntroducesDefinition(flc.Item(k)) And k < flc.Count Then AppendLine restLines, restCount, ""
    Next

    firstLine = "(defmacro"
    restLines(restCount) = restLines(restCount) & ")"
    endCol = Len(restLines(restCount))
End Sub

' The house vertical layout, used by every form that verticalizes -
' too wide to fit flat, forced (directive-shaped, or containing a
' directive-shaped non-first argument), or both. PpDefmacro's own
' shape, generalized from a fixed signature/docstring/body triple to N
' ordinary arguments: the functor stands alone on its own line; EVERY
' argument, including the first, gets its own line below it, indented
' 4 past wherever this form itself starts (same "past ITS OWN column,
' not the file's left margin" rule PpDefmacro already follows, for the
' identical reason - a form nested inside a generator's own
' (begin ...) must indent relative to where IT starts, not absolute).
' Owner-requested correction, this session's second pass: the first
' version of this rule (then named PpDirective) only applied to the
' directive shapes, and PpForm's OWN generic verticalize branch kept an
' ordinary call's first argument fused to the functor - correct for
' the real directives (english-vla, test-success, ...) but, on a real
' generator like table-property-family, a long functor name left a
' "giant square of wasted space" to the left of every argument
' (owner's own framing). One shape now, for every verticalized form.
' The functor is rendered through PpForm too, not assumed to be a bare
' atom - directive-shaped nodes always are (IsDirectiveShaped rejects
' an object/compound head before matching any suffix), but an ORDINARY
' call reaching this Sub is not guaranteed to be: english.vla's own
' ANTONYM-SWEEP generators produce forms like
' ((symbol "un" name) param), a COMPOUND functor. In practice a
' compound functor here is always itself a short form (never actually
' multi-line) - but the fallback stays correct even if one someday is,
' the same guarantee PpForm's own retired generic-verticalize branch
' already made.
' Never called with fewer than two total elements (one argument): a
' node with zero arguments would already have been routed through
' PpForm's own "bare functor call" branch before this is ever reached.
Private Sub PpVerticalize(ByVal node As Variant, ByVal col As Long, _
                           ByRef firstLine As String, ByRef restLines() As String, ByRef restCount As Long, _
                           ByRef endCol As Long)
    Dim flc As Collection
    Set flc = node
    Dim bodyCol As Long
    bodyCol = col + 4
    restCount = 0

    Dim k As Long
    For k = 2 To flc.Count
        Dim bf As String, br() As String, brc As Long, bec As Long
        PpForm flc.Item(k), bodyCol, bf, br, brc, bec
        AppendLine restLines, restCount, Space$(bodyCol) & bf
        AppendLines restLines, restCount, br, brc
        ' A nested defmacro/deflambda argument gets a blank line after
        ' it too, same as a top-level one (VlaLintFormat's own rule) -
        ' IntroducesDefinition's own header has the full reasoning.
        If IntroducesDefinition(flc.Item(k)) And k < flc.Count Then AppendLine restLines, restCount, ""
    Next

    Dim functorFirst As String, functorRest() As String, functorRestCount As Long, functorEndCol As Long
    PpForm flc.Item(1), col + 1, functorFirst, functorRest, functorRestCount, functorEndCol
    firstLine = "(" & functorFirst
    If functorRestCount > 0 Then
        ' The (never yet exercised in practice) multi-line-functor case:
        ' its own continuation lines belong BEFORE the argument lines
        ' already built above, so merge in that order.
        Dim merged() As String, mergedCount As Long
        AppendLines merged, mergedCount, functorRest, functorRestCount
        AppendLines merged, mergedCount, restLines, restCount
        restLines = merged
        restCount = mergedCount
    End If

    restLines(restCount) = restLines(restCount) & ")"
    endCol = Len(restLines(restCount))
End Sub

' The inline vertical layout - PpForm's own original shape, restored as
' its own named Sub in this session's third pass (see PpForm's own call
' site for the length-threshold decision that routes here): the functor
' and its FIRST argument share the opening line; every remaining
' argument goes on its own line, aligned directly beneath the first
' argument's own opening column. Reads as ordinary Lisp for a short
' functor - if, begin, while, quote-if, deflambda, sub, on-error, and
' the rest of VLA's own control-flow/definition vocabulary - which is
' exactly the set MAX_INLINE_FUNCTOR was measured against.
Private Sub PpVerticalizeInline(ByVal node As Variant, ByVal col As Long, _
                                 ByRef firstLine As String, ByRef restLines() As String, ByRef restCount As Long, _
                                 ByRef endCol As Long)
    Dim flc As Collection
    Set flc = node
    restCount = 0

    Dim functorFirst As String, functorRest() As String, functorRestCount As Long, functorEndCol As Long
    PpForm flc.Item(1), col + 1, functorFirst, functorRest, functorRestCount, functorEndCol
    firstLine = "(" & functorFirst
    AppendLines restLines, restCount, functorRest, functorRestCount

    Dim firstArgCol As Long
    firstArgCol = functorEndCol + 1
    Dim faFirst As String, faRest() As String, faRestCount As Long, lastEndCol As Long
    PpForm flc.Item(2), firstArgCol, faFirst, faRest, faRestCount, lastEndCol

    If restCount = 0 Then
        firstLine = firstLine & " " & faFirst
    Else
        restLines(restCount) = restLines(restCount) & " " & faFirst
    End If
    AppendLines restLines, restCount, faRest, faRestCount
    ' A defmacro/deflambda as the FIRST argument (a real corpus shape -
    ' table-property-family's own generated (begin (defmacro ...)
    ' (english-vla ...) (test-success ...)) puts a defmacro right here)
    ' gets a blank line after it too, whether it rendered flat or
    ' multi-line - IntroducesDefinition's own header has the full
    ' reasoning.
    If IntroducesDefinition(flc.Item(2)) And flc.Count > 2 Then AppendLine restLines, restCount, ""

    Dim m As Long
    For m = 3 To flc.Count
        Dim argFirst As String, argRest() As String, argRestCount As Long, argEndCol As Long
        PpForm flc.Item(m), firstArgCol, argFirst, argRest, argRestCount, argEndCol
        AppendLine restLines, restCount, Space$(firstArgCol) & argFirst
        AppendLines restLines, restCount, argRest, argRestCount
        If IntroducesDefinition(flc.Item(m)) And m < flc.Count Then AppendLine restLines, restCount, ""
        lastEndCol = argEndCol
    Next

    If restCount > 0 Then
        restLines(restCount) = restLines(restCount) & ")"
    Else
        firstLine = firstLine & ")"
    End If
    endCol = lastEndCol + 1
End Sub

' Signature stays on one line, verbatim structure - not run through
' PpForm's recursive rules, matching english.vla's own convention of
' always-one-line signatures. VLA.VlaWriteForm already recurses
' correctly at any depth and re-escapes strings correctly, a strict
' superset of the Perl tool's own flat, one-level-only join_sig - every
' signature in today's corpus is a flat, non-nested parameter list, so
' this produces byte-identical output today, and is also correct for a
' hypothetical future nested-list signature parameter a literal,
' one-level-only port would not be.
Private Function JoinSig(ByVal node As Variant) As String
    JoinSig = VLA.VlaWriteForm(node)
End Function

' Reformat one top-level form (any shape - defmacro, english-vla,
' test-success/test-fail, a generator call like
' (bool-antonym-family ...), ...). isDefmacro is an out param (the
' caller needs the shape to enforce the blank-line-after-defmacro house
' rule; re-deriving it from the formatted TEXT would be strictly harder
' than reading it off the already-parsed node once, here).
Private Function ReformatForm(ByVal rawText As String, ByRef isDefmacro As Boolean) As String
    Dim forms As Collection
    Set forms = VLA.VlaReadForms(rawText)
    If forms.Count <> 1 Then
        VLA_Messages.RaiseMsg "lint-trailing-tokens", "text", rawText
    End If
    If Not IsObject(forms.Item(1)) Then
        VLA_Messages.RaiseMsg "lint-not-a-list", "form", CStr(forms.Item(1))
    End If
    isDefmacro = IsDefmacroShaped(forms.Item(1))

    Dim firstLine As String, restLines() As String, restCount As Long, endCol As Long
    PpForm forms.Item(1), 0, firstLine, restLines, restCount, endCol

    Dim r As String
    r = firstLine
    Dim i As Long
    For i = 1 To restCount
        r = r & vbCrLf & restLines(i)
    Next
    ReformatForm = r
End Function

' =====================================================================
'  Public entry point: reformat a whole .vla source text into house
'  style. Comments and blank lines pass through byte-for-byte at their
'  original position; every top-level form is independently reformatted
'  via ReformatForm. Output uses vbCrLf throughout - this codebase's
'  own established convention for every generated text file (see
'  VLA_Build.bas's WriteTextFile), not a literal match to the Perl
'  tool's own bare "\n" - callers comparing against an on-disk file
'  (which may be LF-only) should normalize both sides first, the same
'  way EnglishToVla already does (VLA_English.bas).
' =====================================================================
Public Function VlaLintFormat(ByVal text As String) As String
    Dim segTypes() As String, segTexts() As String, segCount As Long
    SplitSegments text, segTypes, segTexts, segCount

    Dim buf As String, bufU As Long
    Dim i As Long
    For i = 1 To segCount
        If segTypes(i) = "verbatim" Then
            If bufU > 0 Then SbAdd buf, bufU, vbCrLf
            SbAdd buf, bufU, segTexts(i)
        Else
            Dim isDefmacro As Boolean
            Dim formatted As String
            formatted = ReformatForm(segTexts(i), isDefmacro)
            Dim formLines() As String
            formLines = Split(formatted, vbCrLf)
            Dim j As Long
            For j = LBound(formLines) To UBound(formLines)
                If bufU > 0 Then SbAdd buf, bufU, vbCrLf
                SbAdd buf, bufU, formLines(j)
            Next

            ' Owner house rule: a bare top-level defmacro is always
            ' followed by a blank line. One extra separator here is
            ' enough - the NEXT segment's own leading "If bufU > 0"
            ' separator (above, or the verbatim branch's own) supplies
            ' the second line break that actually makes the gap visible;
            ' adding an empty string here would be a no-op (SbAdd exits
            ' immediately on a zero-length s). Skipped when there is no
            ' next segment (end-of-file gets the trailing-blank-line
            ' rule below instead) or when the next segment is already a
            ' blank verbatim line (don't duplicate).
            If isDefmacro And i < segCount Then
                Dim nextIsBlank As Boolean
                nextIsBlank = (segTypes(i + 1) = "verbatim") And (Len(segTexts(i + 1)) = 0)
                If Not nextIsBlank Then SbAdd buf, bufU, vbCrLf
            End If
        End If
    Next

    ' Owner house rule: the file always ends with exactly two blank
    ' lines (three trailing line breaks) - appending a new form at the
    ' end never needs more than one press of Enter first. Strip
    ' whatever trailing breaks the loop above already emitted FIRST -
    ' a source file's own trailing blank lines each pass through as
    ' their own verbatim segment, so appending three more unconditionally
    ' would accumulate on every relint rather than normalize (caught
    ' live: a real corpus file with existing trailing blanks came back
    ' with SIX, not two - this file's own idempotence, the property
    ' VlaLintCheck's whole comparison depends on, was broken until this
    ' fix).
    Dim result As String
    result = SbText(buf, bufU)
    Do While Len(result) >= 2 And Right$(result, 2) = vbCrLf
        result = Left$(result, Len(result) - 2)
    Loop
    If Len(result) > 0 Then result = result & vbCrLf & vbCrLf & vbCrLf
    VlaLintFormat = result
End Function
