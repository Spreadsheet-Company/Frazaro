Attribute VB_Name = "VLA_Unify"
Option Explicit
Public Const VLA_UNIFY_VERSION As String = "PROLOG.3"
' PROLOG.3: EnvWalkInto promoted Private->Public (VLA_Prolog.bas's own
' query solver needs to read a free variable's own final bound value
' back out of a successful unification), and a new UnifyEnvClone -
' PROLOG.3's own conjunct solver must try every stored fact against the
' SAME query conjunct in turn, each attempt needing its own independent
' env copy so a failed or partial attempt can never leak a binding into
' its own sibling. Neither is a new unification CAPABILITY - both are
' plumbing this module's own first real consumer beyond G-RENDER needed
' to exist as Public.
'
' PROLOG.2: UnifyTwoWay - real unification, variables allowed on BOTH
' sides (the case PROLOG.1's UnifyOneWay structurally cannot serve: its
' bn/bv convention only ever RECORDS a binding, it never APPLIES one
' back into a later comparison, which is fine when concreteForm is
' always fully ground and wrong the moment it might itself carry an
' unbound variable). New code, not a wrapper calling UnifyOneWay twice -
' see this function's own header for exactly why that shortcut doesn't
' work, verified by construction rather than assumed.
'
' A DIFFERENT variable convention than PROLOG.1's own {name} bracket
' syntax, deliberately: {name} is English-template-specific (a rule's
' OWN vocabulary, never anything a stored fact/rule form would contain),
' while PROLOG.2 operates on real program terms, where the classical
' Prolog/Datalog convention already applies - a bare token whose first
' character is A-Z is a VARIABLE, everything else a CONSTANT
' (VLA_Datalog.bas's own IsVariableAtom, repeated here as IsVarAtom
' rather than shared across modules - a two-line rule, not worth a
' cross-module dependency for). UnifyOneWay is UNCHANGED and still the
' only mode UnifyForm calls; UnifyTwoWay is new, separate, additive.
'
' The substitution environment is two parallel Collections (envN/envT),
' PROLOG.1's own bn/bv convention generalized, NOT VLA_Runtime.
' VlaDictNew() - checked, not assumed: that dict's own CompareMode is
' vbTextCompare (case-insensitive) by construction (VlaDictNew,
' VLA_Runtime.bas), and while VLA_Datalog.bas's own module header
' describes SD-8's fold as "correct for a predicate/variable name," its
' ACTUAL code never once folds variable text - every VLA_Identity.Fold
' call in that file (grep-verified) targets a PREDICATE-shaped name
' (fact/rule/query keywords, predicate names), never an X/Y/Z argument.
' Two Prolog variables differing only in case past their first letter
' (Foo vs FOo) are genuinely different variables, not the same
' identifier spelled two ways, so folding them together would be a real
' silent-collision bug, not a harmless convenience - this module follows
' the verified CODE, not the header's own looser prose, and stays
' case-sensitive throughout (VBA's own default Option Compare Binary,
' unchanged everywhere else in this module already). Env sizes here (a
' handful of variables per unification call) make the parallel-
' Collection linear scan a non-concern; revisit only if PROLOG.4's real
' proof search is ever profiled and found slow, never speculatively.
'
' PROLOG.1: new module - the one-way match mode of BETA_ROADMAP2.md's own
' PROLOG entry ("the unifier is substrate, not engine-private"), hoisted
' out of VLA_English.bas's own UnifyForm (G-RENDER's FormSubstitute
' inverse) rather than duplicated. UnifyOneWay below IS UnifyForm's own
' four cases verbatim (a bare {name} slot binds a whole subform; a
' glued slot - make-{d}/xl{d}/format-as-{d} - matches a known
' prefix/suffix and binds what's between; a literal atom matches
' text-for-text; a list recurses element-by-element at equal length)
' plus one real correction, found by reading UnifyForm and its own
' RenderBoundLookup in full while scoping this rather than assumed: the
' original never checked repeated-variable consistency - if a template
' repeated the same {name} slot, each occurrence bound independently
' into bn/bv and RenderBoundLookup silently returned whichever
' occurrence it scanned first, discarding any mismatch a second
' occurrence would represent. Real unification requires repeated
' variables to bind to the SAME value; UnifyBind (below) now enforces
' that - a genuine, deliberate behavior change, not a silent one,
' verified against the loaded English corpus rather than assumed safe:
' english_expanded.vla's own "paste-values" has two templates - one
' with distinct {a}/{b} (a general two-range paste), one reusing {r}
' twice (the paste-in-place shorthand, "convert this range to values").
' Before this fix, G-RENDER could have silently matched the in-place
' template against a concrete call whose two ranges actually DIFFER,
' rendering a sentence that mentions the range only once and quietly
' drops the fact that source and destination differ. See
' VLA_Tests_Query.bas's own TestGRenderUnify (via TestDSLs) for the pin.
'
' Also corrects BETA_ROADMAP2.md's own PROLOG paragraph on one point,
' found while reading UnifyForm rather than trusted from its prose: the
' glued-slot branch never actually read mPatItems (that name belongs to
' a DIFFERENT, later gate - VLA_English.bas's own CandidateShapeOk,
' which re-checks a render candidate's own typed slots AFTER UnifyForm
' has already succeeded, entirely outside this walk). The glued-slot
' branch is pure prefix/suffix string arithmetic over its own two
' arguments, exactly as DSL-ignorant as the other three cases - so
' UnifyForm needed no atom-level hook back into VLA_English at all; it
' is a complete, thin pass-through, not a partial one.
'
' LAYER:     0.5 (no dependency but VLA_Messages, for its own two named
'            refusals below - a malformed multi-slot glued atom and an
'            occurs-check violation are both program-authoring defects,
'            not per-caller argument-shape questions, so this module
'            owns that wording directly rather than deferring it to
'            each caller the way VLA_Relation's own LAYER 0.5 functions
'            do). Deliberately still no VLA_Identity/VLA_Runtime
'            dependency - PROLOG.2's own header above has why.
' MAY CALL:  VLA_Messages (RaiseMsg - two refusals only, see above)
' SHIPS:     add-in only, VLA_Relation.bas's own sibling - never
'            injected into emitted VBA.
' PAYS INTO: G-RENDER (UnifyOneWay, via UnifyForm), VLA_Prolog.bas's own
'            conjunctive query solver (UnifyTwoWay/EnvWalkInto/
'            UnifyEnvClone, PROLOG.3), PROLOG.4's SLD resolution
'            (UnifyTwoWay is its own resolution step - proving a
'            subgoal against a candidate clause head), SOLVE's own
'            grounding step (one-way mode, per
'            BETA_ROADMAP2.md's own PROLOG paragraph).
' REASON:    proving the substrate-sharing claim against a REAL,
'            already-shipped second client (G-RENDER) before anything
'            backtracking-shaped exists to also depend on it - this
'            line's own house methodology, applied to itself.

' One-way structural match: tmplForm's own {name}/prefix{name}suffix
' slots bind against concreteForm, which is assumed fully ground - a
' bare {name} binds the WHOLE matching subform; a glued slot
' (prefix{name}suffix, exactly one embedded slot per atom - a second is
' refused loudly via unify-glue-multiple-slots rather than silently
' mismatched) binds the stripped middle; a literal atom must match
' text-for-text (case-sensitive, no VLA_Identity fold - this compares
' Lisp-form text, not an identifier lookup); a list must be the same
' length, every element unifying in order. Returns False - with bn/bv
' left however far the walk got - the moment any part fails to unify;
' a caller wanting a clean retry must pass a FRESH bn/bv pair each time
' ("As New" inside a loop only auto-instantiates once - the same trap
' VLA_English.bas's own FindRenderRule already documents). A slot name
' repeated within tmplForm must bind the SAME value every time it
' recurs (UnifyBind, below, enforces this) - real unification's own
' rule; see this module's own header for why that's a deliberate change
' from UnifyForm's prior behavior, not merely inherited from it.
Public Function UnifyOneWay(ByVal tmplForm As Variant, ByVal concreteForm As Variant, _
                             bn As Collection, bv As Collection) As Boolean
    If Not IsObject(tmplForm) Then
        Dim s As String
        s = CStr(tmplForm)
        If Left$(s, 1) = "{" And Right$(s, 1) = "}" And Len(s) > 2 And InStr(2, s, "{") = 0 Then
            ' Bare slot - binds the whole concrete subform.
            UnifyOneWay = UnifyBind(bn, bv, Mid$(s, 2, Len(s) - 2), concreteForm)
            Exit Function
        ElseIf InStr(s, "{") > 0 Then
            ' Glued slot (make-{d}, xl{d}, format-as-{d}): the concrete
            ' side must be a bare, non-list atom whose known
            ' prefix/suffix strips clean.
            If IsObject(concreteForm) Then Exit Function
            Dim openPos As Long, closePos As Long
            openPos = InStr(s, "{")
            closePos = InStr(openPos, s, "}")
            If closePos = 0 Then Exit Function
            If InStr(closePos, s, "{") > 0 Then
                VLA_Messages.RaiseMsg "unify-glue-multiple-slots", "atom", s
            End If
            Dim prefix As String, suffix As String, gSlotName As String
            prefix = Left$(s, openPos - 1)
            suffix = Mid$(s, closePos + 1)
            gSlotName = Mid$(s, openPos + 1, closePos - openPos - 1)
            Dim cs As String
            cs = CStr(concreteForm)
            If Len(cs) <= Len(prefix) + Len(suffix) Then Exit Function
            If Left$(cs, Len(prefix)) <> prefix Then Exit Function
            If Len(suffix) > 0 Then
                If Right$(cs, Len(suffix)) <> suffix Then Exit Function
            End If
            UnifyOneWay = UnifyBind(bn, bv, gSlotName, Mid$(cs, Len(prefix) + 1, Len(cs) - Len(prefix) - Len(suffix)))
            Exit Function
        Else
            ' A literal atom - must match the concrete atom text-for-
            ' text. VBA's And does not short-circuit, so the object
            ' check must be its own branch, never combined with a
            ' CStr() of the same value in one expression (a concrete
            ' LIST here would otherwise raise 450 trying to coerce a
            ' Collection through its own indexed default member).
            If IsObject(concreteForm) Then Exit Function
            UnifyOneWay = (CStr(concreteForm) = s)
            Exit Function
        End If
    End If
    ' tmplForm is a list - concreteForm must be one too, same length,
    ' every element unifying in order.
    If Not IsObject(concreteForm) Then Exit Function
    Dim tl As Collection, cl As Collection
    Set tl = tmplForm
    Set cl = concreteForm
    If tl.Count <> cl.Count Then Exit Function
    Dim i As Long
    For i = 1 To tl.Count
        If Not UnifyOneWay(tl.Item(i), cl.Item(i), bn, bv) Then Exit Function
    Next
    UnifyOneWay = True
End Function

' Binds slotName -> boundVal into bn/bv. A FIRST occurrence of a slot name
' always succeeds and is recorded; a REPEATED occurrence must bind the
' identical value (FormsEqual, below) or the whole match fails - real
' unification's own rule, not merely recorded twice the way UnifyForm
' used to (silently, and only ever read back as whichever occurrence
' came first).
Private Function UnifyBind(bn As Collection, bv As Collection, ByVal slotName As String, ByVal boundVal As Variant) As Boolean
    Dim k As Long
    For k = 1 To bn.Count
        If CStr(bn.Item(k)) = slotName Then
            UnifyBind = FormsEqual(bv.Item(k), boundVal)
            Exit Function
        End If
    Next
    bn.Add slotName
    bv.Add boundVal
    UnifyBind = True
End Function

' Ground-form structural equality - used only to check a repeated
' slot's second occurrence against its first (UnifyBind, above), never
' for unification itself (a slot atom "{x}" compared here is literal
' braced text, not re-interpreted as a binding site). Case-sensitive
' text compare for atoms, matching UnifyOneWay's own literal-atom case;
' structural recursion, same length and every element equal in order,
' for lists.
Private Function FormsEqual(ByVal a As Variant, ByVal b As Variant) As Boolean
    If IsObject(a) <> IsObject(b) Then Exit Function
    If Not IsObject(a) Then
        FormsEqual = (CStr(a) = CStr(b))
        Exit Function
    End If
    Dim la As Collection, lb As Collection
    Set la = a
    Set lb = b
    If la.Count <> lb.Count Then Exit Function
    Dim i As Long
    For i = 1 To la.Count
        If Not FormsEqual(la.Item(i), lb.Item(i)) Then Exit Function
    Next
    FormsEqual = True
End Function

' The classical Prolog/Datalog convention (VLA_Datalog.bas's own
' IsVariableAtom, repeated - see this module's own PROLOG.2 header for
' why this is a deliberate small duplication there, not a missed hoist).
' PROLOG.3: promoted Private->Public - unlike VLA_Datalog.bas (an
' independent module with no calling relationship to this one),
' VLA_Prolog.bas already depends on VLA_Unify directly, and its own
' parser needs to check "is this fact fully ground" using the EXACT
' same rule UnifyTwoWay uses internally - a THIRD independent copy here
' would be a real synchronization risk (the two modules silently
' disagreeing on what a variable looks like) rather than the harmless,
' genuinely-decoupled duplication DATALOG's own copy is.
Public Function IsVarAtom(ByVal s As String) As Boolean
    If Len(s) = 0 Then Exit Function
    Dim c As Integer
    c = AscW(Left$(s, 1))
    IsVarAtom = (c >= 65 And c <= 90)
End Function

' Two-way unification: a bare variable atom (IsVarAtom) may now appear
' on EITHER side, not just tmplForm's. Both sides are dereferenced
' through envN/envT first (EnvWalkInto - chases a variable-to-variable
' chain to whatever it's ultimately bound to, or to itself if still
' free); a still-free variable unifies with anything by binding to it
' (occurs-checked first, PROLOG.1's own already-decided policy: refused
' by name, never skipped for speed); two ground atoms must match text-
' for-text; two compound terms (lists) must be the same length, every
' element unifying under the SAME threaded environment; anything else
' (a variable aside) fails. Returns False - with envN/envT left however
' far the walk got - the moment any part fails to unify or raises; a
' caller wanting a clean retry must pass a FRESH envN/envT pair, same
' discipline as UnifyOneWay's own bn/bv.
Public Function UnifyTwoWay(ByVal a As Variant, ByVal b As Variant, envN As Collection, envT As Collection) As Boolean
    Dim aw As Variant, bw As Variant
    EnvWalkInto aw, a, envN, envT
    EnvWalkInto bw, b, envN, envT

    ' G0 hazard, live-caught (TestUnifyTwoWay's own run 450'd on exactly
    ' this line): VBA's And does not short-circuit - the original single
    ' expression "(Not IsObject(aw)) And IsVarAtom(CStr(aw))" evaluates
    ' CStr(aw) unconditionally even when aw IS an object, invoking a
    ' Collection's own default member with no index and raising 450 -
    ' the exact trap UnifyOneWay's own literal-atom case already has a
    ' comment naming, missed here anyway. Every IsObject check below is
    ' its own guarding If, never combined with a same-value CStr() in
    ' one boolean expression.
    Dim aIsVar As Boolean, bIsVar As Boolean
    If IsObject(aw) Then
        aIsVar = False
    Else
        aIsVar = IsVarAtom(CStr(aw))
    End If
    If IsObject(bw) Then
        bIsVar = False
    Else
        bIsVar = IsVarAtom(CStr(bw))
    End If

    If aIsVar And bIsVar Then
        If CStr(aw) = CStr(bw) Then
            UnifyTwoWay = True   ' the same still-free variable, trivially
            Exit Function
        End If
    End If

    If aIsVar Then
        If EnvOccurs(CStr(aw), bw, envN, envT) Then VLA_Messages.RaiseMsg "prolog-occurs-check", "var", CStr(aw)
        EnvBind envN, envT, CStr(aw), bw
        UnifyTwoWay = True
        Exit Function
    End If

    If bIsVar Then
        If EnvOccurs(CStr(bw), aw, envN, envT) Then VLA_Messages.RaiseMsg "prolog-occurs-check", "var", CStr(bw)
        EnvBind envN, envT, CStr(bw), aw
        UnifyTwoWay = True
        Exit Function
    End If

    If IsObject(aw) <> IsObject(bw) Then Exit Function

    If Not IsObject(aw) Then
        UnifyTwoWay = (CStr(aw) = CStr(bw))
        Exit Function
    End If

    Dim la2 As Collection, lb2 As Collection
    Set la2 = aw
    Set lb2 = bw
    If la2.Count <> lb2.Count Then Exit Function
    Dim i2 As Long
    For i2 = 1 To la2.Count
        If Not UnifyTwoWay(la2.Item(i2), lb2.Item(i2), envN, envT) Then Exit Function
    Next
    UnifyTwoWay = True
End Function

' Dereferences term through envN/envT: while term is a bare variable
' atom with an existing binding, follow it (a variable bound to another
' variable chains through, union-find style); stops at either a still-
' free variable or a non-variable term. dest is a ByRef out param, not
' a function return, matching VLA_Datalog.bas's own NthInto - a bare
' "x = objectValue" assignment on a Variant invokes the object's own
' default member instead of copying the reference (this codebase's own
' documented VBA trap), so every reassignment below branches on
' IsObject explicitly. PROLOG.3: promoted Private->Public - a query
' engine built on top of UnifyTwoWay needs this to read a free
' variable's own final bound value back out of a successful
' unification's own env, the identical F.8 precedent (a Private Sub
' promoted the moment a second module needs to call it).
Public Sub EnvWalkInto(ByRef dest As Variant, ByVal term As Variant, envN As Collection, envT As Collection)
    Do While Not IsObject(term)
        Dim s As String
        s = CStr(term)
        If Not IsVarAtom(s) Then Exit Do
        Dim k As Long, found As Boolean
        found = False
        For k = 1 To envN.Count
            If CStr(envN.Item(k)) = s Then
                If IsObject(envT.Item(k)) Then
                    Set term = envT.Item(k)
                Else
                    term = envT.Item(k)
                End If
                found = True
                Exit For
            End If
        Next
        If Not found Then Exit Do
    Loop
    If IsObject(term) Then
        Set dest = term
    Else
        dest = term
    End If
End Sub

' PROLOG.3: a fresh, independent copy of one substitution environment -
' needed the moment a caller enumerates several candidate matches for
' the SAME position (a query conjunct tried against every stored fact
' in turn, PROLOG.3's own case; a choice point in PROLOG.4's own future
' backtracking search) and must not let one candidate's own partial or
' successful bindings leak into a sibling attempt. UnifyTwoWay's own
' contract already requires a FRESH envN/envT per attempt when a retry
' is wanted; this is that fresh copy, built once so no caller re-invents
' the "Collection has no clone method" loop by hand.
Public Sub UnifyEnvClone(ByVal srcN As Collection, ByVal srcT As Collection, ByRef dstN As Collection, ByRef dstT As Collection)
    Set dstN = New Collection
    Set dstT = New Collection
    Dim k As Long
    For k = 1 To srcN.Count
        dstN.Add srcN.Item(k)
        dstT.Add srcT.Item(k)
    Next k
End Sub

' Occurs check: does varName appear anywhere within term's own fully-
' dereferenced structure? Run BEFORE every EnvBind, never after - as
' long as that order holds, envN/envT can never contain a cycle, which
' is what makes EnvWalkInto's own chain-chasing loop above safe to run
' unconditionally with no cycle guard of its own.
Private Function EnvOccurs(ByVal varName As String, ByVal term As Variant, envN As Collection, envT As Collection) As Boolean
    Dim tw As Variant
    EnvWalkInto tw, term, envN, envT
    If Not IsObject(tw) Then
        EnvOccurs = IsVarAtom(CStr(tw)) And (CStr(tw) = varName)
        Exit Function
    End If
    Dim lst As Collection
    Set lst = tw
    Dim i As Long
    For i = 1 To lst.Count
        If EnvOccurs(varName, lst.Item(i), envN, envT) Then
            EnvOccurs = True
            Exit Function
        End If
    Next
    EnvOccurs = False
End Function

' Records varName -> boundTerm. Only ever reached (from UnifyTwoWay,
' above) for a variable EnvWalkInto has just confirmed is still free -
' see EnvWalkInto's own comment - so this never needs to check for or
' overwrite an existing binding.
Private Sub EnvBind(envN As Collection, envT As Collection, ByVal varName As String, ByVal boundTerm As Variant)
    envN.Add varName
    envT.Add boundTerm
End Sub
