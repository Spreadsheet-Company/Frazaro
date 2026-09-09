Attribute VB_Name = "VLA_Interpreter"
Option Explicit
Public Const VLA_INTERPRETER_VERSION As String = "LINTERPOLATE.0"
' LINTERPOLATE.0: (interpolate tpl :key val ...) - EvalExpr's own Case
' "interpolate", below, next to "array". Named "interpolate", not CL's
' "format" - this codebase's own format-as-currency/-percent/-date
' (english.vla) already own "format" for visual cell styling; a second
' "format" doing string substitution would collide in READING even
' with no symbol collision (owner's own call, this session - the
' roadmap item itself was renamed L-FORMAT -> L-INTERPOLATE alongside
' this). A real primitive, not a defmacro (BETA_ROADMAP2.md's own
' L-INTERPOLATE text: a defmacro substitutes into a fixed template FORM
' at define time and cannot parse the CONTENTS of a string literal
' supplied at a call site). tpl must be a literal string node (checked
' structurally, like "new"'s own type-name argument, just above - not
' evaluated, since VLA.bas's own EmitExpr Case "interpolate" needs the
' hole names at COMPILE time and the two backends must accept exactly
' the same programs). Every argument after tpl must be a ":key value"
' pair (EvalKeywordArgs's own all-keyword convention, reused by shape
' only - a bespoke loop, not a shared call, since this primitive owns
' its own refusals). Named holes, not positional {0}/{1}: the two
' DATALOG macros this primitive was built to absorb
' (datalog-filter-place/datalog-chain-place, english.vla) each reuse a
' hole (alias, table) two or three times in one template, which a named
' hole gives back "for free" (supply the value once, reference it as
' often as the template needs) - a positional scheme would make every
' repeat a renumbering hazard instead. Value substitution is plain VBA
' "&" on the evaluated hole value, the exact operator EvalOpChain's own
' "&" case already uses two functions below - no separate coercion
' rule, by the roadmap item's own "confirm before diverging" question,
' answered no. New messages (VLA_Messages.bas): interp-interpolate-
' template-must-be-literal, interp-interpolate-expected-keyword-arg,
' interp-interpolate-unclosed-hole, interp-interpolate-empty-hole-name,
' interp-interpolate-unknown-key, interp-interpolate-unused-argument (a
' dangling ":key" with no value reuses the existing interp-keyword-arg-
' missing-value, the same generic refusal every other keyword-arg call
' site already raises).
' PF4C.0: ExecForEachRow - the interpreter's own twin of VLA.bas's
' EmitForEachRow, PF.4c. Bulk-read once (VLA_Runtime.VlaSlabRead), walk
' rows as a plain VBA array held entirely in this Sub's own locals -
' never boxed into the frame Dictionary, so there's no name-collision
' question to solve the way the compiled side's mSlabCounter scheme
' has to (this bookkeeping lives on the real call stack, invisible to
' the VLA namespace); `row` itself IS stored in the frame under its own
' name each iteration, exactly like ExecFor/ExecForEach's own loop
' variable, so PF.4a's array-element set!/read parity is what makes
' (row 2)/(set! (row 2) ...) work against it. All-or-nothing commit,
' checked against the COMPILED side's REAL behavior rather than assumed
' symmetric: EmitReturn emits a bare Exit Sub/Exit Function, always
' leaving the whole procedure and skipping EmitForEachRow's own write-
' back no matter the nesting depth - so mProcReturn/mGotoLabel are
' checked FIRST here and skip VlaSlabWrite too, matching that; only a
' plain mLoopBreak commits everything computed so far, mirroring the
' compiled side's own Exit For falling through to its write-back line
' normally. New message: interp-for-each-row-needs-range (VLA_Messages.
' bas). Parity proven the standard way (VLA_Tests_Host.TestStmtParity):
' a full round-trip case (write-back actually reaches the live sheet,
' verified via a plain '.'-dot read, not through for-each-row again)
' and a break-commits-partial case (row 3 must survive untouched after
' breaking right after row 2).
' PF4A.0: array-element set!/read parity with the compiled backend -
' PRODUCT · PERFORMANCE's PF.4a, the first of PF.4's three sub-items
' (BETA_ROADMAP1.md/BETA_ROADMAP2.md). Compiled has always handled
' (set! (arr i) v) and (arr i) for free - EmitStmt's/EmitExpr's own
' generic text substitution emits real VBA array-index syntax with no
' compiler awareness needed. The interpreter had neither direction:
' ExecSet's one non-'.' place case required an Object with a default
' Value member (every prior use was a Range), and EvalDynamicHead's
' every tier required h to match something specific, with no tier ever
' checking "is h a local variable bound to a plain array" - so both
' (set! (arr i) v) and (arr i) always failed outright before this,
' never silently wrong, just never exercised (nothing in the existing
' corpus ever assigned into or read a plain array element via a VLA
' variable). Both fixed as new, purely-additive fallback tiers - array-
' place detection in ExecSet ahead of its existing Object-place
' fallback, array-read detection in EvalDynamicHead after every
' existing tier has already failed to claim the head - so neither can
' shadow anything that already worked. VlaDictGet copies an array out
' (a Variant-boxed SafeArray copies on assignment in VBA, unlike an
' object reference), so ExecSet's own write re-stores the mutated copy
' under its own name via VlaDictSet rather than relying on any in-place
' aliasing. Only 1-D indexing - the only shape PF.4c's own for-each-row
' will need. Parity proven the standard way (VLA_Tests_Host.
' TestStmtParity, CheckStmtParity running the same program under both
' backends against a hand-computed expected value), not assumed from
' the emitter's own free ride.
' A real bug caught live (owner-run VlaSelfTestsAll, first pass), not
' reasoned through in advance: both new tiers' first draft called
' VLA_Runtime.VlaDictGet directly to test "does h/the place's head name
' a stored array" - but VlaDictGet is deliberately loud on a miss
' (RaiseRuntimeMsg "rt-dict-key-missing", VLA_Runtime.bas's own "loud
' step error" contract for a genuinely undefined VLA variable), not a
' safe try-get returning Empty. A plain (range "a1") place also has
' Count = 2, and "range" was never a stored variable, so ExecSet's own
' new check raised "there is nothing stored at key 'range'" for every
' ordinary bare-place Range assignment (IN.2/IN.11's own shape,
' `(set! (range ...) v)`) - 10 host failures, all of them either
' TestStmtParity's own cases (every one ends by writing its answer to
' `(range "A1")`) or the one dedicated IN.2 bare-place test, never the
' pure suite, which explained the pattern exactly once traced. Fixed by
' guarding both new tiers with VlaDictHas (VLA_Runtime.bas) first - a
' non-raising existence check already built for exactly this "test
' before falling back" shape, whose own header names this
' interpreter's module-scope read fallback as its intended first
' caller, a precedent this should have followed from the start rather
' than reinventing a riskier check.
' LX2.0: this module's 60 of 61 raw Err.Raise refusal sites now route
' through VLA_Messages.RaiseMsg with a stable id - SD-2/LX.2. The 61st
' (line ~1199, ExecStmtTrapped) is a bare re-raise of an already-caught
' error escaping a local On Error Resume Next, not an origination of new
' English text, and stays raw by design, same category as VLA_IDE.bas's
' line ~1972. Several sites shared one id with another call site
' elsewhere in this file when the rendered English text was exactly
' identical - not a partial migration, the same refusal reached from
' more than one place. Rendered text and Err.Number/Err.Source are
' unchanged (verified by manual trace at migration time for every site).
' This module is the second, independent backend: a tree-walking
' interpreter that runs VLA source directly against a live workbook,
' with no VBA code ever generated. Built so the exported VBA (the
' emitter, VLA.bas) would have something real to be checked against
' instead of trusted on its own word - VerifyReportInterpreter now
' runs instructions.txt to completion with cell-identical output to the
' emitter's own real Run.
'
' The build history - what each dispatch tier cost, and the bugs found
' getting whole-corpus parity - is in docs/BETA_ROADMAP.md's own IN.2/IN.3/
' IN.10/IN.11 entries and, for the harder-won stories, docs/TRENCHES.md.
' =====================================================================
'  VLA_Interpreter - the full evaluator: every core control-flow/
'  arithmetic form, member dispatch against Excel's own object model,
'  user-procedure call/return, and on-error/goto/label/resume, proven
'  against instructions.txt's whole corpus, not just hand-picked forms.
'
'  LAYER:     dev-rig in name only now (not one of REBUILD.md's
'             numbered layers - that split was never revisited when
'             this stopped being a dev-only module)
'  MAY CALL:  VLA (VlaCompileToForms), VLA_English (EnglishToVla and its
'             grammar-loading siblings), VLA_Runtime (VlaDictNew/Set/Get
'             directly for the variable frame; every other VLA_Runtime
'             helper reached dynamically through Application.Run, not a
'             hand-written direct call), VLA_Identity (Fold), Excel's
'             own global objects via CallByName.
'  SHIPS:     yes, as of IN.6 - in VLA_Build.bas's mods array, and
'             called from VLA_IDE.bas's InterpretProgram (the "Interpret
'             Instructions"/"Interpret and Trace" ribbon commands),
'             IN.9's own resolution made real. Left off the shipped
'             list from IN.0.5 through IN.11 while this was dev-rig-only
'             by design; the omission survived unnoticed for several
'             items after that stopped being true, caught only when
'             IN.6 made VLA_IDE.bas the first SHIPPED module to actually
'             reference VLA_Interpreter by name (a fresh add-in build
'             failing Debug > Compile, the same failure shape F5.0/
'             LX3.0/IN1.0 already hit for the same root cause).
'  PAYS INTO: IN.2/IN.3/IN.10/IN.11 (this module IS their build), IN.3.5
'             (VlaInterpreterEffectLog, the interpreter's own runtime-
'             effect record), IN.9 (the interpreter-as-runtime decision
'             this module makes real).
'
'  Duplicates VLA.bas's own Private IsList/HeadSym/Nth/SymText/
'  StrLitContent/AssignVar rather than widen VLA.bas's public surface
'  for a dev-only caller - cross-module Private calls do not exist in
'  VBA, the same move VLA_Runtime.bas already makes for Fold. Each is
'  under fifteen lines; kept in sync by hand.
' =====================================================================

' ---------------------------------------------------------------------
'  The demo: a real, working english.vla program - English text
'  in, live effects out (Debug.Print only; no VBA is ever generated,
'  compiled, or written to any module). Run from the Immediate window:
'    VlaInterpretDemo
'  Ten forms exercised: dim, set!, if, for, begin, while, debug-print,
'  +, -, >. The English half (dim/set!/+/if/for/debug-print) is the
'  "one sentence end to end" proof, unmodified corpus phrasing,
'  compiled through the real EnglishToVla; the trailing hand-written
'  VLA line adds begin/while/- for full ten-form coverage in one run,
'  since no single short English sentence naturally reaches all ten
'  and VLA is itself a legitimate, hand-writable input, not just a
'  compile target.
' ---------------------------------------------------------------------
' IN3_5.0: the effect log for the most recent VlaInterpret call - see
' the LogEffect/VlaInterpreterEffectLog pair below.
Private mEffectLog As Collection

' IN2.3: set by exit-for/exit-do, consumed by whichever loop executor
' (ExecFor/ExecWhile/ExecForEach/ExecDoUntil) owns the nearest enclosing
' loop - see ExecBody's own note for how it propagates through nested
' bodies (if-branches, begin) without each of those needing to know
' anything about loops. Reset at the top of every VlaInterpret call,
' the same S3.1-safe discipline as mEffectLog/mHeadAliases.
Private mLoopBreak As Boolean

' IN.10: the registered user-procedure table - every top-level
' (sub name ...)/(function name ...) form, keyed by VLA_Identity.Fold'd
' name, populated by a first pass over VlaInterpret's forms BEFORE any
' execution (RegisterProc/LookupProc, the same Collection-as-dict shape
' VLA.bas's own mMacros/GetMacro already use). Reset fresh at the top
' of every VlaInterpret call, same S3.1-safe discipline as mEffectLog/
' mLoopBreak.
Private mProcs As Collection

' IN.3: the module-level frame - real VBA module scope, measured against
' the live corpus rather than assumed (IN.10's own text flagged this as
' an open question: "whether this gap actually blocks whole-corpus parity
' ... is exactly the kind of thing to settle by reading the real corpus").
' It does: instructions.txt's "Define hot-pink as ..." compiles to a TRUE
' top-level (const hot-pink ...), a sibling of (sub main ...) rather than
' something inside it (VLA_English.bas's ParseDefine, "Emits a module-
' level Const ... visible in every action"), and "To tidy-up:" - a real
' called procedure, not main - reads hot-pink in its own body. A call
' frame built by CallUserProc is fresh and isolated (IN.10's own named
' non-goal), so without this, that read raises "undefined variable" the
' first time anything actually calls tidy-up. Set once per VlaInterpret
' call to the same frame "main" itself runs in (module-level top forms
' and main share one frame already, by construction - ExecTop inlines
' main into VlaInterpret's own top-level frame), so no separate reset
' path is needed beyond VlaInterpret's own top-level frame creation.
' READ-only: a plain-variable lookup that misses the local frame falls
' back to this one (EvalExpr's atom base case). Write-back (a callee
' assigning to a module-level name with no local shadow) is a named
' non-goal, not built - no real procedure in today's corpus writes one;
' every module-level top form instructions.txt has (hot-pink, and
' vla-problem when the corpus uses Try:) is either a const (cannot be
' assigned, VLA_English.bas's own ParseDefine comment) or, for
' vla-problem, read/written only from within main's own body, which
' already IS this frame, not a call frame.
Private mModuleFrame As Object

' IN.6: "thisworkbook" (save-workbook-as/close-workbook/save-current-
' workbook - real english.vla macros, none exercised by
' instructions.txt, which is why whole-corpus parity never caught this)
' means, in VBA, "the workbook that contains the code now running."
' For emitted VBA that's unambiguously the host - the generated module
' lives there. This interpreter lives in the add-in, so a bare
' ThisWorkbook here would mean the ADD-IN once shipped, a silent
' backend divergence (SD-5) invisible in dev-rig only because the
' dev workbook plays both roles at once. Set once per VlaInterpret
' call, defaulting to ActiveWorkbook (dev-rig-safe, unchanged for
' every existing caller that doesn't pass one), never module-level
' *cached* across calls - the same S3.1-safe discipline mEffectLog/
' mMacros/mHeadAliases already use.
Private mHostWorkbook As Workbook

' IN.10: return/exit-sub/exit-function unwind - mLoopBreak's own shape
' (IN2.3), but scoped to the nearest enclosing CALL FRAME rather than
' the nearest enclosing loop: ExecBody stops on mLoopBreak OR
' mProcReturn, but only a loop executor resets mLoopBreak (its own
' break), while mProcReturn propagates untouched through any number of
' nested loops/if-branches until CallUserProc - the call's own
' boundary - consumes it. mProcReturnValue holds what (return <expr>)
' assigned; EmitReturn's own semantics (VLA.bas:2313) mirrored exactly:
' (return <expr>) assigns AND unwinds immediately, a bare (return) (or
' exit-sub/exit-function) is a plain unwind. Both reset at the top of
' every VlaInterpret call and at every CallUserProc boundary (a call
' frame's own flags never leak to its caller).
Private mProcReturn As Boolean
Private mProcReturnValue As Variant

' IN3.6: on-error/goto/label/resume - a single flat handler, not a
' stack. Measured against the real corpus, not assumed: every live use
' - the three Try:/If-that-fails: blocks (VLA_English.bas's Case
' "try") and english.vla's show-all-rows macro - is body-local,
' non-nested, and never crosses a user-procedure call boundary. The
' goto/label pairs Case "try" emits are always SIBLINGS of the
' (on-error ...) form that arms them, inside the one statement list
' Try: was itself written into (ParseStmt splices its whole multi-form
' output back into the surrounding block in place, not into a new
' nested one) - so ExecBody's own per-statement loop, searching its
' OWN list for a matching label before bubbling the miss up to
' whichever body/loop called it (mLoopBreak/mProcReturn's exact
' existing shape, applied to a label search instead of a flag), is
' enough: no call-frame-crossing handler stack is needed by anything
' this corpus does. That would be real future work if a Try ever
' wraps a call to a user procedure that itself needs recoverable
' errors - named, not built.
Private mErrMode As String     ' "" (no handler armed) | "goto" | "resume-next"
Private mErrLabel As String    ' target label when mErrMode = "goto"
' IN3.6: a pending (goto <label>)/(resume <label>) jump, or a trapped
' error's jump target - set, searched for, and cleared entirely within
' ExecBody's own loop (see its header note); a loop executor
' (ExecFor/ExecForEach/ExecWhile/ExecDoUntil) that finds it still set
' after its own body executes abandons its own iteration the same way
' it already does for mLoopBreak/mProcReturn, letting the search
' continue one level further out.
Private mGotoLabel As String

' IN3.6: a SHADOW of the most recently trapped real error, populated by
' ExecStmtTrapped and read by EvalExpr's own dedicated "err." dotted-
' atom case - NOT the live VBA Err object, on purpose, found by an
' owner-run test failure rather than assumed: VBA resets the Err
' object's own properties on execution of ANY "On Error" statement
' (Resume/Exit too, but not relevant here), not just an explicit
' Err.Clear - so ExecStmtTrapped's own "On Error Resume Next", re-run
' fresh for EVERY statement while a handler stays armed (including the
' harmless no-op (label errf) statement Try:'s own recovery always
' executes between the jump and its own err.description read), was
' silently wiping the real failure before the recovery paragraph ever
' got to read it. Fully interpreter-owned instead: only ExecStmtTrapped
' writes it (on an actual real-error catch) and only ExecResume clears
' it (mirroring real VBA's own "Resume clears Err" semantics, now
' exact rather than accidental), so nothing this interpreter's own
' plumbing does in between can touch it.
Private mCaughtErrNum As Long
Private mCaughtErrSrc As String
Private mCaughtErrDesc As String

' IN.15: TryRuntimeHelper's manifest cache - see HelperManifestCached
' for why it exists (VlaHelperManifest re-reads and folds the whole
' VLA_Runtime source on every call, and the lookup it feeds sits in the
' per-expression dispatch path). Module memory is wiped by any Run or
' scratch injection (S5.1), which simply costs one re-read.
Private mHelperManifest As String
Private mHelperManifestRead As Boolean

Public Sub VlaInterpretDemo()
    Dim english As String
    Dim vla As String
    vla = VlaSkeletonDemoVla(english)

    Debug.Print "===== IN.0.5 WALKING SKELETON ====="
    Debug.Print "English:"
    Debug.Print english
    Debug.Print "----- interpreting; no VBA generated -----"
    Dim frame As Object
    Set frame = VlaInterpret(vla)
    Debug.Print VlaInterpreterEffectLog()
    Debug.Print "===== DONE - total ended at " & CStr(VLA_Runtime.VlaDictGet(frame, "total")) & " ====="
End Sub

' IN3_5.0: factored out of VlaInterpretDemo (unchanged text, unchanged
' behavior) so VLA_Tests.bas's VlaWriteInterpreterGoldens can run the
' exact same program the demo does, rather than a second hand-copied
' string that would drift from it over time. Returns the assembled VLA
' source; englishOut (if supplied) receives the raw English half, for
' callers that also want to print or golden it.
Public Function VlaSkeletonDemoVla(Optional ByRef englishOut As String) As String
    Dim vocabPath As String
    vocabPath = FindDevFile("english.vla")
    EnglishResetGrammar
    EnglishLoadVocabulary vocabPath

    ' Step tracking (vla-step, at-line, the trace-report helper subs)
    ' is on by default (VLA_English.bas's EnsureInit, "default on") and
    ' would wrap every statement in machinery this ten-form skeleton
    ' does not implement - caught on the maiden reload. Off for this
    ' translation only, restored after: EnglishStepTracking has no
    ' reset counterpart in EnglishResetGrammar, so leaving it off would
    ' silently change every OTHER translation in the same session
    ' (the real IDE's Run/Check, VlaSelfTest's own translations) - the
    ' same class of hazard LESSONS.md already names for module state.
    EnglishStepTracking False
    Dim english As String
    english = "Create a number called total." & vbCrLf & _
               "Set total to 0." & vbCrLf & _
               "Increase total by 5." & vbCrLf & _
               "If total is greater than 3, log ""over three""." & vbCrLf & _
               "Repeat 3 times, log total."
    englishOut = english
    Dim englishVla As String
    englishVla = EnglishToVla(english)
    EnglishStepTracking True

    VlaSkeletonDemoVla = englishVla & vbCrLf & _
        "(begin (dim countdown Double) (set! countdown 3) " & _
        "(while (> countdown 0) (debug-print countdown) (set! countdown (- countdown 1))))"
End Function

' The reusable core VlaInterpretDemo calls, also the test entry point
' (VLA_Tests.bas's TestInterpreterSkeleton): compiles VLA source
' through VLA.bas's real Tokenize/ParseAll/ExpandMacros pipeline
' (VlaCompileToForms), then interprets every resulting top-level form
' in one fresh frame, returning it so a caller can inspect the
' results (VLA_Runtime.VlaDictGet(frame, "name")). IN3_5.0: also
' resets and populates mEffectLog, so every call starts its own clean
' record (VlaInterpreterEffectLog reads back after the call returns).
Public Function VlaInterpret(ByVal vlaSource As String, Optional ByVal hostWb As Workbook = Nothing) As Object
    Dim forms As Collection
    Dim frame As Object
    Set frame = PrepareInterpret(vlaSource, hostWb, forms)
    Dim f As Variant
    For Each f In forms
        ExecTop f, frame
        ' Top level is a call-frame boundary too: an exit-sub/return
        ' (or a stray exit-for/exit-do with no enclosing loop) inside a
        ' bare top-level form must not leak into the next top-level form.
        mLoopBreak = False
        mProcReturn = False
        ' IN3.6: a goto/resume target label that was never found - the
        ' search bubbled all the way out of the top-level form itself
        ' without a match. A stuck flag would silently corrupt every
        ' top-level form after it (ExecBody would keep trying to find a
        ' label that does not exist); raised loudly instead, by name,
        ' the same AS.6 discipline this whole module already follows.
        If Len(mGotoLabel) > 0 Then
            Dim badLabel0 As String
            badLabel0 = mGotoLabel
            mGotoLabel = ""
            VLA_Messages.RaiseMsg "interp-goto-label-not-found", "label", badLabel0
        End If
    Next
    Set VlaInterpret = frame
End Function

' IN.7: the compile+reset+register prefix VlaInterpret and
' VlaInterpretEntry both need - split out so the two entry points
' cannot drift on what "start a fresh interpretation" means. Returns
' the new frame and hands back the compiled forms through forms (an
' Object return can't also be ByRef, and every existing S3.1-safe
' reset here is unchanged from VlaInterpret's own original body).
Private Function PrepareInterpret(ByVal vlaSource As String, ByVal hostWb As Workbook, ByRef forms As Collection) As Object
    If hostWb Is Nothing Then
        Set mHostWorkbook = ActiveWorkbook
    Else
        Set mHostWorkbook = hostWb
    End If
    Set forms = VLA.VlaCompileToForms(vlaSource)
    Dim frame As Object
    Set frame = VLA_Runtime.VlaDictNew()
    Set mModuleFrame = frame
    Set mEffectLog = New Collection
    mLoopBreak = False
    mProcReturn = False
    mProcReturnValue = Empty
    mErrMode = ""
    mErrLabel = ""
    mGotoLabel = ""
    mCaughtErrNum = 0
    mCaughtErrSrc = ""
    mCaughtErrDesc = ""
    Set mProcs = New Collection

    ' IN.10: register every top-level (sub ...)/(function ...) BEFORE
    ' executing anything - a call to a procedure declared later in the
    ' same source (or never called at all) still needs to resolve.
    ' Registration alone has no runtime effect; only "main" (English's
    ' own wrapper name for top-level sentences, per VLA_English.bas)
    ' executes inline in VlaInterpret below - every other declared
    ' procedure (including IN.7's "on:sheet-change") stays dormant
    ' until something actually calls it (EvalDynamicHead, or
    ' VlaInterpretEntry's own direct CallUserProc).
    Dim f As Variant
    For Each f In forms
        RegisterProc f
    Next
    Set PrepareInterpret = frame
End Function

' IN.7: run ONE already-declared top-level procedure by name, instead
' of VlaInterpret's own hardcoded "main" entry point - the seam an
' event handler needs, since the caller here (an add-in-resident
' event sink, VLA_EventSink.cls, not itself running any VLA program)
' is OUTSIDE the currently-running program CallUserProc's only other
' call site (EvalDynamicHead) assumes. Deliberately does NOT persist
' mProcs/mModuleFrame across separate calls: every firing re-compiles
' and re-registers the same source text from scratch, the same
' S3.1-safe "nothing survives this call" discipline VlaInterpret
' itself already follows, rather than inventing new cross-call state
' this module has never needed before. entryProc takes no arguments
' today (event handlers are declared with an empty parameter list);
' the synthesized call form below is just the bare head symbol.
Public Function VlaInterpretEntry(ByVal vlaSource As String, ByVal entryProc As String, Optional ByVal hostWb As Workbook = Nothing) As Object
    Dim forms As Collection
    Dim frame As Object
    Set frame = PrepareInterpret(vlaSource, hostWb, forms)
    Dim proc As Collection
    Set proc = LookupProc(entryProc)
    If proc Is Nothing Then
        VLA_Messages.RaiseMsg "interp-entry-proc-not-found", "proc", entryProc
    End If
    Dim callLst As New Collection
    callLst.Add entryProc
    CallUserProc proc, callLst, frame
    If Len(mGotoLabel) > 0 Then
        Dim badLabel1a As String
        badLabel1a = mGotoLabel
        mGotoLabel = ""
        VLA_Messages.RaiseMsg "interp-goto-label-not-found", "label", badLabel1a
    End If
    Set VlaInterpretEntry = frame
End Function

' IN3_5.0: a Public wrapper letting a caller ask what the interpreter
' computes for one bare VLA EXPRESSION - not a statement, not a full
' program. ExecStmt's dispatch would refuse a bare (+ 2 3): it is
' statement-shaped ("not one of the ten forms"). Built for IN.3's
' expression-parity harness (VLA_Tests_Host.bas's TestExprParity),
' which needs this module's own answer for an expression to compare
' against the emitter's live-executed one, on neutral ground (a fresh
' frame, no prior state).
Public Function VlaEvalExpression(ByVal vlaExprSource As String) As Variant
    Dim forms As Collection
    Set forms = VLA.VlaCompileToForms(vlaExprSource)
    If forms.Count <> 1 Then
        VLA_Messages.RaiseMsg "interp-expr-not-single", "count", forms.Count
    End If
    Dim frame As Object
    Set frame = VLA_Runtime.VlaDictNew()
    ' IN2.0: EvalExpr can now return an OBJECT (a place helper like
    ' (range "a1"), or any '.' read) - plain "= EvalExpr(...)" would
    ' invoke that object's default member instead of capturing the
    ' reference (VBA's Let-without-Set gotcha), silently wrong rather
    ' than loudly broken. Harmless before IN2.0 (every expression this
    ' was ever called with was a scalar); AssignVar going forward.
    AssignVar VlaEvalExpression, EvalExpr(Nth(forms, 1), frame)
End Function

' First reload's maiden run caught this: EnglishToVla wraps every
' translated program in (sub main () body...) - VLA_English.bas's own
' comment says so ("sentences outside any 'To ...' become (sub main ()
' ...)") - so VlaCompileToForms's top-level forms are not bare
' statements when the source came from English text; they are ONE
' (sub main () ...) declaration. "sub" is deliberately NOT one of the
' ten interpreted forms (it is a compile-time declaration, not
' something with a runtime effect of its own) - this unwraps it into
' its body instead, mirroring VLA.bas's own EmitProc bodyIdx/optional-
' doc-skip logic (P.L6) so a docstring, if present, is skipped rather
' than executed. A top-level form that is NOT (sub ...) - this
' module's own hand-written VLA, e.g. VlaInterpretDemo's trailing
' (begin ...) line - reaches ExecStmt unchanged.
' IN.10: a top-level (sub ...)/(function ...) form was ALREADY
' registered into mProcs by VlaInterpret's own first pass - it is not
' executed here unless it is literally named "main", English's own
' wrapper name for top-level sentences. Every other declared procedure
' is reached only through a real call (EvalDynamicHead/CallUserProc),
' not by inlining - the bug this item's own roadmap text named (every
' top-level sub's body ran once, param-blind, in file order) is what
' this replaces.
Private Sub ExecTop(ByVal f As Variant, ByVal frame As Object)
    If Not IsList(f) Then
        ExecStmt f, frame
        Exit Sub
    End If
    Dim lst As Collection
    Set lst = f
    Dim h As String
    h = VLA_Identity.Fold(HeadSym(lst))
    If h <> "sub" And h <> "function" Then
        ExecStmt f, frame
        Exit Sub
    End If
    If VLA_Identity.Fold(SymText(Nth(lst, 2))) <> "main" Then Exit Sub
    Dim bodyIdx As Long
    bodyIdx = ProcBodyIdx(lst, (h = "function"))
    ExecBody lst, bodyIdx, frame
End Sub

' IN.10: shared bodyIdx computation for a (sub ...)/(function ...)
' form, mirroring VLA.bas's own EmitProc exactly - a function may have
' a return-type atom at position 4 (skipped, bodyIdx -> 5), then either
' shape may have an optional docstring immediately before the body
' (skipped again). Reused by both RegisterProc (every declared
' procedure) and ExecTop (main's own inline execution) so the two
' never drift.
Private Function ProcBodyIdx(lst As Collection, ByVal isFunc As Boolean) As Long
    Dim bodyIdx As Long
    bodyIdx = 4
    If isFunc And lst.Count >= 4 Then
        Dim rt As Variant
        AssignVar rt, Nth(lst, 4)
        If Not IsList(rt) Then
            If Left$(CStr(rt), 1) <> Chr$(34) Then bodyIdx = 5
        End If
    End If
    If lst.Count >= bodyIdx Then
        Dim dCand As Variant
        AssignVar dCand, Nth(lst, bodyIdx)
        If Not IsList(dCand) Then
            If Left$(CStr(dCand), 1) = Chr$(34) Then bodyIdx = bodyIdx + 1
        End If
    End If
    ProcBodyIdx = bodyIdx
End Function

' IN.10: registers a top-level (sub name ...)/(function name ...) form
' into mProcs - a no-op for any other top-level form (a bare statement,
' or main's own wrapper, which is registered too but only ever run via
' ExecTop's own special case, not through a call - nothing stops
' something from actually CALLING "main" though, and there is no
' reason it should). Redefinition replaces the earlier entry - mMacros'
' own DefineMacro precedent as it stood at the time (VLA.bas), though
' PDICT.0 later moved DefineMacro itself to a direct Dictionary
' overwrite instead of this Remove-then-Add shape (mMacros became a
' Scripting.Dictionary; mProcs here is still a Collection, untouched -
' out of P-DICT's own scoped literal target, not evidenced hot the way
' GetMacro was, and a candidate for the identical fix if it ever is).
Private Sub RegisterProc(ByVal f As Variant)
    If Not IsList(f) Then Exit Sub
    Dim lst As Collection
    Set lst = f
    Dim h As String
    h = VLA_Identity.Fold(HeadSym(lst))
    If h <> "sub" And h <> "function" Then Exit Sub
    Dim isFunc As Boolean
    isFunc = (h = "function")
    Dim name As String
    name = SymText(Nth(lst, 2))
    Dim rec As New Collection
    rec.Add isFunc
    rec.Add lst
    rec.Add ProcBodyIdx(lst, isFunc)
    Dim key As String
    key = VLA_Identity.Fold(name)
    On Error Resume Next
    mProcs.Remove key
    On Error GoTo 0
    mProcs.Add rec, key
End Sub

' IN.10: mMacros/GetMacro's own On-Error-Resume-Next existence-check
' shape AS IT STOOD AT THE TIME (VLA.bas) - Nothing when h names no
' declared procedure, so EvalDynamicHead's remaining tiers (place
' helpers, dotted globals, builtins, VLA_Runtime helpers) keep trying
' exactly as before. PDICT.0 later swapped GetMacro's own mMacros to a
' Dictionary with a throw-free .Exists check - mProcs/LookupProc here
' still use the original Collection/On-Error-Resume-Next shape this
' comment describes, deliberately untouched (out of that item's own
' scoped literal target; the interpreter's own procedure-count scaling
' is a different, unmeasured axis from the transpiler's macro-chain
' depth P-PROF actually measured).
Private Function LookupProc(ByVal name As String) As Collection
    If mProcs Is Nothing Then Exit Function
    On Error Resume Next
    Set LookupProc = mProcs.Item(VLA_Identity.Fold(name))
    On Error GoTo 0
End Function

' IN.7: a Public query over the SAME mProcs a VlaInterpret/
' VlaInterpretEntry call just populated - reads real state left behind
' by the call that just ran, not a guess. Built so a caller (VLA_IDE's
' InterpretProgram, right after its own VlaInterpret call) can ask
' "did the program I just ran declare a procedure named X" without
' this module knowing or caring what X is for - VLA_Events.bas is the
' only place "on:sheet-change" means anything.
Public Function VlaHasProc(ByVal name As String) As Boolean
    VlaHasProc = Not (LookupProc(name) Is Nothing)
End Function

' ---------------------------------------------------------------------
'  The statement forms: IN0.5's original seven, IN2.0's '.' (member
'  call in statement position) and 'obj-set!' (aliased straight to
'  ExecSet - see ExecSet's own note for why that alias is exact for
'  this interpreter's dict-backed frame, not an approximation), and
'  IN2.3's for-each/do-until/select/exit-for/exit-do - real, occasional
'  control flow the corpus actually uses (instructions.txt's "For each r in
'  results", "Repeat until fuel is 5, stop the loop.", "When region is
'  ...:"), unlike 'with' (zero uses anywhere in the corpus today). IN.10
'  adds return/exit-sub/exit-function - mLoopBreak's own shape, scoped
'  to the nearest enclosing CALL FRAME rather than the nearest
'  enclosing loop; see mProcReturn's own module-level note and
'  CallUserProc below. IN3.6 adds on-error/goto/label/resume - real
'  syntax that WAS a named remainder (only ever emitted as the
'  standard vla-fail wrapper BuildSub puts around every procedure when
'  step tracking is on, which every interpreter-targeted compile so
'  far has deliberately compiled with off) until Try:'s own on-error
'  use, unconditional regardless of step tracking, was actually
'  reached by a live run - see mErrMode's own module-level note for
'  the mechanism.
' ---------------------------------------------------------------------
Private Sub ExecStmt(ByVal f As Variant, ByVal frame As Object)
    If Not IsList(f) Then
        ' A bare atom in statement position - degenerate but legal
        ' (mirrors VLA.bas's own EmitStmt, which refuses this; the
        ' skeleton is more permissive on purpose: "no error handling").
        EvalExpr f, frame
        Exit Sub
    End If
    Dim lst As Collection
    Set lst = f
    Dim h As String
    h = VLA_Identity.Fold(HeadSym(lst))
    Select Case h
        Case "dim": ExecDim lst, frame
        Case "const": ExecConst lst, frame
        Case "set!": ExecSet lst, frame
        Case "obj-set!": ExecSet lst, frame
        Case "if": ExecIf lst, frame
        Case "for": ExecFor lst, frame
        Case "for-each": ExecForEach lst, frame
        Case "for-each-row": ExecForEachRow lst, frame
        Case "while": ExecWhile lst, frame
        Case "do-until": ExecDoUntil lst, frame
        Case "select": ExecSelect lst, frame
        Case "exit-for", "exit-do": mLoopBreak = True
        Case "exit-sub", "exit-function": mProcReturn = True
        Case "return": ExecReturn lst, frame
        Case "begin": ExecBegin lst, frame
        Case "debug-print": ExecDebugPrint lst, frame
        Case ".": ExecDotCall lst, frame
        Case "on-error": ExecOnError lst
        Case "goto": mGotoLabel = SymText(Nth(lst, 2))
        Case "label"
            ' IN3.6: a jump target only - ExecBody's own label search
            ' finds this form by its position, never by executing it;
            ' reached directly only when control simply falls through
            ' to it in sequence (a Try with no error, past its own
            ' recovery), where it is correctly a no-op.
        Case "resume": ExecResume lst
        Case Else
            ' Not one of the dedicated statement forms - fall through
            ' to expression dispatch and discard the result, mirroring
            ' VLA.bas's own EmitCallStmt (an unrecognized head in
            ' statement position becomes "Call foo(...)" there; here
            ' it becomes "evaluate foo as an expression"). This is
            ' exactly what most such statements ARE - a VLA_Runtime
            ' helper called for its side effect (VlaEnsureSheet "Foo"),
            ' say - so it reaches EvalDynamicHead's own tiers (place
            ' helpers/dotted globals/builtins/VLA_Runtime helpers)
            ' rather than refusing before ever trying. If nothing
            ' resolves it, EvalDynamicHead's own raise names the head -
            ' more specific than a generic catch-all here could say
            ' (still true for 'with'/named-arg templates - IN.2's
            ' remaining hand-work, unimplemented anywhere, so they
            ' surface the same way).
            EvalExpr f, frame
    End Select
End Sub

' IN.10: (return <expr>) - VLA.bas's own EmitReturn (VLA.bas:2313)
' mirrored exactly: assigns the return value AND unwinds immediately,
' not merely a value-producing form. A bare (return) just unwinds,
' leaving mProcReturnValue at whatever CallUserProc initialized it to
' (Empty) - the same "Exit Function with no prior assignment returns
' the type's default" VBA gives a real compiled function. Unlike
' EmitReturn, this does not refuse (return <expr>) inside a Sub (no
' per-call "am I in a function" context is threaded through ExecBody/
' ExecIf/loop bodies to check against) - more permissive than the
' emitter on purpose, IN0.5's own "no error handling" precedent, not a
' silently different value: CallUserProc only ever reads
' mProcReturnValue back out for isFunc procedures.
Private Sub ExecReturn(lst As Collection, frame As Object)
    If lst.Count >= 2 Then
        AssignVar mProcReturnValue, EvalExpr(Nth(lst, 2), frame)
    End If
    mProcReturn = True
End Sub

' IN3.6: (on-error goto <label>) arms mErrMode/mErrLabel; (on-error
' goto 0) disarms - "0" is VBA's own spelling for "no handler active"
' (RestoreHandlerVla's own literal text, VLA_English.bas), not a real
' label, matched here the same way EmitStmt's own emission relies on
' real VBA's GoTo 0 to mean at runtime, since this interpreter has no
' such primitive of its own to lean on; (on-error resume-next) arms
' mErrMode = "resume-next" (show-all-rows's own shape,
' english.vla). No stack, mirroring VLA.bas's own EmitStmt "on-
' error" case exactly: a fresh (on-error ...) always REPLACES whatever
' handler, if any, was already active, the same way a fresh real VBA
' `On Error` statement does - not layered atop it.
Private Sub ExecOnError(lst As Collection)
    Dim k As String
    k = VLA_Identity.Fold(SymText(Nth(lst, 2)))
    Select Case k
        Case "resume-next"
            mErrMode = "resume-next"
            mErrLabel = ""
        Case "goto"
            Dim tgt As String
            tgt = SymText(Nth(lst, 3))
            If tgt = "0" Then
                mErrMode = ""
                mErrLabel = ""
            Else
                mErrMode = "goto"
                mErrLabel = tgt
            End If
        Case Else
            VLA_Messages.RaiseMsg "interp-on-error-bad-shape"
    End Select
End Sub

' IN3.6: (resume <label>) - the only shape Case "try" ever emits
' (ParseStmt's own header comment: the recovery paragraph "runs under
' the RESTORED step handler", reached via exactly this hop). Two
' things happen, matching real VBA's Resume exactly, not just a goto:
' the jump itself (mGotoLabel, found by ExecBody's own search the same
' way a literal (goto ...) is), AND exiting the "a handler is active"
' state - mCaughtErr* cleared included, per the Try: comment's own
' reasoning ("VBA cannot trap an error raised while a handler is
' active... Resume exits that state (and clears Err) before any
' recovery sentence executes") - this interpreter's own shadow of that
' clear (mCaughtErrDesc's own module-level note has the full reasoning
' for why it is a shadow and not the real Err object). Because this
' runs AFTER the recovery paragraph's own (set! vla-problem
' err.description) has already read it, clearing it here loses
' nothing. Bare (resume) - real VBA's "re-run the failing statement" -
' and (resume next) - "skip to the statement after it" - are named
' non-goals, not built: Case "try" never emits either shape (grepped),
' and "re-run the failing statement" needs state (which statement, in
' which body) nothing else in this interpreter tracks. Refused in
' words rather than silently mishandled.
Private Sub ExecResume(lst As Collection)
    If lst.Count < 2 Then
        VLA_Messages.RaiseMsg "interp-resume-bare-unsupported"
    End If
    Dim tgt As String
    tgt = SymText(Nth(lst, 2))
    If VLA_Identity.Fold(tgt) = "next" Then
        VLA_Messages.RaiseMsg "interp-resume-next-unsupported"
    End If
    mErrMode = ""
    mErrLabel = ""
    mCaughtErrNum = 0
    mCaughtErrSrc = ""
    mCaughtErrDesc = ""
    mGotoLabel = tgt
End Sub

Private Sub ExecDim(lst As Collection, frame As Object)
    Dim nm As String
    nm = SymText(Nth(lst, 2))
    Dim t As String
    If lst.Count >= 3 Then t = VLA_Identity.Fold(SymText(Nth(lst, 3)))
    Dim v As Variant
    Select Case t
        Case "double", "long", "integer", "single", "currency"
            v = 0
        Case "string"
            v = ""
        Case "boolean"
            v = False
        Case Else
            v = Empty
    End Select
    VLA_Runtime.VlaDictSet frame, nm, v
End Sub

' IN.3: (const name value) | (const name type value) - VLA.bas's own
' EmitConstCore (VLA.bas:2600) mirrored: a real head this interpreter had
' NO case for at all (grepped - "const" never appeared in this module
' before this pass), a genuine pre-existing gap this item's own corpus
' read surfaced, unrelated to module-scope visibility itself. Before this
' fix, instructions.txt's own top-level "(const hot-pink ...)" (VLA_English.
' bas's ParseDefine) fell through ExecStmt's Case Else into expression
' dispatch, where "const" resolves as neither a user procedure nor any
' builtin/dot-member/runtime-helper tier - an "unrecognized head" raise
' before the program's very first real statement. The type token at
' position 3, when present, is emitter-only signature noise (VBA needs
' "As String" to fix a Const's storage type; this frame is dynamically
' typed, same reasoning ExecDim's own Select Case already applies) - read
' past, not used to coerce the value.
Private Sub ExecConst(lst As Collection, frame As Object)
    Dim nm As String
    nm = SymText(Nth(lst, 2))
    Dim v As Variant
    If lst.Count >= 4 Then
        AssignVar v, EvalExpr(Nth(lst, 4), frame)
    Else
        AssignVar v, EvalExpr(Nth(lst, 3), frame)
    End If
    VLA_Runtime.VlaDictSet frame, nm, v
End Sub

' IN2.0: the place can now be a plain variable name (IN0.5's original
' scope) OR a computed place - (. obj member ...) or a bare object-
' returning call like (range "a1") - closing IN0.5's own "out of
' scope on purpose" note. Also reached directly for 'obj-set!'
' (ExecStmt aliases it here). *Correction, IN.11, found by a live run
' rather than assumed:* this note used to claim obj-set! and set! to a
' plain variable are the SAME operation here, reasoning that a VlaDict
' needs no Let/Set distinction to STORE a value - true, but incomplete:
' it says nothing about WHICH value the source program's own Let/Set
' choice means to store when v IsObject. Real VBA's plain Let DOES
' still invoke the object's default member even though the interpreter's
' frame could hold the raw reference just fine - `(set! v (range
' "b2"))`, english.vla's own real "value in cell" shape, needs v
' to end up holding the CELL'S VALUE, not the Range object, to match
' real compiled VBA's `v = Range("b2")`. Now honored: `set!` to a plain
' variable coerces an object value via the default member (below,
' this Sub's own note has the mechanism); `obj-set!` never does. The
' distinction still matters, and is honored, for a COMPUTED place too
' (DynamicSet tries Let before Set - see its own note).
' IN.3: a bare place ATOM containing a "." (application.screenupdating,
' activewindow.freezepanes, activesheet.name, ...) is a dotted-global
' chain, the exact same thing a zero-arg CALL form of the same text
' means in expression position ((activesheet.name)) - not a plain
' variable whose name happens to contain a period. Found live, not read:
' every one of these (set-screen-updating/paste-values/freeze-top-row/
' add-sheet-called's own macros) was silently landing in the frame dict
' as a fake variable named e.g. "application.screenupdating" instead of
' ever touching the real Application object - no crash, because
' VlaDictSet/VlaDictSet never raise for an unusual key, which is exactly
' why this one went unnoticed through every reload so far: a WRONG,
' SILENT no-op, not a loud one. "Turn off screen updating." (instructions.txt
' line 20, among the first ten real sentences in the file) never
' actually turned off screen updating under the interpreter, this whole
' session, and nothing said so.
Private Sub ExecSet(lst As Collection, frame As Object)
    Dim place As Variant
    AssignVar place, Nth(lst, 2)
    Dim v As Variant
    AssignVar v, EvalExpr(Nth(lst, 3), frame)

    If Not IsList(place) Then
        Dim plainName As String
        plainName = SymText(place)
        Dim dotAt2 As Long
        dotAt2 = InStr(plainName, ".")
        If dotAt2 > 0 Then
            Dim globalReceiver As Object
            Set globalReceiver = ResolveGlobalReceiver(VLA_Identity.Fold(Left$(plainName, dotAt2 - 1)))
            If Not globalReceiver Is Nothing Then
                WalkMemberSet globalReceiver, Mid$(plainName, dotAt2 + 1), v
                LogEffect "set: " & plainName & " = " & CStr(v)
                Exit Sub
            End If
        End If
        ' IN.11: "Set X to value in cell Y." compiles to (set! x (range
        ' "...")) - the PLACE, not (. (range "...") value) - relying on
        ' real VBA's own implicit Let coercion (english.vla's own
        ' test: "Set v to value in cell B2. => (set! v (range "b2"))",
        ' the exact shape "Set spaced-check to value in cell 'Q1
        ' Data'!A1." compiles to). A real compiled `x = Range("b2")`
        ' (no `Set` keyword) invokes the object's default member (Value,
        ' for a Range) automatically; this module's own header note above
        ' argued a VlaDict needs no Let/Set distinction to STORE a value,
        ' true but incomplete - it says nothing about WHICH value the
        ' source program's own semantics mean to store. `obj-set!` (real
        ' VBA `Set`) still keeps the bare object reference exactly as
        ' before, checked directly by the form's own head rather than
        ' guessed from the value's shape - only `set!` (real VBA `Let`)
        ' coerces. The coercion itself is a plain Let done in THIS
        ' interpreter's own VBA (`coerced = v`), not a hand-picked
        ' Range-specific member: VBA's own assignment operator already
        ' knows how to find an object's default member, so this reuses
        ' that instead of reimplementing it, and correctly falls through
        ' to storing the raw object when there is no default member (a
        ' Collection from `(new Collection)`, VBA's own real behavior
        ' for `Set`-worthy objects that Let cannot coerce - though `(new
        ' Collection)` is always paired with obj-set! in today's corpus,
        ' never reaching this branch at all).
        If IsObject(v) And VLA_Identity.Fold(HeadSym(lst)) = "set!" Then
            Dim coerced As Variant
            On Error Resume Next
            coerced = v
            Dim coerceFailed As Boolean
            coerceFailed = (Err.Number <> 0)
            On Error GoTo 0
            If Not coerceFailed Then
                VLA_Runtime.VlaDictSet frame, plainName, coerced
                Exit Sub
            End If
        End If
        VLA_Runtime.VlaDictSet frame, plainName, v
        Exit Sub
    End If

    Dim placeLst As Collection
    Set placeLst = place
    Dim placeHead As String
    placeHead = VLA_Identity.Fold(HeadSym(placeLst))
    ' PF.4a: (set! (arr i) v) - array-element assignment. Compiled
    ' already handles this for free: EmitStmt's "set!" case is plain
    ' text substitution (EmitExpr(place) = EmitExpr(value)), and
    ' EmitExpr's generic Case Else emits (arr i) as "arr(i)" - real
    ' VBA array-index-assignment syntax, no compiler awareness needed.
    ' The interpreter had no such case until now: every non-'.' place
    ' in the existing corpus was a Range/object with a default Value
    ' member (the Else branch below), so a plain array place always
    ' failed "interp-set-place-not-object" - never exercised, not
    ' silently broken. Checked here, ahead of that fallback, only when
    ' the place has exactly one index and its head names a variable
    ' currently holding a real array - never true for any place that
    ' already worked before this, so this is purely additive.
    ' VlaDictGet copies the array out (a Variant-boxed SafeArray copies
    ' on assignment in VBA, unlike an object reference), so mutating
    ' the local copy alone would never be visible to a later read of
    ' the same variable - the mutated array is explicitly re-stored
    ' under its own name after the write, not mutated "in place".
    ' Only 1-D indexing is handled - nothing in scope today (for-each-
    ' row's own per-row array) needs more, and VBA gives no clean way
    ' to forward a variable number of subscripts to Let-assignment.
    ' Bug caught live (owner-run VlaSelfTestsAll, first pass): VlaDictGet
    ' is deliberately loud on a miss (RaiseRuntimeMsg "rt-dict-key-
    ' missing" - VLA_Runtime.bas's own documented "loud step error"
    ' contract), not a safe try-get returning Empty - a plain (range
    ' "a1") place also has Count = 2, and "range" was never a stored
    ' VLA variable, so calling VlaDictGet unconditionally here raised
    ' "there is nothing stored at key 'range'" for every ordinary Range
    ' place, not just array places. VlaDictHas (VLA_Runtime.bas) is the
    ' existing non-raising existence check built for exactly this "test
    ' before falling back" shape - its own header names this
    ' interpreter's module-scope read fallback as its first caller, the
    ' same precedent this should have followed from the start.
    Dim arrCandidate As Variant
    Dim isArrayPlace As Boolean
    If placeLst.Count = 2 Then
        If VLA_Runtime.VlaDictHas(frame, HeadSym(placeLst)) Then
            arrCandidate = VLA_Runtime.VlaDictGet(frame, HeadSym(placeLst))
            isArrayPlace = IsArray(arrCandidate)
        End If
    End If
    If placeHead = "." Then
        If placeLst.Count < 3 Then VLA_Messages.RaiseMsg "interp-dot-needs-object-member"
        Dim objVal As Variant
        AssignVar objVal, EvalExpr(Nth(placeLst, 2), frame)
        If Not IsObject(objVal) Then
            VLA_Messages.RaiseMsg "interp-dot-needs-object-left"
        End If
        Dim memTok As String
        memTok = SymText(Nth(placeLst, 3))
        WalkMemberSet objVal, memTok, v
        LogEffect "set: ." & memTok & " on " & TypeName(objVal) & " = " & CStr(v)
    ElseIf isArrayPlace Then
        Dim arrName As String
        arrName = HeadSym(placeLst)
        Dim idx As Long
        idx = CLng(EvalExpr(Nth(placeLst, 2), frame))
        arrCandidate(idx) = v
        VLA_Runtime.VlaDictSet frame, arrName, arrCandidate
        LogEffect "set: " & arrName & "(" & idx & ") = " & CStr(v)
    Else
        ' Honest, documented simplification: a bare computed place
        ' that is NOT a '.' form - (set! (range "a1") x) - is one of
        ' the new object-returning place helpers used directly. VBA's
        ' own default-property assignment (Range("a1") = x sets
        ' .Value) is what the emitter's plain text substitution
        ' already relies on for this exact corpus shape; this
        ' interpreter has no compiler to fall back on, so it assigns
        ' the object's "Value" member directly. Every real use in
        ' today's corpus is a Range, whose default member IS Value -
        ' documented, not proven general for every place helper.
        Dim placeObj As Variant
        AssignVar placeObj, EvalExpr(placeLst, frame)
        If Not IsObject(placeObj) Then
            VLA_Messages.RaiseMsg "interp-set-place-not-object"
        End If
        DynamicSet placeObj, "Value", v
        LogEffect "set: " & TypeName(placeObj) & ".Value = " & CStr(v)
    End If
End Sub

Private Sub ExecIf(lst As Collection, frame As Object)
    Dim test As Boolean
    test = CBool(EvalExpr(Nth(lst, 2), frame))
    Dim i As Long
    Dim cl As Collection
    Dim ch As String
    For i = 3 To lst.Count
        Set cl = Nth(lst, i)
        ch = VLA_Identity.Fold(HeadSym(cl))
        Select Case ch
            Case "then"
                If test Then
                    ExecBody cl, 2, frame
                    Exit Sub
                End If
            Case "elseif"
                If test Then Exit Sub    ' an earlier clause already ran
                test = CBool(EvalExpr(Nth(cl, 2), frame))
                If test Then
                    ExecBody cl, 3, frame
                    Exit Sub
                End If
            Case "else"
                If Not test Then ExecBody cl, 2, frame
                Exit Sub
        End Select
    Next
End Sub

Private Sub ExecFor(lst As Collection, frame As Object)
    Dim hdr As Collection
    Set hdr = Nth(lst, 2)
    Dim v As String
    v = SymText(Nth(hdr, 1))
    Dim startN As Double, endN As Double, stepN As Double
    startN = CDbl(EvalExpr(Nth(hdr, 2), frame))
    endN = CDbl(EvalExpr(Nth(hdr, 3), frame))
    stepN = 1
    If hdr.Count >= 4 Then stepN = CDbl(EvalExpr(Nth(hdr, 4), frame))
    Dim i As Double
    i = startN
    Do While (stepN > 0 And i <= endN) Or (stepN < 0 And i >= endN)
        VLA_Runtime.VlaDictSet frame, v, i
        ExecBody lst, 3, frame
        If mProcReturn Or Len(mGotoLabel) > 0 Then Exit Do
        If mLoopBreak Then
            mLoopBreak = False
            Exit Do
        End If
        i = i + stepN
    Loop
End Sub

' IN2.3: (for-each (v coll) body...) - coll is any expression yielding
' an enumerable object (a VLA list is a plain Collection; a dict walk
' arrives the same way, through (vladictpairs d) - EvalDynamicHead's
' existing VLA_Runtime fallback tier already reaches that helper, so
' this form needs no dict-vs-list case of its own). AssignVar, not a
' plain "=", to capture the object reference correctly (the Let-
' without-Set gotcha IN2.0's own history note already found once).
Private Sub ExecForEach(lst As Collection, frame As Object)
    Dim hdr As Collection
    Set hdr = Nth(lst, 2)
    Dim v As String
    v = SymText(Nth(hdr, 1))
    Dim collVal As Variant
    AssignVar collVal, EvalExpr(Nth(hdr, 2), frame)
    Dim itm As Variant
    For Each itm In collVal
        VLA_Runtime.VlaDictSet frame, v, itm
        ExecBody lst, 3, frame
        If mProcReturn Or Len(mGotoLabel) > 0 Then Exit For
        If mLoopBreak Then
            mLoopBreak = False
            Exit For
        End If
    Next
End Sub

' PF.4c: (for-each-row (row rng) body...) - the interpreter's own twin
' of VLA.bas's EmitForEachRow. Bulk-read rng once (VLA_Runtime.
' VlaSlabRead), walk its rows as a plain VBA array held entirely in
' this Sub's own locals - never boxed into the frame Dictionary the way
' a VLA-level variable would be, so there's no name-collision question
' the compiled side's mSlabCounter scheme has to solve at all: this
' bookkeeping lives on the real VBA call stack, invisible to the VLA
' namespace. `row` itself IS stored in the frame under its own name
' each iteration (VlaDictSet), exactly like ExecFor/ExecForEach's own
' loop variable - PF.4a's array-element set!/read parity is what makes
' (row 2)/(set! (row 2) ...) work against it once it's there.
' All-or-nothing commit, matched deliberately against the COMPILED
' side's own real behavior, not assumed symmetric: EmitReturn emits a
' bare Exit Sub/Exit Function (VLA.bas) - unconditionally leaving the
' WHOLE procedure, bypassing EmitForEachRow's own VlaSlabWrite call no
' matter how deep the nesting - so a (return ...) inside a for-each-row
' body discards everything on the compiled side, the SAME as an error,
' never merely "up to the return" the way break does. mProcReturn/
' mGotoLabel are checked FIRST here and skip VlaSlabWrite entirely to
' match that; only a plain mLoopBreak (no pending return/goto) commits
' everything computed so far, mirroring the compiled side's own Exit
' For, which falls through to its write-back line normally.
Private Sub ExecForEachRow(lst As Collection, frame As Object)
    Dim hdr As Collection
    Set hdr = Nth(lst, 2)
    Dim v As String
    v = SymText(Nth(hdr, 1))
    Dim rngVal As Variant
    AssignVar rngVal, EvalExpr(Nth(hdr, 2), frame)
    If Not TypeOf rngVal Is Range Then VLA_Messages.RaiseMsg "interp-for-each-row-needs-range"
    Dim rng As Range
    Set rng = rngVal

    Dim arr As Variant
    arr = VLA_Runtime.VlaSlabRead(rng)
    Dim nCols As Long
    nCols = UBound(arr, 2) - LBound(arr, 2) + 1

    Dim abandonWrite As Boolean
    Dim i As Long, c As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        Dim rowArr() As Variant
        ReDim rowArr(1 To nCols)
        For c = 1 To nCols
            rowArr(c) = arr(i, c)
        Next c
        VLA_Runtime.VlaDictSet frame, v, rowArr
        ExecBody lst, 3, frame
        Dim rowBack As Variant
        rowBack = VLA_Runtime.VlaDictGet(frame, v)
        For c = 1 To nCols
            arr(i, c) = rowBack(c)
        Next c
        If mProcReturn Or Len(mGotoLabel) > 0 Then
            abandonWrite = True
            Exit For
        End If
        If mLoopBreak Then
            mLoopBreak = False
            Exit For
        End If
    Next i
    If Not abandonWrite Then VLA_Runtime.VlaSlabWrite arr, rng
End Sub

Private Sub ExecWhile(lst As Collection, frame As Object)
    Do While CBool(EvalExpr(Nth(lst, 2), frame))
        ExecBody lst, 3, frame
        If mProcReturn Or Len(mGotoLabel) > 0 Then Exit Do
        If mLoopBreak Then
            mLoopBreak = False
            Exit Do
        End If
    Loop
End Sub

' IN2.3: (do-until cond body...) - While's inverted-sense sibling,
' same shape as EmitStmt's own "Do Until ... Loop" (tested before each
' pass, like While, sense flipped - not tested AFTER, which VBA's own
' Do..Loop Until would mean).
Private Sub ExecDoUntil(lst As Collection, frame As Object)
    Do Until CBool(EvalExpr(Nth(lst, 2), frame))
        ExecBody lst, 3, frame
        If mProcReturn Or Len(mGotoLabel) > 0 Then Exit Do
        If mLoopBreak Then
            mLoopBreak = False
            Exit Do
        End If
    Loop
End Sub

' IN2.3: (select expr (case (v1 v2 ...) body...) ... (case-else body...))
' - mirrors EmitSelect's own simplified semantics exactly: equality
' against a comma-list of values per case, first match wins,
' case-else optional and checked last regardless of where it appears
' (matching VBA's own Select Case, which also allows Case Else
' anywhere but only reaches it after every real Case has missed).
' No "Is >"/"To" range forms - EmitSelect doesn't support them either,
' so there is nothing here to be behind on. Plain VBA "=" for the
' match test, not a VLA expression form: select's own case-matching
' has nothing to do with whatever "=" does or does not mean as VLA
' expression syntax (a separate, pre-existing gap - EvalExpr has no
' "=" case today, untouched by this item), since select never needs
' to emit or evaluate a "(= a b)" form to do its own matching.
Private Sub ExecSelect(lst As Collection, frame As Object)
    Dim testVal As Variant
    AssignVar testVal, EvalExpr(Nth(lst, 2), frame)
    Dim elseCl As Collection
    Dim i As Long, j As Long
    Dim cl As Collection, ch As String, vals As Collection
    Dim caseVal As Variant
    For i = 3 To lst.Count
        Set cl = Nth(lst, i)
        ch = VLA_Identity.Fold(HeadSym(cl))
        Select Case ch
            Case "case"
                Set vals = Nth(cl, 2)
                For j = 1 To vals.Count
                    AssignVar caseVal, EvalExpr(Nth(vals, j), frame)
                    If testVal = caseVal Then
                        ExecBody cl, 3, frame
                        Exit Sub
                    End If
                Next
            Case "case-else"
                Set elseCl = cl
            Case Else
                VLA_Messages.RaiseMsg "interp-select-bad-clause", "head", ch
        End Select
    Next
    If Not elseCl Is Nothing Then ExecBody elseCl, 2, frame
End Sub

Private Sub ExecBegin(lst As Collection, frame As Object)
    ExecBody lst, 2, frame
End Sub

Private Sub ExecDebugPrint(lst As Collection, frame As Object)
    Dim i As Long
    Dim parts As String
    For i = 2 To lst.Count
        If Len(parts) > 0 Then parts = parts & " "
        parts = parts & CStr(EvalExpr(Nth(lst, i), frame))
    Next
    Debug.Print parts
    LogEffect "debug-print: " & parts
End Sub

' IN2.0: '.' in statement position - (. obj member ...) called for its
' side effect, the return value discarded (VLA.bas's own EmitStmt "."
' case: "Call range(...).clearcontents" - a statement, not an
' assignment). Logged as an effect: unlike debug-print/set!, this is
' the interpreter's first form whose whole POINT is an action on a
' real object (ws.Activate, range.ClearContents, ...), so IN.3.5's
' effect log widens to record it.
Private Sub ExecDotCall(lst As Collection, frame As Object)
    If lst.Count < 3 Then VLA_Messages.RaiseMsg "interp-dot-call-needs-object-member"
    Dim objVal As Variant
    AssignVar objVal, EvalExpr(Nth(lst, 2), frame)
    If Not IsObject(objVal) Then
        VLA_Messages.RaiseMsg "interp-dot-needs-object-left"
    End If
    Dim memTok As String
    memTok = SymText(Nth(lst, 3))

    ' IN2.5: an all-keyword argument list (the only shape today's
    ' corpus uses - remove-duplicates/autofilter/protect/... - see
    ' EvalKeywordArgs's own note) routes to the named-argument
    ' dispatch table instead of CallByName, which cannot carry a
    ' keyword. Checked by peeking at the first argument only - a call
    ' that starts positional and switches to keyword partway through
    ' still reaches EvalPositionalArgs below, which refuses it in
    ' words rather than silently taking the wrong branch.
    If lst.Count >= 4 Then
        Dim firstArgTok As Variant
        AssignVar firstArgTok, Nth(lst, 4)
        If IsKeywordTok(firstArgTok) Then
            Dim kwArgs As Collection
            Set kwArgs = EvalKeywordArgs(lst, 4, frame)
            WalkMemberNamedCall objVal, memTok, kwArgs
            LogEffect "call: ." & memTok & " on " & TypeName(objVal) & " kwargs=" & JoinKwArgs(kwArgs)
            Exit Sub
        End If
    End If

    Dim argVals As Variant
    argVals = EvalPositionalArgs(lst, 4, frame)
    WalkMemberCall objVal, memTok, argVals
    LogEffect "call: ." & memTok & " on " & TypeName(objVal) & _
              IIf(ArgCount(argVals) > 0, " args=" & JoinArgs(argVals), "")
End Sub

' ---------------------------------------------------------------------
'  IN3_5.0/IN2.0: the effect log - every runtime effect this
'  interpreter performs, in order, as text. Before IN2.0, debug-print
'  was the ONLY externally observable effect the ten-form skeleton
'  could produce (dim/set! to a plain variable only mutate the
'  in-memory frame - internal bookkeeping, the same distinction
'  VerifyReport's own cell reads draw for the emitter: it checks what
'  a program left BEHIND, not every assignment along the way). IN2.0
'  adds two more genuine effects that widen this log the same way: a
'  '.' statement call (ExecDotCall) and a set! to a COMPUTED place
'  (ExecSet) - both act on a real object outside this interpreter's
'  own frame, so both are logged; set! to a plain variable still is
'  not, unchanged. mEffectLog is reset at the top of every
'  VlaInterpret call (S3.1-safe: never module-level *cached* across
'  calls, matching mMacros/mHeadAliases in VLA.bas), so
'  VlaInterpreterEffectLog always answers for the most recent run only.
' ---------------------------------------------------------------------
Private Sub LogEffect(ByVal text As String)
    If mEffectLog Is Nothing Then Set mEffectLog = New Collection
    mEffectLog.Add text
End Sub

' The interpreter's own VlaWriteGoldens: a plain-text, ordered record
' of everything the most recent VlaInterpret call did, formatted so a
' golden file (VLA_Tests.bas's VlaWriteInterpreterGoldens) can commit
' it and a git diff becomes the witness, same discipline as
' instructions_golden.vla/.vba.
Public Function VlaInterpreterEffectLog() As String
    If mEffectLog Is Nothing Or mEffectLog.Count = 0 Then
        VlaInterpreterEffectLog = "(no effects)" & vbCrLf
        Exit Function
    End If
    Dim r As String
    Dim e As Variant
    For Each e In mEffectLog
        r = r & e & vbCrLf
    Next
    VlaInterpreterEffectLog = r
End Function

' IN2.3: stops at mLoopBreak (set by exit-for/exit-do), not just after
' the current statement - VBA's own "Exit For"/"Exit Do" stop
' executing immediately, not at the end of the current pass. Since
' ExecIf/ExecBegin both route their own branch/block bodies through
' this same function, one check here is enough for the flag to bubble
' up through any nesting depth to whichever loop executor owns it,
' with no loop-awareness needed anywhere in between. IN.10: also stops
' at mProcReturn (return/exit-sub/exit-function) - the same bubbling,
' but a loop executor only ever RESETS mLoopBreak (its own break);
' mProcReturn passes through every loop executor untouched, so it
' keeps unwinding past any number of nested loops until CallUserProc,
' the call's own boundary, finally consumes it.
' IN3.6: a goto/resume target (mGotoLabel), set either by a literal
' (goto <label>)/(resume <label>) statement or by a trapped error
' under an armed handler (ExecStmtTrapped), is searched for in THIS
' body's own statement range (FindLabelIndex) before anything else -
' found means this is the right nesting level, so the loop just jumps
' its own index there and keeps going, exactly like a real GoTo; not
' found means the label lives further out (this body is itself nested
' inside the one Try:'s goto/label pair was written into), so this
' body abandons the rest of its own statements and bubbles the still-
' pending flag up to its caller - mLoopBreak/mProcReturn's exact
' existing shape, applied to a search instead of a flag check. A
' pending mErrMode arms ExecStmtTrapped instead of the plain ExecStmt
' fast path, scoped to exactly one statement per iteration so this
' interpreter's own On Error state is never left active across
' statement boundaries (see that function's own header note).
Private Sub ExecBody(lst As Collection, ByVal fromIdx As Long, frame As Object)
    Dim i As Long
    i = fromIdx
    Do While i <= lst.Count
        If Len(mErrMode) > 0 Then
            ExecStmtTrapped Nth(lst, i), frame
        Else
            ExecStmt Nth(lst, i), frame
        End If
        If Len(mGotoLabel) > 0 Then
            Dim tgt As Long
            tgt = FindLabelIndex(lst, fromIdx, mGotoLabel)
            If tgt = 0 Then Exit Do
            mGotoLabel = ""
            i = tgt
        ElseIf mLoopBreak Or mProcReturn Then
            Exit Do
        End If
        i = i + 1
    Loop
End Sub

' IN3.6: wraps exactly one statement's execution with a real VBA On
' Error Resume Next, scoped to this one ExecStmt call and nowhere else
' - the only place this interpreter itself uses On Error. VBA's error
' handling is per-PROCEDURE, not global: none of ExecStmt/EvalExpr/
' ExecIf/the loop executors/etc. have any On Error of their own, so an
' error raised anywhere in a statement's entire evaluation tree - a
' deeply nested expression, a '.' member call, a loop inside the
' statement - propagates up the VBA call stack untouched until it
' reaches THIS frame's active Resume Next, exactly matching what a
' real (on-error goto ...)/(on-error resume-next) needs, with no
' changes required anywhere else in the module.
' The distinction the module's own IN3.6 header names as the real risk
' - trapping this interpreter's own bug/gap signaling as if it were a
' recoverable program error, vs. letting a genuine program error
' escape uncaught - is resolved by Err.Source, checked directly against
' every Err.Raise this project's own interpreter-reachable code
' actually uses (IsOwnDiagnosticSource), not guessed: ONLY
' "VLA-Interpreter" counts as this interpreter's own bug/gap signaling
' - every one of its raise sites (grepped, not sampled) names an
' arity mismatch, a malformed form, or an unsupported construct
' ("IN.2: '...' is not a form...", "this interpreter's dynamic
' dispatch supports up to 4..."), never a program-data condition. A
' real, recoverable program error - a genuine COM/Automation failure
' (activating a sheet that does not exist, instructions.txt's own first
' Try:), a native VBA runtime error (division by zero), OR a
' deliberate SD-10 "refuse in words" business-rule raise from
' VLA_Runtime.bas/VLA_English.bas (VlaCheckSheetName, and IN.12's own
' VlaCheckSheetAbsent - "a sheet called ... already exists" is exactly
' as recoverable as "activating a sheet that does not exist," the
' two are the same failure from opposite directions) - escapes this
' Sub's Resume Next and gets caught below, exactly like the emitter's
' real "On Error Goto Label" already does for these same raises,
' since compiled VBA never distinguishes by Err.Source at all. IN.12
' found this the hard way: VLA_Runtime/VLA_English were ORIGINALLY
' both in this list too, on the assumption that nothing in those
' modules' own call sites was interpreter-reachable at Try:-wrapped
' runtime - true by accident, until VlaCheckSheetAbsent became the
' first VLA_Runtime raise a live Try: ever actually hit, escaping
' uncaught into a hard VBA break instead of being caught the way the
' exact same statement already is under the emitter (SD-5: no
' backend may silently diverge). Narrowed to "VLA-Interpreter" alone
' once found, matching what was already true of the emitted backend.
Private Sub ExecStmtTrapped(ByVal f As Variant, ByVal frame As Object)
    On Error Resume Next
    ExecStmt f, frame
    If Err.Number <> 0 Then
        Dim num As Long, src As String, desc As String
        num = Err.Number: src = Err.Source: desc = Err.Description
        If IsOwnDiagnosticSource(src) Then
            ' Own diagnostics must actually escape, not be reabsorbed
            ' by the very On Error Resume Next this Sub just set - only
            ' disarmed here, immediately before re-raising, never on
            ' the real-error path below (see mCaughtErrDesc's own note
            ' for why: an On Error statement resets the real VBA Err
            ' object, so this Sub must not execute one on a path the
            ' interpreted program still needs Err's shadow to survive).
            On Error GoTo 0
            Err.Raise num, src, desc
        End If
        ' A real, genuinely trappable program error - captured into
        ' this interpreter's OWN shadow (mCaughtErrDesc's own note),
        ' never the live Err object, which this Sub's own On Error
        ' Resume Next would otherwise silently wipe by the time Try:'s
        ' recovery paragraph gets to read it, one statement later.
        mCaughtErrNum = num
        mCaughtErrSrc = src
        mCaughtErrDesc = desc
        Select Case mErrMode
            Case "goto"
                mGotoLabel = mErrLabel
            Case "resume-next"
                ' Swallowed, exactly like real VBA's Resume Next:
                ' ExecBody's own loop simply advances to the next
                ' statement with no jump at all.
        End Select
    End If
End Sub

Private Function IsOwnDiagnosticSource(ByVal src As String) As Boolean
    IsOwnDiagnosticSource = (src = "VLA-Interpreter")
End Function

' IN3.6: scans lst's own statement range for (label <name>) - the
' nesting-level test ExecBody's own header note describes: found here
' means the goto/resume target belongs to THIS body (Try:'s own
' goto/label pairs are always siblings of each other, never split
' across a nested if/loop), so the caller jumps in place; not found
' means it belongs to an enclosing body, so the caller bubbles the
' still-pending mGotoLabel up unchanged. Scans the WHOLE fromIdx..Count
' range regardless of the current position, so a backward jump - not
' needed by anything in today's corpus, every real jump is forward -
' would still resolve correctly.
Private Function FindLabelIndex(lst As Collection, ByVal fromIdx As Long, ByVal name As String) As Long
    Dim folded As String
    folded = VLA_Identity.Fold(name)
    Dim i As Long
    Dim cl As Variant
    Dim clLst As Collection
    For i = fromIdx To lst.Count
        AssignVar cl, Nth(lst, i)
        If IsList(cl) Then
            Set clLst = cl
            If VLA_Identity.Fold(HeadSym(clLst)) = "label" Then
                If VLA_Identity.Fold(SymText(Nth(clLst, 2))) = folded Then
                    FindLabelIndex = i
                    Exit Function
                End If
            End If
        End If
    Next
End Function

' ---------------------------------------------------------------------
'  The expression forms: IN0.5's original three (+/-/>) plus IN2.0's
'  '.' (member access) and EvalDynamicHead's non-core dispatch (place
'  helpers/dotted globals/builtins/VLA_Runtime helpers) - plus the
'  atom base case (number, string literal, true/false, or a variable
'  reference), unchanged.
' ---------------------------------------------------------------------
Private Function EvalExpr(ByVal x As Variant, ByVal frame As Object) As Variant
    If Not IsList(x) Then
        Dim s As String
        s = CStr(x)
        If Left$(s, 1) = Chr$(34) Then
            EvalExpr = StrLitContent(x)
            Exit Function
        End If
        If IsNumeric(s) Then
            EvalExpr = CDbl(s)
            Exit Function
        End If
        Dim folded As String
        folded = VLA_Identity.Fold(s)
        Select Case folded
            Case "true": EvalExpr = True
            Case "false": EvalExpr = False
            Case Else
                ' IN.3: a bare atom containing a "." (rows.count,
                ' english.vla's own last-filled-row macro -
                ' (cells rows.count c)) is a dotted-global READ, the
                ' expression-position sibling of ExecSet's own bare-
                ' dotted-global SET fix (that Sub's own note has the
                ' full reasoning) - not a variable whose name happens to
                ' contain a period. Checked before the constant tables
                ' and the frame lookup chain below since neither an
                ' Excel/VBA constant nor a VLA identifier (hyphens only,
                ' MangleIdent's own territory) is ever spelled with a
                ' dot - no real ambiguity to break by checking first.
                Dim dotAt3 As Long
                dotAt3 = InStr(folded, ".")
                If dotAt3 > 0 Then
                    ' IN3.6: "err.description"/"err.number"/"err.source" -
                    ' Try:'s own recovery reads err.description (VLA_
                    ' English.bas's Case "try") - resolved from this
                    ' interpreter's OWN shadow of the most recently
                    ' trapped error (mCaughtErrDesc's own module-level
                    ' note has the full reasoning), NOT the real VBA Err
                    ' object via ResolveGlobalReceiver/WalkMemberGet: by
                    ' the time recovery code reads it, this interpreter's
                    ' own plumbing (ExecStmtTrapped's per-statement On
                    ' Error Resume Next, re-executed for every statement
                    ' in between, including the harmless (label ...) the
                    ' jump itself lands on) has already reset the real
                    ' Err object's properties - checked directly against
                    ' a live host run, not assumed correct.
                    If Left$(folded, dotAt3 - 1) = "err" Then
                        Select Case Mid$(folded, dotAt3 + 1)
                            Case "description": EvalExpr = mCaughtErrDesc
                            Case "number": EvalExpr = mCaughtErrNum
                            Case "source": EvalExpr = mCaughtErrSrc
                            Case Else
                                VLA_Messages.RaiseMsg "interp-err-field-unsupported", "field", Mid$(folded, dotAt3 + 1)
                        End Select
                        Exit Function
                    End If
                    Dim globalReceiver3 As Object
                    Set globalReceiver3 = ResolveGlobalReceiver(Left$(folded, dotAt3 - 1))
                    If Not globalReceiver3 Is Nothing Then
                        AssignVar EvalExpr, WalkMemberGet(globalReceiver3, Mid$(folded, dotAt3 + 1), Array())
                        Exit Function
                    End If
                End If
                Dim constFound As Boolean
                Dim constVal As Double
                constVal = ResolveExcelConstant(folded, constFound)
                If Not constFound Then constVal = ResolveVbConstant(folded, constFound)
                If constFound Then
                    EvalExpr = constVal
                ElseIf VLA_Runtime.VlaDictHas(frame, s) Then
                    AssignVar EvalExpr, VLA_Runtime.VlaDictGet(frame, s)
                ElseIf Not (mModuleFrame Is Nothing) And Not (frame Is mModuleFrame) _
                        And VLA_Runtime.VlaDictHas(mModuleFrame, s) Then
                    ' IN.3: a call frame (CallUserProc) has no binding for
                    ' s of its own - fall back to module scope (this
                    ' module's own header note on mModuleFrame) before
                    ' giving up. Skipped entirely for the module frame's
                    ' own execution (main's body): frame IS mModuleFrame
                    ' there, so the miss above already covers it.
                    AssignVar EvalExpr, VLA_Runtime.VlaDictGet(mModuleFrame, s)
                Else
                    ' IN.11: a BARE, undotted global receiver name -
                    ' "activewindow" alone, not "activewindow.freezepanes"
                    ' - found live: "Unfreeze panes." compiles to (set!
                    ' (. activewindow freezepanes) false), a real '.'
                    ' LIST form whose first argument is the plain symbol
                    ' "activewindow" - a genuinely different shape from
                    ' the dotted-atom case above (dotAt3 > 0), which never
                    ' fires here since "activewindow" alone has no dot at
                    ' all. ExecSet's own '.'-form branch evaluates that
                    ' argument as an ordinary expression
                    ' (EvalExpr(Nth(placeLst, 2), frame)), which lands
                    ' here, in the atom base case, same as any bare name
                    ' would. Checked LAST, only once neither frame has a
                    ' binding, on purpose: a real corpus variable always
                    ' wins if one happens to share the name (never
                    ' actually reached today - VLA_English.bas's own
                    ' hyphenated identifiers can't collide with these),
                    ' the same fallback-not-shadow precedent mModuleFrame
                    ' itself already sets for the frame chain above.
                    Dim bareReceiver As Object
                    Set bareReceiver = ResolveGlobalReceiver(folded)
                    If Not bareReceiver Is Nothing Then
                        AssignVar EvalExpr, bareReceiver
                    Else
                        ' Neither frame has it, and it is not a known
                        ' global receiver either - VlaDictGet's own raise
                        ' names the miss the same way it always has.
                        AssignVar EvalExpr, VLA_Runtime.VlaDictGet(frame, s)
                    End If
                End If
        End Select
        Exit Function
    End If

    Dim lst As Collection
    Set lst = x
    Dim h As String
    h = VLA_Identity.Fold(HeadSym(lst))
    Select Case h
        Case "+"
            EvalExpr = EvalChain(lst, frame, "+")
        Case "-"
            If lst.Count = 2 Then
                EvalExpr = -CDbl(EvalExpr(Nth(lst, 2), frame))
            Else
                EvalExpr = EvalChain(lst, frame, "-")
            End If
        Case "*", "&", "/", "\", "=", "<>", "<", ">", "<=", ">=", _
             "and", "or", "xor", "mod", "is", "like", "imp", "eqv"
            AssignVar EvalExpr, EvalOpChain(lst, frame, h)
        Case "not"
            Dim notOperand As Variant
            AssignVar notOperand, EvalExpr(Nth(lst, 2), frame)
            EvalExpr = Not notOperand
        Case "."
            AssignVar EvalExpr, EvalDotForm(lst, frame)
        Case "new"
            ' IN.3: (new Collection) - VLA.bas's own EmitExpr Case "new"
            ' (VLA.bas:2661) mirrored: the type-name argument is a
            ' LITERAL symbol, never evaluated as an expression - falling
            ' through to EvalDynamicHead's generic dispatch before this
            ' case existed did exactly that (EvalPositionalArgs treats
            ' every remaining list element as an expression), which is
            ' why "(new Collection)" used to raise "there is nothing
            ' stored at key 'Collection'" - a plain variable-lookup
            ' miss on the type name, not a "new" problem as such. Found
            ' live (VerifyReportInterpreter, english.vla's V3
            ' keyed-memory macros - IN2.3's own header note already
            ' named "(new Collection)" as a real, pre-existing gap this
            ' interpreter had never touched). VBA has no generic runtime
            ' "New <type name string>" - CreateObject only reaches
            ' OLE-registered ProgIDs, and Collection is not one - so
            ' this is a literal `New Collection`, not reflection, scoped
            ' to the one type name today's corpus actually uses (checked
            ' directly: no other `(new ...)` type appears anywhere in
            ' instructions.txt/english.vla/prelude.vla).
            Dim newType As String
            newType = VLA_Identity.Fold(SymText(Nth(lst, 2)))
            Select Case newType
                Case "collection"
                    Set EvalExpr = New Collection
                Case Else
                    VLA_Messages.RaiseMsg "interp-new-unsupported-type", "type", SymText(Nth(lst, 2))
            End Select
        Case "quote"
            ' REPLEVAL.0: found live - eval "(quote (1 2 3))" raised
            ' "'1' is not a form ..." because this Select Case had no
            ' "quote" arm at all, so an unrecognized head fell to
            ' EvalDynamicHead, which tried to CALL "quote" and evaluate
            ' its argument (1 2 3) AS CODE - a list headed by the
            ' number 1, not a form. VLA.bas's own EmitExpr/EmitQuote
            ' has always had this right (P.L5: quote is a DATA
            ' context); the interpreter simply never grew the matching
            ' case, because nothing before REPLEVAL.0's own eval verb
            ' ever asked this interpreter to evaluate a BARE (quote
            ' ...) expression on its own - LISTOPS's own (quote ...)
            ' uses are expand-time only and never reach the
            ' interpreter at all, and every OTHER (quote ...) in the
            ' existing corpus sits inside a larger form some other
            ' Case already handles structurally (e.g. `array`'s own
            ' elements, or a defmacro template LISTOPS itself
            ' resolves) rather than being EVALUATED here as a value in
            ' its own right. A real parity hole, not a regression -
            ' TestExprParity's own job, simply never previously
            ' exercised on this shape.
            If lst.Count <> 2 Then VLA_Messages.RaiseMsg "interp-quote-arity"
            AssignVar EvalExpr, EvalQuoteDatum(Nth(lst, 2))
        Case "array"
            ' G6: mirrors VLA.bas's own EmitExpr Case "array" - a
            ' list-valued slot's runtime value, this interpreter's own
            ' equivalent of VBA's Array(...). Given an explicit case for
            ' the same reason that module's own comment gives (an
            ' unrecognized head falls through to EvalDynamicHead, which
            ' tries a VLA_Runtime helper lookup by name - genuinely
            ' different behavior from the compile side's harmless
            ' case-insensitive coincidence, so backend parity needs this
            ' spelled out here too, not assumed).
            Dim arrCount As Long
            arrCount = lst.Count - 1
            If arrCount <= 0 Then
                EvalExpr = Array()
            Else
                Dim arrVals() As Variant
                ReDim arrVals(0 To arrCount - 1)
                Dim arrIdx As Long
                For arrIdx = 1 To arrCount
                    AssignVar arrVals(arrIdx - 1), EvalExpr(Nth(lst, arrIdx + 1), frame)
                Next
                EvalExpr = arrVals
            End If
        Case "interpolate"
            ' L-INTERPOLATE: mirrors VLA.bas's own EmitExpr Case
            ' "interpolate" - see this module's own LINTERPOLATE.0
            ' header note for the full design (named holes, literal-
            ' template requirement, "&" coercion). Given an explicit
            ' case for the same reason "array" (immediately above) has
            ' one: an unrecognized head would fall through to
            ' EvalDynamicHead's generic dispatch instead of this
            ' primitive's own semantics.
            EvalExpr = EvalInterpolateCall(lst, frame)
        Case Else
            AssignVar EvalExpr, EvalDynamicHead(h, lst, frame)
    End Select
End Function

' REPLEVAL.0: EvalExpr's own "quote" case, above - mirrors VLA.bas's
' QuoteDatum (EmitQuote's own recursive helper) EXACTLY, one important
' difference by necessity: that one builds VBA SOURCE TEXT ("Array(1,
' 2, 3)", to be compiled); this one builds the REAL VBA VALUE (an
' actual Variant array), since the interpreter has no separate compile
' step to hand text to. Never calls EvalExpr on any element - quote's
' whole point is that its datum is DATA, not code to run, the same
' distinction ExpandMacros's own "quote is a DATA context" skip
' protects at expand time. A list becomes an array, recursively; a
' string literal becomes its own content; a number becomes a Double
' (IsNumeric/CDbl, the SAME atom-parsing convention EvalExpr's own base
' case already uses two cases above - not re-litigating LISTOPS-
' EXPAND's own Val()/Str$() locale argument here, since that guardrail
' was scoped for a NEW expand-time primitive family, not this
' interpreter's existing, already-established atom convention); any
' other bare symbol becomes that string verbatim, matching QuoteDatum's
' own "symbols cannot contain a double quote, plain wrapping is exact"
' reasoning - quoting a symbol yields a STRING at runtime here exactly
' as it does in compiled VBA, not some distinct "symbol" value.
Private Function EvalQuoteDatum(ByVal v As Variant) As Variant
    If IsList(v) Then
        Dim lst As Collection
        Set lst = v
        Dim n As Long
        n = lst.Count
        If n = 0 Then
            EvalQuoteDatum = Array()
            Exit Function
        End If
        Dim arr() As Variant
        ReDim arr(0 To n - 1)
        Dim i As Long
        For i = 1 To n
            AssignVar arr(i - 1), EvalQuoteDatum(Nth(lst, i))
        Next
        EvalQuoteDatum = arr
        Exit Function
    End If
    Dim s As String
    s = CStr(v)
    If Left$(s, 1) = Chr$(34) Then
        EvalQuoteDatum = StrLitContent(v)
    ElseIf IsNumeric(s) Then
        EvalQuoteDatum = CDbl(s)
    Else
        EvalQuoteDatum = s
    End If
End Function

' IN2.6: Excel's named enum constants this corpus actually uses, as a
' number - see this module's own IN2.6 header note for how the set was
' measured and why it stops at 13 rather than cataloguing Excel's whole
' constant surface. folded is already VLA_Identity.Fold'd by the caller,
' so this Select Case matches lowercase literals only. found is ByRef
' rather than a sentinel return value (Empty/Null both already carry
' real meaning elsewhere in this interpreter - a plain flag says "not a
' constant" unambiguously instead of borrowing an overloaded value).
Private Function ResolveExcelConstant(ByVal folded As String, ByRef found As Boolean) As Double
    found = True
    Select Case folded
        Case "xlleft": ResolveExcelConstant = -4131            ' XlHAlign / XlDirection
        Case "xlcenter": ResolveExcelConstant = -4108           ' XlHAlign / XlVAlign
        Case "xlright": ResolveExcelConstant = -4152            ' XlHAlign / XlDirection
        Case "xlascending": ResolveExcelConstant = 1            ' XlSortOrder
        Case "xldescending": ResolveExcelConstant = 2           ' XlSortOrder
        Case "xlyes": ResolveExcelConstant = 1                  ' XlYesNoGuess
        Case "xlno": ResolveExcelConstant = 2                   ' XlYesNoGuess
        Case "xlpastevalues": ResolveExcelConstant = -4163      ' XlPasteType
        Case "xlnone": ResolveExcelConstant = -4142             ' XlLineStyle / XlPattern
        Case "xlcontinuous": ResolveExcelConstant = 1           ' XlLineStyle
        Case "xlup": ResolveExcelConstant = -4162                ' XlDirection
        Case "xltypepdf": ResolveExcelConstant = 0              ' XlFixedFormatType
        Case "xlcalculationautomatic": ResolveExcelConstant = -4105  ' XlCalculation
        Case "xlcalculationmanual": ResolveExcelConstant = -4135     ' XlCalculation
        Case "xlsrcrange": ResolveExcelConstant = 1             ' XlListObjectSourceType
        ' G-STRUCT (owner-caught live, interpret-and-trace on
        ' instructions.txt): "there is nothing stored at key
        ' 'xlpasteformats'" - copy-formats-only/-formulas-only/-widths-
        ' only and delete-shift-up/-left all pass a bare xl* symbol this
        ' table didn't have yet. Same one-name-at-a-time discipline as
        ' every entry above: exactly what this pass's own rules use, not
        ' a general PasteType/DeleteShiftDirection catalogue.
        Case "xlpasteformats": ResolveExcelConstant = -4122     ' XlPasteType
        Case "xlpasteformulas": ResolveExcelConstant = -4123    ' XlPasteType
        Case "xlpastecolumnwidths": ResolveExcelConstant = 8    ' XlPasteType
        Case "xlshiftup": ResolveExcelConstant = -4162          ' XlDeleteShiftDirection (same number as xlup/XlDirection above - one Long value, two different enum types, both real)
        Case "xlshifttoleft": ResolveExcelConstant = -4159      ' XlDeleteShiftDirection
        Case Else
            found = False
    End Select
End Function

' IN.3: VBA's OWN intrinsic constants (vbRed, vbTextCompare, ...) -
' distinct from ResolveExcelConstant's xl* Excel-type-library enums
' above, a genuinely different source (the VBA language itself, not
' Excel's object model), so kept separate rather than folded into the
' same table under one misleading name. Found live, not read:
' "Make cell C4 yellow." (english.vla's set-fill-color macro,
' passing the bare symbol vbyellow) raised "there is nothing stored at
' key 'vbyellow'" the first time instructions.txt reached it - this
' interpreter resolved Excel's own named constants (IN2.6) but never
' VBA's. Same one-name-at-a-time discipline, scoped to exactly what is
' measured, not a general VBA-constant catalogue: vbred/vbyellow
' (english.vla's "make cell ... red|yellow") and vbtextcompare
' (english.vla's contains/starts-with tests, "instr ... vbtextcompare").
Private Function ResolveVbConstant(ByVal folded As String, ByRef found As Boolean) As Double
    found = True
    Select Case folded
        Case "vbred": ResolveVbConstant = vbRed
        Case "vbyellow": ResolveVbConstant = vbYellow
        Case "vbtextcompare": ResolveVbConstant = vbTextCompare
        Case Else
            found = False
    End Select
End Function

' (. obj member ...) as an EXPRESSION - reads a property or calls a
' method for its return value (EvalDynamicHead's Dynamic* fallback
' logic decides which). member may itself be a dotted chain
' (font.bold): VLA.bas's own EmitDotText passes such a token through
' as one piece of literal VBA member-access text for the compiler to
' resolve; this interpreter has no compiler, so WalkMemberGet/
' DescendToParent walk it hop by hop instead.
' F.2 G-TABLES: the keyword-argument peek ExecDotCall's own IN2.5 note
' already explains, mirrored here - this function's OWN gap, not a new
' one: a '.' call whose return value is actually used (make-table's
' (. (. activesheet.listobjects add :sourcetype ...) name) - the Add
' call is the OBJECT of an outer property access, not a bare statement)
' used to fall straight to EvalPositionalArgs below regardless of
' keyword tokens, where three keyword pairs count as six positional
' slots and hit that function's own 4-argument ceiling. Every keyword-
' bearing '.' call this corpus had before F.2 happened to sit in
' STATEMENT position (ExecDotCall's own "the only shape today's corpus
' uses" note, now no longer true) - expression position was simply
' never exercised with keyword arguments until this rule.
Private Function EvalDotForm(lst As Collection, frame As Object) As Variant
    If lst.Count < 3 Then VLA_Messages.RaiseMsg "interp-dot-call-needs-object-member"
    Dim objVal As Variant
    AssignVar objVal, EvalExpr(Nth(lst, 2), frame)
    If Not IsObject(objVal) Then
        VLA_Messages.RaiseMsg "interp-dot-needs-object-left"
    End If
    Dim memTok As String
    memTok = SymText(Nth(lst, 3))
    If lst.Count >= 4 Then
        Dim firstArgTok As Variant
        AssignVar firstArgTok, Nth(lst, 4)
        If IsKeywordTok(firstArgTok) Then
            Dim kwArgs As Collection
            Set kwArgs = EvalKeywordArgs(lst, 4, frame)
            AssignVar EvalDotForm, WalkMemberNamedCall(objVal, memTok, kwArgs)
            Exit Function
        End If
    End If
    Dim argVals As Variant
    argVals = EvalPositionalArgs(lst, 4, frame)
    AssignVar EvalDotForm, WalkMemberGet(objVal, memTok, argVals)
End Function

' ---------------------------------------------------------------------
'  IN2.0: resolves a head symbol with no dedicated Case in EvalExpr,
'  in the order the roadmap's own IN.2 measurement laid the non-core
'  dispatch surface out - see this module's own IN2.0 header note for
'  the full reasoning behind each tier. Shared by expression position
'  (here) and statement position indirectly: a bare call used only for
'  its side effect still reaches this same function, since ExecStmt's
'  fallback for an unrecognized head is "evaluate it as an expression
'  and discard the result" (mirrors VLA.bas's EmitCallStmt, which
'  wraps an unrecognized head in "Call"). IN.10: a call to a declared
'  user procedure is checked FIRST, ahead of every other tier and
'  ahead of EvalPositionalArgs below - a named-argument call
'  ((stamp :row-number 2 ...)) would otherwise hit EvalPositionalArgs's
'  own keyword-token refusal before this function ever got a chance to
'  recognize "stamp" as a real procedure. This is the one dispatch
'  point real corpus calls reach either way: statement position
'  ((tidy-up)) arrives via ExecStmt's own Case-Else fallback into
'  EvalExpr; expression position ((tax 100)) arrives directly - both
'  end up here, so no second interception point is needed.
' ---------------------------------------------------------------------
Private Function EvalDynamicHead(ByVal h As String, lst As Collection, frame As Object) As Variant
    Dim proc As Collection
    Set proc = LookupProc(h)
    If Not proc Is Nothing Then
        AssignVar EvalDynamicHead, CallUserProc(proc, lst, frame)
        Exit Function
    End If

    Dim argVals As Variant
    argVals = EvalPositionalArgs(lst, 2, frame)

    ' The six place helpers are called NATIVELY, not through CallByName
    ' - a real, empirically-found gap this pass's first draft got
    ' wrong: CallByName(Application, "range", ...) fails in real Excel
    ' (both Get and Method), because Range/Cells/Rows/Columns/
    ' Worksheets/Workbooks are Application's PARAMETERIZED DEFAULT
    ' members, exposed through a hidden interface late-bound CallByName
    ' cannot reach - a documented Excel/VBA limitation, not a bug in
    ' this interpreter's own dispatch logic. Calling Range(...)/
    ' Cells(...)/etc. directly, unqualified, is also more faithful to
    ' what the emitter itself does (VLA.bas's own EmitExpr Case Else
    ' emits a bare "range(...)" call, never "application.range(...)")
    ' - this was arguably the more correct design even before the
    ' failure surfaced it.
    Select Case h
        Case "range"
            Select Case ArgCount(argVals)
                Case 1: AssignVar EvalDynamicHead, Range(ArgAt(argVals, 0))
                Case 2: AssignVar EvalDynamicHead, Range(ArgAt(argVals, 0), ArgAt(argVals, 1))
                Case Else
                    VLA_Messages.RaiseMsg "interp-range-arity", "n", ArgCount(argVals)
            End Select
            Exit Function
        Case "cells"
            Select Case ArgCount(argVals)
                Case 0: AssignVar EvalDynamicHead, Cells
                Case 2: AssignVar EvalDynamicHead, Cells(ArgAt(argVals, 0), ArgAt(argVals, 1))
                Case Else
                    VLA_Messages.RaiseMsg "interp-cells-arity", "n", ArgCount(argVals)
            End Select
            Exit Function
        Case "rows"
            Select Case ArgCount(argVals)
                Case 0: AssignVar EvalDynamicHead, Rows
                Case 1: AssignVar EvalDynamicHead, Rows(ArgAt(argVals, 0))
                Case Else
                    VLA_Messages.RaiseMsg "interp-rows-arity", "n", ArgCount(argVals)
            End Select
            Exit Function
        Case "columns"
            Select Case ArgCount(argVals)
                Case 0: AssignVar EvalDynamicHead, Columns
                Case 1: AssignVar EvalDynamicHead, Columns(ArgAt(argVals, 0))
                Case Else
                    VLA_Messages.RaiseMsg "interp-columns-arity", "n", ArgCount(argVals)
            End Select
            Exit Function
        Case "worksheets"
            Select Case ArgCount(argVals)
                Case 0: AssignVar EvalDynamicHead, Worksheets
                Case 1: AssignVar EvalDynamicHead, Worksheets(ArgAt(argVals, 0))
                Case Else
                    VLA_Messages.RaiseMsg "interp-worksheets-arity", "n", ArgCount(argVals)
            End Select
            Exit Function
        Case "workbooks"
            Select Case ArgCount(argVals)
                Case 0: AssignVar EvalDynamicHead, Workbooks
                Case 1: AssignVar EvalDynamicHead, Workbooks(ArgAt(argVals, 0))
                Case Else
                    VLA_Messages.RaiseMsg "interp-workbooks-arity", "n", ArgCount(argVals)
            End Select
            Exit Function
        Case "make-button"
            ' IN.7 (button-click half): the English rule
            ' "make a button {c:text} at cell {r:cell}" compiles to
            ' (make-button {c} (range {r})) - a purely imperative
            ' action, no meaningful return value, so it rides
            ' EvalDynamicHead exactly like the six place helpers above
            ' for the same reason (a Form Control button's .Add/.Name/
            ' .OnAction are ordinary Excel object-model calls, not
            ' reachable through the vocabulary/macro path any more
            ' than Range/Cells were).
            Select Case ArgCount(argVals)
                Case 2: MakeVlaButton CStr(ArgAt(argVals, 0)), ArgAt(argVals, 1)
                Case Else
                    VLA_Messages.RaiseMsg "interp-make-button-arity", "n", ArgCount(argVals)
            End Select
            Exit Function

        ' IN.3: Application.WorksheetFunction's own members are the SAME
        ' class of gap the six place helpers above were already fixed
        ' for, one level deeper - found live, not read: "Set grand to sum
        ' of range B2:B3." (sum-of's own macro expansion,
        ' (application.worksheetfunction.sum r)) silently returned a
        ' raw-pointer-shaped garbage number instead of 45, no error at
        ' all - CallByName's own documented failure mode for a
        ' parameterized member exposed through a hidden interface (this
        ' function's own header note on the place helpers above has the
        ' full reasoning: "confuses late-bound CallByName the same way
        ' Application's parameterized Range/Cells/etc. did - just
        ' silently instead of raising"). Every WorksheetFunction member
        ' takes an argument, so every one of them is this same shape, not
        ' an occasional edge case. Dispatched natively for the seven real
        ' members english.vla's own macros use - its own "F.1
        ' residual" comment names exactly these seven
        ' (sum/average/max/min/vlookup/sumif/countif), not a general
        ' WorksheetFunction catalogue, the same one-name-at-a-time
        ' discipline IN2.6's own Excel-constant table already set.
        Case "application.worksheetfunction.sum"
            If ArgCount(argVals) <> 1 Then VLA_Messages.RaiseMsg "interp-sum-arity", "n", ArgCount(argVals)
            AssignVar EvalDynamicHead, Application.WorksheetFunction.Sum(ArgAt(argVals, 0))
            Exit Function
        Case "application.worksheetfunction.average"
            If ArgCount(argVals) <> 1 Then VLA_Messages.RaiseMsg "interp-average-arity", "n", ArgCount(argVals)
            AssignVar EvalDynamicHead, Application.WorksheetFunction.Average(ArgAt(argVals, 0))
            Exit Function
        Case "application.worksheetfunction.max"
            If ArgCount(argVals) <> 1 Then VLA_Messages.RaiseMsg "interp-max-arity", "n", ArgCount(argVals)
            AssignVar EvalDynamicHead, Application.WorksheetFunction.Max(ArgAt(argVals, 0))
            Exit Function
        Case "application.worksheetfunction.min"
            If ArgCount(argVals) <> 1 Then VLA_Messages.RaiseMsg "interp-min-arity", "n", ArgCount(argVals)
            AssignVar EvalDynamicHead, Application.WorksheetFunction.Min(ArgAt(argVals, 0))
            Exit Function
        Case "application.worksheetfunction.countif"
            If ArgCount(argVals) <> 2 Then VLA_Messages.RaiseMsg "interp-countif-arity", "n", ArgCount(argVals)
            AssignVar EvalDynamicHead, Application.WorksheetFunction.CountIf(ArgAt(argVals, 0), ArgAt(argVals, 1))
            Exit Function
        Case "application.worksheetfunction.sumif"
            ' sum-of-where's own macro passes (c e r) - (range, criteria,
            ' sum-range), WorksheetFunction.SumIf's own real signature
            ' (Range, Criteria, [SumRange]) - three args every real call
            ' supplies, not the two-required/one-optional shape SumIf
            ' allows in general (no real corpus use omits SumRange).
            If ArgCount(argVals) <> 3 Then VLA_Messages.RaiseMsg "interp-sumif-arity", "n", ArgCount(argVals)
            AssignVar EvalDynamicHead, Application.WorksheetFunction.SumIf(ArgAt(argVals, 0), ArgAt(argVals, 1), ArgAt(argVals, 2))
            Exit Function
        Case "application.worksheetfunction.vlookup"
            If ArgCount(argVals) <> 4 Then VLA_Messages.RaiseMsg "interp-vlookup-arity", "n", ArgCount(argVals)
            AssignVar EvalDynamicHead, Application.WorksheetFunction.VLookup(ArgAt(argVals, 0), ArgAt(argVals, 1), ArgAt(argVals, 2), ArgAt(argVals, 3))
            Exit Function
    End Select

    Dim dotAt As Long
    dotAt = InStr(h, ".")
    If dotAt > 0 Then
        Dim firstSeg As String
        firstSeg = Left$(h, dotAt - 1)
        Dim receiver As Object
        Set receiver = ResolveGlobalReceiver(firstSeg)
        If Not receiver Is Nothing Then
            Dim restPath As String
            restPath = Mid$(h, dotAt + 1)
            AssignVar EvalDynamicHead, WalkMemberGet(receiver, restPath, argVals)
            Exit Function
        End If
    End If

    Dim builtinResult As Variant
    Dim builtinHandled As Boolean
    AssignVar builtinResult, TryEvalBuiltin(h, argVals, builtinHandled)
    If builtinHandled Then
        AssignVar EvalDynamicHead, builtinResult
        Exit Function
    End If

    Dim runtimeResult As Variant
    Dim runtimeHandled As Boolean
    AssignVar runtimeResult, TryRuntimeHelper(h, argVals, runtimeHandled)
    If runtimeHandled Then
        AssignVar EvalDynamicHead, runtimeResult
        Exit Function
    End If

    ' PF.4a: (arr i) as an EXPRESSION - reading an array element, the
    ' read-side counterpart to ExecSet's own new array-element-
    ' assignment case above. Compiled needs nothing here either -
    ' EmitExpr's generic Case Else already emits "arr(i)", real VBA
    ' array indexing. The interpreter had no path to this at all: every
    ' tier above (user proc, the six place helpers, dotted global
    ' receiver, builtin, runtime helper) requires h to match something
    ' specific, and none of them ever checked "is h a local variable
    ' bound to a plain array" - reading (arr i) always fell all the way
    ' through to interp-head-unresolved. Checked last, only when every
    ' other tier has already failed to claim h, so this can never
    ' shadow an existing resolution (a real proc, a place helper, a
    ' runtime helper) - purely additive, the same safety argument as
    ' ExecSet's own new case. Guarded with VlaDictHas, not a bare
    ' VlaDictGet - the sibling bug ExecSet's own case just caught live:
    ' VlaDictGet raises loudly on a miss (VLA_Runtime.bas's own "loud
    ' step error" contract), so an unrelated, genuinely-unresolved h
    ' reaching this tier would otherwise be reported as "there is
    ' nothing stored at key '<h>'" instead of the correct, more useful
    ' interp-head-unresolved just below.
    If VLA_Runtime.VlaDictHas(frame, h) Then
        Dim arrReadCandidate As Variant
        arrReadCandidate = VLA_Runtime.VlaDictGet(frame, h)
        If IsArray(arrReadCandidate) And ArgCount(argVals) = 1 Then
            AssignVar EvalDynamicHead, arrReadCandidate(CLng(ArgAt(argVals, 0)))
            Exit Function
        End If
    End If

    VLA_Messages.RaiseMsg "interp-head-unresolved", "head", h
End Function

' IN.7 (button-click half): draws a Form Control button at a cell's
' position and wires it to this add-in's own single dispatch entry
' point, VLA_Events.VlaButtonClickDispatch - never mHostWorkbook's
' module, and never anything living in the user's own file. The
' OnAction string must name the ADD-IN's own workbook (bare
' ThisWorkbook, this code running physically inside it): the exact
' inverse mistake IN.6/IN.7's sheet-change half each hit once already
' (this module's own header note on why hb is threaded explicitly
' everywhere else), so it gets called out here rather than trusted to
' be obviously right a second time. `place` arrives already evaluated
' (a real Range - "(range {r})" in the compiled template), so its own
' .Worksheet, not ActiveSheet, decides where the button lands -
' correct regardless of which sheet happens to be active when this
' statement runs, the same "let the value carry its own sheet" trust
' EvalDynamicHead's other place helpers already place in a Range
' argument. Re-running the same caption on the same sheet replaces
' the old button outright (VlaRegisterSheetChangeHandler's own
' "current run wins" precedent) rather than piling up duplicates -
' without this, a second "Interpret Instructions" click would refuse
' outright the moment .Name is set to an already-used caption (Excel
' requires unique shape names per sheet).
Private Sub MakeVlaButton(ByVal caption As String, ByVal place As Variant)
    If Not TypeOf place Is Range Then
        VLA_Messages.RaiseMsg "interp-make-button-not-cell", "type", TypeName(place)
    End If
    Dim r As Range
    Set r = place
    Dim ws As Worksheet
    Set ws = r.Worksheet
    On Error Resume Next
    ws.Buttons(caption).Delete
    On Error GoTo 0
    Dim b As Button
    Set b = ws.Buttons.Add(r.Left, r.Top, 120, 24)
    b.Caption = caption
    b.Name = caption
    b.OnAction = "'" & ThisWorkbook.Name & "'!VlaButtonClickDispatch"
End Sub

' Excel's own well-known global objects, the receivers a dotted-global
' head's first segment (application.*, activesheet.*, ...) resolves
' against. Nothing here needs VBProject trust - these are ordinary
' Excel object-model reads, the same ones EN.1's capability probe
' would report as always available.
Private Function ResolveGlobalReceiver(ByVal name As String) As Object
    Select Case name
        Case "application": Set ResolveGlobalReceiver = Application
        Case "activesheet": Set ResolveGlobalReceiver = ActiveSheet
        Case "activeworkbook": Set ResolveGlobalReceiver = ActiveWorkbook
        Case "thisworkbook": Set ResolveGlobalReceiver = mHostWorkbook
        Case "activewindow": Set ResolveGlobalReceiver = ActiveWindow
        ' IN.3: the zero-argument shape of the five place helpers
        ' (EvalDynamicHead's own Select Case handles 0 AND N-argument
        ' forms for these same five names when CALLED, (cells)/(rows)/
        ' etc. - those never reach here, an exact-text Select Case match
        ' on the bare name intercepts them first) - added so a bare
        ' dotted atom rooted at one of them (rows.count, spreadsheet.
        ' vocab's own last-filled-row macro) resolves the same way
        ' application.screenupdating already does, rather than needing a
        ' second, parallel receiver table. Conceptually the same kind of
        ' thing either way: a global name with no arguments, usable as
        ' the root of a dotted chain.
        Case "cells": Set ResolveGlobalReceiver = Cells
        Case "rows": Set ResolveGlobalReceiver = Rows
        Case "columns": Set ResolveGlobalReceiver = Columns
        Case "worksheets": Set ResolveGlobalReceiver = Worksheets
        Case "workbooks": Set ResolveGlobalReceiver = Workbooks
        ' IN3.6: deliberately NO "err" case here - see EvalExpr's own
        ' dedicated "err." handling, checked before this function is
        ' ever reached for that root. The real VBA Err object cannot be
        ' trusted to still hold the failure it is read for by the time
        ' Try:'s recovery gets to read it - see mCaughtErrDesc's own
        ' module-level note for why.
    End Select
End Function

' A small, EXPLICITLY BOUNDED subset of the roadmap's own 17-name
' bare-VBA-builtin count - real VBA global functions with no object
' receiver, so CallByName cannot reach them at all (IN.0.5's own
' cost-measurement finding). Hand-written on purpose, one at a time,
' exactly as that measurement said this tier would cost - NOT the
' full catalog (mod/timer/timeserial are not here; mod is already the
' core '-'-style operator's sibling and was out of scope for this
' pass, timer/timeserial are lower-value and deferred). handled is
' False (return value Empty) for anything outside this subset, so
' EvalDynamicHead keeps trying its remaining tiers. msgbox/inputbox
' are deliberately NEVER exercised by the automated test suite - a
' modal dialog mid-VlaSelfTest would hang the whole run untrappably,
' the same "VBA's OWN modal during Application.Run" hazard this
' codebase already guards against elsewhere (VLA.bas's EmitCallStmt
' comment, S4's verdict).
Private Function TryEvalBuiltin(ByVal h As String, ByVal argVals As Variant, ByRef handled As Boolean) As Variant
    handled = True
    Select Case h
        Case "msgbox"
            MsgBox CStr(ArgAt(argVals, 0))
            LogEffect "msgbox: " & CStr(ArgAt(argVals, 0))
            TryEvalBuiltin = Empty
        Case "inputbox"
            TryEvalBuiltin = InputBox(CStr(ArgAt(argVals, 0)))
        Case "len"
            TryEvalBuiltin = Len(CStr(ArgAt(argVals, 0)))
        Case "trim"
            TryEvalBuiltin = Trim$(CStr(ArgAt(argVals, 0)))
        Case "lcase"
            TryEvalBuiltin = LCase$(CStr(ArgAt(argVals, 0)))
        Case "ucase"
            TryEvalBuiltin = UCase$(CStr(ArgAt(argVals, 0)))
        Case "left"
            TryEvalBuiltin = Left$(CStr(ArgAt(argVals, 0)), CLng(ArgAt(argVals, 1)))
        Case "right"
            TryEvalBuiltin = Right$(CStr(ArgAt(argVals, 0)), CLng(ArgAt(argVals, 1)))
        Case "round"
            If ArgCount(argVals) >= 2 Then
                TryEvalBuiltin = Round(CDbl(ArgAt(argVals, 0)), CLng(ArgAt(argVals, 1)))
            Else
                TryEvalBuiltin = Round(CDbl(ArgAt(argVals, 0)))
            End If
        Case "instr"
            TryEvalBuiltin = InStr(CStr(ArgAt(argVals, 0)), CStr(ArgAt(argVals, 1)))
        Case "isempty"
            TryEvalBuiltin = IsEmpty(ArgAt(argVals, 0))
        Case "now"
            TryEvalBuiltin = Now
        Case "date"
            TryEvalBuiltin = Date
        Case "time"
            TryEvalBuiltin = Time
        Case Else
            handled = False
    End Select
End Function

' Nearly all VLA_Runtime.bas helpers (VlaColor, VlaDictGet, VlaCount,
' ...) through ONE mechanism - Application.Run against the module by
' name - rather than 17 hand branches, since (the roadmap's own words)
' "they are already enumerated in one module with one calling
' convention." The three named in IN.12's own native tier just above
' are the declared exception, for a reason specific to THIS mechanism,
' not a retreat from it. CallByName cannot reach a standard-module
' Function at all (it
' dispatches on OBJECTS, not modules); Application.Run is VBA's own
' answer to that, and needs no VBProject trust (an existing, already-
' documented fact elsewhere in this codebase - VLA_IDE.bas's own
' VlaTrace comment). A call to a name that isn't a real VLA_Runtime
' function raises VBA error 1004 ("cannot run the macro..."), caught
' here as simply "not handled" so EvalDynamicHead's final raise names
' it cleanly instead of surfacing that cryptic text.
' IN.12: a small, native fast path for the VLA_Runtime helpers whose
' whole job is to validate and maybe raise (VlaCheckSheetName,
' VlaCheckSheetAbsent, VlaCheckRangeName) - needed for the same reason
' IN.11's own native Cases were: the generic mechanism below breaks for
' this specific shape. Proven with a standalone, zero-dependency repro
' (tools/VLA_Diag2.bas, scenario 1 - the SIMPLEST possible case, one
' Sub, one Err.Raise, one level of On Error Resume Next in the caller)
' that Application.Run does NOT propagate a target macro's own
' Err.Raise back to the caller's On Error Resume Next at all - it
' breaks straight through to an unhandled VBA error, unconditionally,
' regardless of how many levels of Resume Next sit above it. That
' means every VLA_Runtime raise ever routed through this function's
' own On Error Resume Next below (two lines down) has always had this
' bug; VlaCheckSheetName's five raise conditions simply never fired
' live before IN.12's own VlaCheckSheetAbsent became the first to
' actually hit it, inside a live Try: block, where the difference
' between "caught, exactly like the emitter" and "hard crash into the
' VBE debugger" is the whole point. Dispatched here as a direct Sub
' call - no Application.Run, no local error trap - so a raise
' propagates exactly like EvalDynamicHead's own raises already do:
' normally, up the real VBA call stack, to whichever caller has an
' active handler (ExecStmtTrapped's On Error Resume Next, IN3.6).
' Every OTHER VLA_Runtime helper this function dispatches to (VlaColor,
' VlaDictGet, VlaCount, ...) returns a normal value on success and is
' not yet known to have hit this live - the same latent class, named
' here rather than fixed blindly, matching this project's own "count
' before patching the next crash" lesson (TRENCHES.md IV) without
' expanding this pass past what the corpus actually exercises today.
Private Function TryRuntimeHelper(ByVal h As String, ByVal argVals As Variant, ByRef handled As Boolean) As Variant
    Select Case h
        Case "vlachecksheetname"
            VLA_Runtime.VlaCheckSheetName CStr(ArgAt(argVals, 0))
            handled = True
            Exit Function
        Case "vlachecksheetabsent"
            VLA_Runtime.VlaCheckSheetAbsent CStr(ArgAt(argVals, 0))
            handled = True
            Exit Function
        Case "vlacheckrangename"
            VLA_Runtime.VlaCheckRangeName CStr(ArgAt(argVals, 0))
            handled = True
            Exit Function
        ' SEC.8: a FOURTH native Case, and it joins the three above for
        ' precisely the reason IN.12's comment gives for them - this
        ' helper can now raise. Before SEC.8, VlaSendMail only ever
        ' displayed a draft or reported a missing attachment; the
        ' provenance guard at the top of it made "validate and maybe
        ' raise" part of its job, which is exactly the shape the generic
        ' Application.Run mechanism below is PROVEN to break
        ' (tools/VLA_Diag2.bas scenario 1). Live-caught, owner-run,
        ' 2026-09-08: the refusal surfaced as a raw "Run-time error '5'"
        ' VBE dialog with a Debug button instead of Frazaro's own modal,
        ' while the identical refusal on protect/unprotect (raised
        ' inline in DynamicNamedCall) presented correctly. Both of the
        ' generic path's possible outcomes are wrong here: the raise
        ' breaks through unhandled, and had it NOT, the
        ' "If Err.Number <> 0 Then handled = False" below would have
        ' swallowed the refusal and mis-reported it as
        ' interp-head-unresolved ("'vlasendmail' is not a form..."). A
        ' direct Sub call puts the raise back on the normal VBA call
        ' stack, where InterpretProgram's own handler catches it and
        ' LX.8's refuse-in-words doctrine actually holds.
        Case "vlasendmail"
            Select Case ArgCount(argVals)
                Case 3
                    VLA_Runtime.VlaSendMail ArgAt(argVals, 0), ArgAt(argVals, 1), ArgAt(argVals, 2)
                Case 4
                    VLA_Runtime.VlaSendMail ArgAt(argVals, 0), ArgAt(argVals, 1), ArgAt(argVals, 2), ArgAt(argVals, 3)
                Case Else
                    ' Any other arity is not this helper - fall through
                    ' to EvalDynamicHead's own clean refusal, the same
                    ' answer the generic path's Case Else gives.
                    handled = False
                    Exit Function
            End Select
            handled = True
            Exit Function

        ' IN.15: the remaining SIXTEEN. IN.11/IN.12/SEC.8 each added a
        ' native Case after a crash was reported, one helper at a time;
        ' this closes the class instead, and pins it with
        ' tools/check_runtime_raise_dispatch.ps1 so a seventeenth cannot
        ' arrive unnoticed. The rule the pin enforces is the one those
        ' three items were each following by hand: a helper needs a
        ' native Case exactly when IT CAN RAISE, because a raise routed
        ' through Application.Run below does not reach any caller's
        ' handler at all (this function's own header, and
        ' tools/VLA_Diag2.bas scenario 1).
        '
        ' SIXTEEN, NOT EIGHT, and the eight-way difference is the whole
        ' argument for mechanizing the count. Scoping this item by hand
        ' found the eight helpers with a raise site in their own body.
        ' The other eight raise only THROUGH RequirePivotTableByName,
        ' whose rt-pivot-not-found fires on a misspelled pivot table
        ' name - as reachable as any refusal in the module, and invisible
        ' to a scan that reads each helper's body alone.
        '
        ' Arguments are passed exactly as the generic path passed them -
        ' CStr where the helper declares String (the shape the four
        ' Cases above already use), the raw Variant everywhere else (the
        ' shape vlasendmail uses for its own Variant parameters). This
        ' change is about WHERE the error goes, and deliberately changes
        ' nothing about what the helper receives.
        Case "vlacolor"
            If Not ArityIs(argVals, 1, handled) Then Exit Function
            AssignVar TryRuntimeHelper, VLA_Runtime.VlaColor(ArgAt(argVals, 0))
            handled = True
            Exit Function
        Case "vladictget"
            If Not ArityIs(argVals, 2, handled) Then Exit Function
            AssignVar TryRuntimeHelper, VLA_Runtime.VlaDictGet(ArgAt(argVals, 0), ArgAt(argVals, 1))
            handled = True
            Exit Function
        Case "vlafreezepanes"
            If Not ArityIs(argVals, 1, handled) Then Exit Function
            VLA_Runtime.VlaFreezePanes ArgAt(argVals, 0)
            handled = True
            Exit Function
        Case "vlafillseries"
            If Not ArityIs(argVals, 4, handled) Then Exit Function
            VLA_Runtime.VlaFillSeries ArgAt(argVals, 0), ArgAt(argVals, 1), ArgAt(argVals, 2), CStr(ArgAt(argVals, 3))
            handled = True
            Exit Function
        Case "vlapivotrefresh"
            If Not ArityIs(argVals, 1, handled) Then Exit Function
            VLA_Runtime.VlaPivotRefresh CStr(ArgAt(argVals, 0))
            handled = True
            Exit Function
        Case "vlapivotdelete"
            If Not ArityIs(argVals, 1, handled) Then Exit Function
            VLA_Runtime.VlaPivotDelete CStr(ArgAt(argVals, 0))
            handled = True
            Exit Function
        Case "vlapivotclear"
            If Not ArityIs(argVals, 1, handled) Then Exit Function
            VLA_Runtime.VlaPivotClear CStr(ArgAt(argVals, 0))
            handled = True
            Exit Function
        Case "vlapivotrename"
            If Not ArityIs(argVals, 2, handled) Then Exit Function
            VLA_Runtime.VlaPivotRename CStr(ArgAt(argVals, 0)), CStr(ArgAt(argVals, 1))
            handled = True
            Exit Function
        Case "vlapivotsetrowlayout"
            If Not ArityIs(argVals, 2, handled) Then Exit Function
            VLA_Runtime.VlaPivotSetRowLayout CStr(ArgAt(argVals, 0)), CStr(ArgAt(argVals, 1))
            handled = True
            Exit Function
        Case "vlapivotchangesource"
            If Not ArityIs(argVals, 2, handled) Then Exit Function
            VLA_Runtime.VlaPivotChangeSource CStr(ArgAt(argVals, 0)), ArgAt(argVals, 1)
            handled = True
            Exit Function
        Case "vlapivotsetorientation"
            If Not ArityIs(argVals, 3, handled) Then Exit Function
            VLA_Runtime.VlaPivotSetOrientation CStr(ArgAt(argVals, 0)), ArgAt(argVals, 1), CStr(ArgAt(argVals, 2))
            handled = True
            Exit Function
        Case "vlapivotaddvalues"
            If Not ArityIs(argVals, 3, handled) Then Exit Function
            VLA_Runtime.VlaPivotAddValues CStr(ArgAt(argVals, 0)), ArgAt(argVals, 1), CStr(ArgAt(argVals, 2))
            handled = True
            Exit Function
        Case "vlapivotsetshowdetail"
            If Not ArityIs(argVals, 3, handled) Then Exit Function
            VLA_Runtime.VlaPivotSetShowDetail CStr(ArgAt(argVals, 0)), ArgAt(argVals, 1), ArgAt(argVals, 2)
            handled = True
            Exit Function
        Case "vlapivotsetsubtotals"
            If Not ArityIs(argVals, 3, handled) Then Exit Function
            VLA_Runtime.VlaPivotSetSubtotals CStr(ArgAt(argVals, 0)), ArgAt(argVals, 1), ArgAt(argVals, 2)
            handled = True
            Exit Function
        Case "vlapivotsetblankline"
            If Not ArityIs(argVals, 3, handled) Then Exit Function
            VLA_Runtime.VlaPivotSetBlankLine CStr(ArgAt(argVals, 0)), ArgAt(argVals, 1), ArgAt(argVals, 2)
            handled = True
            Exit Function
        Case "vlapivotsort"
            If Not ArityIs(argVals, 4, handled) Then Exit Function
            VLA_Runtime.VlaPivotSort CStr(ArgAt(argVals, 0)), CStr(ArgAt(argVals, 1)), CStr(ArgAt(argVals, 2)), CStr(ArgAt(argVals, 3))
            handled = True
            Exit Function
    End Select

    ' IN.15 (S3): "is this a real helper?" is answered HERE, BEFORE the
    ' call, by name. It used to be inferred AFTER the call from "did an
    ' error happen" - the "If Err.Number <> 0 Then handled = False" line
    ' that used to sit at the bottom of this function - and that one line
    ' conflated two unrelated questions: whether the name is a helper at
    ' all, and whether a helper that IS real deliberately refused. The
    ' first is a name-resolution fact, knowable before any code runs; the
    ' second is a result. Answering the first with the second is what made
    ' a refusal capable of surfacing as "'vlacolor' is not a form...",
    ' which is a confident wrong answer - worse than a crash.
    '
    ' VlaHelperManifest is the same list VLA_SentenceEngine's own
    ' EnglishResolveCheck already trusts for exactly this question, and
    ' HelperKey below reuses that function's own normalization so the two
    ' cannot drift into disagreeing about what a helper is called.
    '
    ' SOFT-FAILURE IS DELIBERATE AND MUST STAY THIS DIRECTION. The
    ' manifest returns "" when it cannot read its own source - no
    ' VBProject trust, no VLAr_Source sheet. Empty means "cannot verify,
    ' so fall back to trying the call", NEVER "nothing is a helper": the
    ' latter would refuse every runtime helper in the language on a host
    ' where the project is locked. manifestUsable carries that distinction
    ' past the call, so the recovery below can tell "a known helper
    ' failed" from "we never knew in the first place".
    Dim manifest As String
    On Error Resume Next
    manifest = HelperManifestCached()
    On Error GoTo 0
    Dim manifestUsable As Boolean
    manifestUsable = (Len(Trim$(manifest)) > 0)
    If manifestUsable Then
        If InStr(manifest, " " & HelperKey(h) & " ") = 0 Then
            ' Not a helper. A name-resolution answer, given without
            ' running anything - EvalDynamicHead's own clean
            ' interp-head-unresolved refusal names it.
            handled = False
            Exit Function
        End If
    End If

    Dim target As String
    target = "'" & ThisWorkbook.Name & "'!VLA_Runtime." & MangleIdent(h)
    Dim n As Long
    n = ArgCount(argVals)
    handled = True
    On Error Resume Next
    Err.Clear
    Select Case n
        Case 0: AssignVar TryRuntimeHelper, Application.Run(target)
        Case 1: AssignVar TryRuntimeHelper, Application.Run(target, ArgAt(argVals, 0))
        Case 2: AssignVar TryRuntimeHelper, Application.Run(target, ArgAt(argVals, 0), ArgAt(argVals, 1))
        Case 3: AssignVar TryRuntimeHelper, Application.Run(target, ArgAt(argVals, 0), ArgAt(argVals, 1), ArgAt(argVals, 2))
        Case 4: AssignVar TryRuntimeHelper, Application.Run(target, ArgAt(argVals, 0), ArgAt(argVals, 1), ArgAt(argVals, 2), ArgAt(argVals, 3))
        Case Else
            ' Arity this mechanism cannot express. Still a resolution
            ' answer, not a refusal - unchanged.
            handled = False
    End Select
    ' IN.15 (S3): what is left here is NOT "was this a helper" - the
    ' manifest settled that above. Per this function's own header, a
    ' target's Err.Raise never arrives here at all (it breaks straight
    ' through), and after the sixteen Cases above no helper that can
    ' raise reaches this tier anyway. So a non-zero Err at this point
    ' means the call itself could not be made against a name the manifest
    ' vouched for - VBA's 1004 "cannot run the macro", or an argument
    ' VBA would not coerce. That is worth saying plainly instead of
    ' claiming the name was never a helper.
    Dim failNum As Long
    Dim failDesc As String
    failNum = Err.Number
    failDesc = Err.Description
    ' Read BOTH before clearing the handler: On Error GoTo 0 resets Err.
    On Error GoTo 0
    If failNum <> 0 Then
        If Not manifestUsable Then
            ' The manifest could not be read, so this error is the only
            ' evidence available and it means what it always meant:
            ' probably not a helper. Exactly the pre-IN.15 behaviour, kept
            ' for exactly the path that still needs it.
            handled = False
            Exit Function
        End If
        VLA_Messages.RaiseMsg "interp-runtime-helper-call-failed", _
                              "head", h, "num", CStr(failNum), "desc", failDesc
    End If
End Function

' IN.15: the arity guard the sixteen native Cases above share. A call
' whose shape does not match the helper's own signature is not that
' helper - so it takes the same exit an unknown name takes (handled =
' False, then EvalDynamicHead's clean refusal), which is the answer
' vlasendmail's own Case Else and the generic path's Case Else both
' already give. Factored out because sixteen copies of the same
' four-line If is not clearer than one named predicate, and a predicate
' can be read once and trusted sixteen times.
Private Function ArityIs(ByVal argVals As Variant, ByVal n As Long, ByRef handled As Boolean) As Boolean
    If ArgCount(argVals) = n Then
        ArityIs = True
    Else
        ArityIs = False
        handled = False
    End If
End Function

' IN.15: VlaHelperManifest's own key shape. Copied deliberately from
' VLA_SentenceEngine's EnglishResolveCheck rather than invented here -
' the manifest strips underscores when it builds its names, so a lookup
' that did not strip them too would miss, and silently answer "not a
' helper" for a helper that exists. Both punctuation forms are removed
' for the same reason that function removes both.
Private Function HelperKey(ByVal s As String) As String
    HelperKey = VLA_Identity.Fold(Replace(Replace(s, "-", ""), "_", ""))
End Function

' IN.15: one manifest read per session, VLA_SentenceEngine's own
' mManifestCache precedent (its "V8: one manifest read per session"
' note) and for the same reason, which matters more here than there:
' VlaHelperManifest reads and folds the whole ~1600-line VLA_Runtime
' source every call, and this function sits in the interpreter's
' per-expression dispatch path, where IN.8 already measured 100x-2700x
' slower than compiled. Reading it per dispatch would be a real
' regression. An unreadable manifest is NOT cached as a result - the
' attempt simply repeats next time, so a transient failure cannot pin
' the interpreter into fallback mode for the rest of the session.
Private Function HelperManifestCached() As String
    If Not mHelperManifestRead Then
        mHelperManifest = VLA_Runtime.VlaHelperManifest()
        If Len(Trim$(mHelperManifest)) > 0 Then mHelperManifestRead = True
    End If
    HelperManifestCached = mHelperManifest
End Function

' ---------------------------------------------------------------------
'  IN2.0: the general CallByName dispatch mechanism - a heuristic, not
'  a type system. A member whose Get and Method forms are BOTH valid
'  but mean something different could pick the wrong one; none of
'  this project's own corpus does that (checked directly against the
'  actual dot-forms instructions.txt and english.vla use), and VBA's
'  own late-bound CallByName has exactly this ambiguity by design -
'  this is not a gap invented here, only inherited honestly.
' ---------------------------------------------------------------------

' Read position: try Get, then Method (covers a zero-arg method called
' for its return value, or a Get VBA itself resolves through method
' dispatch under the hood).
' IN.3: the READ-position sibling of DynamicSet's own native-dispatch
' story (that Sub's own header note has the full history). "end" is the
' first, and so far only, real GET-position member this corpus reaches
' that takes an ARGUMENT (Range.End(xlUp), english.vla's
' last-filled-row macro) - found live, the same way size/color/
' screenupdating/rows.count were, and root-caused the same way: every
' member tried through CallByName WITH an argument has failed one way
' or another, whether the argument was for a Get (this, and the
' WorksheetFunction census in EvalDynamicHead's own IN.3 note) or a
' Let/Set (DynamicSet's 20-member census). "Get: Overflow; Method:
' Object doesn't support this property or method" - a new symptom pair,
' same underlying lesson. Dispatched natively rather than added as a
' third generic argument-taking tier (EvalDynamicHead's WorksheetFunction
' Cases and this one solve the same problem in the two different places
' arguments can arrive, not two different problems).
' IN.11: three more native GET cases, found by reading the real
' corpus's own multi-segment dotted paths rather than guessed -
' "borders"/"tab"/"entirecolumn" are each the FIRST time this
' interpreter has ever fetched THAT specific intermediate segment
' (add-border's "borders.linestyle", set-tab-color's "tab.color",
' fit-all-columns's "entirecolumn.autofit" - english.vla's own
' defmacro text), reached only through DescendToParent's own generic
' CallByName Get fallback until now, the same never-audited shape this
' item's own header note names. `Range.Borders`/`Worksheet.Tab` are
' both PARAMETERIZED default properties in the real Excel object model
' (Borders takes an optional Index; a bare zero-arg late-bound
' CallByName Get against a parameterized default property is a
' documented COM/CallByName sharp edge, distinct from the plain-
' property Let/Set ambiguity DynamicSet's own census already fixed) -
' dispatched with a direct late-bound property read instead, the same
' "obj.Whatever" shape DynamicSet's own Select Case already uses.
Private Function DynamicGet(ByVal obj As Object, ByVal member As String, ByVal argVals As Variant) As Variant
    Select Case VLA_Identity.Fold(member)
        Case "end"
            If ArgCount(argVals) = 1 Then
                AssignVar DynamicGet, obj.End(CLng(ArgAt(argVals, 0)))
                Exit Function
            End If
        Case "borders": AssignVar DynamicGet, obj.Borders: Exit Function
        Case "tab": AssignVar DynamicGet, obj.Tab: Exit Function
        Case "entirecolumn": AssignVar DynamicGet, obj.EntireColumn: Exit Function
        ' SEC.1 Tier 2: promoted from the CallByName fallback removed
        ' below, after a full-repo census (every "(. obj member...)" and
        ' bare dotted-global shape across scripts/*.vla, scripts/
        ' polyglotta/*.vla, and the src/*.bas test suites) of every
        ' member currently reached only through it. Each of these
        ' already ran, silently, through late-bound CallByName before
        ' this pass - nothing here is new capability, only a new fixed,
        ' audited line of VBA for a member real shipped macros or the
        ' green host-test suite already exercise (found-items/for-each's
        ' own ".value" read, G-TABLES's own listobjects/listrows/
        ' listcolumns chain, "font."/"interior."/"entirerow." as
        ' intermediate DescendToParent hops, "rows.count"/".row" for
        ' last-filled-row, "worksheets.add", cell-of-sheet's own
        ' "(worksheets s) range r").
        Case "value": AssignVar DynamicGet, obj.Value: Exit Function
        ' SEC.1 Tier 2 correction (found live, host self-test): every
        ' member below was ALREADY native for SET, before this pass ever
        ' started - DynamicSet's own IN.3 census only ever covered the
        ' write direction ((set! (. r member) v)). Reading the SAME
        ' member back ((. r member) as an expression, or a bare dotted-
        ' global like activesheet.name/thisworkbook.name) was NEVER
        ' native on the get side and so, like "value" above before its
        ' own fix, silently rode the CallByName fallback the whole time
        ' - invisible until that fallback came out. The corpus-and-test
        ' census this pass ran caught "value" specifically because a
        ' host test read it explicitly; it did not generalize to "every
        ' Set-native property is presumably also Get-native somewhere,"
        ' which the same host suite's own "activesheet.name"/
        ' "font.bold"/"thisworkbook.name" pins then caught live. Fixed
        ' by mirroring DynamicSet's own list wholesale rather than
        ' waiting for each to fail individually - the same "once a tier
        ' is this unaudited, close it in one pass" lesson IN.11 already
        ' recorded for DynamicCall's own zero-arg census.
        Case "size": AssignVar DynamicGet, obj.Size: Exit Function
        Case "color": AssignVar DynamicGet, obj.Color: Exit Function
        Case "formula": AssignVar DynamicGet, obj.Formula2: Exit Function
        Case "bold": AssignVar DynamicGet, obj.Bold: Exit Function
        Case "italic": AssignVar DynamicGet, obj.Italic: Exit Function
        Case "horizontalalignment": AssignVar DynamicGet, obj.HorizontalAlignment: Exit Function
        Case "columnwidth": AssignVar DynamicGet, obj.ColumnWidth: Exit Function
        Case "colorindex": AssignVar DynamicGet, obj.ColorIndex: Exit Function
        Case "numberformat": AssignVar DynamicGet, obj.NumberFormat: Exit Function
        Case "linestyle": AssignVar DynamicGet, obj.LineStyle: Exit Function
        Case "hidden": AssignVar DynamicGet, obj.Hidden: Exit Function
        Case "rowheight": AssignVar DynamicGet, obj.RowHeight: Exit Function
        Case "name": AssignVar DynamicGet, obj.Name: Exit Function
        Case "freezepanes": AssignVar DynamicGet, obj.FreezePanes: Exit Function
        Case "wraptext": AssignVar DynamicGet, obj.WrapText: Exit Function
        Case "calculation": AssignVar DynamicGet, obj.Calculation: Exit Function
        Case "cutcopymode": AssignVar DynamicGet, obj.CutCopyMode: Exit Function
        Case "displayalerts": AssignVar DynamicGet, obj.DisplayAlerts: Exit Function
        Case "screenupdating": AssignVar DynamicGet, obj.ScreenUpdating: Exit Function
        Case "statusbar": AssignVar DynamicGet, obj.StatusBar: Exit Function
        Case "tablestyle": AssignVar DynamicGet, obj.TableStyle: Exit Function
        Case "showtotals": AssignVar DynamicGet, obj.ShowTotals: Exit Function
        Case "range"
            If ArgCount(argVals) = 1 Then
                AssignVar DynamicGet, obj.Range(ArgAt(argVals, 0))
                Exit Function
            End If
        Case "entirerow": AssignVar DynamicGet, obj.EntireRow: Exit Function
        Case "font": AssignVar DynamicGet, obj.Font: Exit Function
        Case "interior": AssignVar DynamicGet, obj.Interior: Exit Function
        Case "row": AssignVar DynamicGet, obj.Row: Exit Function
        Case "count": AssignVar DynamicGet, obj.Count: Exit Function
        Case "databodyrange": AssignVar DynamicGet, obj.DataBodyRange: Exit Function
        Case "add": AssignVar DynamicGet, obj.Add: Exit Function
        Case "listobjects"
            If ArgCount(argVals) = 0 Then
                AssignVar DynamicGet, obj.ListObjects
            Else
                AssignVar DynamicGet, obj.ListObjects(ArgAt(argVals, 0))
            End If
            Exit Function
        Case "listrows"
            If ArgCount(argVals) = 0 Then
                AssignVar DynamicGet, obj.ListRows
            Else
                AssignVar DynamicGet, obj.ListRows(ArgAt(argVals, 0))
            End If
            Exit Function
        Case "listcolumns"
            If ArgCount(argVals) = 1 Then
                AssignVar DynamicGet, obj.ListColumns(ArgAt(argVals, 0))
                Exit Function
            End If
    End Select
    ' SEC.1 Tier 2: the CallByName fallback that used to sit here is
    ' removed - anything not in the Select Case above is refused in
    ' words (LX.8's doctrine), not attempted via arbitrary late-bound
    ' dispatch against whatever obj happens to be at runtime. This
    ' closes THREAT_MODEL.md SS1.2's own finding. A real, legitimate new
    ' member belongs in the Select Case above, reviewed and added by
    ' name, the same way every member already there got there (IN.3/
    ' IN.11's own history).
    VLA_Messages.RaiseMsg "interp-dynamic-member-refused", "member", member
End Function

' Statement position: try Method, then Get (a bare '.' call is almost
' always an action; Get is the fallback, not the default).
' IN.11: the real, zero-argument statement-position method census,
' counted from english.vla's own macros the way DynamicSet's own
' trust-question audit counted its SET census - not one member patched
' reactively per crash. "activate" (activate-sheet, "Go to sheet ...")
' has run successfully through the untested CallByName path below many
' times already this session, so a blanket "the whole tier is broken"
' claim would be overstated; converted to native anyway for the same
' reason DynamicSet converted members that had not yet individually
' failed - once a tier is this unaudited, closing it in one pass beats
' leaving some members still resting on an unproven heuristic.
' Correction, found by the owner's own second live run rather than this
' pass's own first census: "select" (freeze-top-row's own
' "(. (rows 2) select) ...") was missed the first time - the first
' census read grepped for a macro body starting directly with "(. ",
' which silently excludes any macro whose "." call sits inside its own
' "(begin ...)" wrapper (freeze-top-row's exact shape); paste-values
' happened to be caught anyway by a manual re-read, freeze-top-row was
' not. Re-read exhaustively, line by line, not by pattern this time -
' "filldown"/"fillright" (fill-down/fill-right) added alongside it,
' unused by today's corpus but the same free, cheap completeness
' DynamicSet's own census already established as worth it once a tier
' is open.
Private Sub DynamicCall(ByVal obj As Object, ByVal member As String, ByVal argVals As Variant)
    If ArgCount(argVals) = 0 Then
        Select Case VLA_Identity.Fold(member)
            Case "activate": obj.Activate: Exit Sub
            Case "clearcontents": obj.ClearContents: Exit Sub
            Case "clearformats": obj.ClearFormats: Exit Sub
            Case "merge": obj.Merge: Exit Sub
            Case "unmerge": obj.UnMerge: Exit Sub
            Case "delete": obj.Delete: Exit Sub
            Case "insert": obj.Insert: Exit Sub
            Case "autofit": obj.AutoFit: Exit Sub
            Case "copy": obj.Copy: Exit Sub
            Case "showalldata": obj.ShowAllData: Exit Sub
            Case "select": obj.Select: Exit Sub
            Case "filldown": obj.FillDown: Exit Sub
            Case "fillright": obj.FillRight: Exit Sub
            ' SEC.1 Tier 2: promoted from the CallByName fallback
            ' (removed below) - table-to-range/table-add-row/clear-
            ' everything-from/refresh-everything/save-current-workbook/
            ' group-rows/ungroup-rows/(new Collection)'s own zero-arg
            ' ".add" all already reached these natively-absent members
            ' through late-bound dispatch before this.
            Case "unlist": obj.Unlist: Exit Sub
            Case "add": obj.Add: Exit Sub
            Case "clear": obj.Clear: Exit Sub
            Case "refreshall": obj.RefreshAll: Exit Sub
            Case "save": obj.Save: Exit Sub
            Case "group": obj.Group: Exit Sub
            Case "ungroup": obj.Ungroup: Exit Sub
        End Select
    ElseIf ArgCount(argVals) = 1 Then
        ' SEC.1 Tier 2: the one-positional-argument statement-call
        ' census - open-workbook/save-workbook-as/save-copy-as/
        ' wait-seconds, plus (new Collection)'s own one-arg ".add"
        ' (IN.11's own "Set pick-check to item 2 of found-items." pin),
        ' each previously reaching CallByName's VbMethod path with
        ' exactly one argument.
        Select Case VLA_Identity.Fold(member)
            Case "add": obj.Add ArgAt(argVals, 0): Exit Sub
            ' SEC.8: the three members here that reach OUTSIDE this
            ' workbook - a file read, and two file writes to a path the
            ' program chooses. Each is gated on the provenance of the
            ' workbook that CARRIED this program, captured once at
            ' VLA_IDE.CaptureHost. The guard is inside each Case rather
            ' than once at the top of this Sub deliberately: "add" and
            ' "wait" are not external effect and must stay ungated, and
            ' a per-Case guard is what check_sec8_provenance_gate.ps1
            ' can actually verify site by site.
            Case "open"
                VLA_Provenance.VlaProvenanceGuardCaptured "open a workbook"
                obj.Open ArgAt(argVals, 0): Exit Sub
            Case "saveas"
                VLA_Provenance.VlaProvenanceGuardCaptured "save a workbook under a new name"
                obj.SaveAs ArgAt(argVals, 0): Exit Sub
            Case "savecopyas"
                VLA_Provenance.VlaProvenanceGuardCaptured "save a copy of a workbook"
                obj.SaveCopyAs ArgAt(argVals, 0): Exit Sub
            Case "wait": obj.Wait ArgAt(argVals, 0): Exit Sub
        End Select
    End If
    ' SEC.1 Tier 2: the CallByName fallback that used to sit here is
    ' removed - anything not in the Select Case above is refused in
    ' words (LX.8's doctrine), not attempted via arbitrary late-bound
    ' dispatch. This closes THREAT_MODEL.md SS1.2's own finding.
    VLA_Messages.RaiseMsg "interp-dynamic-member-refused", "member", member
End Sub

' Write position: NATIVE fast paths, not a CallByName heuristic - and,
' as of this pass, the DEFAULT posture for every member this interpreter
' has actual corpus evidence for, not an exception list grown one crash
' at a time. History, briefly: IN2.1's own host run found CallByName +
' VbLet (and its VbSet fallback) UNRELIABLE for Range.Value specifically
' - no raise, but the wrong value written (a raw-pointer-shaped garbage
' 13-digit number), root-caused to Range.Value being a parameterized
' property in Excel's own type library, which confuses late-bound
' CallByName the same way Application's parameterized Range/Cells/etc.
' did (EvalDynamicHead's own header note). "value" got a native fast
' path; everything else kept going through CallByName on the working
' assumption that Value was the unusual case. IN.3 (three more live
' reloads, same session) found "size" (loud - object-defined error),
' then "color" (loud - Overflow), each patched the same way as found -
' 3 real members tried, 3 real members broken. THE OWNER'S OWN QUESTION,
' asked plainly after the third: can this heuristic be trusted for
' anything, or are these three data points, not three coincidences?
' Answered by counting rather than arguing: every DISTINCT `.`-set-position
' member english.vla's own macros actually use, across the whole
' corpus (formula/bold/italic/horizontalalignment/columnwidth/color/
' size/colorindex/numberformat/linestyle/hidden/rowheight/name/
' freezepanes/wraptext - 15 members, not 3), all switched to this same
' native shape in one pass rather than waiting for each to crash on its
' own reload. CallByName's Let/Set path used to sit below as a fallback
' for a member no real corpus use had reached yet - genuinely unproven,
' not "probably fine": every member this project ever actually tried
' through it failed, one way or another (silently for Value, loudly for
' Size/Color). SEC.1 Tier 2 later removed that fallback entirely (see
' this Sub's own tail below) once a full-repo census confirmed every
' member still reaching it by name, promoting each the same way.
' IN.3 (the very next host run, same audit, an incomplete count rather
' than a new failure mode): TestInterpreterObjectDispatch's own new
' application.screenupdating pin FAILED - "still True" - not a crash,
' the exact silent-wrong-value shape Value's own original bug had.
' Root cause: the 15-member census above was built from `.`-SET forms
' only ((set! (. r member) v)); "screenupdating" arrives through
' ExecSet's OTHER new path instead - a bare dotted-global place
' ((set! application.screenupdating v), no "." form at all) - which
' WalkMemberSet still hands to THIS SAME DynamicSet, member=
' "screenupdating", not in the 15-member list, so it fell straight
' through to the CallByName fallback and failed the same way Value did
' before it had a native path. Not a new class of bug, an incomplete
' count: the bare-dotted-global census (ExecSet's own header note) has
' five real members, not the two ("name"/"freezepanes") that happened to
' already overlap the `.`-form list - calculation/cutcopymode/
' displayalerts/screenupdating/statusbar all added below.
' SEC.4: a program-written cell value beginning with =, +, -, or @
' becomes a live formula (or, in some readers, a DDE/command trigger)
' the instant Excel evaluates it - the same class of bug CSV-export
' tooling has been bitten by industry-wide for a decade, and nothing
' guarded the interpreter's own single Value-write choke point
' (DynamicSet's own "value" case, below - confirmed the only .Value-
' setting site in this whole codebase before writing this) until now.
' The four trigger characters are OWASP's own well-known set for this
' exact vulnerability class - the roadmap item's own text names only
' "=" as its illustrative case, but leaving the other three out would
' be an incomplete fix for the class of bug it names, not a narrower
' one, so all four are guarded here, named explicitly rather than
' silently expanded.
'
' NEUTRALIZED, never refused: a computed value legitimately starting
' with one of these (an imported note, a negative-number-as-text field,
' an email starting with "@") is ordinary DATA, not an authoring
' mistake - refusing would break real, legitimate programs on every run
' whose data happens to shape up this way. Case "formula" (below, in
' the same Select Case, untouched by this fix) is the interpreter's own
' explicit, stated-intent path for a REAL live formula; "value"'s own
' contract has always been "write this value as seen," never "maybe
' execute it," so restoring that contract here is a bug fix, not a new
' restriction.
'
' Only VarType vbString is ever in scope - a genuine VBA number, date,
' or boolean can never literally begin with one of these four
' characters the way Excel evaluates real cell content, so every other
' type reaching this Sub is left completely untouched, not an
' oversight.
'
' The neutralization itself (a leading apostrophe, forcing Excel's own
' "treat as text" interpretation) is the standard, widely-documented
' fix for this exact class - but genuinely NEEDS A LIVE CONFIRMATION
' this session cannot give itself: whether a leading apostrophe set
' programmatically via .Value strips cleanly (matching manual cell
' entry) or survives as a literal, visible character depends on
' Excel/VBA behavior this session has no way to execute and observe.
' Flagged for the owner's own live check, not assumed either way.
' Public, not Private - the one deliberate exception to this Sub's own
' otherwise-internal helpers, made specifically so VLA_Tests.bas can
' pin the guard's own decision logic directly and purely: DynamicSet's
' own "value" case only ever fires against a real Range-like object
' (this module's own header note above, confirmed, not assumed), which
' a pure test has no way to construct - PrologRun's own PrologRun-made-
' Public-for-testability precedent (VLA_Prolog.bas), applied here to
' let the injection-guard decision itself be tested with zero host
' dependency, rather than skipped or faked through a stub object.
Public Function NeutralizeFormulaInjection(ByVal v As Variant) As Variant
    ' IsObject-branched before ANY assignment, never a bare `= v` first
    ' - this codebase's own already-documented trap: this function's
    ' own only caller today never passes an object (already excluded at
    ' the call site), but the function's own signature (ByVal v As
    ' Variant) invites a future caller that might, and a bare
    ' assignment from an object Variant invokes its own default member
    ' instead of copying the reference. Caught before this ever ran
    ' live, not live-caught this time - the exact class this project's
    ' own PROLOG line has hit more than once, checked for deliberately
    ' here rather than trusted away.
    If IsObject(v) Then
        Set NeutralizeFormulaInjection = v
        Exit Function
    End If
    NeutralizeFormulaInjection = v
    If VarType(v) <> vbString Then Exit Function
    Dim s As String
    s = CStr(v)
    If Len(s) = 0 Then Exit Function
    Select Case Left$(s, 1)
    Case "=", "+", "-", "@"
        NeutralizeFormulaInjection = "'" & s
    End Select
End Function

Private Sub DynamicSet(ByVal obj As Object, ByVal member As String, ByVal v As Variant)
    Dim foldedMember As String
    foldedMember = VLA_Identity.Fold(member)
    Select Case foldedMember
        Case "value"
            If IsObject(v) Then
                Set obj.Value = v
            Else
                Dim guardedVal As Variant
                guardedVal = NeutralizeFormulaInjection(v)
                obj.Value = guardedVal
                ' CStr-compared, not Variant "=" - a bare Variant "="
                ' comparison between two Strings is fine here (never an
                ' object on this branch, IsObject(v) already excluded
                ' that above), but CStr keeps the comparison's own type
                ' unambiguous rather than leaning on Variant coercion.
                If CStr(guardedVal) <> CStr(v) Then
                    LogEffect "sec.4: neutralized a leading formula-trigger character before writing .Value (was: " & Left$(CStr(v), 40) & ")"
                End If
            End If
            Exit Sub
        Case "size": obj.Size = v: Exit Sub
        Case "color": obj.Color = v: Exit Sub
        ' Formula2, not Formula: Range.Formula auto-inserts "@" (implicit
        ' intersection) on anything that could spill, silently breaking
        ' every query engine's own "returns a spilled array" promise
        ' (SQL.1's own SD-4 freeze) the moment Frazaro itself writes the
        ' formula - not just a person's stray Ctrl+Shift+Enter. Formula2
        ' (Excel 2019+/365, already required by deflambda's own LAMBDA)
        ' behaves identically to Formula for an ordinary scalar formula,
        ' so this is a strict improvement. VLA.bas's own emitter carries
        ' the matching fix, scoped identically narrow (its "set!" case,
        ' exactly (. obj formula) as the target) - AS.8 parity.
        Case "formula": obj.Formula2 = v: Exit Sub
        Case "bold": obj.Bold = v: Exit Sub
        Case "italic": obj.Italic = v: Exit Sub
        Case "horizontalalignment": obj.HorizontalAlignment = v: Exit Sub
        Case "columnwidth": obj.ColumnWidth = v: Exit Sub
        Case "colorindex": obj.ColorIndex = v: Exit Sub
        Case "numberformat": obj.NumberFormat = v: Exit Sub
        Case "linestyle": obj.LineStyle = v: Exit Sub
        Case "hidden": obj.Hidden = v: Exit Sub
        Case "rowheight": obj.RowHeight = v: Exit Sub
        Case "name": obj.Name = v: Exit Sub
        Case "freezepanes": obj.FreezePanes = v: Exit Sub
        Case "wraptext": obj.WrapText = v: Exit Sub
        Case "calculation": obj.Calculation = v: Exit Sub
        Case "cutcopymode": obj.CutCopyMode = v: Exit Sub
        Case "displayalerts": obj.DisplayAlerts = v: Exit Sub
        Case "screenupdating": obj.ScreenUpdating = v: Exit Sub
        Case "statusbar": obj.StatusBar = v: Exit Sub
        ' SEC.1 Tier 2: promoted from the CallByName fallback (removed
        ' below) - table-style/table-totals-on/table-totals-off's own
        ' two members, previously reaching CallByName's VbLet path.
        Case "tablestyle": obj.TableStyle = v: Exit Sub
        Case "showtotals": obj.ShowTotals = v: Exit Sub
    End Select

    ' SEC.1 Tier 2: the CallByName fallback that used to sit here is
    ' removed - anything not in the Select Case above is refused in
    ' words (LX.8's doctrine), not attempted via arbitrary late-bound
    ' dispatch. This closes THREAT_MODEL.md SS1.2's own finding.
    VLA_Messages.RaiseMsg "interp-dynamic-member-refused", "member", member
End Sub

' Walks all but the last dot-separated segment of memberPath via
' DynamicGet with no args, landing on the object the FINAL segment
' should be read/called/set on - member may be a multi-hop chain like
' "font.bold" (VLA.bas's own EmitDotText treats such a token as one
' piece of literal VBA member-access text for the compiler to resolve;
' this interpreter walks it hop by hop instead, since it has no
' compiler to lean on). lastSeg receives the final (mangled) segment.
Private Function DescendToParent(ByVal obj As Object, ByVal memberPath As String, ByRef lastSeg As String) As Object
    Dim segs() As String
    segs = Split(MangleIdent(memberPath), ".")
    Dim cur As Object
    Set cur = obj
    Dim i As Long
    For i = LBound(segs) To UBound(segs) - 1
        Dim mid0 As Variant
        AssignVar mid0, DynamicGet(cur, segs(i), Array())
        If Not IsObject(mid0) Then
            VLA_Messages.RaiseMsg "interp-member-chain-not-object", "segment", segs(i), "path", memberPath
        End If
        Set cur = mid0
    Next i
    lastSeg = segs(UBound(segs))
    Set DescendToParent = cur
End Function

Private Function WalkMemberGet(ByVal obj As Object, ByVal memberPath As String, ByVal argVals As Variant) As Variant
    Dim lastSeg As String
    Dim parent As Object
    Set parent = DescendToParent(obj, memberPath, lastSeg)
    AssignVar WalkMemberGet, DynamicGet(parent, lastSeg, argVals)
End Function

Private Sub WalkMemberCall(ByVal obj As Object, ByVal memberPath As String, ByVal argVals As Variant)
    Dim lastSeg As String
    Dim parent As Object
    Set parent = DescendToParent(obj, memberPath, lastSeg)
    DynamicCall parent, lastSeg, argVals
End Sub

' IN2.5: named-argument sibling of WalkMemberCall - same multi-hop
' descent (DescendToParent), the final segment dispatched through
' DynamicNamedCall instead of DynamicCall/CallByName. No real corpus
' use needs the multi-hop case, but there is no reason this one
' shouldn't have it too, since DescendToParent already does the work.
' F.2 G-TABLES: Function now, propagating DynamicNamedCall's own return
' (see that function's header) - ExecDotCall's existing call to this,
' `WalkMemberNamedCall objVal, memTok, kwArgs`, still works unchanged
' as a bare statement (a Function called without capturing its result
' discards it, same as a Sub); EvalDotForm's new keyword branch below
' is the first caller that actually reads what comes back.
Private Function WalkMemberNamedCall(ByVal obj As Object, ByVal memberPath As String, kwArgs As Collection) As Variant
    Dim lastSeg As String
    Dim parent As Object
    Set parent = DescendToParent(obj, memberPath, lastSeg)
    AssignVar WalkMemberNamedCall, DynamicNamedCall(parent, lastSeg, kwArgs)
End Function

Private Sub WalkMemberSet(ByVal obj As Object, ByVal memberPath As String, ByVal v As Variant)
    Dim lastSeg As String
    Dim parent As Object
    Set parent = DescendToParent(obj, memberPath, lastSeg)
    DynamicSet parent, lastSeg, v
End Sub

' Evaluates lst's elements from fromIdx to the end as positional
' expression arguments (always returns a genuine array, via VBA's own
' Array() for the zero-arg case, so ArgCount/ArgAt never need a
' separate "was this ever assigned" check). Refuses, in words, if any
' element is a keyword token (':name value') - named-argument dispatch
' is a dedicated, separate mechanism (EvalKeywordArgs/DynamicNamedCall,
' below - IN2.5), since VBA's CallByName has no keyword-argument
' mechanism to hand them to and this generic CallByName path never
' will. Only ExecDotCall (statement position) routes there today - see
' its own IN2.5 note for why expression-position '.' reads do not need
' it yet.
Private Function EvalPositionalArgs(lst As Collection, ByVal fromIdx As Long, frame As Object) As Variant
    Dim n As Long
    n = lst.Count - fromIdx + 1
    If n < 0 Then n = 0
    If n > 4 Then
        VLA_Messages.RaiseMsg "interp-positional-args-too-many", "n", n
    End If
    If n = 0 Then
        EvalPositionalArgs = Array()
        Exit Function
    End If
    Dim vals() As Variant
    ReDim vals(0 To n - 1)
    Dim i As Long
    For i = 1 To n
        Dim tok As Variant
        AssignVar tok, Nth(lst, fromIdx + i - 1)
        If IsKeywordTok(tok) Then
            VLA_Messages.RaiseMsg "interp-named-args-not-supported-here", "tok", CStr(tok)
        End If
        AssignVar vals(i - 1), EvalExpr(tok, frame)
    Next i
    EvalPositionalArgs = vals
End Function

' IN2.5: named-argument dispatch. Evaluates lst's elements from
' fromIdx as ':key value' pairs, returning a Collection of (foldedKey,
' value) pairs - not a VlaDict/Scripting.Dictionary, since every real
' caller (DynamicNamedCall, below) only ever does a handful of
' straight-line lookups by a known literal key on 1-3 pairs; reusing
' VLA_Runtime's own keyed-collection machinery for that would be the
' premature-reuse mistake this project's own docs warn against
' elsewhere, not the free win it looks like. Every real use in today's
' corpus is all-keyword (grepped directly: english.vla/
' prelude.vla, ten call sites, not one mixes a plain positional
' argument in before the first ':'), so that is the only shape this
' function accepts; a mix refuses in words rather than guessing which
' half is positional.
Private Function EvalKeywordArgs(lst As Collection, ByVal fromIdx As Long, frame As Object) As Collection
    Dim outc As New Collection
    Dim i As Long
    i = fromIdx
    Do While i <= lst.Count
        Dim tok As Variant
        AssignVar tok, Nth(lst, i)
        If Not IsKeywordTok(tok) Then
            VLA_Messages.RaiseMsg "interp-expected-keyword-arg", "pos", (i - fromIdx + 1)
        End If
        If i + 1 > lst.Count Then
            VLA_Messages.RaiseMsg "interp-keyword-arg-missing-value", "tok", CStr(tok)
        End If
        Dim v As Variant
        AssignVar v, EvalExpr(Nth(lst, i + 1), frame)
        ' IN.11: an EXPLICIT 2-slot array, built with indexed
        ' assignment (Set when v IsObject, matching AssignVar's own
        ' precedent), NOT Array(key, v) - the first real theory this
        ' item chased for "Copy range G1:I4 to range K1:M4."'s own
        ' "Type mismatch", and NOT actually where the bug was
        ' (JoinKwArgs's own IN.11 note has the real story) - kept
        ' anyway, since it is still strictly safer than Array(key, v)
        ' for the same reason AssignVar exists everywhere else in this
        ' module, at no cost. Array() is documented to
        ' apply implicit Let-style coercion to its own arguments for at
        ' least some object shapes - this codebase's own AssignVar
        ' exists specifically because a plain, uncontrolled assignment
        ' of an object can silently substitute its default member
        ' instead of the reference; Array(key, v) never went through
        ' that same discipline. Turned out NOT to be where the real
        ' "Type mismatch" came from (JoinKwArgs's own IN.11 note has
        ' the real mechanism - a LOGGING function downstream, not this
        ' storage step), but this is still the more defensible shape.
        Dim pair(0 To 1) As Variant
        pair(0) = VLA_Identity.Fold(Mid$(CStr(tok), 2))
        If IsObject(v) Then
            Set pair(1) = v
        Else
            pair(1) = v
        End If
        outc.Add pair
        i = i + 2
    Loop
    Set EvalKeywordArgs = outc
End Function

Private Function KwArg(kwArgs As Collection, ByVal key As String) As Variant
    Dim e As Variant
    For Each e In kwArgs
        If e(0) = key Then
            AssignVar KwArg, e(1)
            Exit Function
        End If
    Next
    VLA_Messages.RaiseMsg "interp-missing-keyword-arg", "key", key
End Function

' G-STRUCT (owner-caught live: three refusals in a row on the same
' underlying gap - an unresolved Excel constant, then two keyword-shape
' mismatches on ALREADY-whitelisted methods). KwArg's own unconditional
' raise-if-absent is right for a genuinely required keyword, but real
' PasteSpecial/Cut/Delete/Resize each have OPTIONAL named parameters,
' and every Case below that touches one of them was, until now,
' hardcoded to the ONE keyword combination whichever rule first needed
' it - not because the method doesn't support more, but because nobody
' had asked yet. KwArgOptional lets a Case ask "is this key present"
' without raising, so one Case can genuinely cover a method's real
' optional-parameter surface instead of silently narrowing to whatever
' the first caller happened to pass.
Private Function KwArgOptional(kwArgs As Collection, ByVal key As String, ByVal defaultVal As Variant) As Variant
    Dim e As Variant
    For Each e In kwArgs
        If e(0) = key Then
            AssignVar KwArgOptional, e(1)
            Exit Function
        End If
    Next
    AssignVar KwArgOptional, defaultVal
End Function

' Best-effort, log-only rendering - JoinArgs's named-argument sibling.
' IN.11: the REAL root cause behind "Copy range G1:I4 to range K1:M4."'s
' own "Type mismatch" - found only after DynamicNamedCall's "copy" case
' was instrumented end to end and never once fired (IN.11 DIAG A-H),
' proving the copy itself was already succeeding. `ExecDotCall` calls
' this AFTER WalkMemberNamedCall returns, purely to build the effect
' log's own text - `CStr(e(1))` applied directly to an object value
' invokes its default member (the same Let-coercion AssignVar exists
' to control elsewhere in this module), which for a MULTI-CELL Range
' is a 2D ARRAY of cell values, not a scalar; CStr on an array is
' exactly "Type mismatch" - AFTER the real work already completed
' correctly, purely while formatting a log line nobody's own logic
' depends on. A single-cell range's default member is a harmless
' scalar, which is why every prior range-valued keyword arg (sort's
' key1, the single-cell "in2.5" copy pin) never surfaced this. Fixed
' the same way LogEffect's own receiver text already does (TypeName,
' not CStr, for an object) - this function's job is a readable log
' line, not a value the interpreter itself ever reads back.
Private Function JoinKwArgs(kwArgs As Collection) As String
    Dim e As Variant
    Dim r As String
    For Each e In kwArgs
        If Len(r) > 0 Then r = r & ", "
        If IsObject(e(1)) Then
            r = r & e(0) & "=" & TypeName(e(1))
        Else
            r = r & e(0) & "=" & CStr(e(1))
        End If
    Next
    JoinKwArgs = r
End Function

' IN.10: existence check to go with KwArg (which raises when the key
' is missing, the right behavior for its own IN2.5 callers but wrong
' here - a missing keyword against a user procedure's OPTIONAL
' parameter just means "use the default", not an error).
Private Function KwArgHas(kwArgs As Collection, ByVal key As String) As Boolean
    Dim e As Variant
    For Each e In kwArgs
        If e(0) = key Then
            KwArgHas = True
            Exit Function
        End If
    Next
End Function

' IN.10: calls a registered user procedure - proc is a RegisterProc
' record (isFunc, the whole (sub/function name params ...) form,
' bodyIdx), lst is the CALL form ((stamp :row-number 2 ...) or
' (tax 100)), callerFrame is whatever frame the call expression is
' being evaluated in. Builds one fresh, ISOLATED call frame (no
' visibility into callerFrame's own bindings beyond what argument
' binding copies in - a named non-goal, not a reuse of F.5's VlaFrame,
' see the roadmap's own IN.10 entry), binds parameters either
' positionally or by keyword, runs the body, and returns whatever
' (return <expr>) assigned (Empty for a Sub, or a Function that never
' returned explicitly - VBA's own default-return-value behavior).
' mLoopBreak/mProcReturn are both hard-reset at the end: a call is its
' own unwind boundary, so a degenerate exit-for/exit-do with no
' enclosing loop inside the callee cannot leak a break into whatever
' loop the CALLER happens to be inside. IN3.6: mErrMode/mErrLabel are
' SAVED and restored around the call, not hard-reset like
' mLoopBreak/mProcReturn - real VBA gives every procedure its own
' fresh On Error state (Off) on entry (so a handler the callee arms
' must not leak out, the callee's own fresh "" on the way in), but
' unlike a break/return flag, a caller's ALREADY-armed handler is not
' something an ordinary, error-free call should consume - the caller's
' own Try: (if any) must still be armed for whatever statement comes
' next, exactly as if the call had never happened. This covers the
' success path only, on purpose: if the callee's own body raises
' UNCAUGHT (mErrMode = "" during the callee, so ExecBody used the
' plain untrapped ExecStmt path), the error unwinds this function
' before the restore lines below ever run - VLA.bas's own EmitStmt
' "at-line" wrapper has the exact same documented shape ("the restore
' is deliberately skipped [on error]... so the failure message can
' name the source line it was in"). Whether that error should then be
' caught by whatever handler the CALLER had armed before making this
' call is exactly the call-crossing question this item's own header
' note already defers - nothing in today's corpus calls a user
' procedure from inside an armed handler, so this gap costs nothing
' today. Named, not hidden.
Private Function CallUserProc(ByVal proc As Collection, ByVal lst As Collection, ByVal callerFrame As Object) As Variant
    Dim isFunc As Boolean
    isFunc = proc.Item(1)
    Dim defLst As Collection
    Set defLst = proc.Item(2)
    Dim bodyIdx As Long
    bodyIdx = proc.Item(3)
    Dim params As Collection
    Set params = Nth(defLst, 3)

    Dim callFrame As Object
    Set callFrame = VLA_Runtime.VlaDictNew()
    BindProcArgs params, lst, callerFrame, callFrame

    mLoopBreak = False
    mProcReturn = False
    mProcReturnValue = Empty
    Dim savedErrMode As String, savedErrLabel As String
    Dim savedCaughtNum As Long, savedCaughtSrc As String, savedCaughtDesc As String
    savedErrMode = mErrMode: savedErrLabel = mErrLabel
    savedCaughtNum = mCaughtErrNum: savedCaughtSrc = mCaughtErrSrc: savedCaughtDesc = mCaughtErrDesc
    mErrMode = ""
    mErrLabel = ""
    ExecBody defLst, bodyIdx, callFrame
    mLoopBreak = False
    mProcReturn = False
    mErrMode = savedErrMode
    mErrLabel = savedErrLabel
    mCaughtErrNum = savedCaughtNum
    mCaughtErrSrc = savedCaughtSrc
    mCaughtErrDesc = savedCaughtDesc

    ' IN3.6: the same stuck-flag safety net as VlaInterpret's own
    ' top-level loop - a goto/resume target never found anywhere up to
    ' and including the callee's own top-level body is a real bug
    ' (an unreachable label, or a label spelled differently than its
    ' goto), not something that should silently corrupt whatever the
    ' caller does next.
    If Len(mGotoLabel) > 0 Then
        Dim badLabel1 As String
        badLabel1 = mGotoLabel
        mGotoLabel = ""
        VLA_Messages.RaiseMsg "interp-goto-label-not-found", "label", badLabel1
    End If

    If isFunc Then AssignVar CallUserProc, mProcReturnValue
End Function

' IN.10: positional vs. named is detected the same way EvalKeywordArgs
' already implies (IN2.5) - the corpus's real calls never mix the two,
' so the FIRST argument token alone decides which binder runs, exactly
' as EvalPositionalArgs's own keyword-token refusal already assumes
' for the generic dispatch path.
Private Sub BindProcArgs(params As Collection, callLst As Collection, callerFrame As Object, callFrame As Object)
    Const firstArgIdx As Long = 2
    Dim isNamedCall As Boolean
    If callLst.Count >= firstArgIdx Then
        Dim firstTok As Variant
        AssignVar firstTok, Nth(callLst, firstArgIdx)
        isNamedCall = IsKeywordTok(firstTok)
    End If
    If isNamedCall Then
        BindNamedArgs params, callLst, firstArgIdx, callerFrame, callFrame
    Else
        BindPositionalArgs params, callLst, firstArgIdx, callerFrame, callFrame
    End If
End Sub

' (tax 100), (commission 600 (/ 5 100)) - positional binding against
' the DECLARED parameter list (not EvalPositionalArgs's generic
' 4-argument-cap array, which exists for CallByName's own dynamic-
' dispatch limitation and has nothing to do with a user procedure
' whose real arity is already known). Argument expressions evaluate in
' the CALLER's frame (the caller's own variables are what such an
' expression can reference); a default expression for an omitted
' trailing parameter evaluates in the new callFrame instead (VBA's own
' Optional default is part of the callee's signature, not the
' caller's scope - and every real default in today's corpus is a
' literal or constant arithmetic that could not reference a caller
' variable regardless).
Private Sub BindPositionalArgs(params As Collection, callLst As Collection, ByVal firstArgIdx As Long, callerFrame As Object, callFrame As Object)
    Dim nGiven As Long
    nGiven = callLst.Count - firstArgIdx + 1
    If nGiven < 0 Then nGiven = 0
    If nGiven > params.Count Then
        VLA_Messages.RaiseMsg "interp-too-many-positional-args", "given", nGiven, "declared", params.Count
    End If
    Dim i As Long
    For i = 1 To params.Count
        Dim pname As String, isOptional As Boolean, hasDefaultNode As Boolean, defaultNode As Variant
        ParamSpec Nth(params, i), pname, isOptional, hasDefaultNode, defaultNode
        Dim v As Variant
        If i <= nGiven Then
            AssignVar v, EvalExpr(Nth(callLst, firstArgIdx + i - 1), callerFrame)
        ElseIf isOptional Then
            If hasDefaultNode Then
                AssignVar v, EvalExpr(defaultNode, callFrame)
            Else
                v = Empty
            End If
        Else
            VLA_Messages.RaiseMsg "interp-missing-required-arg", "name", pname
        End If
        VLA_Runtime.VlaDictSet callFrame, pname, v
    Next
End Sub

' (stamp :row-number 2 :value "beta"), (commission :sale 600) - named
' binding, reusing IN2.5's own EvalKeywordArgs to parse+evaluate the
' ':key value' pairs (against the CALLER's frame, same reasoning as
' the positional binder above) rather than re-implementing that
' parsing a second time.
Private Sub BindNamedArgs(params As Collection, callLst As Collection, ByVal firstArgIdx As Long, callerFrame As Object, callFrame As Object)
    Dim kwArgs As Collection
    Set kwArgs = EvalKeywordArgs(callLst, firstArgIdx, callerFrame)
    Dim i As Long
    For i = 1 To params.Count
        Dim pname As String, isOptional As Boolean, hasDefaultNode As Boolean, defaultNode As Variant
        ParamSpec Nth(params, i), pname, isOptional, hasDefaultNode, defaultNode
        Dim foldedName As String
        foldedName = VLA_Identity.Fold(pname)
        Dim v As Variant
        If KwArgHas(kwArgs, foldedName) Then
            AssignVar v, KwArg(kwArgs, foldedName)
        ElseIf isOptional Then
            If hasDefaultNode Then
                AssignVar v, EvalExpr(defaultNode, callFrame)
            Else
                v = Empty
            End If
        Else
            VLA_Messages.RaiseMsg "interp-missing-required-named-arg", "name", pname
        End If
        VLA_Runtime.VlaDictSet callFrame, pname, v
    Next
End Sub

' One parameter spec, VLA.bas's own EmitParams shapes (VLA.bas:2079):
' name | (name type) | (byval name type) | (byref name type) |
' (optional name type [default]) | (paramarray name). isOptional
' means "may be omitted"; hasDefaultNode/defaultNode carry an EXPLICIT
' default expression when the spec has one (an (optional ...) with no
' 4th element is still omittable, just defaults to Empty - the same
' "Optional As Variant with no = clause" VBA itself gives). paramarray
' refuses in words - out of scope for this pass, named rather than
' silently mis-bound, same discipline as IN2.5's own out-of-scope
' shapes.
Private Sub ParamSpec(ByVal spec As Variant, ByRef pname As String, ByRef isOptional As Boolean, _
                       ByRef hasDefaultNode As Boolean, ByRef defaultNode As Variant)
    isOptional = False
    hasDefaultNode = False
    defaultNode = Empty
    If Not IsList(spec) Then
        pname = SymText(spec)
        Exit Sub
    End If
    Dim pl As Collection
    Set pl = spec
    Dim ph As String
    ph = VLA_Identity.Fold(HeadSym(pl))
    Select Case ph
        Case "byval", "byref"
            pname = SymText(Nth(pl, 2))
        Case "optional"
            pname = SymText(Nth(pl, 2))
            isOptional = True
            If pl.Count >= 4 Then
                hasDefaultNode = True
                AssignVar defaultNode, Nth(pl, 4)
            End If
        Case "paramarray"
            VLA_Messages.RaiseMsg "interp-paramarray-unsupported"
        Case Else                              ' (name type)
            pname = SymText(Nth(pl, 1))
    End Select
End Sub

' IN2.5: the eleven real named-argument '.' calls this corpus actually
' uses (english.vla/prelude.vla, grepped directly, not guessed -
' this is not a generic mechanism, because none can exist: CallByName
' has no keyword-argument mechanism to hand a name to, at all, for any
' object). Each dispatches through an EARLY-BOUND typed local instead -
' VBA resolves named arguments at compile time for a KNOWN type, the
' same reason the place helpers (range/cells/...) already call
' Range(...)/Cells(...) natively instead of through CallByName. A call
' whose object does not actually support the assumed type fails with
' VBA's own honest type-mismatch error - the same failure a hand-
' written VBA program calling the wrong method would get, not a new
' kind of silent wrongness this interpreter introduces.
' F.2 G-TABLES: a twelfth Case, "add" (ListObjects.Add), and the reason
' this became a Function instead of a Sub - the first named-argument
' call this corpus needs FOR ITS RETURN VALUE (make-table's macro reads
' .Name off the new ListObject), not just its side effect. Every
' existing Case still works unchanged: a Function called as a bare
' statement (WalkMemberNamedCall's own body, and ExecDotCall's call to
' THAT) discards its return exactly like a Sub would - nothing here
' requires an existing caller to change.
Private Function DynamicNamedCall(ByVal obj As Object, ByVal member As String, kwArgs As Collection) As Variant
    Dim m As String
    m = VLA_Identity.Fold(member)
    Select Case m
        Case "add"
            Dim lobjs As ListObjects
            Set lobjs = obj
            Dim newLo As ListObject
            Set newLo = lobjs.Add(SourceType:=KwArg(kwArgs, "sourcetype"), _
                                   Source:=KwArg(kwArgs, "source"), _
                                   XlListObjectHasHeaders:=KwArg(kwArgs, "xllistobjecthasheaders"))
            Set DynamicNamedCall = newLo
        ' SEC.8: five gated members in this Sub. Two kinds, and the
        ' roadmap entry names them separately rather than blurring them:
        ' printout and exportasfixedformat are EGRESS (content leaves,
        ' to a printer or to a file at a path the program picks), while
        ' close-with-save and protect/unprotect are DESTRUCTIVE to data
        ' already on this machine - no attacker-chosen path, but a
        ' password-protect an internet-sourced program applies is a
        ' person locked out of their own sheet. Gating both widens SEC.8
        ' from "provenance gates egress" to "provenance gates external
        ' AND destructive effect", which is a deliberate scope call.
        Case "protect"
            VLA_Provenance.VlaProvenanceGuardCaptured "protect a sheet with a password"
            Dim wsProtect As Worksheet
            Set wsProtect = obj
            wsProtect.Protect Password:=KwArg(kwArgs, "password")
        Case "unprotect"
            VLA_Provenance.VlaProvenanceGuardCaptured "remove a sheet's password protection"
            Dim wsUnprotect As Worksheet
            Set wsUnprotect = obj
            wsUnprotect.Unprotect Password:=KwArg(kwArgs, "password")
        Case "close"
            VLA_Provenance.VlaProvenanceGuardCaptured "close a workbook"
            Dim wbClose As Workbook
            Set wbClose = obj
            wbClose.Close SaveChanges:=KwArg(kwArgs, "savechanges")
        Case "printout"
            VLA_Provenance.VlaProvenanceGuardCaptured "print a sheet"
            Dim wsPrintOut As Worksheet
            Set wsPrintOut = obj
            wsPrintOut.PrintOut Preview:=KwArg(kwArgs, "preview")
        Case "exportasfixedformat"
            VLA_Provenance.VlaProvenanceGuardCaptured "export a sheet to a PDF or XPS file"
            Dim wsExport As Worksheet
            Set wsExport = obj
            wsExport.ExportAsFixedFormat Type:=KwArg(kwArgs, "type"), Filename:=KwArg(kwArgs, "filename")
        Case "copy"
            ' IN.11: the real work here was never broken - the "Type
            ' mismatch" traced all the way to JoinKwArgs's own IN.11
            ' note (this file), a LOGGING function ExecDotCall calls
            ' AFTER this Case already succeeds. Kept the early-bound
            ' rDest As Range (this Case's own prior form inlined the
            ' KwArg(...) result straight into Destination:=), matching
            ' every sibling Case's own declare-then-Set shape - harmless
            ' either way, but consistent now.
            Dim rCopy As Range
            Set rCopy = obj
            Dim rDest As Range
            Set rDest = KwArg(kwArgs, "destination")
            rCopy.Copy Destination:=rDest
        Case "sort"
            Dim rSort As Range
            Set rSort = obj
            rSort.Sort Key1:=KwArg(kwArgs, "key1"), Order1:=KwArg(kwArgs, "order1"), Header:=KwArg(kwArgs, "header")
        Case "pastespecial"
            ' G-STRUCT: widened from "always requires paste" (paste-
            ' values' own original shape) to genuinely optional Paste/
            ' Transpose, each defaulting to PasteSpecial's own real
            ' default when the caller didn't ask for it - -4104 =
            ' xlPasteAll (ResolveExcelConstant's own numeric-literal-
            ' plus-comment convention, not the symbolic name, matching
            ' every other entry in that table).
            Dim rPaste As Range
            Set rPaste = obj
            rPaste.PasteSpecial Paste:=KwArgOptional(kwArgs, "paste", -4104), _
                                 Transpose:=KwArgOptional(kwArgs, "transpose", False)
        Case "replace"
            Dim rReplace As Range
            Set rReplace = obj
            rReplace.Replace What:=KwArg(kwArgs, "what"), Replacement:=KwArg(kwArgs, "replacement")
        Case "removeduplicates"
            Dim rDedupe As Range
            Set rDedupe = obj
            rDedupe.RemoveDuplicates Columns:=KwArg(kwArgs, "columns"), Header:=KwArg(kwArgs, "header")
        Case "autofilter"
            Dim rFilter As Range
            Set rFilter = obj
            rFilter.AutoFilter Field:=KwArg(kwArgs, "field"), Criteria1:=KwArg(kwArgs, "criteria1")
        ' G-STRUCT: cut/delete/resize, added instead of leaving cut-
        ' range/delete-shift-up/delete-shift-left/insert-n-rows-at on
        ' plain positional calls - Destination/Shift/RowSize+ColumnSize
        ' are all genuinely optional in real Excel, each defaulting to
        ' Excel's own behavior (Cut with no Destination just marks the
        ' clipboard; Delete with no Shift lets Excel infer a direction;
        ' Resize with either dimension omitted leaves it unchanged) when
        ' the caller didn't ask for it.
        Case "cut"
            Dim rCut As Range
            Set rCut = obj
            ' IN.11's own lesson, missed here originally: "destination" is
            ' a Range, and a plain cutDest = KwArgOptional(...) silently
            ' takes the Range's default member (its .Value) instead of the
            ' reference - a currently-empty destination cell then reads as
            ' Empty, IsEmpty below goes True, and Cut runs with NO
            ' Destination at all (arms the clipboard, never pastes) even
            ' though the caller passed one. AssignVar (Set for objects)
            ' fixes it the same way it already does everywhere else in
            ' this file.
            Dim cutDest As Variant
            AssignVar cutDest, KwArgOptional(kwArgs, "destination", Empty)
            ' Confirmed live, 2026-08-28 (DIAG cut probe): AssignVar above
            ' already does its job correctly - TypeName(cutDest)="Range",
            ' IsObject(cutDest)=True when a destination was passed. The
            ' bug was THIS check: IsEmpty() on an object argument invokes
            ' its DEFAULT MEMBER (Range.Value) rather than testing whether
            ' a reference is actually present - a destination cell that
            ' happens to be blank right now (the ordinary case: nothing's
            ' been cut into it yet) reads back as IsEmpty=True even though
            ' a perfectly real Range was supplied, so this used to take
            ' the no-destination branch and silently arm-but-never-paste
            ' the clipboard. IsObject is the right test: it asks what
            ' cutDest IS, not what its default property currently
            ' evaluates to.
            If Not IsObject(cutDest) Then
                rCut.Cut
            Else
                rCut.Cut Destination:=cutDest
            End If
        Case "delete"
            Dim rDelete As Range
            Set rDelete = obj
            Dim delShift As Variant
            delShift = KwArgOptional(kwArgs, "shift", Empty)
            If IsEmpty(delShift) Then
                rDelete.Delete
            Else
                rDelete.Delete Shift:=delShift
            End If
        Case "resize"
            ' Resize is a property returning a Range, not a void method -
            ' this Case has to hand that Range back (Set DynamicNamedCall
            ' = ...) so a chained ".Insert" (insert-n-rows-at's own
            ' shape) has something real to call next, the same return-
            ' value convention the "add" Case above already established.
            Dim rResize As Range
            Set rResize = obj
            Dim rzRows As Variant, rzCols As Variant
            rzRows = KwArgOptional(kwArgs, "rowsize", Empty)
            rzCols = KwArgOptional(kwArgs, "columnsize", Empty)
            If Not IsEmpty(rzRows) And Not IsEmpty(rzCols) Then
                Set DynamicNamedCall = rResize.Resize(RowSize:=rzRows, ColumnSize:=rzCols)
            ElseIf Not IsEmpty(rzRows) Then
                Set DynamicNamedCall = rResize.Resize(RowSize:=rzRows)
            ElseIf Not IsEmpty(rzCols) Then
                Set DynamicNamedCall = rResize.Resize(ColumnSize:=rzCols)
            Else
                Set DynamicNamedCall = rResize
            End If
        Case Else
            VLA_Messages.RaiseMsg "interp-named-args-unsupported-member", "member", member
    End Select
End Function

' IN.11: argVals - the array-holding Variant EvalPositionalArgs builds
' and every dispatch tier threads onward - is ByRef here and in every
' OTHER function in this file that receives it (TryEvalBuiltin/
' TryRuntimeHelper/DynamicGet/DynamicCall/CallByNameArgs/
' CallByNameArgsVoid/WalkMemberGet/WalkMemberCall/JoinArgs), not ByVal.
' This is the actual fix behind "Set pick-check to item 2 of found-
' items."'s own raw-pointer garbage (H22, VerifyReportInterpreter) -
' found by a controlled bisection (TestInterpreterNewCollection's own
' six throwaway hand-written-VBA controls, VLA_Tests.bas), not
' reasoned out in advance: a literal value through CallByName, fine; a
' single array element, fine; the SAME array threaded through THREE
' nested ByVal Variant parameter hops before being indexed at the
' last one (mirroring this file's own WalkMemberCall -> DynamicCall ->
' CallByNameArgsVoid depth exactly) - corrupted, reproduced with both
' a dynamic (ReDim) and a fixed-size array, ruling the array's own
' kind out; zero hops - clean. Changing every hop from ByVal to ByRef
' closed it, confirmed by an isolated sixth control that changes
' nothing else. This is a pure, zero-risk mechanical change - none of
' these functions ever WRITE to argVals, only read/index it, so ByRef
' changes nothing about what this module does, only how VBA passes the
' Variant between procedures. Present since IN.2 first built this
' chain; invisible until now because every real corpus member this
' session has natively cased (font.bold, sort, removeduplicates, ...)
' stopped reaching this generic multi-hop fallback tier at all -
' Collection.Add (unmapped, no native Case, CallByNameArgsVoid's own
' first real live exercise of a multi-argument call this deep) is what
' finally reached it.
Private Function ArgCount(ByVal argVals As Variant) As Long
    ArgCount = UBound(argVals) - LBound(argVals) + 1
End Function

Private Function ArgAt(ByVal argVals As Variant, ByVal i As Long) As Variant
    If i >= ArgCount(argVals) Then
        VLA_Messages.RaiseMsg "interp-dispatch-missing-arg", "n", (i + 1)
    End If
    AssignVar ArgAt, argVals(LBound(argVals) + i)
End Function

' Best-effort, log-only rendering of a dispatch call's arguments -
' never used for anything but LogEffect text. IN.11: "a value that
' cannot render cleanly just shows via CStr's own default text" turned
' out wrong, not merely imprecise - CStr on an OBJECT invokes its
' default member (JoinKwArgs's own IN.11 note has the full mechanism,
' found live via that function's own sibling bug), and for a multi-
' cell Range that member is a 2D array, which CStr cannot render at
' all - it raises, rather than falling back to any "default text".
' TypeName for an object, same fix as JoinKwArgs, before a multi-cell
' range ever reaches this positional-argument path too.
Private Function JoinArgs(ByVal argVals As Variant) As String
    Dim n As Long
    n = ArgCount(argVals)
    Dim r As String
    Dim i As Long
    Dim argVal As Variant
    For i = 0 To n - 1
        If i > 0 Then r = r & ", "
        AssignVar argVal, ArgAt(argVals, i)
        If IsObject(argVal) Then
            r = r & TypeName(argVal)
        Else
            r = r & CStr(argVal)
        End If
    Next i
    JoinArgs = r
End Function

Private Function IsKeywordTok(ByVal v As Variant) As Boolean
    If IsObject(v) Then Exit Function
    Dim s As String
    s = CStr(v)
    IsKeywordTok = (Left$(s, 1) = ":" And Len(s) > 1)
End Function

' SymName's hyphen-to-underscore half (VLA.bas's own SymName, Private
' there too - R7 duplicate, one line): a macro or built-in name spelled
' with a hyphen at the VLA level needs the same mangling the emitter
' already gives it before it becomes a real VBA-callable name. Skips
' SymName's transliteration half on purpose - every name this function
' sees is either hardcoded here or drawn from VLA_Runtime's own plain-
' ASCII function names, never a user-authored identifier LX.6 already
' handles at the emitter.
Private Function MangleIdent(ByVal s As String) As String
    MangleIdent = Replace(s, "-", "_")
End Function

Private Function EvalChain(lst As Collection, frame As Object, ByVal op As String) As Double
    Dim acc As Double
    acc = CDbl(EvalExpr(Nth(lst, 2), frame))
    Dim i As Long
    For i = 3 To lst.Count
        If op = "+" Then
            acc = acc + CDbl(EvalExpr(Nth(lst, i), frame))
        Else
            acc = acc - CDbl(EvalExpr(Nth(lst, i), frame))
        End If
    Next
    EvalChain = acc
End Function

' IN2.7: the 18-operator family VLA.bas's own EmitExpr shares through ONE
' mechanism (EmitChain, fed every operator in its Case list alike, not
' just the arithmetic ones) - AS.8's own scan found only +/-/> were ever
' given interpreter cases; this closes the rest the same way, one shared
' function rather than 18 near-duplicates. A LITERAL LEFT-TO-RIGHT FOLD,
' matching EmitChain's generated text exactly - "(< a b c)" means
' "((a < b) < c)" under real VBA's own left-to-right evaluation and
' Boolean/Integer coercion, NOT a check that a<b AND b<c. Deliberately
' NOT forced through CDbl the way EvalChain (+/-) is: comparisons on
' strings/dates are real VBA usage the emitter already supports for free
' by emitting bare operator text, so native VBA operators on the raw
' Variant are what this interpreter must use too, for the same programs
' to agree - CDbl-forcing here would silently narrow what this backend
' accepts relative to the emitter, the exact asymmetry AS.8 exists to
' catch. ">" used to be its own bespoke 2-operand-only, CDbl-forced
' Case - folded in here instead (not left inconsistent beside its five
' siblings), which also fixes a real, pre-existing gap of its own: it
' never supported EmitChain's own N-ary "(> a b c)" shape, only exactly
' two operands. Every VBA operator here is used AS a VBA operator, not
' hand-derived: 'is' on non-object Variants, or '\'/'mod' on non-numeric
' ones, raises the SAME runtime type-mismatch real generated VBA would
' raise for the same misuse - correct parity, not a gap this function
' should paper over with its own type checking.
Private Function EvalOpChain(lst As Collection, frame As Object, ByVal opFolded As String) As Variant
    Dim acc As Variant
    AssignVar acc, EvalExpr(Nth(lst, 2), frame)
    Dim i As Long
    Dim rhs As Variant
    For i = 3 To lst.Count
        AssignVar rhs, EvalExpr(Nth(lst, i), frame)
        Select Case opFolded
            Case "*": acc = acc * rhs
            Case "&": acc = acc & rhs
            Case "/": acc = acc / rhs
            Case "\": acc = acc \ rhs
            Case "=": acc = (acc = rhs)
            Case "<>": acc = (acc <> rhs)
            Case "<": acc = (acc < rhs)
            Case ">": acc = (acc > rhs)
            Case "<=": acc = (acc <= rhs)
            Case ">=": acc = (acc >= rhs)
            Case "and": acc = (acc And rhs)
            Case "or": acc = (acc Or rhs)
            Case "xor": acc = (acc Xor rhs)
            Case "mod": acc = (acc Mod rhs)
            Case "is": acc = (acc Is rhs)
            Case "like": acc = (acc Like rhs)
            Case "imp": acc = (acc Imp rhs)
            Case "eqv": acc = (acc Eqv rhs)
            Case Else
                VLA_Messages.RaiseMsg "interp-opchain-unknown-operator", "op", opFolded
        End Select
    Next
    EvalOpChain = acc
End Function

' ---------------------------------------------------------------------
'  R7 duplicates of VLA.bas's Private form-walking helpers (IsList,
'  HeadSym, Nth, SymText, StrLitContent, AssignVar) - cross-module
'  Private calls do not exist in VBA, and each of these is under
'  fifteen lines. Faithful copies as of IN0.5; VLA.bas's own copies
'  remain authoritative if the two ever drift.
' ---------------------------------------------------------------------
Private Function IsList(ByVal v As Variant) As Boolean
    IsList = IsObject(v)
End Function

Private Function Nth(lst As Collection, ByVal i As Long) As Variant
    If i > lst.Count Then VLA_Messages.RaiseMsg "interp-form-missing-element", "n", i
    If IsObject(lst.Item(i)) Then
        Set Nth = lst.Item(i)
    Else
        Nth = lst.Item(i)
    End If
End Function

Private Function SymText(ByVal v As Variant) As String
    If IsObject(v) Then VLA_Messages.RaiseMsg "interp-expected-symbol-got-list"
    Dim s As String
    s = CStr(v)
    If Left$(s, 1) = Chr$(34) Then VLA_Messages.RaiseMsg "interp-expected-symbol-got-string"
    SymText = s
End Function

Private Function HeadSym(lst As Collection) As String
    If lst.Count = 0 Then VLA_Messages.RaiseMsg "interp-empty-form"
    HeadSym = SymText(Nth(lst, 1))
End Function

Private Function StrLitContent(ByVal v As Variant) As String
    If IsObject(v) Then VLA_Messages.RaiseMsg "interp-expected-string-got-list"
    Dim s As String
    s = CStr(v)
    If Left$(s, 1) <> Chr$(34) Then VLA_Messages.RaiseMsg "interp-expected-string-got-other", "value", s
    StrLitContent = Mid$(s, 2)
End Function

' L-INTERPOLATE: EvalExpr's own "interpolate" case. tpl is required to
' be a literal string NODE (checked structurally below, like SymText's
' own literal-symbol check just above), never evaluated as an
' expression - VLA.bas's EmitExpr Case "interpolate" has no runtime to
' hand a computed value to, so both backends refuse the same programs
' rather than the interpreter quietly accepting something the compiler
' cannot.
Private Function EvalInterpolateCall(lst As Collection, frame As Object) As Variant
    Dim tmplNode As Variant
    AssignVar tmplNode, Nth(lst, 2)
    If IsObject(tmplNode) Then
        VLA_Messages.RaiseMsg "interp-interpolate-template-must-be-literal", "got", "a nested form"
    End If
    Dim tmplTok As String
    tmplTok = CStr(tmplNode)
    If Left$(tmplTok, 1) <> Chr$(34) Then
        VLA_Messages.RaiseMsg "interp-interpolate-template-must-be-literal", "got", tmplTok
    End If
    Dim tpl As String
    tpl = Mid$(tmplTok, 2)
    Dim pieces As Collection
    Set pieces = EvalParseInterpolateTemplate(tpl)

    ' Every argument after tpl is a ":key value" pair (EvalKeywordArgs's
    ' own all-keyword convention, reused by shape only). Each value is
    ' evaluated exactly ONCE here, up front, then reused for every
    ' occurrence of its hole in the template below - never re-evaluated
    ' per occurrence, so an argument with a side effect (a function
    ' call) cannot silently run twice just because its hole is written
    ' twice.
    Dim keyRaw() As String, keyFold() As String
    Dim valArr() As Variant, usedArr() As Boolean
    Dim kCount As Long
    kCount = 0
    Dim i As Long
    i = 3
    Do While i <= lst.Count
        Dim tok As Variant
        AssignVar tok, Nth(lst, i)
        If Not IsKeywordTok(tok) Then
            VLA_Messages.RaiseMsg "interp-interpolate-expected-keyword-arg", "pos", (i - 2)
        End If
        If i + 1 > lst.Count Then
            VLA_Messages.RaiseMsg "interp-keyword-arg-missing-value", "tok", CStr(tok)
        End If
        Dim rawK As String
        rawK = Mid$(CStr(tok), 2)
        Dim foldK As String
        foldK = VLA_Identity.Fold(rawK)
        Dim v As Variant
        AssignVar v, EvalExpr(Nth(lst, i + 1), frame)
        Dim slot As Long
        slot = InterpolateKeyIndex(keyFold, kCount, foldK)
        If slot < 0 Then
            ReDim Preserve keyRaw(0 To kCount)
            ReDim Preserve keyFold(0 To kCount)
            ReDim Preserve valArr(0 To kCount)
            ReDim Preserve usedArr(0 To kCount)
            keyRaw(kCount) = rawK
            keyFold(kCount) = foldK
            AssignVar valArr(kCount), v
            usedArr(kCount) = False
            kCount = kCount + 1
        Else
            keyRaw(slot) = rawK
            AssignVar valArr(slot), v
        End If
        i = i + 2
    Loop

    Dim result As String
    Dim p As Variant
    For Each p In pieces
        Dim ps As String
        ps = CStr(p)
        If Left$(ps, 1) = Chr$(1) Then
            Dim hName As String
            hName = Mid$(ps, 2)
            Dim idx As Long
            idx = InterpolateKeyIndex(keyFold, kCount, VLA_Identity.Fold(hName))
            If idx < 0 Then
                VLA_Messages.RaiseMsg "interp-interpolate-unknown-key", "key", hName
            End If
            usedArr(idx) = True
            result = result & valArr(idx)
        Else
            result = result & ps
        End If
    Next p

    Dim j As Long
    For j = 0 To kCount - 1
        If Not usedArr(j) Then
            VLA_Messages.RaiseMsg "interp-interpolate-unused-argument", "key", keyRaw(j)
        End If
    Next j

    EvalInterpolateCall = result
End Function

' A simple linear scan (BETA_ROADMAP2.md's own L-INTERPOLATE scoping
' text: "not a second grammar") over tpl, splitting it into
' literal-text pieces and named-hole pieces. A hole piece is marked
' internally with a leading Chr$(1) (unlikely to appear in real text;
' the same "prefix character marks a token kind" trick this module's
' own tokenizer uses for string literals, Chr$(34)) so the pieces
' Collection can hold both kinds as plain strings. "{{" escapes to a
' literal "{"; a bare "}" (or a doubled "}}") is always literal, since
' a "}" is only ever AMBIGUOUS while a hole is open, and this scanner
' already closes a hole the moment it sees one - no separate escape is
' structurally needed for it, so this parser accepts one for symmetry
' rather than requiring it.
Private Function EvalParseInterpolateTemplate(ByVal tpl As String) As Collection
    Dim outc As New Collection
    Dim n As Long
    n = Len(tpl)
    Dim i As Long
    i = 1
    Dim lit As String
    lit = ""
    Do While i <= n
        Dim c As String
        c = Mid$(tpl, i, 1)
        If c = "{" Then
            If Mid$(tpl, i + 1, 1) = "{" Then
                lit = lit & "{"
                i = i + 2
            Else
                Dim closeAt As Long
                closeAt = InStr(i + 1, tpl, "}")
                If closeAt = 0 Then
                    VLA_Messages.RaiseMsg "interp-interpolate-unclosed-hole"
                End If
                Dim holeName As String
                holeName = Mid$(tpl, i + 1, closeAt - i - 1)
                If Len(holeName) = 0 Then
                    VLA_Messages.RaiseMsg "interp-interpolate-empty-hole-name"
                End If
                If Len(lit) > 0 Then
                    outc.Add lit
                    lit = ""
                End If
                outc.Add Chr$(1) & holeName
                i = closeAt + 1
            End If
        ElseIf c = "}" Then
            If Mid$(tpl, i + 1, 1) = "}" Then
                lit = lit & "}"
                i = i + 2
            Else
                lit = lit & "}"
                i = i + 1
            End If
        Else
            lit = lit & c
            i = i + 1
        End If
    Loop
    If Len(lit) > 0 Then outc.Add lit
    Set EvalParseInterpolateTemplate = outc
End Function

' Linear search, not a keyed Collection/VlaDict lookup - an interpolate
' call's own keyword-argument count is always small (the corpus's own
' worst case, datalog-chain-place, has three), so this costs nothing
' a human would notice, and it sidesteps a keyed Collection's own
' "remove-then-re-add to update" dance for the common "same key twice"
' path in EvalInterpolateCall above.
Private Function InterpolateKeyIndex(arr() As String, ByVal n As Long, ByVal target As String) As Long
    Dim k As Long
    For k = 0 To n - 1
        If arr(k) = target Then
            InterpolateKeyIndex = k
            Exit Function
        End If
    Next k
    InterpolateKeyIndex = -1
End Function

Private Sub AssignVar(ByRef target As Variant, ByVal v As Variant)
    If IsObject(v) Then
        Set target = v
    Else
        target = v
    End If
End Sub

' S1.1-style dev-file probe, duplicated from VLA_Tests.bas's
' FindDevFile / VLA_DevRig.bas's DevFindFile per the same R7 logic -
' three copies, not a fourth cross-module Private call.
Private Function FindDevFile(ByVal fileName As String) As String
    Dim sep As String
    sep = Application.PathSeparator
    Dim c1 As String, c2 As String, c3 As String
    c1 = ThisWorkbook.Path & sep & fileName
    c2 = ThisWorkbook.Path & sep & "scripts" & sep & fileName
    c3 = ThisWorkbook.Path & sep & "scripts" & sep & "polyglotta" & sep & fileName
    If Len(Dir$(c1)) > 0 Then
        FindDevFile = c1
    ElseIf Len(Dir$(c2)) > 0 Then
        FindDevFile = c2
    ElseIf Len(Dir$(c3)) > 0 Then
        FindDevFile = c3
    Else
        VLA_Messages.RaiseMsg "interp-dev-file-not-found", "filename", fileName, "beside", c1, "scripts", c2, "polyglotta", c3
    End If
End Function
