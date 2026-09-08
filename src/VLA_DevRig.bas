Attribute VB_Name = "VLA_DevRig"
Option Explicit
Public Const VLA_DEVRIG_VERSION As String = "PPROF.0"
' PPROF.0: VlaProfileAll added - P-PROF's own dial
' (BETA_ROADMAP2.md, MACHINE + OPTIMIZATION), VlaTimeIt's family shape
' applied to a per-phase timing switch (VLA.mProfileOn, VLA.bas/
' VLA_SentenceEngine.bas) instead of a single number: turns it on,
' runs VLA_Tests_Host.VlaSelfTestsAll (the scale-included suite - the
' 2000/5500-deep TestListopsBudget/TestListopsDepthSafety/
' TestTablespecDepthSafety cases, VLA_Tests.bas), turns it off, prints
' the six-bucket breakdown (tokenize/parse/expand/emit compile-side,
' trans-tokenize/trans-build translate-side) against PREVIOUS. See its
' own header comment, right above the Sub, for the full mechanism and
' gating reasoning.
' SEC11.0: mods array gains VLA_Digest (SEC.11's pure SHA-256). Added to
' VLA_Build.bas's own array in the SAME edit, which is now mechanically
' required rather than merely remembered - tools/check_devrig_mods_
' parity.ps1 pins the two lists against each other.
' PROLOG.3: mods array gains VLA_Prolog (ships - the first real
' =PROLOG(...) worksheet function). Test pin (TestProlog) lands in
' VLA_Tests_Query.bas, no new test module - the same "one file per
' roadmap SECTION" rule PROLOG.1's own note below already applies.
' PROLOG.1: mods array gains VLA_Unify (ships - VLA_Relation's own
' sibling). No new test module: TestUnify/TestGRenderUnify landed in
' VLA_Tests_Query.bas instead (that file's own "one file per roadmap
' SECTION" rule - VLA_Unify.bas is PROLOG.1), dispatched from TestDSLs,
' not VlaSelfTest - corrected mid-session from an initial wrong call.
' AS7.0: VlaTimeItSelfTest - VlaTimeIt's own suite-runtime counterpart
' (AS.7, docs/BETA_ROADMAP.md). Built as a new dial in this module,
' modeled on VlaTimeIt's own shape (workbook-Name baselines, CURRENT
' next to PREVIOUS), rather than as anything wired into VlaSelfTest
' itself - see its own header note, right above the Sub, for why.
' REPLEVAL.0: eval(source) added beside expand (renamed from lisp,
' same session, same reasoning - see expand's own comment below) - the
' E the REPL was missing: expands (prelude loaded, defmacros consumed)
' then computes via VLA_Interpreter.VlaEvalExpression and prints the
' value, Lisp-flavored (arrays as parenthesized lists, true/false
' lowercase, Str$'s locale-proof period). NOT the eval the roadmap
' vetoed - that veto targets an expansion-time primitive whose computed
' values flow back into emitted code; this runs strictly post-
' expansion, dev-rig only, and its output is ink in the Immediate pane,
' never compiler input. See its own comment (right after expand) for
' the full wall adjudication and the interpreter-vs-compiled "whose
' answer" caveat.
' VlaEvalDisplay (Public, pinned by TestEvalDisplay) is its display
' formatter - deliberately not VlaTryValueLit, whose *1/*2/*3 recall
' contract must refuse the arrays eval exists to print.
' LISTOPS.0: lisp(source) added - a one-word Immediate-window REPL,
' Debug.Print(VLA.VlaExpandText(source, True, fired)) under the hood.
' Requested live, same session LISTOPS.0 shipped in VLA.bas - the
' owner's own words, "I can never remember what the command is to use
' Lisp in the immediate pane, so why not just implement a 'vla' command
' that takes a string of VLA." Two names tried and retired before this
' one, both live, both first real use: "vla" collided with the VLA
' MODULE itself ("Expected variable or procedure, not module" - VBA
' identifier lookup is case-insensitive, so a bare "vla" resolved to
' VLA.bas's own VB_Name before it ever looked for a Sub); "vlae" ("vla
' eval") fixed that but the owner's own second thought landed on the
' actually-obvious name - it IS Lisp in the immediate pane, so the
' command is just "lisp". See its own comment (right after VlaTry) for
' how it differs from VlaTry (actually runs source) and VlaExpandStep
' (narrates every intermediate application).
' IN8.0: VlaTimeItInterpreter (VlaTimeIt's own interpreter-side
' counterpart, same instructions.txt corpus, same four-leg shape) and
' VlaBenchmarkLoop (a pure dispatch-overhead microbenchmark needing no
' corpus or worksheet) added - IN.8's "measure, not estimate" harness.
' F5.0: VlaFrame (F.5's new class module, VLA.bas's push/pop context)
' added to the reload list. The list's file-path construction tried only
' ".bas" - harmless while every entry was a standard module, wrong for
' this one - so it now tries ".bas" first and falls back to ".cls",
' unchanged for every existing (still ".bas") entry.
' LX3.0: VLA_Identity added to the reload list (below); every
' identity-deciding LCase$ site in this module now folds through
' VLA_Identity.Fold - R6/SD-8.
' IO6.0: VlaTimeIt's benchmark module was hardcoded "EN_TimeIt", now
' "Frazaro_EN_TimeIt" - matches the OUT_MODULE/RT_MODULE rename in
' VLA_IDE.bas/VLA_Runtime.bas.
' LX3.1: VlaDevReload's default folder now probes "src" beside the
' workbook first (the Beta reorg), falling back to the workbook's own
' folder - the same c1/c2 shape DevFindFile already uses for scripts/.
' IN1.0: VLA_HeadTable added to the reload list (below) and to
' VlaDiagnostics's printout.
' IN0.5: VLA_Interpreter (the walking-skeleton interpreter, dev-rig
' only) added to the reload list and to VlaDiagnostics's printout,
' same treatment as VLA_HeadTable got at IN.1.
' F8.0: VLA_Tests_Grammar (F.8's split-out half of VLA_Tests, holding
' the grammar/phrasebook feature pins) added to the reload list and
' to VlaDiagnostics's printout, same treatment.
' F11.0: VLA_Tests_Host (F.11's split-out half of VLA_Tests, holding
' every test that needs a live workbook - VerifyReport included)
' added to the reload list and to VlaDiagnostics's printout, same
' treatment.

' =====================================================================
'  VLA_DevRig - the developer rig's loop half. Dev-only: NOT shipped
'  in the add-in build. Requires "Trust access to the VBA project
'  object model" (the dev machine has it).
'
'  VlaDevReload    re-imports every VLA .bas found in a folder in one
'                  call - retires the manual remove/import ritual and
'                  the half-updated-project bug class. Then: Debug >
'                  Compile (no API can trigger that for you).
'
'  VlaTry          L2: the scratch evaluation loop - try VLA forms
'                  from the Immediate window without editing a
'                  program or hand-building a module (full notes at
'                  its section below).
'
'  VlaDiagnostics  one pasteable blob: module versions, grammar
'                  counts, paths, workbook names. The answer to
'                  "which version is actually loaded?"
' =====================================================================

Public Sub VlaDevReload(Optional ByVal folder As String = "")
    On Error GoTo failed
    If Len(folder) = 0 Then
        ' Beta reorg: the .bas sources moved into a "src" folder
        ' sibling to the workbook - the same shape every other
        ' file-finder here already handles for scripts/ (DevFindFile
        ' below, VLA_Build's vocab path, VLA_IDE's vocab candidates).
        ' Probe for the new layout first; fall back to the old flat
        ' one so a workbook sitting directly beside the sources still
        ' reloads unchanged.
        Dim srcFolder As String
        srcFolder = ThisWorkbook.Path & "\src"
        If Len(Dir$(srcFolder & "\VLA.bas")) > 0 Then
            folder = srcFolder
        Else
            folder = ThisWorkbook.Path
        End If
    End If
    Dim mods As Variant
    ' IN.7: VLA_Events and VLA_EventSink added - left off this list
    ' would mean VlaDevReload silently never refreshes them, the same
    ' failure shape F5.0/LX3.0/IN1.0/IN.6 already document for
    ' VLA_Build.bas's own separate mods array (this one is independent
    ' of that one - it also covers dev-only modules, like VLA_Tests,
    ' that never ship).
    ' LX2.0: VLA_Messages added - reproduced the exact failure this
    ' comment already warned about: added to VLA_Build.bas's mods array
    ' but not here, so the new module sat on disk, unimported, and the
    ' first live compile of a call site referencing it failed with
    ' "Variable not defined" - live-caught, not found by inspection.
    ' CLI.0: frmCLI reloads through this same loop, plain, no special
    ' case. It briefly wasn't: a UserForm Import auto-opening its
    ' designer window in the VBIDE threw "Can't move focus to the
    ' control..." on three straight live runs, in a shape no amount of
    ' On Error scoping in this loop could reliably catch - looked like a
    ' genuine VBIDE limitation. It wasn't: all three failures rode the
    ' same session as repeated forced `taskkill /F /IM EXCEL.EXE` calls
    ' against the dev workbook's own live Excel process (verification
    ' tooling, external to this file) - live-confirmed the actual cause,
    ' not just suspected, by a full clean Excel restart alone clearing it
    ' with zero code change. Left in plainly rather than re-adding dead
    ' defensive scaffolding for a problem that was never in this loop.
    ' PNTH.0: VlaSlice added (VlaSlice.cls) - cdr/cddr's own O(1) view
    ' type (VLA.bas's own ListTail), F5.0's own ".cls fallback (tries
    ' .bas first, falls back to .cls, unchanged for every existing
    ' entry)" precedent covers it with no loop change needed.
    ' SEC.8: VLA_Provenance added - and this list's own warning above
    ' came true a SIXTH time on the way in, in the exact shape LX2.0
    ' recorded for VLA_Messages: added to VLA_Build.bas's mods array,
    ' missed here, so the module sat on disk unimported and the first
    ' live compile of a call site referencing it failed with "Variable
    ' not defined" - live-caught by the owner, not found by inspection,
    ' again. Two independent arrays that must agree, with nothing
    ' mechanical holding them together, is the actual defect; the
    ' repeated comments are a workaround for it, not a fix.
    mods = Array("VLA_Identity", "VLA_Messages", "VLA_Digest", "VLA_HeadTable", "VLA", "VlaFrame", "VlaSlice", "VLA_Loader", "VLA_Provenance", "VLA_English", "VLA_SentenceEngine", "VLA_Runtime", "VLA_IDE", "VLA_Build", "VLA_Lint", "VLA_Tests", "VLA_Tests_Grammar", "VLA_Tests_Host", "VLA_Tests_Query", "VLA_Interpreter", "VLA_Events", "VLA_EventSink", "VLA_Unify", "VLA_Relation", "VLA_Datalog", "VLA_Sql", "VLA_Prolog", "VLA_Browser", "frmCLI")
    Dim i As Long
    Dim fp As String
    Dim comp As Object
    Dim n As Long
    For i = LBound(mods) To UBound(mods)
        fp = folder & "\" & mods(i) & ".bas"
        If Len(Dir$(fp)) = 0 Then fp = folder & "\" & mods(i) & ".cls"   ' F5.0: VlaFrame
        If Len(Dir$(fp)) = 0 Then fp = folder & "\" & mods(i) & ".frm"   ' CLI.0: frmCLI, on trial
        If Len(Dir$(fp)) > 0 Then
            ' VBA defers component removal, so removing then importing
            ' the same name can yield "Name1". Rename the old one
            ' aside first; it is gone by the time this sub returns.
            On Error Resume Next
            Set comp = Nothing
            Set comp = ThisWorkbook.VBProject.VBComponents(CStr(mods(i)) & "_old")
            If Not comp Is Nothing Then ThisWorkbook.VBProject.VBComponents.Remove comp
            Set comp = Nothing
            Set comp = ThisWorkbook.VBProject.VBComponents(CStr(mods(i)))
            On Error GoTo failed
            If Not comp Is Nothing Then
                comp.Name = CStr(mods(i)) & "_old"
                ThisWorkbook.VBProject.VBComponents.Remove comp
            End If
            ThisWorkbook.VBProject.VBComponents.Import fp
            n = n + 1
            Debug.Print "reloaded: " & mods(i)
        End If
    Next
    If Len(Dir$(folder & "\VLA_DevRig.bas")) > 0 Then
        Debug.Print "note: VLA_DevRig.bas found - a module cannot replace itself while running; update it manually if it changed."
    End If
    Debug.Print "VlaDevReload: " & n & " module(s) refreshed from " & folder & ". Now: Debug > Compile."
    Exit Sub
failed:
    MsgBox Err.Description, vbExclamation, "VLA DevRig"
End Sub

' V5.3 (added after a real reload collision): report every Public
' procedure defined in MORE than one module - the "Ambiguous name
' detected" compile error, found in one call instead of a hunt.
' Import churn creates these three known ways: a hand-import beside
' a reload-import (VBE silently names the second copy "Name1"), an
' interrupted reload stranding a helper-bearing "_old" module
' (removal is deferred - hazard 11), and files from an abandoned
' session. Run it whenever Compile says "Ambiguous".
Public Sub VlaAuditDuplicates()
    Dim seen As New Collection      ' name -> first module
    Dim dupes As String
    Dim comp As Object
    Dim i As Long, n As Long
    Dim ln As String, nm As String, p As Long
    For Each comp In ThisWorkbook.VBProject.VBComponents
        If comp.Type = 1 Then       ' standard modules only
            n = comp.CodeModule.CountOfLines
            For i = 1 To n
                ln = Trim$(comp.CodeModule.Lines(i, 1))
                nm = ""
                If Left$(ln, 16) = "Public Function " Then
                    nm = Mid$(ln, 17)
                ElseIf Left$(ln, 11) = "Public Sub " Then
                    nm = Mid$(ln, 12)
                End If
                If Len(nm) > 0 Then
                    p = InStr(nm, "(")
                    If p > 1 Then
                        nm = Trim$(Left$(nm, p - 1))
                        Dim firstMod As Variant
                        Dim hit As Boolean
                        hit = False
                        On Error Resume Next
                        firstMod = seen.Item(VLA_Identity.Fold(nm))
                        hit = (Err.Number = 0)
                        On Error GoTo 0
                        If hit Then
                            dupes = dupes & "  " & nm & ": " & firstMod & " AND " & comp.Name & vbCrLf
                        Else
                            seen.Add comp.Name, VLA_Identity.Fold(nm)
                        End If
                    End If
                End If
            Next
        End If
    Next
    If Len(dupes) = 0 Then
        Debug.Print "VlaAuditDuplicates: clean - every Public defined once."
    Else
        Debug.Print "VlaAuditDuplicates: DUPLICATE Public definitions:" & vbCrLf & dupes & _
                    "Remove the stale module(s) (likely *_old, *1, or an abandoned import), then Debug > Compile."
    End If
End Sub

' =====================================================================
'  L2: VlaTry - the scratch evaluation loop. VLA's answer to the
'  Immediate window: try one or more VLA forms without editing a
'  program, hand-building a module, or clicking anything.
'
'    VlaTry "(debug-print (+ 1 2))"
'
'  builds a temporary Sub around the statements, transpiles, injects
'  a scratch module (VLA_Scratch) into THIS workbook through the
'  D1-threaded VlaCompileToModule (host passed explicitly - a focus
'  change can never redirect the injection), runs it, and deletes
'  it. Forms that must live at module level (sub, function,
'  defmacro, public, private, type, enum) are hoisted above the
'  scratch Sub, so a macro or helper can be defined and exercised
'  in one string:
'
'    VlaTry "(defmacro (twice x) (begin x x)) (twice (debug-print 7))"
'
'  (dim ...) and (const ...) are NOT hoisted - inside the scratch
'  they read as procedure locals, the natural scratch meaning, and
'  statements keep their relative order. Hoisting a definition
'  above the statements is harmless by construction: the transpiler
'  collects defmacros and module-level forms order-independently.
'
'  VlaTry source, True   also prints the scratch VLA and the VBA it
'                        became - and any emitter "near vla line N"
'                        indexes the SCRATCH text, so this view is
'                        that message's decoder.
'
'  VlaTryBuild(source)   the pure half (split + classify + wrap),
'                        returning the scratch VLA text. The
'                        self-test pins its contract there; the
'                        inject/run/delete half mutates the live
'                        project, so only the Immediate-window
'                        smoke in the verification loop can judge
'                        it (the trace's precedent).
'
'  Honest limits, recorded:
'  - A compile-invalid scratch (an undeclared name under the emitted
'    Option Explicit, say) surfaces VBA's OWN modal during
'    Application.Run - compile errors cross Run untrappably (S4's
'    empirical verdict). Dev-rig territory: the dev-side compile
'    gate in miniature, and the dev can read the modal. A stranded
'    VLA_Scratch such a modal may leave behind is swept at the
'    START of the next VlaTry, and the rename-aside delete means
'    it can never collide with the next injection.
'  - Injecting into this project resets ITS module-level state
'    (loaded grammar, S2 usage counters, capture flags) - the same
'    force every dev-topology Run exerts (S5.1's lesson). Reload
'    vocabulary before the next Check/Run. VlaTry itself keeps no
'    module state, so it survives its own injection.
'  Dev-rig only, never in the add-in - exactly like this module.
' =====================================================================

Public Sub VlaTry(ByVal source As String, Optional ByVal showCode As Boolean = False)
    Dim phase As String
    Dim scratch As String
    Debug.Print "=== VlaTry"
    On Error GoTo oops
    phase = "recall history (*1/*2/*3)"
    source = TryExpandStars(source)
    phase = "build the scratch"
    scratch = VlaTryBuild(source)
    If showCode Then
        phase = "transpile the scratch"
        Debug.Print "  scratch VLA:"
        DevPrintIndented scratch
        Debug.Print "  VBA:"
        DevPrintIndented VlaTranspile(scratch)
    End If
    phase = "inject the scratch module"
    TryDropScratch                       ' sweep any stranded scratch first
    VlaCompileToModule scratch, "VLA_Scratch", ThisWorkbook
    phase = "run the scratch"
    Application.Run "'" & ThisWorkbook.Name & "'!VLA_Scratch.vla_scratch"
    phase = "clean up the scratch module"
    TryDropScratch
    Debug.Print "  (ran and cleaned up)"
    Exit Sub
oops:
    Dim msg As String
    msg = Err.Description
    On Error Resume Next
    TryDropScratch
    On Error GoTo 0
    Debug.Print "  (could not " & phase & ": " & msg & ")"
    If Not showCode Then Debug.Print "  (VlaTry source, True shows the scratch text a 'vla line N' indexes)"
End Sub

' LISTOPS.0: the one-word REPL - "what's the command again" solved by
' making the command a word short enough not to forget. Expands SOURCE
' to its fixpoint and prints the result, nothing else - the terse
' sibling of VlaExpandStep (narrates every intermediate application,
' teaching material by design) and VlaTry above (actually COMPILES AND
' RUNS source against the live workbook, Trust access and scratch-
' module injection included). expand touches no workbook state at all -
' it's VlaExpandText under the hood, the same pure macro-expansion path
' every self-test assertion already exercises - so it's always safe to
' fire at anything from the Immediate window, no cleanup, no injected
' module, no risk of a stranded scratch. Exactly the right tool for the
' LISTOPS primitives this pass added (car/cdr/cons/list/null?/eq?/
' equal?/quote-if are pure expand-time computation - there is nothing
' to RUN, only something to EXPAND):
'   expand "(car (quote (1 2 3)))"
'   expand "(cons 1 (quote (2 3)))"
'   expand "(defmacro (double x) (+ x x)) (double 21)"
' For a program with real side effects (Set!, Range writes, MsgBox),
' use VlaTry instead - expand will happily show you the expanded FORM
' but never executes anything.
' RENAMED to "expand" (REPLEVAL.0) from "lisp" - the owner's own
' insistence, for consistency with eval, its new sibling: "lisp"
' didn't say what it does, and now that there are TWO one-word verbs
' side by side (one shows what a form EXPANDS to, one shows what it
' YIELDS), a cute name stops pulling its weight - the pair reads
' correctly at a glance only if both names are actions. "lisp" was
' itself the THIRD name, not the first: two shorter ones were tried
' and retired before it, both live, both first real use. "vla"
' collided with the MODULE named VLA (VLA.bas's own VB_Name) - VBA
' identifier lookup is case-insensitive, so a bare "vla" resolved to
' the module before it ever looked for a Sub, and a module referenced
' bare, not qualifying a member, is not a callable expression
' ("Expected variable or procedure, not module"). "vlae" ("vla eval")
' sidestepped that collision cleanly, and "lisp" - "this is Lisp,
' running in the Immediate pane" - lived the longest of the three, one
' whole session's worth of real use, before losing to the same
' clarity argument that named eval in the first place. Three retired
' names now on the record, not two - a future session should not
' re-litigate "vla"/"vlae"/"lisp" as if they were never tried.
Public Sub expand(ByVal source As String)
    Dim fired As Long
    Debug.Print VLA.VlaExpandText(source, True, fired)
End Sub

' REPLEVAL.0: the E the REPL was missing - requested the same session
' COND.0 shipped. expand (above) answers "what does it EXPAND to"; eval
' answers "what does it YIELD" - ordinary runtime arithmetic is the
' clean way to see them differ, since macro-expansion has no reason to
' touch a plain "+" call (only +expand's own eager-fold family is
' expand-time):
'   expand "(+ 1 2)"   (+ 1 2), unchanged - nothing to expand
'   eval "(+ 1 2)"   3
' (length (quote (a b c d))) is NOT this pair's own example anymore,
' by the way - length itself now folds via +expand (LISTOPS-EXPAND
' follow-up, prelude.vla), so expand and eval agree on it too, both
' printing 4 directly.
' VLA_Interpreter.VlaEvalExpression under the hood (IN3_5.0's public
' wrapper: VlaCompileToForms - prelude loaded, defmacros consumed,
' macros fully expanded - then EvalExpr on a fresh frame), so a
' defmacro may precede the expression:
'   eval "(defmacro (double x) (+ x x)) (double 21)"   42
' Expressions only, exactly one body form - a statement ((set! ...),
' (debug-print ...)) belongs to VlaTry, and a compiled-VBA answer with
' *1/*2/*3 recall belongs to VlaTryValue (below), the heavy sibling
' (scratch injection, Application.Run, module state wiped). eval
' injects nothing and wipes nothing - safe to fire at anything from
' the Immediate window, same as expand.
' THE WALL, stated plainly so nobody ever "promotes" this: this is NOT
' the eval the roadmap vetoed. The veto (LISPIMPORT: "eval never was
' [on the table]") targets an EXPANSION-TIME primitive - a computed
' value flowing back into emitted code, breaking pencil-auditability.
' This Sub runs strictly AFTER expansion finishes; its output is ink
' in the Immediate pane, never compiler input - the same post-compile
' exemption LISTOPS-PURITY's own verification already carved for
' EnglishCompileToModule's Application.Run. It lives in the dev rig,
' is not a recognized head in VLA.bas, and must never become one.
' WHOSE answer, on the record: the INTERPRETER's, not compiled VBA's.
' The two are one semantics by intent and TestExprParity hunts
' divergence, but if they ever disagree, this prints the interpreter's
' side - VlaTryValue prints the compiled side.
' NAMED "eval" after the same collision check that cost expand two
' renames on its own way to its name: no module, procedure, or
' Excel-VBA builtin claims the bare name (Access's Eval does not exist
' in Excel VBA), verified by grep and reserved-word list before
' writing, not assumed.
Public Sub eval(ByVal source As String)
    Debug.Print VlaEvalDisplay(VLA_Interpreter.VlaEvalExpression(source))
End Sub

' Display formatter for eval - Public, pinned (TestEvalDisplay,
' VLA_Tests_Grammar.bas), VlaTryValueLit's own precedent. Deliberately
' NOT VlaTryValueLit itself: that one's contract is RE-PARSEABLE
' source text for *1/*2/*3 recall, where an array must refuse; here
' arrays are the whole point - quote data comes back from the
' interpreter as real VBA arrays (CStr on one raises Type mismatch,
' the G6.0 test note), and a REPL that cannot print (1 2 3) would be
' mute on half of LISTOPS. Lisp-flavored on purpose: arrays print as
' parenthesized lists, recursively; Booleans print as VLA's own
' source literals true/false, not VBA's True/False; numbers ride
' Trim$(Str$()) - the locale-proof period, VlaTryValueLit's own
' precedent; strings re-quote with the tokenizer's \" and \\ escapes
' so a string result never masquerades as a number or a symbol.
' Objects and dates keep the sibling's #<...> convention - real
' values, just not VLA-literal-printable. Branch order is load-
' bearing twice: Boolean before IsNumeric (IsNumeric(True) is True in
' VBA), String before IsNumeric (a numeric-looking string stays
' quoted) - the same order VlaTryValueLit already holds.
Public Function VlaEvalDisplay(v As Variant) As String
    If IsObject(v) Then
        VlaEvalDisplay = "#<" & TypeName(v) & ">"
    ElseIf IsArray(v) Then
        Dim parts As String
        Dim i As Long
        For i = LBound(v) To UBound(v)
            If Len(parts) > 0 Then parts = parts & " "
            parts = parts & VlaEvalDisplay(v(i))
        Next
        VlaEvalDisplay = "(" & parts & ")"
    ElseIf IsEmpty(v) Then
        VlaEvalDisplay = "Empty"
    ElseIf IsNull(v) Then
        VlaEvalDisplay = "Null"
    ElseIf VarType(v) = vbBoolean Then
        VlaEvalDisplay = IIf(v, "true", "false")
    ElseIf VarType(v) = vbDate Then
        VlaEvalDisplay = "#<date " & CStr(v) & ">"
    ElseIf VarType(v) = vbString Then
        VlaEvalDisplay = """" & Replace(Replace(CStr(v), "\", "\\"), """", "\""") & """"
    ElseIf IsNumeric(v) Then
        VlaEvalDisplay = Trim$(Str$(v))
    Else
        VlaEvalDisplay = "#<" & TypeName(v) & ">"
    End If
End Function

' The pure half: split the input into balanced top-level chunks -
' honoring "..." strings with \" and \\ escapes and ; comments, the
' tokenizer's own rules, duplicated per rule 12 - hoist the
' module-level-only heads, and wrap everything else in the scratch
' Sub. Raises only on paren imbalance; every other judgement belongs
' to the transpiler, whose errors teach better than a pre-guess.
' Chunks are raw substrings, so a comment INSIDE a form survives
' into the chunk and the real tokenizer strips it later - this
' scanner only refuses to COUNT parens inside strings and comments.
Public Function VlaTryBuild(ByVal source As String) As String
    Dim tops As New Collection
    Dim body As New Collection
    TrySplit source, tops, body

    Dim r As String
    Dim v As Variant
    For Each v In tops
        r = r & CStr(v) & vbCrLf
    Next
    If body.Count = 0 Then
        r = r & "(sub vla-scratch ())" & vbCrLf
    Else
        ' L8.2: the scratch catches its OWN runtime errors and speaks
        ' them to the Immediate window - runtime errors do not cross
        ' Application.Run into the rig's handler (the same seam that
        ' gave translated programs their vla-fail instrumentation and
        ' V6.1 its step-38 dialog), so without this a failing check
        ' ended in VBA's modal instead of words. Fixed label name
        ' (vla-try-oops) - the rig owns it; a user body reusing it
        ' meets VBA's duplicate-label compile step, the C1 failure.
        r = r & "(sub vla-scratch ()" & vbCrLf
        r = r & "  (on-error goto vla-try-oops)" & vbCrLf
        For Each v In body
            r = r & "  " & CStr(v) & vbCrLf
        Next
        r = r & "  (exit-sub)" & vbCrLf
        r = r & "  (label vla-try-oops)" & vbCrLf
        r = r & "  (debug-print (& ""  (scratch error: "" err.description "")""))" & vbCrLf
        ' S1.2 house style: stack the closer onto the last body line.
        r = Left$(r, Len(r) - 2) & ")" & vbCrLf
    End If
    VlaTryBuild = r
End Function

' L6: the scanner, extracted from VlaTryBuild so both builders ride
' ONE splitter - same balanced-chunk walk, same string/comment rules,
' same imbalance refusals. A second scanner would be a second set of
' edge cases; the G10 design doctrine (one expansion, reused) applies
' to scanners too.
Private Sub TrySplit(ByVal source As String, tops As Collection, body As Collection)
    Dim n As Long, i As Long, depth As Long
    Dim formStart As Long, bareStart As Long
    Dim inString As Boolean
    Dim c As String, d As String
    n = Len(source)
    i = 1
    Do While i <= n
        c = Mid$(source, i, 1)
        If inString Then
            If c = "\" Then
                d = Mid$(source, i + 1, 1)
                If d = """" Or d = "\" Then i = i + 2 Else i = i + 1
            ElseIf c = """" Then
                inString = False
                i = i + 1
            Else
                i = i + 1
            End If
        ElseIf c = ";" Then
            If depth = 0 And bareStart > 0 Then
                TryAddChunk tops, body, Mid$(source, bareStart, i - bareStart)
                bareStart = 0
            End If
            Do While i <= n
                c = Mid$(source, i, 1)
                If c = vbCr Or c = vbLf Then Exit Do
                i = i + 1
            Loop
        ElseIf c = """" Then
            inString = True
            If depth = 0 And bareStart = 0 Then bareStart = i
            i = i + 1
        ElseIf c = "(" Then
            If depth = 0 Then
                If bareStart > 0 Then
                    TryAddChunk tops, body, Mid$(source, bareStart, i - bareStart)
                    bareStart = 0
                End If
                formStart = i
            End If
            depth = depth + 1
            i = i + 1
        ElseIf c = ")" Then
            depth = depth - 1
            If depth < 0 Then
                Err.Raise 5, "VlaTry", "unbalanced ')' - more closers than openers (strings and ; comments were honored)"
            End If
            If depth = 0 Then
                TryAddChunk tops, body, Mid$(source, formStart, i - formStart + 1)
                formStart = 0
            End If
            i = i + 1
        ElseIf c = " " Or c = vbTab Or c = vbCr Or c = vbLf Then
            If depth = 0 And bareStart > 0 Then
                TryAddChunk tops, body, Mid$(source, bareStart, i - bareStart)
                bareStart = 0
            End If
            i = i + 1
        Else
            If depth = 0 And bareStart = 0 Then bareStart = i
            i = i + 1
        End If
    Loop
    If depth > 0 Then
        Err.Raise 5, "VlaTry", "unclosed '(' - the input ended " & depth & " level(s) deep (strings and ; comments were honored)"
    End If
    If bareStart > 0 Then TryAddChunk tops, body, Mid$(source, bareStart, n - bareStart + 1)
End Sub

' Route one chunk: module-level-ONLY heads hoist above the scratch
' Sub; everything else - statements, dim/const (procedure locals in
' a scratch), raw rows, begin blocks, bare atoms the transpiler will
' judge - rides the body in original order. dim/const/raw/begin are
' legal at BOTH levels; the body reading is the scratch's natural
' one, and type/enum, legal only at module level (the emitter's own
' teaching refusal says so), are hoisted along with the definitions.
Private Sub TryAddChunk(tops As Collection, body As Collection, ByVal chunk As String)
    If Len(Trim$(chunk)) = 0 Then Exit Sub
    Select Case TryChunkHead(chunk)
        Case "sub", "function", "defmacro", "public", "private", "type", "enum"
            tops.Add chunk
        Case Else
            body.Add chunk
    End Select
End Sub

' Lowercased head symbol of a "(...)" chunk, or "" for a bare chunk.
Private Function TryChunkHead(ByVal chunk As String) As String
    If Left$(chunk, 1) <> "(" Then Exit Function
    Dim i As Long, n As Long, start As Long
    Dim c As String
    n = Len(chunk)
    i = 2
    Do While i <= n
        c = Mid$(chunk, i, 1)
        If c <> " " And c <> vbTab And c <> vbCr And c <> vbLf Then Exit Do
        i = i + 1
    Loop
    start = i
    Do While i <= n
        c = Mid$(chunk, i, 1)
        If c = "(" Or c = ")" Or c = " " Or c = vbTab _
           Or c = vbCr Or c = vbLf Or c = ";" Or c = """" Then Exit Do
        i = i + 1
    Loop
    TryChunkHead = VLA_Identity.Fold(Mid$(chunk, start, i - start))
End Function

' Delete the scratch module. Module removal is deferred (hazard 11),
' so a plain Remove of "VLA_Scratch" could collide with the next
' injection inside one call chain - VlaDevReload's rename-aside
' pattern, duplicated per rule 12: rename to _old, remove; the name
' is free immediately. Best-effort throughout: a scratch that is not
' there is already what we want.
Private Sub TryDropScratch()
    Dim comp As Object
    On Error Resume Next
    Set comp = Nothing
    Set comp = ThisWorkbook.VBProject.VBComponents("VLA_Scratch_old")
    If Not comp Is Nothing Then ThisWorkbook.VBProject.VBComponents.Remove comp
    Set comp = Nothing
    Set comp = ThisWorkbook.VBProject.VBComponents("VLA_Scratch")
    If Not comp Is Nothing Then
        comp.Name = "VLA_Scratch_old"
        ThisWorkbook.VBProject.VBComponents.Remove comp
    End If
    On Error GoTo 0
End Sub

' =====================================================================
'  L6: VlaTryValue - the true eval, riding L2's machinery. VlaTry
'  answers "did it run"; this answers "what does it YIELD":
'
'    VlaTryValue "(+ 1 2)"           =>  3
'    VlaTryValue "(set! x 3) (* x x)"    =>  9
'
'  The input splits through the SAME scanner as VlaTry (TrySplit -
'  one splitter, one set of edge cases): module-level definitions
'  hoist, every body chunk but the LAST is a statement, and the last
'  chunk is the value - wrapped as (return ...) inside
'  (function vla-scratch-value () Variant ...). Application.Run
'  returns the function's value; the rig prints it and stores it.
'
'  *1 / *2 / *3 - the last three results, recallable in the NEXT
'  input, VlaTry or VlaTryValue alike:
'
'    VlaTryValue "(* *1 2)"          doubles the last result
'    VlaTry "(debug-print *1)"       prints it
'
'  Recall is a TEXTUAL substitution before the scratch is built -
'  token-boundary aware, never inside "..." strings or ; comments
'  (the scanner rules again). History rides workbook names
'  (VLAt_Star1/2/3) - the VlaTimeIt pattern, and the REASON is the
'  same: every scratch injection wipes module-level state (S5.1),
'  so module state cannot remember anything across evals; Names
'  survive the wipe, the reloads, and the saves.
'
'  VlaTryValueBuild(source)      the pure half (split + wrap), pinned
'  VlaTryExpandStars(...)        the pure substitution, pinned
'  VlaTryValueLit(v, ok)         the pure value formatter, pinned
'
'  Honest limits, recorded:
'  - The LAST form must be an EXPRESSION. A Sub-shaped statement
'    there ((debug-print 3), say) emits an illegal assignment and
'    surfaces VBA's OWN modal during Application.Run - compile
'    errors cross Run untrappably (S4's verdict), VlaTry's recorded
'    limit in a new coat. The dev can read the modal.
'  - Recallable values are scalars: numbers (written with Str$'s
'    locale-proof period), strings (re-quoted with the tokenizer's
'    \" and \\ escapes, ceiling ~180 chars), True/False. Objects,
'    arrays, Empty, Null, dates store as #<...> markers that REFUSE
'    substitution with words - *1 never silently injects garbage.
'  - Only VlaTryValue writes history; VlaTry only reads it (a Sub
'    run yields no value to store).
'  - The injection wipes loaded grammar exactly like VlaTry - reload
'    the vocabulary before the next Check/Run.
'  Dev-rig only, never in the add-in - exactly like this module.
' =====================================================================

Public Sub VlaTryValue(ByVal source As String, Optional ByVal showCode As Boolean = False)
    Dim phase As String
    Dim scratch As String
    Debug.Print "=== VlaTryValue"
    On Error GoTo oops
    phase = "recall history (*1/*2/*3)"
    source = TryExpandStars(source)
    phase = "build the scratch"
    scratch = VlaTryValueBuild(source)
    If showCode Then
        phase = "transpile the scratch"
        Debug.Print "  scratch VLA:"
        DevPrintIndented scratch
        Debug.Print "  VBA:"
        DevPrintIndented VlaTranspile(scratch)
    End If
    phase = "inject the scratch module"
    TryDropScratch                       ' sweep any stranded scratch first
    VlaCompileToModule scratch, "VLA_Scratch", ThisWorkbook
    phase = "run the scratch"
    Dim v As Variant
    v = Application.Run("'" & ThisWorkbook.Name & "'!VLA_Scratch.vla_scratch_value")
    phase = "clean up the scratch module"
    TryDropScratch
    Dim ok As Boolean
    Dim lit As String
    lit = VlaTryValueLit(v, ok)
    TryShiftHistory lit
    Debug.Print "  => " & lit
    Debug.Print "  (recall it as *1; older values shifted to *2, *3)"
    Exit Sub
oops:
    Dim msg As String
    msg = Err.Description
    On Error Resume Next
    TryDropScratch
    On Error GoTo 0
    Debug.Print "  (could not " & phase & ": " & msg & ")"
    If Not showCode Then Debug.Print "  (VlaTryValue source, True shows the scratch text a 'vla line N' indexes)"
End Sub

' The pure half: split through TrySplit, hoist the tops, wrap the
' body in the value function with the last chunk as the return.
' Raises with words when there is nothing to return.
Public Function VlaTryValueBuild(ByVal source As String) As String
    Dim tops As New Collection
    Dim body As New Collection
    TrySplit source, tops, body
    If body.Count = 0 Then
        Err.Raise 5, "VlaTryValue", "give me an expression to evaluate - definitions alone have no value (the last form must be an expression)"
    End If
    Dim r As String
    Dim v As Variant
    Dim i As Long
    For Each v In tops
        r = r & CStr(v) & vbCrLf
    Next
    r = r & "(function vla-scratch-value () Variant" & vbCrLf
    ' L8.2: same self-handler as VlaTryBuild (see the note there).
    ' The error path prints the words and falls out returning Empty,
    ' so history stores the #<empty> marker - which already refuses
    ' recall with words.
    r = r & "  (on-error goto vla-try-oops)" & vbCrLf
    For i = 1 To body.Count - 1
        r = r & "  " & CStr(body.Item(i)) & vbCrLf
    Next
    r = r & "  (return " & CStr(body.Item(body.Count)) & ")" & vbCrLf
    r = r & "  (label vla-try-oops)" & vbCrLf
    r = r & "  (debug-print (& ""  (scratch error: "" err.description "")""))" & vbCrLf
    ' S1.2 house style: stack the closer onto the last body line.
    r = Left$(r, Len(r) - 2) & ")" & vbCrLf
    VlaTryValueBuild = r
End Function

' The pure substitution, pinned by the self-test: replace *1/*2/*3
' at token boundaries with the given texts, honoring strings and
' ; comments exactly as TrySplit does. On a problem (an empty slot,
' a #<...> marker), sets problem to a teaching sentence and returns
' the input untouched - the caller raises with those words.
Public Function VlaTryExpandStars(ByVal source As String, ByVal s1 As String, _
                                  ByVal s2 As String, ByVal s3 As String, _
                                  ByRef problem As String) As String
    problem = ""
    VlaTryExpandStars = source
    Dim r As String
    Dim n As Long, i As Long
    Dim inString As Boolean
    Dim c As String, d As String
    Dim prevDelim As Boolean
    n = Len(source)
    i = 1
    prevDelim = True
    Do While i <= n
        c = Mid$(source, i, 1)
        If inString Then
            If c = "\" Then
                d = Mid$(source, i + 1, 1)
                If d = """" Or d = "\" Then
                    r = r & c & d
                    i = i + 2
                Else
                    r = r & c
                    i = i + 1
                End If
            Else
                If c = """" Then inString = False
                r = r & c
                i = i + 1
            End If
            prevDelim = False
        ElseIf c = ";" Then
            Do While i <= n
                c = Mid$(source, i, 1)
                If c = vbCr Or c = vbLf Then Exit Do
                r = r & c
                i = i + 1
            Loop
            prevDelim = True
        ElseIf c = """" Then
            inString = True
            r = r & c
            i = i + 1
            prevDelim = False
        ElseIf c = "*" And prevDelim And i < n And InStr("123", Mid$(source, i + 1, 1)) > 0 _
               And (i + 1 = n Or InStr(" ()" & vbTab & vbCr & vbLf & ";", Mid$(source, i + 2, 1)) > 0) Then
            Dim slot As String
            Dim sv As String
            slot = Mid$(source, i + 1, 1)
            If slot = "1" Then
                sv = s1
            ElseIf slot = "2" Then
                sv = s2
            Else
                sv = s3
            End If
            If Len(sv) = 0 Then
                problem = "there is no *" & slot & " yet - evaluate something with VlaTryValue first"
                Exit Function
            End If
            If Left$(sv, 2) = "#<" Then
                problem = "*" & slot & " holds " & sv & " - that value could not be stored for recall"
                Exit Function
            End If
            r = r & sv
            i = i + 2
            prevDelim = False
        Else
            prevDelim = (c = " " Or c = "(" Or c = ")" Or c = vbTab Or c = vbCr Or c = vbLf)
            r = r & c
            i = i + 1
        End If
    Loop
    VlaTryExpandStars = r
End Function

' The impure wrapper: read the three slots from the Names and raise
' with the substitution's own words on a problem.
Private Function TryExpandStars(ByVal source As String) As String
    Dim prob As String
    TryExpandStars = VlaTryExpandStars(source, DevNameGetText("VLAt_Star1"), _
                                       DevNameGetText("VLAt_Star2"), _
                                       DevNameGetText("VLAt_Star3"), prob)
    If Len(prob) > 0 Then Err.Raise 5, "VlaTry", prob
End Function

' The pure value formatter, pinned: a scalar becomes the VLA literal
' that re-reads as itself (ok = True); everything else becomes a
' #<...> marker that the substitution refuses (ok = False). Numbers
' go through Str$ - period decimal on every locale, unlike CStr.
Public Function VlaTryValueLit(v As Variant, ByRef ok As Boolean) As String
    ok = False
    If IsObject(v) Then
        VlaTryValueLit = "#<" & TypeName(v) & ">"
    ElseIf IsArray(v) Then
        VlaTryValueLit = "#<array>"
    ElseIf IsEmpty(v) Then
        VlaTryValueLit = "#<empty>"
    ElseIf IsNull(v) Then
        VlaTryValueLit = "#<null>"
    ElseIf VarType(v) = vbBoolean Then
        ok = True
        VlaTryValueLit = IIf(v, "True", "False")
    ElseIf VarType(v) = vbDate Then
        VlaTryValueLit = "#<date " & CStr(v) & ">"
    ElseIf VarType(v) = vbString Then
        If Len(v) > 180 Then
            VlaTryValueLit = "#<a " & Len(v) & "-character string - too long to recall>"
        Else
            ok = True
            VlaTryValueLit = """" & Replace(Replace(CStr(v), "\", "\\"), """", "\""") & """"
        End If
    ElseIf IsNumeric(v) Then
        ok = True
        VlaTryValueLit = Trim$(Str$(v))
    Else
        VlaTryValueLit = "#<" & TypeName(v) & ">"
    End If
End Function

' Shift the history: *3 <- *2 <- *1 <- the new literal. Names, not
' module state - the injection that produced the value has already
' proven why (it wipes module state on every eval).
Private Sub TryShiftHistory(ByVal lit As String)
    DevNameSetText "VLAt_Star3", DevNameGetText("VLAt_Star2")
    DevNameSetText "VLAt_Star2", DevNameGetText("VLAt_Star1")
    DevNameSetText "VLAt_Star1", lit
End Sub

' Text-valued workbook names - the VLAt_Time* pattern's sibling,
' storing an arbitrary short text as ="..." with doubled quotes.
' A refused Add (odd content, protection) loses one history slot,
' never the eval - On Error says so.
Private Sub DevNameSetText(ByVal nm As String, ByVal txt As String)
    On Error Resume Next
    ThisWorkbook.Names(nm).Delete
    On Error GoTo 0
    If Len(txt) = 0 Then Exit Sub
    On Error Resume Next
    ThisWorkbook.Names.Add nm, "=" & Chr$(34) & Replace(txt, Chr$(34), Chr$(34) & Chr$(34)) & Chr$(34)
    On Error GoTo 0
End Sub

Private Function DevNameGetText(ByVal nm As String) As String
    Dim v As String
    On Error Resume Next
    v = ThisWorkbook.Names(nm).RefersTo
    On Error GoTo 0
    If Len(v) >= 3 And Left$(v, 2) = "=" & Chr$(34) And Right$(v, 1) = Chr$(34) Then
        DevNameGetText = Replace(Mid$(v, 3, Len(v) - 3), Chr$(34) & Chr$(34), Chr$(34))
    End If
End Function

Private Sub DevPrintIndented(ByVal s As String)
    Dim lines As Variant
    Dim i As Long
    s = Replace(s, vbCrLf, vbLf)
    lines = Split(s, vbLf)
    For i = LBound(lines) To UBound(lines)
        If Len(Trim$(CStr(lines(i)))) > 0 Then Debug.Print "    " & lines(i)
    Next
End Sub

' =====================================================================
'  V6: VlaTimeIt - the timing harness (Alpha 4, Band 0; adopted at
'  the owner's word). Four dials, milliseconds:
'
'    vocab load   - reset + EnglishLoadVocabulary (the Reload click)
'    translate    - EnglishToVla over the whole instructions.txt corpus
'    transpile    - VlaTranspile of that translation
'    run (full)   - EnglishRunProgram end-to-end, INCLUSIVE by
'                   design: reset + load + translate + inject +
'                   execute is what a user's Run click costs, and
'                   the empty-Run finding is about exactly that
'                   inclusive number.
'
'  VlaTimeIt            averages load/translate/transpile over 3
'                       reps (Timer granularity is ~16 ms, so small
'                       legs need averaging); the Run leg runs ONCE.
'  VlaTimeIt 10         more reps for steadier small-leg numbers.
'
'  Every number is stored in workbook names (VLAt_TimeLoad /
'  Translate / Transpile / Run - visible in Name Manager, surviving
'  grammar wipes, module reloads, and workbook saves), and each
'  line prints CURRENT next to PREVIOUS - the before/after dial the
'  V7/V8 promotion gate reads. The standing rule from Alpha 4: no
'  optimization proceeds without a before-number from this dial.
'
'  Honest notes: the Run leg runs instructions.txt FOR REAL - the same
'  sheet effects as clicking Run, module Frazaro_EN_TimeIt left behind like
'  any program module, and the dev project's module-level state
'  reset by the injection (S5.1) - reload the vocabulary before the
'  next Check, and run this dial LAST in any session's measuring.
'  The Run leg targets the ACTIVE workbook, exactly like the IDE's
'  Run click - the operator aims by activating the instructions.txt
'  workspace workbook (the one with the Output sheet) first. A
'  pre-flight refuses with directions, before anything is touched,
'  when the aim is wrong (V6.2; V6.1's activate-the-dev-workbook
'  was a disproven guess - the dev workbook has no Output either).
'  Timer wraps at midnight (handled); resolution is ~16 ms, fine
'  for these macro legs and documented so nobody micro-benchmarks
'  with it (that is L7's (time ...) territory, when adopted).
'  Dev-rig only, never in the add-in. No self-test pins, with the
'  reasoning recorded: the Run leg cannot execute inside the suite
'  (it runs the corpus and wipes module state mid-suite), and the
'  tool's printed numbers ARE its verification - VlaDevReload's
'  precedent for dev-loop tools.
' =====================================================================

Public Sub VlaTimeIt(Optional ByVal reps As Long = 3)
    If reps < 1 Then reps = 1
    Dim vocabPath As String, progPath As String
    vocabPath = DevFindFile("english.vla")
    progPath = DevFindFile("instructions.txt")
    Dim progText As String
    progText = DevReadTextFileUtf8(progPath)

    ' previous readings first - the set below overwrites them
    Dim pLoad As String, pTrans As String, pEmit As String, pRun As String
    pLoad = DevNameGet("VLAt_TimeLoad")
    pTrans = DevNameGet("VLAt_TimeTranslate")
    pEmit = DevNameGet("VLAt_TimeTranspile")
    pRun = DevNameGet("VLAt_TimeRun")

    Debug.Print "=== VlaTimeIt (reps=" & reps & " for load/translate/transpile; Run once; Timer ~16 ms)"

    Dim msLoad As Double, msTrans As Double, msEmit As Double, msRun As Double
    Dim r As Long
    Dim t0 As Double
    Dim vla As String, vbaOut As String

    For r = 1 To reps
        EnglishResetGrammar
        t0 = Timer
        EnglishLoadVocabulary vocabPath
        msLoad = msLoad + DevMs(t0)
    Next
    msLoad = msLoad / reps

    For r = 1 To reps
        t0 = Timer
        vla = EnglishToVla(progText)
        msTrans = msTrans + DevMs(t0)
    Next
    msTrans = msTrans / reps

    For r = 1 To reps
        t0 = Timer
        vbaOut = VlaTranspile(vla)
        msEmit = msEmit + DevMs(t0)
    Next
    msEmit = msEmit / reps

    ' store the three isolated legs BEFORE the Run leg wipes module
    ' state - the names survive everything
    DevNameSet "VLAt_TimeLoad", msLoad
    DevNameSet "VLAt_TimeTranslate", msTrans
    DevNameSet "VLAt_TimeTranspile", msEmit

    ' V6.2 (second maiden-run incident - the V6.1 fix was a wrong
    ' guess, and the second failure disproved it): step 38 died with
    ' the DEV workbook active too, so the dev workbook has no Output
    ' sheet either - the corpus's real contract is "run in a
    ' workbook containing a sheet named Output" (line 62 is its
    ' first sheet-qualified reference), and only the OPERATOR knows
    ' which workbook that is: the owner's instructions.txt runs happen in
    ' a separate workspace workbook, aimed by activating it before
    ' clicking Run. So the harness stops guessing: no activation -
    ' the Run leg targets whatever the operator aimed at, exactly
    ' like the IDE's click - and a PRE-FLIGHT check refuses with
    ' directions BEFORE the run when the active workbook lacks
    ' Output, so a wrong aim costs zero scribbled steps instead of
    ' thirty-seven. The three measured legs still print and store.
    If Not DevHasSheet(ActiveWorkbook, "Output") Then
        Debug.Print "  vocab load:  " & Format$(msLoad, "0") & " ms" & DevPrevTag(pLoad)
        Debug.Print "  translate:   " & Format$(msTrans, "0") & " ms" & DevPrevTag(pTrans)
        Debug.Print "  transpile:   " & Format$(msEmit, "0") & " ms" & DevPrevTag(pEmit)
        Debug.Print "  run (full):  SKIPPED - the active workbook ('" & ActiveWorkbook.Name & _
                    "') has no sheet named Output, which instructions.txt requires (line 62 writes Output!H16)."
        Debug.Print "               Activate your instructions.txt workspace workbook and run VlaTimeIt again - nothing was touched."
        Exit Sub
    End If
    t0 = Timer
    EnglishRunProgram progPath, "Frazaro_EN_TimeIt", vocabPath
    msRun = DevMs(t0)
    DevNameSet "VLAt_TimeRun", msRun

    Debug.Print "  vocab load:  " & Format$(msLoad, "0") & " ms" & DevPrevTag(pLoad)
    Debug.Print "  translate:   " & Format$(msTrans, "0") & " ms" & DevPrevTag(pTrans)
    Debug.Print "  transpile:   " & Format$(msEmit, "0") & " ms" & DevPrevTag(pEmit)
    Debug.Print "  run (full):  " & Format$(msRun, "0") & " ms" & DevPrevTag(pRun) & "   <- inclusive: reset+load+translate+inject+execute"
    Debug.Print "  (stored in workbook names VLAt_Time*; V7/V8 read this dial for their before-numbers)"
    Debug.Print "  (the Run leg ran instructions.txt for real - reload the vocabulary before the next Check)"
End Sub

' =====================================================================
'  AS.7: VlaTimeItSelfTest - VlaTimeIt's own suite-runtime counterpart.
'  Two legs, milliseconds, each run ONCE (a full suite pass, unlike
'  VlaTimeIt's small legs averaged over reps - Timer's ~16 ms
'  resolution is noise next to a multi-second suite run):
'
'    pure   - VLA_Tests.VlaSelfTest() (host-independent)
'    host   - VLA_Tests_Host.VlaSelfTestHost() (needs this workbook live)
'
'  Same shape as VlaTimeIt's other dials: stored in workbook names
'  (VLAt_TimeSelfTestPure / VLAt_TimeSelfTestHost - Name Manager,
'  survives grammar wipes/reloads/saves), CURRENT prints next to
'  PREVIOUS.
'
'  What this is actually for, and what it honestly cannot be: AS.3
'  (mutation testing, unbuilt) deliberately breaks the emitter/
'  interpreter to prove the suite notices - a mutant that negates a
'  comparison or flips an increment can turn a terminating loop
'  infinite. VBA has no per-statement timeout, so a hung mutant just
'  looks like a frozen VBE with no way to tell "slow" from "stuck."
'  This Sub cannot watch a run WHILE it hangs - if the suite never
'  returns, Timer never reports and nothing below prints, the same
'  limit VlaTimeIt's own Run leg has always had. What it gives instead
'  is the number to hold in your head going in: "ordinary is ~N
'  seconds" - so a VBE still frozen at 10x that during AS.3 reads as
'  stuck, not slow, without needing to have just run this a moment
'  before.
'
'  No self-test pin, no automated pass/fail threshold on the numbers
'  themselves - same reasoning already on record for VlaTimeIt: a
'  threshold could only fire AFTER a hang a human already has to kill
'  by hand, so it would add noise on a slower dev machine, never real
'  signal. Dev-rig only, never in the add-in, like VlaTimeIt. Both legs
'  print their own PASS/FAIL-per-check output as they run (VlaSelfTest/
'  VlaSelfTestHost's own existing behavior, unchanged) - this Sub only
'  adds the elapsed-time line after each.
' =====================================================================

Public Sub VlaTimeItSelfTest()
    Dim pPure As String, pHost As String
    pPure = DevNameGet("VLAt_TimeSelfTestPure")
    pHost = DevNameGet("VLAt_TimeSelfTestHost")

    Debug.Print "=== VlaTimeItSelfTest (each leg runs once; Timer ~16 ms)"

    Dim t0 As Double
    Dim pureOk As Boolean, hostOk As Boolean
    Dim msPure As Double, msHost As Double

    t0 = Timer
    pureOk = VLA_Tests.VlaSelfTest()
    msPure = DevMs(t0)
    DevNameSet "VLAt_TimeSelfTestPure", msPure

    t0 = Timer
    hostOk = VLA_Tests_Host.VlaSelfTestHost()
    msHost = DevMs(t0)
    DevNameSet "VLAt_TimeSelfTestHost", msHost

    Debug.Print "  pure (VlaSelfTest):      " & Format$(msPure, "0") & " ms" & DevPrevTag(pPure) & "   " & IIf(pureOk, "PASS", "FAIL")
    Debug.Print "  host (VlaSelfTestHost):  " & Format$(msHost, "0") & " ms" & DevPrevTag(pHost) & "   " & IIf(hostOk, "PASS", "FAIL")
    Debug.Print "  (stored in workbook names VLAt_TimeSelfTest*; a hang during AS.3 mutation testing means neither line above ever prints - hold the PREVIOUS number as the 'ordinary' baseline while watching a frozen VBE)"
    Debug.Print "  (the pure leg reset the grammar - Reload vocabulary before the next Run)"
End Sub

' =====================================================================
'  IN.8: VlaTimeItInterpreter - VlaTimeIt's own interpreter-side
'  counterpart, the same shape VerifyReportInterpreter already is to
'  VerifyReport: runs the SAME instructions.txt corpus through the SAME
'  vocab-load and translate legs (both backends read it through the
'  same EnglishToVla), diverging only where the backends actually
'  diverge - "transpile" becomes "compile-to-forms" (VlaCompileToForms
'  instead of VlaTranspile) and "run (full)" becomes "interpret
'  (full)" (VlaInterpret instead of inject-then-Application.Run - no
'  module is ever created on this path, IN.9's architecture made
'  visible in the timing itself, not just the roadmap).
'
'  Prints its own four legs AND, when VlaTimeIt has already been run
'  in this workbook (its number lives in a workbook Name, so a prior
'  session counts), IN.8's own headline number: the measured factor,
'  read straight out of the Name VlaTimeIt already wrote, never
'  re-derived or estimated.
'
'  Same aim discipline as VlaTimeIt's own Run leg (V6.2): targets
'  whatever workbook is ACTIVE, refusing with directions before
'  touching anything if it lacks a sheet named Output.
'  Dev-rig only, never in the add-in - exactly like VlaTimeIt.
' =====================================================================

Public Sub VlaTimeItInterpreter(Optional ByVal reps As Long = 3)
    If reps < 1 Then reps = 1
    Dim vocabPath As String, progPath As String
    vocabPath = DevFindFile("english.vla")
    progPath = DevFindFile("instructions.txt")
    Dim progText As String
    progText = DevReadTextFileUtf8(progPath)

    Dim pLoad As String, pTrans As String, pForms As String, pRun As String
    pLoad = DevNameGet("VLAt_TimeLoadI")
    pTrans = DevNameGet("VLAt_TimeTranslateI")
    pForms = DevNameGet("VLAt_TimeFormsI")
    pRun = DevNameGet("VLAt_TimeRunI")

    Debug.Print "=== VlaTimeItInterpreter (reps=" & reps & " for load/translate/compile-to-forms; Interpret once; Timer ~16 ms)"

    Dim msLoad As Double, msTrans As Double, msForms As Double, msRun As Double
    Dim r As Long
    Dim t0 As Double
    Dim vla As String
    Dim forms As Collection

    For r = 1 To reps
        EnglishResetGrammar
        t0 = Timer
        EnglishLoadVocabulary vocabPath
        msLoad = msLoad + DevMs(t0)
    Next
    msLoad = msLoad / reps

    ' VerifyReportInterpreter's own precondition, missed on the first
    ' pass here and caught live (a real crash, not a guess): step
    ' tracking (default on) wraps every statement in the emitter's own
    ' vlatraceon/vlatracestep trace scaffold, which is not part of the
    ' interpreter's dispatch surface at all (VlaInterpreterEffectLog is
    ' the interpreter's own, separate tracing) - so a traced compile
    ' run through VlaInterpret dies with "'begin' is not a form ...".
    ' Off for every EnglishToVla call this sub makes, restored
    ' immediately after each, the same discipline VerifyReportInterpreter
    ' and VlaSkeletonDemoVla already established for this exact reason.
    EnglishStepTracking False
    For r = 1 To reps
        t0 = Timer
        vla = EnglishToVla(progText)
        msTrans = msTrans + DevMs(t0)
    Next
    EnglishStepTracking True
    msTrans = msTrans / reps

    For r = 1 To reps
        t0 = Timer
        Set forms = VlaCompileToForms(vla)
        msForms = msForms + DevMs(t0)
    Next
    msForms = msForms / reps

    DevNameSet "VLAt_TimeLoadI", msLoad
    DevNameSet "VLAt_TimeTranslateI", msTrans
    DevNameSet "VLAt_TimeFormsI", msForms

    If Not DevHasSheet(ActiveWorkbook, "Output") Then
        Debug.Print "  vocab load:        " & Format$(msLoad, "0") & " ms" & DevPrevTag(pLoad)
        Debug.Print "  translate:         " & Format$(msTrans, "0") & " ms" & DevPrevTag(pTrans)
        Debug.Print "  compile-to-forms:  " & Format$(msForms, "0") & " ms" & DevPrevTag(pForms)
        Debug.Print "  interpret (full):  SKIPPED - the active workbook ('" & ActiveWorkbook.Name & _
                    "') has no sheet named Output, which instructions.txt requires (line 62 writes Output!H16)."
        Debug.Print "               Activate your instructions.txt workspace workbook and run VlaTimeItInterpreter again - nothing was touched."
        Exit Sub
    End If

    ' IN.9: no module injection on this leg at all - VlaInterpret walks
    ' the forms directly against the ACTIVE workbook, passed explicitly
    ' so "thisworkbook" inside the corpus resolves the way a real
    ' install would, not the add-in (IN.6's own fix, exercised here
    ' rather than just relied upon).
    t0 = Timer
    EnglishResetGrammar
    EnglishLoadVocabulary vocabPath
    ' Same step-tracking guard as the translate leg above - tracking
    ' is restored to True the moment translation finishes, BEFORE
    ' VlaInterpret ever runs, so a real interpreter failure during
    ' execution (unrelated to this) can never leave it stuck off for
    ' the rest of the session (LESSONS.md's own named module-state
    ' hazard).
    EnglishStepTracking False
    Dim vlaRun As String
    vlaRun = EnglishToVla(progText)
    EnglishStepTracking True
    VLA_Interpreter.VlaInterpret vlaRun, ActiveWorkbook
    msRun = DevMs(t0)
    DevNameSet "VLAt_TimeRunI", msRun

    Debug.Print "  vocab load:        " & Format$(msLoad, "0") & " ms" & DevPrevTag(pLoad)
    Debug.Print "  translate:         " & Format$(msTrans, "0") & " ms" & DevPrevTag(pTrans)
    Debug.Print "  compile-to-forms:  " & Format$(msForms, "0") & " ms" & DevPrevTag(pForms)
    Debug.Print "  interpret (full):  " & Format$(msRun, "0") & " ms" & DevPrevTag(pRun) & _
                "   <- inclusive: reset+load+translate+interpret, no module ever created"
    Debug.Print "  (stored in workbook names VLAt_Time*I; the Run leg ran instructions.txt for real - reload the vocabulary before the next Check)"

    Dim compiledRun As String
    compiledRun = DevNameGet("VLAt_TimeRun")
    If Len(compiledRun) = 0 Then
        Debug.Print "  IN.8 factor: run VlaTimeIt in this workbook first - its 'run (full)' number is the other half of this ratio."
    Else
        Dim ratio As Double
        ratio = msRun / CDbl(compiledRun)
        Debug.Print "  IN.8 factor (instructions.txt, whole corpus): interpreter is " & Format$(ratio, "0.0") & _
                    "x the compiled-VBA 'run (full)' time (" & compiledRun & " ms compiled vs " & Format$(msRun, "0") & " ms interpreted)."
    End If
End Sub

' =====================================================================
'  IN.8: VlaBenchmarkLoop - the dispatch-overhead half of "measure, not
'  estimate": a pure arithmetic loop (IN.0.5's own ten-core-form shape
'  - dim/set!/for/+/debug-print, no object model at all) run through
'  BOTH backends on the IDENTICAL source text, at a few loop sizes,
'  isolating the walking-forms tax itself from instructions.txt's own
'  fixed, mixed-corpus weight (vocab load, translate, the object-model
'  dispatch its real macros also pay for). This is the number that
'  names which workloads the interpreter is unacceptable for, per this
'  item's own opening line - a loop-shaped program is exactly PF.4's
'  own "two orders of magnitude" concern and the shape the export
'  exists for.
'
'  Compiled leg reuses VlaTry's own inject/run/delete idiom (a
'  throwaway module named "VLA_Bench", not "VLA_Scratch", so a
'  benchmark run mid-VlaTry-session never touches the scratch module
'  that session owns). Interpreted leg calls VlaInterpret directly -
'  no module, by construction (IN.9). Needs no worksheet, no
'  instructions.txt, no vocabulary load - safe in any workbook, any
'  session, unlike VlaTimeIt/VlaTimeItInterpreter above.
'
'  VlaBenchmarkLoop                    default sweep: 300 / 1,000 /
'                                       3,000 iterations
'  VlaBenchmarkLoop Array(200000)       a single, larger size - see the
'                                       up-front time estimate this
'                                       prints before running it
'
'  Every size's pair is stored in workbook Names (VLAt_BenchLoop{C,I}_N)
'  so a rerun prints before/after, VlaTimeIt's own precedent.
'
'  Honest number, measured rather than guessed on the FIRST real run of
'  this tool: the interpreted leg costs roughly 1-2 ms per iteration
'  (dim/set!/for/+/debug-print, nothing else), which made this sub's own
'  original default sweep (1,000/10,000/50,000) silently freeze Excel
'  for about 100 seconds with no warning printed first - Excel is
'  unresponsive for the whole call, no DoEvents inside the
'  interpreter's own loop. The default sweep is now small enough to
'  return in well under a second; a custom, larger `sizes` argument
'  gets an ESTIMATED-duration warning up front instead of silence.
' =====================================================================

Public Sub VlaBenchmarkLoop(Optional ByVal sizes As Variant)
    If IsMissing(sizes) Then sizes = Array(300, 1000, 3000)
    Debug.Print "=== VlaBenchmarkLoop (pure arithmetic loop: dim/set!/for/+/debug-print only, no object model)"
    Dim totalN As Double
    Dim v As Variant
    For Each v In sizes
        totalN = totalN + CDbl(v)
    Next
    ' The same "refuse/warn before touching anything" instinct
    ' VlaTimeIt's own pre-flight check already uses (V6.2), applied
    ' here to an estimated duration instead of a missing sheet: 2
    ' ms/iteration is this session's own measured ceiling, not a
    ' guess, and Excel gives no progress feedback until the whole
    ' call returns.
    If totalN * 2 > 10000 Then
        Debug.Print "  heads up: ~" & Format$(totalN, "#,##0") & " total interpreted iterations at " & _
                    "roughly 1-2 ms/iteration (this session's own measured range) is an ESTIMATED " & _
                    Format$(totalN * 2 / 1000, "0.0") & "s or more - Excel will be unresponsive the " & _
                    "whole time. Ctrl+Break to abort."
    End If
    For Each v In sizes
        BenchOneLoopSize CLng(v)
    Next
End Sub

Private Sub BenchOneLoopSize(ByVal n As Long)
    Dim src As String
    src = "(sub main ()" & vbCrLf & _
          "  (dim total Double)" & vbCrLf & _
          "  (dim i Double)" & vbCrLf & _
          "  (set! total 0)" & vbCrLf & _
          "  (for (i 1 " & n & ") (set! total (+ total i)))" & vbCrLf & _
          "  (debug-print total))" & vbCrLf

    Dim pC As String, pI As String
    pC = DevNameGet("VLAt_BenchLoopC_" & n)
    pI = DevNameGet("VLAt_BenchLoopI_" & n)

    Dim t0 As Double, msCompiled As Double, msInterp As Double

    BenchDropScratch
    VlaCompileToModule src, "VLA_Bench", ThisWorkbook
    t0 = Timer
    Application.Run "'" & ThisWorkbook.Name & "'!VLA_Bench.main"
    msCompiled = DevMs(t0)
    BenchDropScratch

    t0 = Timer
    VLA_Interpreter.VlaInterpret src, ThisWorkbook
    msInterp = DevMs(t0)

    DevNameSet "VLAt_BenchLoopC_" & n, msCompiled
    DevNameSet "VLAt_BenchLoopI_" & n, msInterp

    ' IIf is not short-circuiting in VBA - both branches evaluate
    ' regardless of the condition, so a bare IIf(msCompiled > 0,
    ' msInterp / msCompiled, ...) still divides by zero when the
    ' compiled leg reads under Timer resolution, exactly the case
    ' this guard exists for. A real If/Else avoids it.
    Dim factorText As String
    If msCompiled > 0 Then
        factorText = Format$(msInterp / msCompiled, "0.0") & "x"
    Else
        factorText = "n/a (compiled leg under Timer resolution)"
    End If
    Debug.Print "  N=" & Format$(n, "#,##0") & _
                "   compiled: " & Format$(msCompiled, "0") & " ms" & DevPrevTag(pC) & _
                "   interpreted: " & Format$(msInterp, "0") & " ms" & DevPrevTag(pI) & _
                "   factor: " & factorText
End Sub

' =====================================================================
'  IN.8 follow-up: VlaBenchmarkRowLoop - the member-dispatch half of
'  the loop-shaped workload, named but not built when IN.8 first
'  closed. VlaBenchmarkLoop already measured pure arithmetic dispatch
'  (dim/set!/for/+/debug-print, no object model); this measures the
'  OTHER real cost PF.4 names for row loops - a late-bound property
'  SET through Excel's own object model each iteration
'  (`(set! (. (cells i 1) value) i)`, exactly the shape
'  english.vla's own real macros use, per this interpreter's own
'  header note on WHY the six place helpers and CallByName dispatch
'  exist at all). The two costs are genuinely different mechanisms -
'  ExecStmt/EvalExpr's own per-call tax vs. WalkMemberSet's late-bound
'  Cells(...).Value write - and nothing about the pure-arithmetic
'  number said anything about this one.
'
'  Runs entirely in ThisWorkbook (the dev workbook), on a dedicated,
'  disposable sheet ("VlaBenchRows") created fresh and deleted before
'  AND after each leg - IN3.2's own "never coast on a prior run's
'  leftover cells" discipline (VerifyReportInterpreter's precedent),
'  so the compiled leg's N cells can never be mistaken for the
'  interpreted leg's own N cells, and nothing is left behind in the
'  dev workbook when this returns. Never touches ActiveWorkbook, so -
'  unlike VlaTimeIt/VlaTimeItInterpreter - this needs no "aim" step.
'
'  Deliberately NOT screen-updating-bracketed (PF.2's own future
'  optimization): both legs pay Excel's default per-write screen cost
'  equally, so the RATIO stays a fair backend comparison even though
'  neither leg's raw number reflects what optimized emitted code
'  would eventually cost.
'
'  Defaults are deliberately small (50/200/500) and conservative,
'  UNTIMED before the first real run of this exact shape existed -
'  late-bound member dispatch is a plausibly HEAVIER per-iteration
'  cost than the pure-arithmetic loop's own measured ~1-2 ms, not a
'  lighter one, so the same 100-second-freeze mistake VlaBenchmarkLoop
'  already made once is guarded against here from the start rather
'  than learned the same way twice. The estimated-duration warning
'  reuses that same measured 2 ms/iteration floor for anything larger.
' =====================================================================

Public Sub VlaBenchmarkRowLoop(Optional ByVal sizes As Variant)
    If IsMissing(sizes) Then sizes = Array(50, 200, 500)
    Debug.Print "=== VlaBenchmarkRowLoop (late-bound Cells(...).Value writes, one object-model call per iteration)"
    Dim totalN As Double
    Dim v As Variant
    For Each v In sizes
        totalN = totalN + CDbl(v)
    Next
    If totalN * 2 > 10000 Then
        Debug.Print "  heads up: ~" & Format$(totalN, "#,##0") & " total interpreted iterations, EACH a late-bound " & _
                    "Cells(...).Value write - likely heavier than VlaBenchmarkLoop's own ~1-2 ms/iteration pure- " & _
                    "arithmetic cost, not lighter. Estimated " & Format$(totalN * 2 / 1000, "0.0") & _
                    "s or more, Excel unresponsive throughout. Ctrl+Break to abort."
    End If
    For Each v In sizes
        BenchOneRowLoopSize CLng(v)
    Next
End Sub

Private Sub BenchOneRowLoopSize(ByVal n As Long)
    Dim src As String
    src = "(sub main ()" & vbCrLf & _
          "  (dim i Double)" & vbCrLf & _
          "  (for (i 1 " & n & ") (set! (. (cells i 1) value) i))" & vbCrLf & _
          "  (debug-print (. (cells " & n & " 1) value)))" & vbCrLf

    Dim pC As String, pI As String
    pC = DevNameGet("VLAt_BenchRowC_" & n)
    pI = DevNameGet("VLAt_BenchRowI_" & n)

    Dim t0 As Double, msCompiled As Double, msInterp As Double

    BenchDropScratch
    VlaCompileToModule src, "VLA_Bench", ThisWorkbook
    BenchFreshRowSheet.Activate
    t0 = Timer
    Application.Run "'" & ThisWorkbook.Name & "'!VLA_Bench.main"
    msCompiled = DevMs(t0)
    BenchDropScratch

    ' Fresh sheet again before the interpreted leg - IN3.2's own
    ' discipline: the compiled leg's own N cells must never be
    ' mistaken for cells this leg wrote itself.
    BenchFreshRowSheet.Activate
    t0 = Timer
    VLA_Interpreter.VlaInterpret src, ThisWorkbook
    msInterp = DevMs(t0)

    BenchDropRowSheet

    DevNameSet "VLAt_BenchRowC_" & n, msCompiled
    DevNameSet "VLAt_BenchRowI_" & n, msInterp

    Dim factorText As String
    If msCompiled > 0 Then
        factorText = Format$(msInterp / msCompiled, "0.0") & "x"
    Else
        factorText = "n/a (compiled leg under Timer resolution)"
    End If
    Debug.Print "  N=" & Format$(n, "#,##0") & _
                "   compiled: " & Format$(msCompiled, "0") & " ms" & DevPrevTag(pC) & _
                "   interpreted: " & Format$(msInterp, "0") & " ms" & DevPrevTag(pI) & _
                "   factor: " & factorText
End Sub

' Delete any existing "VlaBenchRows" scratch sheet, then create a
' fresh one - the same delete-then-recreate shape
' VerifyReportInterpreter already uses for "Output" (IN3.2's
' precedent), applied to a benchmark's own disposable sheet instead.
Private Function BenchFreshRowSheet() As Worksheet
    BenchDropRowSheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "VlaBenchRows"
    Set BenchFreshRowSheet = ws
End Function

Private Sub BenchDropRowSheet()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = Nothing
    Set ws = ThisWorkbook.Worksheets("VlaBenchRows")
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If
    On Error GoTo 0
End Sub

' Rule-12 duplicate of TryDropScratch (above), same rename-aside-then-
' remove shape, for "VLA_Bench" instead of "VLA_Scratch" - so a
' benchmark run mid-VlaTry-session never collides with the scratch
' module that session owns.
Private Sub BenchDropScratch()
    Dim comp As Object
    On Error Resume Next
    Set comp = Nothing
    Set comp = ThisWorkbook.VBProject.VBComponents("VLA_Bench_old")
    If Not comp Is Nothing Then ThisWorkbook.VBProject.VBComponents.Remove comp
    Set comp = Nothing
    Set comp = ThisWorkbook.VBProject.VBComponents("VLA_Bench")
    If Not comp Is Nothing Then
        comp.Name = "VLA_Bench_old"
        ThisWorkbook.VBProject.VBComponents.Remove comp
    End If
    On Error GoTo 0
End Sub

Private Function DevHasSheet(wb As Object, ByVal nm As String) As Boolean
    Dim ws As Object
    On Error Resume Next
    Set ws = wb.Worksheets(nm)
    DevHasSheet = Not ws Is Nothing
    On Error GoTo 0
End Function

Private Function DevMs(ByVal t0 As Double) As Double
    Dim d As Double
    d = Timer - t0
    If d < 0 Then d = d + 86400#   ' midnight wrap
    DevMs = d * 1000#
End Function

Private Function DevPrevTag(ByVal prev As String) As String
    If Len(prev) > 0 Then
        DevPrevTag = "   (prev " & prev & " ms)"
    Else
        DevPrevTag = "   (prev n/a)"
    End If
End Function

Private Function DevNameGet(ByVal nm As String) As String
    On Error Resume Next
    Dim v As String
    v = ThisWorkbook.Names(nm).RefersTo   ' stored as "=412"
    On Error GoTo 0
    If Len(v) > 1 Then DevNameGet = Mid$(v, 2)
End Function

Private Sub DevNameSet(ByVal nm As String, ByVal ms As Double)
    On Error Resume Next
    ThisWorkbook.Names(nm).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add nm, "=" & Format$(ms, "0")
End Sub

' Rule-12 duplicates of VLA_Tests' private resolver and reader (the
' S1.1 probing pattern; a Private in another module is invisible).
Private Function DevFindFile(ByVal fileName As String) As String
    Dim sep As String
    sep = Application.PathSeparator
    Dim c1 As String, c2 As String, c3 As String
    c1 = ThisWorkbook.Path & sep & fileName
    c2 = ThisWorkbook.Path & sep & "scripts" & sep & fileName
    c3 = ThisWorkbook.Path & sep & "scripts" & sep & "polyglotta" & sep & fileName
    If Len(Dir$(c1)) > 0 Then
        DevFindFile = c1
    ElseIf Len(Dir$(c2)) > 0 Then
        DevFindFile = c2
    ElseIf Len(Dir$(c3)) > 0 Then
        DevFindFile = c3
    Else
        Err.Raise 53, "VLA_DevRig", fileName & " not found - I looked beside the workbook (" & _
                  c1 & "), in its scripts folder (" & c2 & "), and in scripts/polyglotta (" & c3 & ")"
    End If
End Function

Private Function DevReadTextFileUtf8(ByVal filePath As String) As String
    On Error GoTo ansiFallback
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2                              ' adTypeText
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile filePath
    DevReadTextFileUtf8 = st.ReadText(-1)
    st.Close
    Exit Function
ansiFallback:
    On Error GoTo 0
    Dim f As Integer
    Dim b() As Byte
    f = FreeFile
    Open filePath For Binary Access Read As #f
    If LOF(f) > 0 Then
        ReDim b(1 To LOF(f))
        Get #f, , b
        DevReadTextFileUtf8 = StrConv(b, vbUnicode)
    End If
    Close #f
End Function

' =====================================================================
'  PPROF.0: VlaProfileAll - P-PROF's own dial (BETA_ROADMAP2.md,
'  MACHINE + OPTIMIZATION), run against VLA_Tests_Host.VlaSelfTestsAll -
'  the full pure+host suite WITH mRunScaleTests forced on internally, so
'  TestListopsBudget/TestListopsDepthSafety/TestTablespecDepthSafety's own
'  2000/5500-deep chain and row-walk cases actually run and actually
'  dominate the numbers below (several minutes at 5500 rows, already
'  confirmed live by wall clock alone - this dial's own job is turning
'  that single number into a phase breakdown, not re-discovering it).
'  The tranche's own gating rule (BETA_ROADMAP1.md: "All of it gated on
'  P-PROF's before-numbers") is what this exists to satisfy - a real
'  before-number for P-DICT/P-NTH (GetMacro's per-step lookup, cdr/
'  ListTail's O(n^2) tail copy, both VLA.bas), not the wall-clock total
'  already on record, but that cost's actual SHARE of transpile time.
'
'  Six buckets, all behind VLA.mProfileOn (one switch, VLA.bas's own
'  header comment has the full mechanism):
'    compile-side (VlaTranspile, per call):  tokenize / parse / expand
'      / emit - expand is P-DICT/P-NTH's own target.
'    translate-side (EnglishToVla, VLA_SentenceEngine.bas, per call):
'      trans-tokenize / trans-build - trans-tokenize is P-TOK's own
'      target (TokAt's positional Collection indexing).
'  Every bucket sums across EVERY call the suite makes while the switch
'  is on, thousands of trivial ones included - by design: the 2000/5500
'  cases are minutes-long against a suite that otherwise runs in
'  seconds, so they dominate the sums on their own, no per-call
'  bucketing needed to see them.
'
'  VlaTimeIt's own house style throughout: workbook Names (VLAt_Prof*),
'  CURRENT printed next to PREVIOUS, Immediate window, dev-rig only, no
'  self-test pin (the printed numbers ARE the verification). The switch
'  is restored to False on every exit path, including a mid-run failure
'  (On Error GoTo cleanup below) - it must never leak into an ordinary
'  Check/Compile/Run left running afterward with a stale profiling cost
'  (small but real: a couple of Timer reads and accumulator adds per
'  VlaTranspile/EnglishToVla call).
'
'  Honest notes, VlaTimeItSelfTest's own precedent: this runs the ENTIRE
'  suite for real, several minutes, not seconds - mRunScaleTests forced
'  on inside VlaSelfTestsAll itself - and the suite resets the grammar
'  along the way, so reload the vocabulary before the next Check/Run.
'  Dev-rig only, never in the add-in. Run on demand - whenever a change
'  touches ExpandMacros/cdr/ListTail/GetMacro/TokAt/tablespec's own walk,
'  or a new performance feature needs a real before-number rather than a
'  guess - not part of any routine suite run.
' =====================================================================

Public Sub VlaProfileAll()
    Dim pTok As String, pParse As String, pExp As String, pEmit As String
    Dim pTTok As String, pTBuild As String, pTotal As String
    pTok = DevNameGet("VLAt_ProfTokenize")
    pParse = DevNameGet("VLAt_ProfParse")
    pExp = DevNameGet("VLAt_ProfExpand")
    pEmit = DevNameGet("VLAt_ProfEmit")
    pTTok = DevNameGet("VLAt_ProfTransTokenize")
    pTBuild = DevNameGet("VLAt_ProfTransBuild")
    pTotal = DevNameGet("VLAt_ProfTotal")

    Debug.Print "=== VlaProfileAll (per-phase timing, VLA.mProfileOn, against the full scale-included suite - several minutes)"

    VLA.mProfMsTokenize = 0
    VLA.mProfMsParse = 0
    VLA.mProfMsExpand = 0
    VLA.mProfMsEmit = 0
    VLA.mProfCalls = 0
    VLA_SentenceEngine.mProfMsTransTokenize = 0
    VLA_SentenceEngine.mProfMsTransBuild = 0
    VLA_SentenceEngine.mProfTransCalls = 0
    VLA.mProfileOn = True

    On Error GoTo cleanup
    Dim t0 As Double, msTotal As Double, ok As Boolean
    t0 = Timer
    ok = VLA_Tests_Host.VlaSelfTestsAll()
    msTotal = DevMs(t0)
    VLA.mProfileOn = False

    DevNameSet "VLAt_ProfTokenize", VLA.mProfMsTokenize
    DevNameSet "VLAt_ProfParse", VLA.mProfMsParse
    DevNameSet "VLAt_ProfExpand", VLA.mProfMsExpand
    DevNameSet "VLAt_ProfEmit", VLA.mProfMsEmit
    DevNameSet "VLAt_ProfTransTokenize", VLA_SentenceEngine.mProfMsTransTokenize
    DevNameSet "VLAt_ProfTransBuild", VLA_SentenceEngine.mProfMsTransBuild
    DevNameSet "VLAt_ProfTotal", msTotal

    Dim compileTotal As Double
    compileTotal = VLA.mProfMsTokenize + VLA.mProfMsParse + VLA.mProfMsExpand + VLA.mProfMsEmit

    Debug.Print "  suite result: " & IIf(ok, "PASS", "FAIL") & "   wall time: " & Format$(msTotal, "0") & " ms" & DevPrevTag(pTotal)
    Debug.Print "  --- compile-side (VlaTranspile, " & VLA.mProfCalls & " calls) ---"
    Debug.Print "  tokenize:  " & Format$(VLA.mProfMsTokenize, "0") & " ms" & DevPrevTag(pTok) & ProfShare(VLA.mProfMsTokenize, compileTotal)
    Debug.Print "  parse:     " & Format$(VLA.mProfMsParse, "0") & " ms" & DevPrevTag(pParse) & ProfShare(VLA.mProfMsParse, compileTotal)
    Debug.Print "  expand:    " & Format$(VLA.mProfMsExpand, "0") & " ms" & DevPrevTag(pExp) & ProfShare(VLA.mProfMsExpand, compileTotal) & "   <- P-DICT/P-NTH's own target"
    Debug.Print "  emit:      " & Format$(VLA.mProfMsEmit, "0") & " ms" & DevPrevTag(pEmit) & ProfShare(VLA.mProfMsEmit, compileTotal)
    Debug.Print "  --- translate-side (EnglishToVla, " & VLA_SentenceEngine.mProfTransCalls & " calls) ---"
    Debug.Print "  trans-tokenize:  " & Format$(VLA_SentenceEngine.mProfMsTransTokenize, "0") & " ms" & DevPrevTag(pTTok) & "   <- P-TOK's own target (TokAt)"
    Debug.Print "  trans-build:     " & Format$(VLA_SentenceEngine.mProfMsTransBuild, "0") & " ms" & DevPrevTag(pTBuild)
    Debug.Print "  (stored in workbook names VLAt_Prof*; grammar was reset along the way - reload vocabulary before the next Check)"
    Exit Sub
cleanup:
    Dim n As Long, s As String, d As String
    n = Err.Number: s = Err.Source: d = Err.Description
    VLA.mProfileOn = False
    Err.Raise n, s, d
End Sub

' PPROF.0: part's share of total as a percentage tag, "" when total is
' not yet meaningful (nothing profiled) - avoids a divide-by-zero the
' way BenchOneLoopSize's own factorText guard already does for a
' different ratio.
Private Function ProfShare(ByVal part As Double, ByVal total As Double) As String
    If total <= 0 Then Exit Function
    ProfShare = "  (" & Format$(part / total * 100#, "0.0") & "% of compile)"
End Function

Public Sub VlaDiagnostics()
    On Error Resume Next
    Debug.Print "===== VLA DIAGNOSTICS ====="
    Debug.Print "  VLA_Identity:  " & VLA_IDENTITY_VERSION
    Debug.Print "  VLA_HeadTable: " & VLA_HEADTABLE_VERSION & " (" & VlaHeadTableSummary() & ")"
    Debug.Print "  VLA core:      " & VLA_CORE_VERSION
    Debug.Print "  VLA_Loader:    " & VLA_LOADER_VERSION
    Debug.Print "  VLA_English:   " & VLA_ENGLISH_VERSION
    Debug.Print "  VLA_Runtime:   " & VLA_RUNTIME_VERSION
    Debug.Print "  VLA_IDE:       " & VLA_IDE_VERSION
    Debug.Print "  VLA_Build:     " & VLA_BUILD_VERSION
    Debug.Print "  VLA_Tests:     " & VLA_TESTS_VERSION
    Debug.Print "  VLA_Tests_Grammar: " & VLA_TESTS_GRAMMAR_VERSION
    Debug.Print "  VLA_Tests_Host: " & VLA_TESTS_HOST_VERSION
    Debug.Print "  VLA_DevRig:    " & VLA_DEVRIG_VERSION
    Debug.Print "  VLA_Interpreter: " & VLA_INTERPRETER_VERSION & " (IN.6: ships in the add-in)"
    Debug.Print "  VLA_Events:    " & VLA_EVENTS_VERSION & " (IN.7: sheet-change registry, ships in the add-in)"
    Debug.Print "  rules loaded:  " & EnglishRuleCount()
    Debug.Print "  function words:" & FnWordCount()
    Debug.Print "  this workbook: " & ThisWorkbook.Name
    Debug.Print "  active workbook: " & ActiveWorkbook.Name
    Debug.Print "  " & VlaIdeInfo()
    Debug.Print "===== (paste everything above into a bug report) ====="
    On Error GoTo 0
End Sub

Private Function FnWordCount() As Long
    Dim t As String
    t = EnglishListFunctionWords()
    If Len(Trim$(t)) = 0 Then Exit Function
    FnWordCount = UBound(Split(Trim$(t), vbCrLf)) + 1
End Function
