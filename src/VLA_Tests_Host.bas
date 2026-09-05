Attribute VB_Name = "VLA_Tests_Host"
Option Explicit
Public Const VLA_TESTS_HOST_VERSION As String = "PF4C.0"
' PF4C.0: TestStmtParity gains two for-each-row cases (PF.4c) - a full
' round-trip (doubles a column in place, verified via a plain '.'-dot
' read that write-back actually reached the live sheet, not just that
' the in-memory array looked right) and a break-commits-partial case
' (breaks right after row 2 - row 3 must survive untouched in the
' write-back, proving "up to and including the break," never beyond,
' live rather than assumed from the shape alone).
' PF4B.0: TestArraySlabHelpers - new coverage for VLA_Runtime.
' VlaSlabRead/VlaSlabWrite (PF.4b), wired into VlaSelfTestHost right
' after TestEmbeddedTextFallback. Multi-row/col round-trip WITH a real
' mutate-then-verify (an unchanged echo could pass even if write-back
' silently did nothing), plus 1x1/1-row-N-col/N-row-1-col - the same
' shape-coverage discipline PF.7's own edge-case tests already
' established for the identical scalar-collapse quirk.
' PF4A.0: TestStmtParity gains one case proving set!/read parity into an
' array element ((arr 1) written then read back) - PF.4a, the array-
' element set!/read fallback tiers new to VLA_Interpreter.ExecSet/
' EvalDynamicHead this pass. Compiled needed no change (its own generic
' text substitution already emits real VBA array-index syntax); the
' point of this case is proving the interpreter's brand-new tiers agree
' with it, not merely that the interpreter alone stopped raising.
' PF7.0: TestEmbeddedTextFallback gains three states alongside VLA.bas's
' VlaEmbeddedText bulk-read/Join rewrite (PF.7) - n=1 (Range.Value
' collapses to a scalar on a single-cell range, IdeVocabFileName's own
' VLAe_Name shape), n=1-with-a-never-written cell (End(xlUp) returns row
' 1, never row 0, on a wholly blank column - the real "no content"
' shape), and an mRunScaleTests-gated round-trip at english.vla's own
' real order of magnitude (2,849 rows), printing elapsed ms informationally
' the same way TestVocabMacroProbe's own scale case does - the only
' repeatable, dev-side number for a loop body PF.7's own scoping pass
' found otherwise unreachable from this workbook (PreludeMacros/
' IdeVocabPath both find an external file first here, so the embedded-
' sheet path this test now exercises directly never fires through normal
' use in dev).
' AS7.2: VlaSelfTests and VerifyReports both gain a comprehensive final
' summary (counts + full failure text for BOTH halves) at the very end,
' owner request while scoping AS.7 - previously each printed only a
' binary "X PASS, Y PASS" verdict line, so a failure on the FIRST half
' was already scrolled off the Immediate pane's finite buffer by the
' time you reached the combined line and had to be found by scrolling
' back up. VlaSelfTests' pure half needed a new cross-module accessor
' (VLA_Tests.VlaSelfTestLastResult, AS7.1) since its mPass/mFail/
' mFailedNames are Private to that module; its host half, and both
' VerifyReports halves, already live in this module so are read
' directly. Two correctness fixes made along the way, both required for
' the new summary to be accurate rather than silently misleading:
'   - VerifyReport's "no Output sheet" early exit used to bypass
'     Report()/mFail entirely (a bare Debug.Print), so it returned False
'     with mFail = 0 - a comprehensive summary reading mFail alone would
'     have shown "0 failed" under a FAIL verdict. Now flows through
'     Report() like VerifyReportInterpreter's own three early exits
'     already did, so the reason is always in mFailedNames too.
'   - VerifyReport and VerifyReportInterpreter both now reset
'     mFailedNames at their own top (matching the mPass/mFail reset
'     already there, and VlaSelfTestHost's own precedent) - previously
'     only VlaSelfTestHost ever reset the Collection the three of them
'     share, so a VerifyReport call made after a VlaSelfTestHost run in
'     the same session could otherwise inherit stale host failure names.
' GPIVOT.4: VerifyReportChecks gains real state assertions for the
' CRUD-completeness round: pivot-sort/pivot-sort-by-value (SalesPivot/
' Region and FullPivot/Product each sorted descending then ascending -
' the FINAL, checkable state is always ascending, a genuine transition
' from the descending sort immediately before it), pivot-rename
' (RenameMePivot/RenamedPivot - old name gone, new name resolves),
' pivot-clear (ClearMePivot survives but its one row field and one
' value field are both gone - a real, distinct state from pivot-delete),
' pivot-remove-field (RemoveFieldMePivot - Product hidden, Region
' untouched, proving selective removal not a wholesale reset), and
' pivot-source (SourceTestPivot - built from N1:S7 missing a new eighth
' "West" row on purpose, then repointed to N1:S8; West actually showing
' up as a Region PivotItem is the real proof ChangePivotCache-plus-
' refresh pulled in genuinely new data). Each on its own dedicated
' pivot, deliberately, so none of these destructive/identity-changing
' proofs could collide with SalesPivot's or FullPivot's own already-
' established checks.
' GPIVOT.3: VerifyReportChecks gains real state assertions for
' pivot-subtotals-hide/-show/pivot-blank-line-add/-remove, and one for
' pivot-layout - the report-layout knobs closing out G-PIVOT.
' FIRST VERSION also checked PivotField.LayoutForm on Region/Product,
' expecting it to mirror RowAxisLayout's own compact/tabular/outline
' state - wrong, caught live (both checks read back the same value
' regardless of which layout had actually been applied): LayoutForm is
' a different, older per-field enum (XlLayoutFormType - Microsoft's own
' docs give only xlOutline/xlTabular, no compact option) entirely
' unrelated to RowAxisLayout's own XlLayoutRowType parameter, and per
' Microsoft's own reference there is no object-model readback for
' XlLayoutRowType at all. Both checks removed - pivot-layout's own live
' proof is necessarily thinner than every other pivot verb's, the same
' honestly-named limit VlaPivotRefresh's own comment already accepts.
' What stays checkable and real: SalesPivot's own Region field (already
' asserted collapsed above) gets a fresh post-layout-switch ShowDetail
' re-check, not an assumption that RowAxisLayout (table-wide, re-forms
' every row field at once) leaves an already-collapsed field's per-
' PivotItem state alone - the same class of Compact-Form interaction
' that broke VlaPivotSetShowDetail's first version.
' Subtotals/LayoutBlankLine each get one real distinguishing-change
' proof (SalesPivot/Region) plus one hide-then-show or add-then-remove
' round trip (FullPivot/Product), matching the collapse/expand
' precedent's own honestly-named limit: end-state alone can't
' distinguish "reversed" from "never touched."
' GPIVOT.2: VerifyReportChecks gains real state assertions for
' pivot-refresh/-refresh-all/-delete/-collapse/-expand - SalesPivot's
' own Region field ends collapsed, FullPivot's own Product field ends
' expanded after a collapse-then-expand round trip (a real limit named
' rather than overclaimed: end-state alone can't distinguish "reversed"
' from "untouched"), and TempPivot - a throwaway pivot, TempTable's
' own precedent - is confirmed actually removed, not just cleared.
' Checked per PIVOT ITEM, not PivotField.ShowDetail directly - the
' first version's field-level read crashed live (run-time error 1004)
' before this version was ever committed, caught the same live run
' that found VlaPivotSetShowDetail's own first-version bug
' (VLA_Runtime.bas has the full reasoning: Compact Form, the default
' row layout, makes field-level ShowDetail unreliable).
' IN.3: VerifyReportInterpreter - the interpreter-side counterpart to
' VerifyReport this item's own opening line always meant ("instructions.txt
' produces identical Output sheets under both backends"). VerifyReport's
' own 30-odd CheckV/Report cell checks are factored out, unchanged, into
' VerifyReportChecks(ws) so both backends are checked against the same
' hand-computed expected values (IN3.2's own two-sided precedent, widened
' from four hand-picked statement forms to the whole corpus). Unlike
' VerifyReport, needs no prior manual Run - it compiles instructions.txt
' itself (step tracking off, VLA_Interpreter.bas's own IN.3 note has the
' full reasoning) and calls VlaInterpret for real against a freshly
' deleted-and-recreated Output sheet, so no check can coast on a prior
' run's leftover cell. AS.6-guarded in three stages (compile/run/check).
' Deliberately NOT dispatched from VlaSelfTestHost - see this function's
' own header note for why (mirrors VerifyReport's own reasoning for
' staying manual). Found, and names rather than builds, real structured-
' exception-handling remainder - see docs/BETA_ROADMAP.md's IN.3 entry and
' VLA_Interpreter.bas's own IN.3 note for the full reasoning. Also adds
' VerifyReports, VlaSelfTests's own convenience shape - VerifyReport then
' VerifyReportInterpreter, one call, order load-bearing (see its own
' header note).
' IN.3 (three live reloads later, same session): the owner asked
' directly whether DynamicSet's CallByName Let/Set heuristic
' (VLA_Interpreter.bas) could still be trusted after three straight real
' members broke it (value/size/color) - answered by counting the WHOLE
' corpus's own real usage rather than patching the next crash, 15
' members total, all converted to native dispatch in one pass
' (DynamicSet's own header note has the full reasoning). Two more live
' pins here (font.bold, columns.hidden - different parent types than
' size/color already covered) plus one for a separate, silent bug the
' same audit surfaced (ExecSet treating a bare dotted-global place atom,
' application.screenupdating among them, as a plain variable - "Turn off
' screen updating." never touched the real Application object, all
' session, no error ever raised for it).
' IN2.5: TestInterpreterNamedArgs - four representative live pins for
' IN.2's named-argument dispatch (VLA_Interpreter.bas's own IN2.5
' header note has the full census and reasoning): protect/unprotect
' (Worksheet), copy :destination (Range, a Range-valued keyword
' argument), removeduplicates :columns :header (Range, two keyword
' arguments including an Excel enum constant). close/printout/
' exportasfixedformat are deliberately not exercised, for the same
' reason msgbox/inputbox never are. Also fixed a labeling mistake this
' pass's own review caught: TestInterpreterControlFlow's (VLA_Tests.
' bas) and TestInterpreterForEach's Report/CheckV name strings said
' "in2.1:" - the real IN2.1 tag (a different, prior host-run bugfix) -
' when they meant IN2.3, cosmetic only (display labels, not behavior)
' but worth fixing rather than compounding a homograph SD-9 already
' names as this project's own founding incident.
' IN2.4: TestStmtParity widened - for-each/do-until/select added once
' IN2.3 gave the interpreter its own implementation of each to check
' against the emitter's independently-implemented one (native VBA
' control flow vs. hand-rolled dispatch + the mLoopBreak flag) - see
' this module's own IN2.4 note on TestStmtParity for the full
' reasoning. Each of the three programs also exercises exit-for/
' exit-do, IN2.3's own required companion, so a loop that cannot be
' broken early is not left half-proven. No harness change needed:
' CheckStmtParity/DropParityScratch/the probe-cell mechanism are
' unchanged, exactly the "the mechanism was already the right shape"
' case F.1's ABI is meant to produce.
' A real bug, caught by the owner's own live host run, not by this
' pass's own static self-review: the for-each program's loop variable,
' 'c', was never declared - the exact same undeclared-loop-variable
' mistake this item's own IN.3 predecessor already found and fixed
' once for 'for's counter, recurring here because that earlier fix was
' not generalized into a habit of checking EVERY loop variable a new
' raw-VLA test program introduces, only the one that had just failed.
' Fixed by adding '(dim c)' alongside the other declarations. The
' interpreter side was never at risk - VlaInterpret's frame is a
' dynamically-typed dict, so dim is bookkeeping there, not a
' requirement - only the emitter's live VBA compile, gated by Option
' Explicit, could ever have caught this, which is exactly why this
' harness runs the program for real instead of only interpreting it.
' IN2.3: TestInterpreterForEach - for-each over a real (range "A1:A3"),
' live, summing each cell's value - IN.2's remainder (VLA_Interpreter.
' bas's own IN2.3 header note has the full reasoning). Lives here, not
' in the pure suite, because the corpus's own for-each usage IS a range
' ("Remember range B2:B4 as results. For each r in results...") and
' this interpreter has no (new Collection) support to build one purely
' - a separate, pre-existing gap this item does not touch. Dispatched
' from VlaSelfTestHost right after TestInterpreterObjectDispatch.
' IN3.2: TestStmtParity - IN.3's statement-level parity widening, now
' that IN.2 gives the interpreter a real place to write. Five programs
' (if x2, for, while, set!) run live under both backends against a
' shared probe cell, read back natively (never through either backend's
' own read path) and checked two ways: each backend against a
' hand-computed expected value, and the two backends against each
' other - IN.3's own text is explicit that parity alone proves
' consistency, not truth. The emitter side reuses VLA_DevRig's
' VlaTryBuild (the statement sibling of TestExprParity's
' VlaTryValueBuild) under its own scratch module,
' VLA_Scratch_StmtParity; DropParityScratch is now parameterized on
' the module name so both parity tests share it instead of a second
' near-duplicate. Dispatched from VlaSelfTestHost right after
' TestExprParity.
' IN2.0: TestInterpreterObjectDispatch - proves IN.2's new object-
' place/dot-dispatch/computed-set! mechanism (VLA_Interpreter.bas)
' against a REAL worksheet: (range addr) as a place helper, '.' reads
' and multi-hop chains (font.bold), '.' in statement position as a
' real action (ClearContents), set! to both a bare computed place and
' an explicit '.' place, a dotted-global chain (activesheet.name)
' settling IN.0.5's own "open question" the measurement raised but
' did not settle, and the effect log widening to record the new,
' genuinely-external effects. This is exactly what IN0.5's own "out of
' scope on purpose" note reserved for a live Excel object model - why
' it lives here, not in the pure suite. Dispatched from
' VlaSelfTestHost, after TestExprParity (and, since IN3.2, TestStmtParity).
' IN3.1: VlaSelfTests - a convenience wrapper (owner request: "I'm
' lazy") that runs VLA_Tests.VlaSelfTest then VlaSelfTestHost in one
' call and prints a combined verdict. Lives here, not in VLA_Tests.bas
' - F.11 split this module OUT specifically so the pure half never
' depends on the host half, so a combined runner has to sit on the
' host side of that line. Deliberately does not also run VerifyReport
' - that needs a prior real Run (the Output sheet populated), which
' this function cannot assume happened.
' IN3.0: TestExprParity - IN.3's dual-mode parity harness, scoped
' honestly (see VLA_Interpreter.bas's IN3_5.0 note for the full
' reasoning): the interpreter cannot write worksheet ranges yet
' (IN.0.5 reserved that for IN.2's CallByName dispatch), so "identical
' Output sheets" cannot be checked today. What CAN be checked live is
' a bare arithmetic expression - VLA_Interpreter.VlaEvalExpression's
' in-memory answer against the SAME expression compiled by the real
' emitter, injected into a scratch module, and run for real, via
' VLA_DevRig's already-proven L6 scratch mechanism (VlaTryValueBuild +
' VlaCompileToModule + Application.Run) rather than new plumbing.
' DropParityScratch is an R7 duplicate of VLA_DevRig's own
' TryDropScratch, scoped to its own module name (VLA_Scratch_Parity)
' so this test can never collide with a developer's own live
' VlaTryValue session in the Immediate window. Dispatched from
' VlaSelfTestHost, right after TestRuntimeModule.
' F11.0: new module (F.11 - "a human at a Windows box is the velocity
' ceiling and the bus factor", the roadmap's own words for why this
' item was promoted to first). Everything here needs a LIVE workbook
' with real sheet/range/VBProject/Names state, as opposed to
' VLA_Tests.bas and VLA_Tests_Grammar.bas, which now hold only pure
' compiler-pipeline pins (tokenize/parse/expand/emit/interpret,
' asserted on returned strings and Collections, no Excel object model
' touched). Moved here, verbatim except where noted: VerifyReport (the
' whole function - it was already its own host-only entry point, just
' not yet its own file); TestRuntimeModule + its DevModuleText helper
' (whole subs - VBProject/CodeModule reads need "Trust access to the
' VBA project object model"); and two extracted fragments that used to
' be the host-dependent TAIL of an otherwise-pure test - TestHelpers'
' scratch-sheet section (VlaFindRow/VlaEnsureSheet/range-backed
' VlaCount/VlaItem/VlaFirst/VlaLast) is now TestHelpersHost, and
' TestRuntimeTrace's workbook-Name-persistence section is now
' TestRuntimeTraceHost; the pure halves of both stayed under their
' original names in VLA_Tests.bas, unchanged, still in VlaSelfTest's
' dispatch list. New dispatcher here, VlaSelfTestHost, mirrors
' VlaSelfTest's shape but is a genuinely separate run (own counters,
' own summary) rather than a size-only split like F.8's - the two
' halves are meant to run independently: VlaSelfTest wherever VBA
' executes at all, VlaSelfTestHost/VerifyReport only at a Windows box
' with this workbook live and (for TestRuntimeModule) VBProject trust
' granted. Report/CheckV are duplicated here (R7: self-containment
' across a real architectural boundary, not just a size split) rather
' than promoted Public on VLA_Tests.bas the way F.8's shared assertion
' helpers were - sharing counters across a pure run and a host run
' would conflate two reports that are meant to answer different
' questions on different machines.

' =====================================================================
'  VLA_Tests_Host - every test that needs a live Excel: real sheet/
'  range state (TestHelpersHost), VBProject/CodeModule access
'  (TestRuntimeModule), a live emitter-vs-interpreter comparison
'  (TestExprParity, IN.3), workbook Name persistence
'  (TestRuntimeTraceHost), and the post-Run output check
'  (VerifyReport). Dev-only, like VLA_Tests.bas and
'  VLA_Tests_Grammar.bas: NOT shipped in the add-in build.
'
'  VlaSelfTestHost  runs the five dispatched host pins and prints its
'                 own PASS/FAIL summary, same shape as VlaSelfTest.
'                 Requires: this workbook open and live; VBProject
'                 trust granted (TestRuntimeModule and TestExprParity
'                 both fail loudly, not silently, if that trust is
'                 absent - neither needs it to report, only to pass).
'
'  VlaSelfTests   convenience: VlaSelfTest (VLA_Tests.bas) then
'                 VlaSelfTestHost, one call. Does not also run
'                 VerifyReport - that needs a prior real Run.
'
'  VerifyReport   unchanged in behavior and calling convention from
'                 before this split - machine-checks instructions.txt's
'                 end state on the Output sheet. Run it right after
'                 clicking Run, same as always.
'
'  VerifyReportInterpreter  IN.3's interpreter-side counterpart -
'                 VerifyReport's own checks (VerifyReportChecks, shared)
'                 run against a live interpreter execution of
'                 instructions.txt itself. Needs no prior Run; compiles and
'                 executes it directly. Also manual, not dispatched from
'                 VlaSelfTestHost - see its own header note.
'
'  VerifyReports  convenience: VerifyReport then VerifyReportInterpreter,
'                 one call, VlaSelfTests's own shape. Order matters -
'                 VerifyReport must run first, before VerifyReportInterpreter's
'                 own run deletes and rebuilds the Output sheet it reads.
'                 Still needs a prior real emitter Run for its own half.
'
'  LAYER:     Harness (dev-only; never ships - see REBUILD.md SS3
'             Layer 5)
'  MAY CALL:  VLA_Runtime, VLA_English (DevModuleText's text probes),
'             VLA (VlaCompileToModule, for TestExprParity/TestStmtParity's
'             live scratch injection), VLA_Interpreter (VlaEvalExpression
'             and VlaInterpret, IN.3's in-memory half), VLA_Events (IN.7,
'             the sheet-change registry), VLA_DevRig
'             (VlaTryValueBuild and, since IN3.2, VlaTryBuild - IN.3's
'             emitter-side scratch wrappers, L6's already-proven
'             mechanism, reused rather than duplicated), VLA_Tests
'             (VlaSelfTest, called once by VlaSelfTests - the one
'             intentional exception to this module's "host only" scope,
'             since a pure call adds no requirement a host module
'             doesn't already have)
'  SHIPS:     nowhere. Dev-rig only, like VLA_DevRig.
'  PAYS INTO: F.11 (this split - the roadmap's own words: "pays into
'             every section, forever," since every future pass can
'             now iterate against VlaSelfTest alone without a human
'             at a Windows box, and only needs this module's slower
'             loop when a change actually touches host state)
'  REASON:    see the F11.0 history note above.
' =====================================================================

Private mPass As Long
Private mFail As Long
Private mFailedNames As Collection

Public Function VlaSelfTestHost() As Boolean
    mPass = 0
    mFail = 0
    Set mFailedNames = New Collection
    Debug.Print "===== VLA SELF-TEST (HOST-REQUIRED) ====="

    TestHelpersHost
    TestRuntimeModule
    TestExprParity
    TestStmtParity
    TestInterpreterObjectDispatch
    TestInterpreterHostWorkbook
    TestSetFormulaSpillsWithoutImplicitIntersection
    TestInterpreterSheetChangeEvent
    TestInterpreterButtonClickEvent
    TestCompiledButtonClickParity
    TestExportSkeleton
    TestInterpreterForEach
    TestInterpreterNamedArgs
    TestRuntimeTraceHost
    TestEmbeddedTextFallback
    TestArraySlabHelpers
    TestAutoLoadRegistration

    Debug.Print "===== HOST SELF-TEST: " & mPass & " passed, " & mFail & " failed ====="
    If mFail > 0 Then
        Debug.Print "===== FAILED TESTS (" & mFailedNames.Count & ") ====="
        Dim fn As Variant
        For Each fn In mFailedNames
            Debug.Print "  FAIL  " & fn
        Next fn
        Debug.Print "===== (paste the block above into a bug report) ====="
    End If
    VlaSelfTestHost = (mFail = 0)
End Function

' Convenience: both runs in one call, for a live Windows box with the
' workbook open (the only place VlaSelfTestHost can run anyway, so
' calling VlaSelfTest first from here adds no new requirement). Lives
' here rather than in VLA_Tests.bas on purpose - F.11 split this module
' OUT of VLA_Tests.bas specifically so the pure half never depends on
' the host half; a combined runner has to sit on the host side of that
' line, not the pure side. Prints both summaries in sequence (each
' function already prints its own) plus one combined verdict line.
' Does NOT run VerifyReport - that needs a prior, real Run (the Output
' sheet populated), which is a manual step this function cannot assume
' happened, so bundling it would just fail with "Run the program
' first" on a workbook nobody has run yet.
Public Function VlaSelfTests() As Boolean
    Dim pureOk As Boolean, hostOk As Boolean
    Dim purePass As Long, pureFail As Long, pureFailed As Collection
    Dim hostPass As Long, hostFail As Long, hostFailed As Collection

    pureOk = VLA_Tests.VlaSelfTest()
    VLA_Tests.VlaSelfTestLastResult purePass, pureFail, pureFailed

    hostOk = VlaSelfTestHost()
    hostPass = mPass : hostFail = mFail : Set hostFailed = mFailedNames

    Debug.Print "===== VLA SELF-TESTS: " & IIf(pureOk, "pure PASS", "pure FAIL") & _
                " (" & purePass & "/" & (purePass + pureFail) & ")" & _
                ", " & IIf(hostOk, "host PASS", "host FAIL") & _
                " (" & hostPass & "/" & (hostPass + hostFail) & ") ====="
    Dim fn As Variant
    If pureFail > 0 Then
        Debug.Print "----- pure failures (" & pureFailed.Count & ") -----"
        For Each fn In pureFailed
            Debug.Print "  FAIL  " & fn
        Next fn
    End If
    If hostFail > 0 Then
        Debug.Print "----- host failures (" & hostFailed.Count & ") -----"
        For Each fn In hostFailed
            Debug.Print "  FAIL  " & fn
        Next fn
    End If
    Debug.Print "(VerifyReport is separate - Run the program first, then VerifyReport)"
    VlaSelfTests = pureOk And hostOk
End Function

' Convenience: literally everything that's excluded from the two runs
' above purely for being SLOW, not for needing a prior manual step or
' crossing a deliberate architecture boundary - so nobody has to
' remember a test's name to run it. VlaSelfTestScale's own shape
' (mRunScaleTests, VLA_Tests.bas), just wrapping the combined pure+host
' runner instead of the pure half alone, plus TestF4RealCorpusShadow
' (VLA_Tests_Grammar.bas) - the one real-corpus F.4 audit deliberately
' kept out of the routine suite for its own O(N^2) cost (its own
' header, and BETA_ROADMAP.md's F.4 entry, have the full reasoning).
' Several minutes slower than VlaSelfTests; use before a release, or
' after touching AuditCrossRuleShadow/ExpandMacros/cdr/ListTail/
' tablespec's own walk.
' NOT included, on purpose, for DIFFERENT reasons than "too slow, easy
' to forget": VerifyReport/VerifyReportInterpreter need a prior manual
' Run first (an Output sheet this function has no way to populate for
' you - see VlaSelfTests's own header); VlaLintCheck (VLA_Tests.bas)
' does real disk I/O by design, a boundary VlaSelfTest/VlaSelfTests
' deliberately hold that this convenience function has no business
' quietly breaking. Those three stay exactly as manual as they already
' were - this only sweeps up the ones whose sole reason for exclusion
' was cost.
Public Function VlaSelfTestsAll() As Boolean
    VLA_Tests.mRunScaleTests = True
    On Error GoTo cleanup

    Dim baseOk As Boolean
    baseOk = VlaSelfTests()   ' prints its own full pure+host summary above

    Dim beforePass As Long, beforeFail As Long, failedNames As Collection
    VLA_Tests.VlaSelfTestLastResult beforePass, beforeFail, failedNames
    Dim beforeFailedCount As Long
    beforeFailedCount = failedNames.Count

    VLA_Tests_Grammar.TestF4RealCorpusShadow

    VLA_Tests.mRunScaleTests = False

    Dim afterPass As Long, afterFail As Long
    VLA_Tests.VlaSelfTestLastResult afterPass, afterFail, failedNames   ' same live Collection, possibly longer now

    Dim extraFail As Long
    extraFail = afterFail - beforeFail
    Debug.Print "----- plus F.4's real-corpus audit (scale cases already folded into the pure count above): " & _
                IIf(extraFail = 0, "PASS", "FAIL (" & extraFail & ")") & " -----"
    If extraFail > 0 Then
        Dim i As Long
        For i = beforeFailedCount + 1 To failedNames.Count
            Debug.Print "  FAIL  " & failedNames.Item(i)
        Next i
    End If

    VlaSelfTestsAll = baseOk And (extraFail = 0)
    Exit Function
cleanup:
    Dim n As Long, s As String, d As String
    n = Err.Number: s = Err.Source: d = Err.Description
    VLA_Tests.mRunScaleTests = False
    Err.Raise n, s, d
End Function

' F11.0: the host-dependent tail of VLA_Tests.bas's TestHelpers - a
' scratch sheet's VlaFindRow/VlaEnsureSheet, and range-backed
' VlaCount/VlaItem/VlaFirst/VlaLast (the pure list-backed half of each
' stayed in TestHelpers). Needs a live workbook to create, populate,
' and delete a real worksheet.
Private Sub TestHelpersHost()
    ' VlaFindRow / VlaEnsureSheet on a scratch sheet.
    Dim prior As Worksheet
    Set prior = ActiveSheet
    VlaEnsureSheet "VlaSelfTestSheet"
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets("VlaSelfTestSheet")
    Report "VlaEnsureSheet created it", Not ws Is Nothing, "missing"
    ws.Activate
    ws.Cells(2, 1).Value = "Widget"
    Report "VlaFindRow finds", VlaFindRow("Widget", "a") = 2, "got " & VlaFindRow("Widget", "a")
    Report "VlaFindRow miss is 0", VlaFindRow("Nope", "a") = 0, "got " & VlaFindRow("Nope", "a")

    ' B1: VlaCount answers lists directly and delegates ranges to
    ' Excel's COUNT (numbers only) - the pre-B1 behavior, unchanged.
    Dim tc As New Collection
    tc.Add 10
    tc.Add 20
    tc.Add 30
    Report "VlaCount on a list", VlaCount(tc) = 3, "got " & VlaCount(tc)
    ws.Range("B1").Value = 1
    ws.Range("B2").Value = 2
    ws.Range("B3").Value = "text"
    Report "VlaCount on a range counts numbers", VlaCount(ws.Range("B1:B3")) = 2, _
           "got " & VlaCount(ws.Range("B1:B3"))

    ' B7: element access - lists and ranges through one helper family.
    Report "VlaItem on a list", VlaItem(tc, 2) = 20, "got " & VlaItem(tc, 2)
    Report "VlaFirst on a list", VlaFirst(tc) = 10, "got " & VlaFirst(tc)
    Report "VlaLast on a list", VlaLast(tc) = 30, "got " & VlaLast(tc)
    Report "VlaFirst on a range", VlaFirst(ws.Range("B1:B3")) = 1, _
           "got " & VlaFirst(ws.Range("B1:B3"))
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
    prior.Activate
End Sub

' DI.2 Pass 3: VLA_IDE.VlaRegisterForAutoLoad/VlaUnregisterAutoLoad -
' the "Register for Auto-Load" ribbon command's own registry logic.
' Host-required: reads/writes real HKCU registry state via
' CreateObject("WScript.Shell"), and needs Application.Version.
'
' Deliberately uses a synthetic, never-real path rather than
' ThisWorkbook.FullName - registering the actual dev workbook for
' real auto-load as a side effect of running the self-test would be
' exactly the kind of surprising, hard-to-notice state mutation this
' project's own testing culture avoids (same reasoning as
' TestEmbeddedTextFallback just below choosing not to touch this
' workbook's real prelude.vla/english.vla). The synthetic path is
' harmless even if a failed run somehow left it behind: Excel would at
' worst report "file not found" for it once, not silently misbehave.
'
' Verifies against a direct, independent registry read (not just the
' functions' own True/False reports) - WScript.Shell's RegRead/
' RegWrite/RegDelete with the "HKCU\..." prefix format was confirmed
' live before this test was written, not assumed from memory.
Private Sub TestAutoLoadRegistration()
    Dim testPath As String
    testPath = "C:\VLA_Test_NoSuchFile\FrazaroSelfTestOnly.xlam"
    Dim sh As Object
    Set sh = CreateObject("WScript.Shell")

    ' Clean slate, in case a prior failed run left this behind.
    VLA_IDE.VlaUnregisterAutoLoad testPath

    Dim added As Boolean
    added = VLA_IDE.VlaRegisterForAutoLoad(testPath)
    Report "autoload: first registration reports True (added)", added, "got False"

    ' Independent check: read every slot directly and confirm testPath
    ' is genuinely sitting in the real registry, not just that the
    ' function claimed success.
    Dim found As Boolean
    Dim i As Long
    For i = 0 To 10
        Dim slot As String
        slot = "HKCU\Software\Microsoft\Office\" & Application.Version & "\Excel\Options\"
        If i = 0 Then slot = slot & "OPEN" Else slot = slot & "OPEN" & i
        Dim v As String
        On Error Resume Next
        v = sh.RegRead(slot)
        On Error GoTo 0
        If v = testPath Then found = True
    Next
    Report "autoload: independently readable in the real registry", found, "not found in OPEN..OPEN10"

    Dim addedAgain As Boolean
    addedAgain = VLA_IDE.VlaRegisterForAutoLoad(testPath)
    Report "autoload: registering the same path twice reports False (already present)", _
           Not addedAgain, "got True"

    Dim removed As Boolean
    removed = VLA_IDE.VlaUnregisterAutoLoad(testPath)
    Report "autoload: unregistering removes it, reports True", removed, "got False"

    Dim removedAgain As Boolean
    removedAgain = VLA_IDE.VlaUnregisterAutoLoad(testPath)
    Report "autoload: unregistering an absent path reports False", Not removedAgain, "got True"
End Sub

' DI2.1: VLA.VlaEmbeddedText - the reader half of VLA_Build's new
' EmbedTextAsSheet ("bundle prelude.vla/english.vla into the built
' add-in"). Host-required because it reads ThisWorkbook.Worksheets, the
' same reason TestHelpersHost sits here rather than in VLA_Tests.bas.
' Scoped to the reader alone, deliberately not the full
' PreludeMacros/IdeLoadVocab fallback path: exercising THAT would mean
' making this dev workbook's own real prelude.vla/english.vla
' temporarily invisible mid-test, a state hazard for no real gain once
' the reader itself is proven - the branching around it is a two-line
' If/Else, not new risk.
' PF.7: three more states added alongside the bulk-read/Join rewrite,
' none covered before this pass despite each being a real state the
' old cell-by-cell loop could hit too - the rewrite just makes them
' load-bearing instead of incidental. n=1 (`Range.Value` collapses to a
' bare scalar on a single-cell range - not hypothetical, it's
' IdeVocabFileName's own VLAe_Name shape on every Check in a built
' edition). n=1-but-blank (a sheet that exists with column A never
' written - End(xlUp) still returns row 1, not 0, so this is the real
' shape of "no content", never n=0). And an mRunScaleTests-gated
' round-trip at english.vla's own real order of magnitude (2,849 rows,
' scripts/polyglotta/english.vla's own line count) - PF.7's own scoping
' pass found this loop body unreachable from the dev workbook/self-test
' suite as it stood (both PreludeMacros and IdeVocabPath always find an
' external file first here), so this is the only repeatable, dev-side
' way to see a real-scale number without a full build+install cycle;
' the printed elapsed time is informational only, same discipline as
' TestVocabMacroProbe's own scale case (VLA_Tests_Grammar.bas).
Private Sub TestEmbeddedTextFallback()
    Dim sheetName As String
    sheetName = "VLAtest_Embed"
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(sheetName).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets.Add
    sh.Name = sheetName
    sh.Cells(1, 1).Value = "line one"
    sh.Cells(2, 1).Value = "line two"

    Dim got As String
    got = VLA.VlaEmbeddedText(sheetName)
    Report "embedtext: round-trips written lines", _
           got = "line one" & vbCrLf & "line two" & vbCrLf, got

    Report "embedtext: missing sheet returns empty string", _
           Len(VLA.VlaEmbeddedText("VLAtest_DoesNotExist")) = 0, "(n/a)"

    Application.DisplayAlerts = False
    sh.Delete
    Application.DisplayAlerts = True

    ' n=1: Range.Value on a single-cell range is a scalar, not a 2-D
    ' array - IdeVocabFileName's own VLAe_Name shape.
    Dim sheetOne As String
    sheetOne = "VLAtest_EmbedOne"
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(sheetOne).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Dim shOne As Worksheet
    Set shOne = ThisWorkbook.Worksheets.Add
    shOne.Name = sheetOne
    shOne.Cells(1, 1).Value = "solo line"

    Dim gotOne As String
    gotOne = VLA.VlaEmbeddedText(sheetOne)
    Report "embedtext: single-row sheet round-trips (n=1 scalar case)", _
           gotOne = "solo line" & vbCrLf, gotOne

    Application.DisplayAlerts = False
    shOne.Delete
    Application.DisplayAlerts = True

    ' n=1-but-blank: sheet exists, column A never written. End(xlUp)
    ' still returns row 1 (not 0) on a wholly blank column - this is
    ' the real shape "no content" takes, not a zero-row read.
    Dim sheetBlank As String
    sheetBlank = "VLAtest_EmbedBlank"
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(sheetBlank).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Dim shBlank As Worksheet
    Set shBlank = ThisWorkbook.Worksheets.Add
    shBlank.Name = sheetBlank

    Dim gotBlank As String
    gotBlank = VLA.VlaEmbeddedText(sheetBlank)
    Report "embedtext: sheet with no content returns a bare line break (n=1 blank cell)", _
           gotBlank = vbCrLf, "len=" & Len(gotBlank)

    Application.DisplayAlerts = False
    shBlank.Delete
    Application.DisplayAlerts = True

    ' Scale: english.vla's own real order of magnitude (2,849 rows),
    ' gated the same way TestVocabMacroProbe gates its own 150-macro
    ' case - the point is a repeatable, dev-side number for the loop
    ' body PF.7's scoping pass found otherwise unreachable from here.
    If VLA_Tests.mRunScaleTests Then
        Const scaleN As Long = 2849
        Dim sheetScale As String
        sheetScale = "VLAtest_EmbedScale"
        On Error Resume Next
        Application.DisplayAlerts = False
        ThisWorkbook.Worksheets(sheetScale).Delete
        Application.DisplayAlerts = True
        On Error GoTo 0

        Dim shScale As Worksheet
        Set shScale = ThisWorkbook.Worksheets.Add
        shScale.Name = sheetScale

        Dim writeArr() As Variant
        ReDim writeArr(1 To scaleN, 1 To 1)
        Dim k As Long
        For k = 1 To scaleN
            writeArr(k, 1) = "embedded scale line " & k
        Next k
        shScale.Range(shScale.Cells(1, 1), shScale.Cells(scaleN, 1)).Value = writeArr

        Dim t0 As Double, msElapsed As Double
        t0 = Timer
        Dim gotScale As String
        gotScale = VLA.VlaEmbeddedText(sheetScale)
        msElapsed = (Timer - t0) * 1000#
        Debug.Print "  pf7 scale: " & scaleN & "-row embedded sheet read in " & _
                    Format$(msElapsed, "0") & " ms"

        Dim gotLines() As String
        gotLines = Split(gotScale, vbCrLf)
        Report "embedtext scale: round-trips " & scaleN & " rows correctly", _
               (UBound(gotLines) - LBound(gotLines) = scaleN) And _
               gotLines(LBound(gotLines)) = "embedded scale line 1" And _
               gotLines(LBound(gotLines) + scaleN - 1) = "embedded scale line " & scaleN, _
               "row count " & (UBound(gotLines) - LBound(gotLines) + 1) & _
               ", first=" & gotLines(LBound(gotLines))

        Application.DisplayAlerts = False
        shScale.Delete
        Application.DisplayAlerts = True
    End If
End Sub

' PF.4b: VLA_Runtime.VlaSlabRead/VlaSlabWrite - PF.4's array-slab bulk
' read/write-back helpers. Host-required (real Worksheet ranges).
' Covers every shape PF.7's own VlaEmbeddedText fix already found
' load-bearing for the identical 1-cell-collapses-to-scalar quirk, plus
' the write-back side that fix never needed: multi-row/multi-col round-
' trip with a genuine mutate-then-verify (not just an unchanged echo,
' which could pass even if write-back silently did nothing), 1x1,
' 1-row/N-col, and N-row/1-col.
Private Sub TestArraySlabHelpers()
    Dim sheetName As String
    sheetName = "VLAtest_Slab"
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(sheetName).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets.Add
    sh.Name = sheetName

    ' Multi-row x multi-col (3 rows, 2 columns).
    sh.Cells(1, 1).Value = 1: sh.Cells(1, 2).Value = "a"
    sh.Cells(2, 1).Value = 2: sh.Cells(2, 2).Value = "b"
    sh.Cells(3, 1).Value = 3: sh.Cells(3, 2).Value = "c"
    Dim rng As Range
    Set rng = sh.Range("A1:B3")

    Dim arr As Variant
    arr = VLA_Runtime.VlaSlabRead(rng)
    Report "slab: multi-row/col read has the right shape", _
           UBound(arr, 1) = 3 And UBound(arr, 2) = 2, _
           "UBound1=" & UBound(arr, 1) & " UBound2=" & UBound(arr, 2)
    Report "slab: multi-row/col read round-trips content", _
           arr(1, 1) = 1 And arr(2, 2) = "b" And arr(3, 1) = 3, _
           "arr(1,1)=" & arr(1, 1) & " arr(2,2)=" & arr(2, 2) & " arr(3,1)=" & arr(3, 1)

    arr(2, 2) = "CHANGED"
    VLA_Runtime.VlaSlabWrite arr, rng
    Report "slab: write-back changes only the mutated cell", _
           sh.Cells(2, 2).Value = "CHANGED" And sh.Cells(1, 2).Value = "a" And sh.Cells(3, 2).Value = "c", _
           "B2=" & sh.Cells(2, 2).Value & " B1=" & sh.Cells(1, 2).Value & " B3=" & sh.Cells(3, 2).Value

    ' 1x1: the scalar-collapse quirk, both directions.
    sh.Cells(5, 1).Value = 42
    Dim one As Variant
    one = VLA_Runtime.VlaSlabRead(sh.Cells(5, 1))
    Report "slab: 1x1 read wraps the scalar quirk into a real array", _
           IsArray(one) And one(1, 1) = 42, "IsArray=" & IsArray(one)
    one(1, 1) = 99
    VLA_Runtime.VlaSlabWrite one, sh.Cells(5, 1)
    Report "slab: 1x1 write-back lands correctly", sh.Cells(5, 1).Value = 99, sh.Cells(5, 1).Value

    ' 1 row, N columns.
    sh.Cells(7, 1).Value = 10: sh.Cells(7, 2).Value = 20: sh.Cells(7, 3).Value = 30
    Dim rowArr As Variant
    rowArr = VLA_Runtime.VlaSlabRead(sh.Range("A7:C7"))
    rowArr(1, 2) = 999
    VLA_Runtime.VlaSlabWrite rowArr, sh.Range("A7:C7")
    Report "slab: 1-row/N-col round-trips and writes back correctly", _
           sh.Cells(7, 1).Value = 10 And sh.Cells(7, 2).Value = 999 And sh.Cells(7, 3).Value = 30, _
           "A7=" & sh.Cells(7, 1).Value & " B7=" & sh.Cells(7, 2).Value & " C7=" & sh.Cells(7, 3).Value

    ' N rows, 1 column.
    sh.Cells(9, 1).Value = 100: sh.Cells(10, 1).Value = 200: sh.Cells(11, 1).Value = 300
    Dim colArr As Variant
    colArr = VLA_Runtime.VlaSlabRead(sh.Range("A9:A11"))
    colArr(2, 1) = 888
    VLA_Runtime.VlaSlabWrite colArr, sh.Range("A9:A11")
    Report "slab: N-row/1-col round-trips and writes back correctly", _
           sh.Cells(9, 1).Value = 100 And sh.Cells(10, 1).Value = 888 And sh.Cells(11, 1).Value = 300, _
           "A9=" & sh.Cells(9, 1).Value & " A10=" & sh.Cells(10, 1).Value & " A11=" & sh.Cells(11, 1).Value

    Application.DisplayAlerts = False
    sh.Delete
    Application.DisplayAlerts = True
End Sub

' F11.0: the host-dependent tail of VLA_Tests.bas's TestRuntimeTrace -
' the flag's durable half persists as a hidden workbook Name (S5.1: a
' program's own EN_<program> module wipes module-level state before
' its first step runs, so the in-memory flag alone cannot survive a
' Run). The pure in-memory tracing half (VlaTrace/VlaTraceStep/
' VlaTraceReport/VlaTraceOn) stayed in TestRuntimeTrace.
Private Sub TestRuntimeTraceHost()
    ' S5.1 (owner smoke-test catch): the flag's durable half is a
    ' hidden workbook Name, because injecting EN_<program> into the
    ' project hosting its own run WIPES module state before main's
    ' first step - the in-memory flag died between the Immediate
    ' flip and the first guard. The Name survives recompiles; the
    ' cache-reload path is exercised for real only by an actual Run
    ' (the verification loop's smoke is the true verdict), but the
    ' durable half is pinnable here.
    Dim nv As String
    VlaTrace True
    On Error Resume Next
    nv = CStr(ThisWorkbook.Names("VLAt_TraceOn").RefersTo)
    On Error GoTo 0
    Report "trace: the flag persists as a workbook Name (on)", _
           InStr(nv, "1") > 0, "RefersTo was '" & nv & "'"
    VlaTrace False
    On Error Resume Next
    nv = CStr(ThisWorkbook.Names("VLAt_TraceOn").RefersTo)
    On Error GoTo 0
    Report "trace: the flag persists as a workbook Name (off)", _
           InStr(nv, "0") > 0, "RefersTo was '" & nv & "'"
End Sub

' ---------------------------------------------------------------------
'  V5.3: the runtime topology - pinned against the third field
'  report, the first STANDALONE run: with the dev workbook closed,
'  generated code's vlacolor call failed to compile, because the
'  helpers lived only inside the (locked) add-in and VBA has no
'  cross-project call. The dev machine masked this for the entire
'  project history - same project, everything resolved. These pins
'  hold the fix's invariants as far as a same-project test can:
'  every helper the vocabulary and engine emit is DEFINED in
'  VLA_Runtime (the injectable module), none remains defined in
'  VLA_English (a helper re-added there would work on dev and
'  break standalone again - the exact regression), and the inject
'  text stops at the boundary so add-in machinery never enters a
'  user's workbook. The FULL proof needs the standalone topology
'  itself - now the headline scenario of the D3 fresh-machine
'  script. Uses VBProject reads (dev trust setting, present).
' ---------------------------------------------------------------------
Private Sub TestRuntimeModule()
    Dim names As Variant
    names = Array("VlaColor", "VlaCount", "VlaItem", "VlaFirst", "VlaLast", _
                  "VlaFindRow", "VlaSendMail", "VlaEnsureSheet", "VlaDictNew", _
                  "VlaDictSet", "VlaDictGet", "VlaDictKeys", "VlaDictPairs", _
                  "VlaPairKey", "VlaPairValue")
    Dim rt As String, en As String, se As String
    rt = DevModuleText("VLA_Runtime")
    en = DevModuleText("VLA_English")
    se = DevModuleText("VLA_SentenceEngine")   ' LX5.2: split out of VLA_English
    Dim i As Long
    Dim missing As String, strays As String
    For i = LBound(names) To UBound(names)
        If InStr(1, rt, "Function " & names(i) & "(", vbTextCompare) = 0 _
           And InStr(1, rt, "Sub " & names(i) & "(", vbTextCompare) = 0 Then
            missing = missing & " " & names(i)
        End If
        If InStr(1, en, "Function " & names(i) & "(", vbTextCompare) > 0 _
           Or InStr(1, en, "Sub " & names(i) & "(", vbTextCompare) > 0 _
           Or InStr(1, se, "Function " & names(i) & "(", vbTextCompare) > 0 _
           Or InStr(1, se, "Sub " & names(i) & "(", vbTextCompare) > 0 Then
            strays = strays & " " & names(i)
        End If
    Next
    Report "runtime: every emitted helper is defined in VLA_Runtime", Len(missing) = 0, "missing:" & missing
    Report "runtime: no helper definition remains in VLA_English/VLA_SentenceEngine", Len(strays) = 0, "strayed:" & strays
    Dim itx As String
    itx = VlaRuntimeInjectText()
    Report "runtime: inject text is fenced (no machinery, not empty)", _
           InStr(1, itx, "VlaInjectRuntime") = 0 And Len(itx) > 1000, _
           "boundary leak or empty text (" & Len(itx) & " chars)"
End Sub

Private Function DevModuleText(ByVal moduleName As String) As String
    On Error Resume Next
    Dim cm As Object
    Set cm = ThisWorkbook.VBProject.VBComponents(moduleName).CodeModule
    If Not cm Is Nothing Then
        If cm.CountOfLines > 0 Then DevModuleText = cm.Lines(1, cm.CountOfLines)
    End If
    On Error GoTo 0
End Function

' ---------------------------------------------------------------------
'  IN.3: the dual-mode expression-parity harness. See this module's
'  IN3.0 history note (top of file) and VLA_Interpreter.bas's IN3_5.0
'  note for the scoping reasoning - a bare expression is what both
'  backends can produce a live, comparable answer for today.
' ---------------------------------------------------------------------
Private Sub TestExprParity()
    CheckExprParity "(+ 2 3)"
    CheckExprParity "(- 10 4)"
    CheckExprParity "(- 7)"
    CheckExprParity "(> 5 3)"
    CheckExprParity "(> 2 9)"
    CheckExprParity "(+ 1 2 3 4)"
    CheckExprParity "(- 5 2 1)"
    ' L-INTERPOLATE: named-hole substitution, including a hole reused
    ' twice (the shape datalog-filter-place/datalog-chain-place,
    ' english.vla, both need) and the "&"-matching numeric coercion.
    CheckExprParity "(interpolate ""Hello {name}!"" :name ""World"")"
    CheckExprParity "(interpolate ""{a}-{a}"" :a ""x"")"
    CheckExprParity "(interpolate ""n={n}"" :n 42)"
End Sub

' Runs exprVla through BOTH backends - the interpreter in-memory
' (VLA_Interpreter.VlaEvalExpression), the emitter live (VLA_DevRig's
' L6 scratch: VlaTryValueBuild wraps the expression as a real VBA
' Function's (return ...), VlaCompileToModule injects it, Application.Run
' executes it) - and reports whether they agree. CStr comparison
' throughout, matching this module's own CheckV convention: both sides
' return Variant (Double or Boolean here), and CStr is what a diff
' report can print. AS.6-guarded: a live compile/inject/run crossing
' Application.Run is exactly the class of call that raises untrappably
' if left bare (S4's verdict, cited throughout this codebase) - one bad
' expression reports FAIL under its own name instead of killing the
' rest of the host suite.
Private Sub CheckExprParity(ByVal exprVla As String)
    Dim d As String
    Dim interpResult As Variant
    Dim emitterResult As Variant

    On Error Resume Next
    interpResult = VLA_Interpreter.VlaEvalExpression(exprVla)
    If Err.Number <> 0 Then d = "interpreter: " & Err.Description
    On Error GoTo 0
    If Len(d) > 0 Then
        Report "expr parity: " & exprVla, False, d
        Exit Sub
    End If

    DropParityScratch "VLA_Scratch_Parity"
    On Error Resume Next
    Dim scratch As String
    scratch = VLA_DevRig.VlaTryValueBuild(exprVla)
    VLA.VlaCompileToModule scratch, "VLA_Scratch_Parity", ThisWorkbook
    emitterResult = Application.Run("'" & ThisWorkbook.Name & "'!VLA_Scratch_Parity.vla_scratch_value")
    If Err.Number <> 0 Then d = "emitter: " & Err.Description
    On Error GoTo 0
    DropParityScratch "VLA_Scratch_Parity"
    If Len(d) > 0 Then
        Report "expr parity: " & exprVla, False, d
        Exit Sub
    End If

    Report "expr parity: " & exprVla & " (interpreter agrees with live-executed VBA)", _
           CStr(interpResult) = CStr(emitterResult), _
           "interpreter=" & CStr(interpResult) & " emitter=" & CStr(emitterResult)
End Sub

' R7 duplicate of VLA_DevRig's Private TryDropScratch (cross-module
' Private calls do not exist in VBA), scoped to this test's own
' module names so it can never collide with a developer's live
' VlaTryValue session. Same rename-then-remove shape: module removal
' is deferred (hazard 11), so a plain Remove could collide with the
' very next injection in this same call. IN3.2: parameterized on the
' module name (was hardcoded to "VLA_Scratch_Parity") so TestStmtParity
' can reuse it under its own module name, "VLA_Scratch_StmtParity",
' rather than a second near-identical copy - both callers still get
' their own scratch module, never each other's.
Private Sub DropParityScratch(ByVal modName As String)
    Dim comp As Object
    On Error Resume Next
    Set comp = Nothing
    Set comp = ThisWorkbook.VBProject.VBComponents(modName & "_old")
    If Not comp Is Nothing Then ThisWorkbook.VBProject.VBComponents.Remove comp
    Set comp = Nothing
    Set comp = ThisWorkbook.VBProject.VBComponents(modName)
    If Not comp Is Nothing Then
        comp.Name = modName & "_old"
        ThisWorkbook.VBProject.VBComponents.Remove comp
    End If
    On Error GoTo 0
End Sub

' ---------------------------------------------------------------------
'  IN.3: statement-level parity, widening past bare expressions now
'  that IN.2 gives the interpreter a real place to write (set! to a
'  computed range). Each program writes its own answer into a fixed
'  probe cell - the interpreter runs it live (VLA_Interpreter.
'  VlaInterpret), the emitter runs it live too (VLA_DevRig's statement
'  scratch: VlaTryBuild wraps the program as a real VBA Sub - the
'  statement sibling of the expression harness's VlaTryValueBuild -
'  VLA.VlaCompileToModule injects it, Application.Run executes it) -
'  and the probe cell is read back with a PLAIN native Range read
'  between the two runs, never through either backend's own read path,
'  so a bug shared by both backends' reads could never manufacture a
'  false agreement (the same "consistency, not truth" caveat IN.3's own
'  roadmap text names). The probe is cleared before each run, so a
'  backend that raises before writing can never coast on the other
'  backend's leftover value. if/for/while/set! - the four forms IN.3's
'  own text names - each get one program; each check is two-sided
'  (both backends against a hand-computed expected value, not just
'  against each other), for the same "consistency, not truth" reason.
'  IN2.4: widened to for-each/do-until/select once IN2.3 gave the
'  interpreter its own implementation of each - the natural next check,
'  since those three are independently implemented on both sides
'  (native VBA control flow here, hand-rolled dispatch and the
'  mLoopBreak flag there) rather than sharing any code, which is
'  exactly the class of divergence parity testing exists to catch.
'  PF.4a: (set! (arr i) v) - array-element assignment, added once the
'  interpreter gained its own ExecSet case for it (VLA_Interpreter.bas)
'  to check against the emitter, which needed no change at all - its
'  own generic "set!" text substitution already emits real VBA array-
'  index-assignment syntax for free. Exactly the class of asymmetry
'  parity testing exists to catch: one backend got new code, the other
'  didn't, and only a live run proves they still agree.
' ---------------------------------------------------------------------
Private Sub TestStmtParity()
    CheckStmtParity "set! to a real place", _
        "(begin (set! (range ""A1"") 42))", 42
    CheckStmtParity "if, true branch", _
        "(begin (if (> 5 3) (then (set! (range ""A1"") 1)) (else (set! (range ""A1"") 0))))", 1
    CheckStmtParity "if, false branch", _
        "(begin (if (> 2 9) (then (set! (range ""A1"") 1)) (else (set! (range ""A1"") 0))))", 0
    CheckStmtParity "for, accumulating a sum", _
        "(begin (dim total Double) (dim i Double) (set! total 0) (for (i 1 5) (set! total (+ total i))) (set! (range ""A1"") total))", 15
    CheckStmtParity "while, counting down", _
        "(begin (dim n Double) (set! n 5) (while (> n 0) (set! n (- n 1))) (set! (range ""A1"") n))", 0

    ' IN2.4: for-each/do-until/select, added once IN2.3 gave the
    ' interpreter its own implementation of each to check against the
    ' emitter's - built independently (native VBA "For Each"/"Do
    ' Until"/"Select Case" on this side, hand-rolled dispatch and the
    ' mLoopBreak flag on the interpreter's), which is exactly the case
    ' parity testing earns its keep on. Each program also exercises
    ' exit-for/exit-do - IN2.3's own required companion, not named in
    ' IN.2's original text - since a loop that cannot be broken early
    ' is only half proven.
    CheckStmtParity "for-each over a real range, broken early with exit-for", _
        "(begin (set! (range ""B1"") 10) (set! (range ""B2"") 20) (set! (range ""B3"") 30)" & _
        " (dim total Double) (dim cnt Double) (dim c) (set! total 0) (set! cnt 0)" & _
        " (for-each (c (range ""B1:B3""))" & _
        "   (set! total (+ total (. c value)))" & _
        "   (set! cnt (+ cnt 1))" & _
        "   (if (> cnt 1) (then (exit-for))))" & _
        " (set! (range ""A1"") total))", 30
    CheckStmtParity "do-until, broken early with exit-do from a nested if", _
        "(begin (dim n Double) (set! n 1)" & _
        " (do-until (> n 100)" & _
        "   (set! n (+ n 1))" & _
        "   (if (> n 3) (then (exit-do))))" & _
        " (set! (range ""A1"") n))", 4
    CheckStmtParity "select, matching a multi-value case", _
        "(begin (dim lbl)" & _
        " (select 2" & _
        "   (case (1) (set! lbl ""one""))" & _
        "   (case (2 3) (set! lbl ""two-or-three""))" & _
        "   (case-else (set! lbl ""other"")))" & _
        " (set! (range ""A1"") lbl))", "two-or-three"

    ' PF.4a: set! into an array element - (arr 0) untouched (10) plus
    ' (arr 1) after being overwritten (99) sums to 109 regardless of
    ' Array()'s own base (this codebase declares no Option Base
    ' anywhere, so it stays native VBA 0-based) - one assertion proving
    ' both the write landed at the right index AND nothing else in the
    ' array was disturbed by it.
    CheckStmtParity "set! into an array element (PF.4a)", _
        "(begin (dim arr) (set! arr (array 10 20 30)) (set! (arr 1) 99)" & _
        " (set! (range ""A1"") (+ (arr 0) (arr 1))))", 109

    ' PF.4c: for-each-row's own full round trip, not just the in-memory
    ' array element access PF.4a already proved. Doubles every value in
    ' a single-column range in place, then reads the range back via a
    ' plain '.'-dot Value access - NOT through for-each-row again - so
    ' this proves the bulk write-back actually reached the live sheet,
    ' not just that the in-memory array looked right.
    CheckStmtParity "for-each-row: doubles a column in place, write-back lands on the sheet (PF.4c)", _
        "(begin (set! (range ""B1"") 10) (set! (range ""B2"") 20) (set! (range ""B3"") 30)" & _
        " (dim row)" & _
        " (for-each-row (row (range ""B1:B3""))" & _
        "   (set! (row 1) (* (row 1) 2)))" & _
        " (set! (range ""A1"") (+ (. (range ""B1"") value) (. (range ""B2"") value) (. (range ""B3"") value))))", _
        120

    ' PF.4c: break commits everything computed UP TO AND INCLUDING the
    ' break, never beyond - the settled design decision (3), proven
    ' live rather than assumed from the shape alone. Breaks after row 2
    ' is mutated; row 3 is never reached, so it must survive UNCHANGED
    ' in the write-back (20 + 40 + 30 = 90, not 20 + 40 + 60 = 120).
    CheckStmtParity "for-each-row: exit-for commits rows up to the break, not the untouched one after it (PF.4c)", _
        "(begin (set! (range ""B1"") 10) (set! (range ""B2"") 20) (set! (range ""B3"") 30)" & _
        " (dim row) (dim cnt Double) (set! cnt 0)" & _
        " (for-each-row (row (range ""B1:B3""))" & _
        "   (set! (row 1) (* (row 1) 2))" & _
        "   (set! cnt (+ cnt 1))" & _
        "   (if (> cnt 1) (then (exit-for))))" & _
        " (set! (range ""A1"") (+ (. (range ""B1"") value) (. (range ""B2"") value) (. (range ""B3"") value))))", _
        90
End Sub

' Runs stmtVla through BOTH backends against a shared probe cell (A1
' on its own scratch sheet, cleared before each run), reading each
' backend's result back natively (never through either backend's own
' read path - see the header note above). AS.6-guarded throughout, the
' same discipline as CheckExprParity: a raise anywhere reports FAIL
' under this program's own name instead of crossing Application.Run
' untrappably and killing the rest of the host suite. Its own scratch
' module, VLA_Scratch_StmtParity, can never collide with
' CheckExprParity's VLA_Scratch_Parity or a developer's live VlaTry
' session (DropParityScratch, IN3.2, takes the module name as a
' parameter for exactly this reuse).
Private Sub CheckStmtParity(ByVal name As String, ByVal stmtVla As String, ByVal wantVal As Variant)
    Dim prior As Worksheet
    Set prior = ActiveSheet
    VlaEnsureSheet "VlaStmtParitySheet"
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets("VlaStmtParitySheet")
    ws.Activate

    Dim d As String
    Dim gotInterp As Variant, gotEmitter As Variant

    ws.Range("A1").ClearContents
    On Error Resume Next
    VLA_Interpreter.VlaInterpret stmtVla
    If Err.Number <> 0 Then d = "interpreter: " & Err.Description
    On Error GoTo 0
    gotInterp = ws.Range("A1").Value

    If Len(d) = 0 Then
        ws.Range("A1").ClearContents
        DropParityScratch "VLA_Scratch_StmtParity"
        On Error Resume Next
        Dim scratch As String
        scratch = VLA_DevRig.VlaTryBuild(stmtVla)
        VLA.VlaCompileToModule scratch, "VLA_Scratch_StmtParity", ThisWorkbook
        Application.Run "'" & ThisWorkbook.Name & "'!VLA_Scratch_StmtParity.vla_scratch"
        If Err.Number <> 0 Then d = "emitter: " & Err.Description
        On Error GoTo 0
        DropParityScratch "VLA_Scratch_StmtParity"
        gotEmitter = ws.Range("A1").Value
    End If

    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
    prior.Activate

    If Len(d) > 0 Then
        Report "stmt parity: " & name, False, d
        Exit Sub
    End If

    Report "stmt parity: " & name & " (interpreter matches the expected value)", _
           CStr(gotInterp) = CStr(wantVal), "got " & CStr(gotInterp) & ", wanted " & CStr(wantVal)
    Report "stmt parity: " & name & " (emitter matches the expected value)", _
           CStr(gotEmitter) = CStr(wantVal), "got " & CStr(gotEmitter) & ", wanted " & CStr(wantVal)
    Report "stmt parity: " & name & " (interpreter agrees with live-executed VBA)", _
           CStr(gotInterp) = CStr(gotEmitter), "interpreter=" & CStr(gotInterp) & " emitter=" & CStr(gotEmitter)
End Sub

' ---------------------------------------------------------------------
'  IN.2: object-place/dot-dispatch/computed-set!, proven against a
'  REAL worksheet - exactly what IN0.5's own "out of scope on purpose"
'  note reserved for a live Excel object model. AS.6-guarded the same
'  way TestExprParity is: a raise anywhere here reports FAIL under its
'  own name and still cleans up the scratch sheet, rather than killing
'  the rest of the host suite or leaving debris behind.
' ---------------------------------------------------------------------
Private Sub TestInterpreterObjectDispatch()
    Dim prior As Worksheet
    Set prior = ActiveSheet
    VlaEnsureSheet "VlaInterpDispatchSheet"
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets("VlaInterpDispatchSheet")
    ws.Activate

    Dim d As String

    ' (range "...") as a place helper, read through '.' + value - the
    ' tier-1 dispatch and a zero-hop '.' read in one check.
    ws.Range("B2").Value = 99
    On Error Resume Next
    Dim gotB2 As Variant
    gotB2 = VLA_Interpreter.VlaEvalExpression("(. (range ""B2"") value)")
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    If Len(d) > 0 Then
        Report "in2: (range addr) + '.' value read", False, d
    Else
        CheckV "in2: (range addr) + '.' value read", gotB2, 99
    End If

    ' set! to a bare computed place - the "assign .Value" default-
    ' member simplification (ExecSet's own documented limit).
    d = ""
    On Error Resume Next
    VLA_Interpreter.VlaInterpret "(begin (set! (range ""B3"") 123))"
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    If Len(d) > 0 Then
        Report "in2: set! to a bare (range addr) place", False, d
    Else
        CheckV "in2: set! to a bare (range addr) place writes .Value", ws.Range("B3").Value, 123
        Report "in2: effect log records a computed set! (bare place)", _
               InStr(1, VLA_Interpreter.VlaInterpreterEffectLog(), "set: Range.Value = 123", vbTextCompare) > 0, _
               VLA_Interpreter.VlaInterpreterEffectLog()
    End If

    ' set! to an explicit '.' place - the general WalkMemberSet path,
    ' a genuinely different code path from the bare-place branch above.
    d = ""
    On Error Resume Next
    VLA_Interpreter.VlaInterpret "(begin (set! (. (range ""B4"") value) 456))"
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    If Len(d) > 0 Then
        Report "in2: set! to (. (range addr) value)", False, d
    Else
        CheckV "in2: set! to (. (range addr) value) writes the same cell", ws.Range("B4").Value, 456
        Report "in2: effect log records a computed set! ('.' place)", _
               InStr(1, VLA_Interpreter.VlaInterpreterEffectLog(), "set: .value on Range = 456", vbTextCompare) > 0, _
               VLA_Interpreter.VlaInterpreterEffectLog()
    End If

    ' '.' in statement position - a real ACTION (ClearContents), not a
    ' read; proves ExecDotCall and the Method-before-Get DynamicCall
    ' ordering, and that the effect log records it.
    ws.Range("B5").Value = "gone"
    ws.Range("B6").Value = "also gone"
    d = ""
    On Error Resume Next
    VLA_Interpreter.VlaInterpret "(begin (. (range ""B5:B6"") clearcontents))"
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    If Len(d) > 0 Then
        Report "in2: '.' statement call (ClearContents)", False, d
    Else
        Report "in2: '.' statement call runs a real method (ClearContents)", _
               Len(Trim$(CStr(ws.Range("B5").Value))) + Len(Trim$(CStr(ws.Range("B6").Value))) = 0, _
               "B5='" & ws.Range("B5").Value & "' B6='" & ws.Range("B6").Value & "'"
        Report "in2: effect log records a '.' statement call", _
               InStr(1, VLA_Interpreter.VlaInterpreterEffectLog(), "call: .clearcontents on Range", vbTextCompare) > 0, _
               VLA_Interpreter.VlaInterpreterEffectLog()
    End If

    ' A dotted-global chain (activesheet.name) - settles IN.0.5's own
    ' "open question... this measurement raises but does not settle":
    ' the first segment resolves to a real global object and the rest
    ' walks like a '.' form, no per-name branch required.
    d = ""
    On Error Resume Next
    Dim gotName As Variant
    gotName = VLA_Interpreter.VlaEvalExpression("(activesheet.name)")
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    If Len(d) > 0 Then
        Report "in2: dotted-global chain (activesheet.name)", False, d
    Else
        CheckV "in2: dotted-global chain (activesheet.name)", gotName, ws.Name
    End If

    ' A multi-hop '.' member chain (font.bold) - proves DescendToParent
    ' actually descends through an intermediate object rather than
    ' only ever handling one hop.
    ws.Range("B7").Font.Bold = True
    d = ""
    On Error Resume Next
    Dim gotBold As Variant
    gotBold = VLA_Interpreter.VlaEvalExpression("(. (range ""B7"") font.bold)")
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    If Len(d) > 0 Then
        Report "in2: multi-hop '.' chain (font.bold)", False, d
    Else
        CheckV "in2: multi-hop '.' chain (font.bold)", gotBold, True
    End If

    ' IN.3: the SET-position sibling of the read above - a multi-hop
    ' computed place (font.size), against a real object, for the first
    ' time (font.bold above only ever proved the READ side; the one
    ' prior computed-set! pin used "value", which bypasses DynamicSet's
    ' CallByName heuristic entirely). Found broken by VerifyReportInterpreter's
    ' own live run against instructions.txt itself ("Set font size of cell A1
    ' to 14." raised "'size' could not be assigned"), not by this pin -
    ' this pin is the regression proof for that fix (DynamicSet's own
    ' IN.3 note has the full reasoning), not the discovery.
    d = ""
    On Error Resume Next
    VLA_Interpreter.VlaInterpret "(begin (set! (. (range ""B8"") font.size) 18))"
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    If Len(d) > 0 Then
        Report "in.3: set! to a multi-hop computed place (font.size)", False, d
    Else
        CheckV "in.3: set! to a multi-hop computed place (font.size) writes the real font", _
               ws.Range("B8").Font.Size, 18
    End If

    ' IN.3: the THIRD real member DynamicSet's CallByName heuristic broke
    ' on - "color" (Interior.Color/Font.Color/Tab.Color all land here,
    ' DynamicSet's own IN.3 note has the full reasoning). Found broken by
    ' VerifyReportInterpreter's own live run ("Make cell C4 yellow."
    ' raised "'color' could not be assigned"), not by this pin.
    d = ""
    On Error Resume Next
    VLA_Interpreter.VlaInterpret "(begin (set! (. (range ""B9"") interior.color) 255))"
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    If Len(d) > 0 Then
        Report "in.3: set! to a multi-hop computed place (interior.color)", False, d
    Else
        CheckV "in.3: set! to a multi-hop computed place (interior.color) writes the real fill", _
               ws.Range("B9").Interior.Color, 255
    End If

    ' IN.3: the owner's own challenge, answered by counting rather than
    ' arguing (DynamicSet's own header note has the full reasoning) - the
    ' full real census (15 members, not the 3 that happened to crash)
    ' switched to native dispatch in one pass. Two more pinned live here,
    ' deliberately against DIFFERENT parent object types than size/color
    ' already covered (Font, Interior/Font/Tab) - "bold" (Font again, but
    ' a Boolean rather than a Double) and "hidden" (Columns, a type
    ' neither prior pin touched) - not exhaustive of all 15, the same
    ' "a few checks" precedent already set for the eleven named-arg
    ' members.
    d = ""
    On Error Resume Next
    VLA_Interpreter.VlaInterpret "(begin (set! (. (range ""B10"") font.bold) true))"
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    If Len(d) > 0 Then
        Report "in.3: set! to a multi-hop computed place (font.bold)", False, d
    Else
        Report "in.3: set! to a multi-hop computed place (font.bold) writes the real font", _
               ws.Range("B10").Font.Bold = True, "not bold"
    End If

    d = ""
    On Error Resume Next
    VLA_Interpreter.VlaInterpret "(begin (set! (. (columns ""E"") hidden) true))"
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    If Len(d) > 0 Then
        Report "in.3: set! to a computed place on a different parent type (columns.hidden)", False, d
    Else
        Report "in.3: set! to a computed place on a different parent type (columns.hidden) writes the real column", _
               ws.Columns("E").Hidden = True, "not hidden"
        ws.Columns("E").Hidden = False   ' undo - a hidden scratch column would outlive ws.Delete's own cleanup scope until then
    End If

    ' IN.3: the SEPARATE, silent bug the same challenge surfaced while
    ' investigating - not a CallByName failure at all. A bare place ATOM
    ' containing a "." (application.screenupdating, ...) is a dotted-
    ' global chain reference, the same as the zero-arg CALL form of the
    ' same text in expression position - ExecSet's own IN.3 note has the
    ' full reasoning. Before this fix, "Turn off screen updating."
    ' silently stored a fake frame variable literally named
    ' "application.screenupdating" and never touched the real Application
    ' object - no error, ever, at any point in this whole session's
    ' reloads, because a wrong dict entry doesn't raise.
    Dim priorScreenUpdating As Boolean
    priorScreenUpdating = Application.ScreenUpdating
    d = ""
    On Error Resume Next
    VLA_Interpreter.VlaInterpret "(begin (set! application.screenupdating false))"
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    If Len(d) > 0 Then
        Report "in.3: set! to a bare dotted-global place (application.screenupdating)", False, d
    Else
        Report "in.3: set! to a bare dotted-global place (application.screenupdating) touches the real Application object", _
               Application.ScreenUpdating = False, "still " & Application.ScreenUpdating
    End If
    Application.ScreenUpdating = priorScreenUpdating

    ' IN.3: the READ-position sibling of the fix above, found the same
    ' way VerifyReportInterpreter's own live run found everything else
    ' this session - "there is nothing stored at key 'rows.count'"
    ' (english.vla's last-filled-row macro, (cells rows.count c) -
    ' "rows.count" arrives as a bare EXPRESSION atom, not a variable
    ' reference). EvalExpr's own IN.3 note has the full reasoning -
    ' ResolveGlobalReceiver widened to cover the five place helpers'
    ' own zero-argument shape (rows/cells/columns/worksheets/workbooks),
    ' not just the five original Excel globals, so a bare dotted atom
    ' rooted at any of the ten now resolves the same way.
    d = ""
    On Error Resume Next
    Dim gotRowsCount As Variant
    gotRowsCount = VLA_Interpreter.VlaEvalExpression("rows.count")
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    If Len(d) > 0 Then
        Report "in.3: a bare dotted-global read rooted at a place helper (rows.count)", False, d
    Else
        CheckV "in.3: a bare dotted-global read rooted at a place helper (rows.count) matches the real Rows.Count", _
               gotRowsCount, Rows.Count
    End If

    ' IN.3: the GET-position sibling of DynamicSet's own native-dispatch
    ' story - "end" (Range.End(xlUp)), the first real GET-position member
    ' this corpus reaches that takes an argument. Found broken by
    ' VerifyReportInterpreter's own live run (english.vla's
    ' last-filled-row macro, "IN.2: 'end' is neither a readable property
    ' nor a callable method... (Get: Overflow; Method: Object doesn't
    ' support this property or method)"), not by this pin - mirrors the
    ' real macro's own shape, (. (. (cells rows.count c) end xlup) row),
    ' using a bare (range addr) instead of (cells rows.count c) only to
    ' keep this pin's own setup small.
    ws.Range("D1").Value = "x"
    ws.Range("D2").Value = "y"
    d = ""
    On Error Resume Next
    Dim gotLastRow As Variant
    gotLastRow = VLA_Interpreter.VlaEvalExpression("(. (. (range ""D1048576"") end xlup) row)")
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    If Len(d) > 0 Then
        Report "in.3: a GET-position member taking an argument (range.end xlup)", False, d
    Else
        CheckV "in.3: a GET-position member taking an argument (range.end xlup) finds the real last-filled row", _
               gotLastRow, 2
    End If

    ' IN.3: Application.WorksheetFunction's own members, native dispatch -
    ' EvalDynamicHead's own IN.3 note has the full reasoning. Found broken
    ' by VerifyReportInterpreter's own live run ("Set grand to sum of
    ' range B2:B3." silently returned a raw-pointer-shaped garbage number
    ' instead of 45, no error at all), not by this pin - this pin is the
    ' regression proof, covering both the 1-argument shape (sum, the one
    ' that actually broke) and the 2-argument shape (countif), not
    ' exhaustive of all seven real members.
    ws.Range("C1").Value = 10
    ws.Range("C2").Value = 20
    ws.Range("C3").Value = 30
    d = ""
    On Error Resume Next
    Dim gotSum As Variant
    gotSum = VLA_Interpreter.VlaEvalExpression("(application.worksheetfunction.sum (range ""C1:C3""))")
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    If Len(d) > 0 Then
        Report "in.3: application.worksheetfunction.sum over a real range", False, d
    Else
        CheckV "in.3: application.worksheetfunction.sum over a real range", gotSum, 60
    End If

    d = ""
    On Error Resume Next
    Dim gotCount As Variant
    gotCount = VLA_Interpreter.VlaEvalExpression( _
        "(application.worksheetfunction.countif (range ""C1:C3"") 20)")
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    If Len(d) > 0 Then
        Report "in.3: application.worksheetfunction.countif (2-argument shape)", False, d
    Else
        CheckV "in.3: application.worksheetfunction.countif (2-argument shape)", gotCount, 1
    End If

    ' IN.11: found by VerifyReportInterpreter's own live run past the
    ' on-error boundary, into instructions.txt's own "breadth pass" -
    ' "borders"/"tab"/"entirecolumn" (DynamicGet's own IN.11 note has
    ' the full reasoning) are each the first time this interpreter has
    ' fetched THAT specific intermediate GET segment in a multi-hop '.'
    ' chain, on a REAL object, for the first time - font.bold/font.size
    ' above only ever proved "font" as an intermediate segment works.
    d = ""
    On Error Resume Next
    VLA_Interpreter.VlaInterpret "(begin (set! (. (range ""B11"") borders.linestyle) xlcontinuous))"
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    If Len(d) > 0 Then
        Report "in.11: set! to a multi-hop computed place (borders.linestyle)", False, d
    Else
        Report "in.11: set! to a multi-hop computed place (borders.linestyle) writes the real border", _
               ws.Range("B11").Borders(xlEdgeBottom).LineStyle = xlContinuous, "border not set"
    End If

    d = ""
    On Error Resume Next
    VLA_Interpreter.VlaInterpret "(begin (set! (. (worksheets ""VlaInterpDispatchSheet"") tab.color) 255))"
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    If Len(d) > 0 Then
        Report "in.11: set! to a multi-hop computed place (tab.color)", False, d
    Else
        CheckV "in.11: set! to a multi-hop computed place (tab.color) writes the real tab", ws.Tab.Color, 255
    End If

    ' The GET+CALL sibling: "entirecolumn" as an intermediate segment,
    ' "autofit" as the final zero-arg method call on it (fit-all-
    ' columns's own shape, english.vla) - a smaller range than
    ' the real macro's whole-sheet "cells" on purpose, same code path.
    d = ""
    On Error Resume Next
    VLA_Interpreter.VlaInterpret "(begin (. (range ""B1:B2"") entirecolumn.autofit))"
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    Report "in.11: '.' statement call through a multi-hop GET (entirecolumn.autofit)", Len(d) = 0, d

    ' IN.11: the real zero-argument statement-position method census
    ' (DynamicCall's own IN.11 note has the full reasoning) - merge/
    ' unmerge (merge-range/unmerge-range) and insert/delete (insert-
    ' row-at/delete-row), against a real object, on a throwaway row
    ' far from every other pin's own cells.
    d = ""
    On Error Resume Next
    VLA_Interpreter.VlaInterpret "(begin (. (range ""B20:C20"") merge))"
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    If Len(d) > 0 Then
        Report "in.11: '.' statement call (merge)", False, d
    Else
        Report "in.11: '.' statement call runs a real method (merge)", ws.Range("B20").MergeCells, "not merged"
        d = ""
        On Error Resume Next
        VLA_Interpreter.VlaInterpret "(begin (. (range ""B20:C20"") unmerge))"
        If Err.Number <> 0 Then d = Err.Description
        On Error GoTo 0
        Report "in.11: '.' statement call runs a real method (unmerge)", Len(d) = 0 And Not ws.Range("B20").MergeCells, d
    End If

    d = ""
    On Error Resume Next
    VLA_Interpreter.VlaInterpret "(begin (. (rows 50) insert))"
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    If Len(d) > 0 Then
        Report "in.11: '.' statement call (insert row)", False, d
    Else
        d = ""
        On Error Resume Next
        VLA_Interpreter.VlaInterpret "(begin (. (rows 50) delete))"
        If Err.Number <> 0 Then d = Err.Description
        On Error GoTo 0
        Report "in.11: '.' statement call runs a real method (insert row, then delete row)", Len(d) = 0, d
    End If

    ' IN.11: "select" - freeze-top-row's own shape, missed by this
    ' item's first census read (the header comment on DynamicCall has
    ' the full reasoning: a macro body wrapped in "(begin ...)" was not
    ' matched by that first pass's own grep pattern), the actual gap
    ' the owner's second live run surfaced. Regression-pinned directly
    ' rather than only through the real macro, since freeze-top-row's
    ' own real shape also does a genuine window-state change
    ' (FreezePanes) this test sheet should not leave behind.
    d = ""
    On Error Resume Next
    VLA_Interpreter.VlaInterpret "(begin (. (rows 2) select))"
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    Report "in.11: '.' statement call runs a real method (select)", _
           Len(d) = 0 And TypeName(Selection) = "Range" And Selection.Row = 2, d

    ' IN.11: the actual root cause behind "Object doesn't support this
    ' property or method" past the breadth pass - not a missing native
    ' dispatch case at all, found only after EffectLogTail (VLA_Tests_
    ' Host.bas) named the exact statement in progress. "Set X to value
    ' in cell Y." compiles to (set! x (range "...")) - the PLACE, not
    ' (. (range "...") value) - relying on real VBA's own implicit Let
    ' coercion (a plain "x = Range(...)", no Set keyword, invokes the
    ' Range's default member, Value). ExecSet's own IN.11 note has the
    ' full mechanism. Two-sided: set! coerces (spaced-check's own real
    ' shape), obj-set! still keeps the bare object reference - the
    ' regression proof that this fix did not just make set! always
    ' behave like obj-set! used to, only the other way for the object
    ' case.
    ' A plain "gotX = VlaDictGet(...)" here would risk the exact Let-
    ' without-Set gotcha this pin exists to catch, one level up in this
    ' TEST's own code - IsObject/TypeName take the raw value as a
    ' function ARGUMENT, never a bare assignment, so the frame is
    ' captured once with Set and every inspection below stays safe
    ' regardless of which case (object or coerced scalar) it actually
    ' holds.
    ws.Range("B21").Value = "raw cell text"
    d = ""
    On Error Resume Next
    Dim frameSetCoerce As Object
    Set frameSetCoerce = VLA_Interpreter.VlaInterpret("(begin (dim v) (set! v (range ""B21"")))")
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    If Len(d) > 0 Then
        Report "in.11: set! to a plain variable coerces an object value (default member)", False, d
    Else
        Report "in.11: set! to a plain variable coerces an object value (default member)", _
               Not IsObject(VLA_Runtime.VlaDictGet(frameSetCoerce, "v")) And _
               CStr(VLA_Runtime.VlaDictGet(frameSetCoerce, "v")) = "raw cell text", _
               "got " & TypeName(VLA_Runtime.VlaDictGet(frameSetCoerce, "v")) & _
               " = " & CStr(VLA_Runtime.VlaDictGet(frameSetCoerce, "v"))
    End If

    d = ""
    On Error Resume Next
    Dim frameObjSet As Object
    Set frameObjSet = VLA_Interpreter.VlaInterpret("(begin (dim v) (obj-set! v (range ""B21"")))")
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    If Len(d) > 0 Then
        Report "in.11: obj-set! to a plain variable still keeps the object reference", False, d
    Else
        Report "in.11: obj-set! to a plain variable still keeps the object reference", _
               TypeName(VLA_Runtime.VlaDictGet(frameObjSet, "v")) = "Range", _
               "got " & TypeName(VLA_Runtime.VlaDictGet(frameObjSet, "v"))
    End If

    ' IN.11: a BARE, undotted global receiver name used as an ordinary
    ' expression - "Unfreeze panes." compiles to (set! (. activewindow
    ' freezepanes) false), a real '.' form whose first argument is the
    ' plain symbol "activewindow", not the fused "activewindow.
    ' freezepanes" atom the dotted-global fix already covered.
    ' EvalExpr's own IN.11 note has the full mechanism (a new fallback,
    ' checked only after both frames miss, so a real corpus variable
    ' would always win a same-name collision - never actually reached
    ' today).
    d = ""
    On Error Resume Next
    VLA_Interpreter.VlaInterpret "(begin (set! (. activewindow freezepanes) false))"
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    If Len(d) > 0 Then
        Report "in.11: set! to a '.' place whose object is a bare global receiver name (activewindow)", False, d
    Else
        Report "in.11: set! to a '.' place whose object is a bare global receiver name (activewindow)", _
               ActiveWindow.FreezePanes = False, "freezepanes not set"
    End If

    ' IN.11: "Copy range G1:I4 to range K1:M4." (copy-range-to,
    ' english.vla) - a real corpus statement that raised a bare
    ' "Type mismatch" for seven straight live runs before the actual
    ' cause was found: JoinKwArgs's own IN.11 note (VLA_Interpreter.bas)
    ' has the full story - a LOGGING function called AFTER this exact
    ' call already succeeds, CStr'ing an object value whose default
    ' member (for a multi-cell Range) is a 2D array. Two throwaway A/B
    ' controls that helped localize it (plain hand-written VBA, with
    ' and without Variant-function indirection, both confirmed the
    ' Excel API itself was never the problem) are not kept here - they
    ' never touched this project's own code, so they would not catch a
    ' regression; this pin, against the corpus's own real shape, would.
    ws.Range("D1:F4").Value = "src"
    d = ""
    On Error Resume Next
    VLA_Interpreter.VlaInterpret "(begin (. (range ""D1:F4"") copy :destination (range ""D10:F13"")))"
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    If Len(d) > 0 Then
        Report "in.11: '.' named-arg call (copy :destination), a computed multi-cell range value", False, d
    Else
        Report "in.11: '.' named-arg call (copy :destination), a computed multi-cell range value", _
               ws.Range("D10").Value = "src", "D10 = '" & ws.Range("D10").Value & "'"
    End If

    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
    prior.Activate
End Sub

' ---------------------------------------------------------------------
'  IN.6: "thisworkbook" must mean the workbook running the program, not
'  the add-in this interpreter physically lives in - a distinction
'  invisible in the dev-rig, where they are the same object, and
'  therefore untestable with ThisWorkbook alone (a bug that fell back
'  to it would pass a same-workbook check by accident). Proven the only
'  way that actually distinguishes them: a second, genuinely different
'  workbook, passed as VlaInterpret's own hostWb argument, checked
'  against BY NAME - if this regressed to a bare ThisWorkbook, the
'  probe would read the dev workbook's name instead and this pin would
'  fail loudly, not silently.
' ---------------------------------------------------------------------
Private Sub TestInterpreterHostWorkbook()
    Dim priorWb As Workbook
    Set priorWb = ActiveWorkbook
    Dim scratchWb As Workbook
    Set scratchWb = Workbooks.Add
    Dim d As String
    Dim frame As Object
    On Error Resume Next
    Set frame = VLA_Interpreter.VlaInterpret( _
        "(begin (dim probe) (set! probe thisworkbook.name))", scratchWb)
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    If Len(d) > 0 Then
        Report "in.6: thisworkbook resolves to the passed host workbook, not the add-in", False, d
    Else
        Report "in.6: thisworkbook resolves to the passed host workbook, not the add-in", _
               VLA_Runtime.VlaDictGet(frame, "probe") = scratchWb.Name, _
               "got '" & VLA_Runtime.VlaDictGet(frame, "probe") & "', wanted '" & scratchWb.Name & "'"
    End If
    Application.DisplayAlerts = False
    scratchWb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    priorWb.Activate
End Sub

' Live-caught during G-DATALOG grammar work, not by any prior test:
' set-formula ((set! (. r formula) f)) writes through Range.Formula,
' which auto-inserts "@" (implicit intersection) on anything that
' could spill - silently defeating every query engine's own "returns a
' spilled array" promise (SQL.1's SD-4 freeze) the moment FRAZARO
' ITSELF writes the formula, not just a person's stray Ctrl+Shift+Enter.
' First surfaced by =DATALOG(...) written into a live cell coming back
' as a single "X" (the array's own top-left/header cell) instead of a
' spilled range. Fixed by routing exactly the (. obj formula) shape
' through Range.Formula2 instead - both here (the interpreter, this
' test) and in the emitter (VLA.bas's own "set!" case, pinned by
' TestG10's ".formula2 =" fragment check, VLA_Tests_Grammar.bas).
' SEQUENCE(3), not DATALOG(...), on purpose: the bug is in set-formula
' itself, not any one query engine, so the regression test stays
' engine-independent - any native dynamic-array formula reproduces it.
Private Sub TestSetFormulaSpillsWithoutImplicitIntersection()
    Dim priorWb As Workbook
    Set priorWb = ActiveWorkbook
    Dim scratchWb As Workbook
    Set scratchWb = Workbooks.Add
    Dim ws As Worksheet
    Set ws = scratchWb.Worksheets(1)
    ws.Activate

    Dim d As String
    Dim frame As Object
    On Error Resume Next
    Set frame = VLA_Interpreter.VlaInterpret( _
        "(set! (. (range ""b2"") formula) ""=SEQUENCE(3)"")", scratchWb)
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0

    Dim gotFormula As String
    gotFormula = ws.Range("B2").Formula2
    Report "set-formula writes an array formula via Formula2, not Formula: no auto-inserted ""@""", _
           Len(d) = 0 And Left$(gotFormula, 2) <> "=@", _
           "error: '" & d & "'; got formula: '" & gotFormula & "'"

    Report "set-formula's array formula actually spills (B2:B4 all populated)", _
           Len(d) = 0 And Not IsEmpty(ws.Range("B2").Value) And _
           Not IsEmpty(ws.Range("B3").Value) And Not IsEmpty(ws.Range("B4").Value), _
           "b2=" & CStr(ws.Range("B2").Value) & " b3=" & CStr(ws.Range("B3").Value) & " b4=" & CStr(ws.Range("B4").Value)

    Application.DisplayAlerts = False
    scratchWb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    priorWb.Activate
End Sub

' ---------------------------------------------------------------------
'  IN.7: events, interpreter-native. "When the sheet changes:" compiles
'  to a top-level (sub on:sheet-change () ...) - the same declaration
'  shape a "To ...:" action already gets, dormant until something
'  calls it by name (VlaHasProc/VlaInterpretEntry) or, for real, until
'  VLA_Events' own registered entry fires from VLA_EventSink's
'  WithEvents Application_SheetChange. That last hop is untestable
'  from the dev-rig: instantiating VLA_EventSink here would attach a
'  SECOND live sink to this session's own Application object, on top
'  of whatever the built add-in may already have running, double-
'  firing every handler for the rest of the session - so this proves
'  the registry end to end (register/dispatch/unregister, by real
'  object identity, against two genuinely separate host workbooks,
'  IN.6's own TestInterpreterHostWorkbook discipline, immediately
'  above) and leaves the WithEvents wiring itself to a manual check,
'  the same class of gap TestExportSkeleton's own header already names
'  for the add-in-closed case.
' ---------------------------------------------------------------------
Private Sub TestInterpreterSheetChangeEvent()
    Dim priorWb As Workbook
    Set priorWb = ActiveWorkbook
    Dim wbA As Workbook, wbB As Workbook
    Set wbA = Workbooks.Add
    Set wbB = Workbooks.Add

    ' English grammar: translates, and is a top-level declaration - not
    ' folded into main, so a normal run must leave it dormant.
    Dim d As String, vla As String
    On Error Resume Next
    EnglishStepTracking False
    vla = EnglishToVla("Create a number called total." & vbCrLf & _
                        "Set total to 1." & vbCrLf & _
                        "When the sheet changes:" & vbCrLf & _
                        "Set total to 99.")
    EnglishStepTracking True
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    If Len(d) > 0 Then
        Report "in.7: 'When the sheet changes:' translates", False, d
    Else
        Report "in.7: 'When the sheet changes:' translates", _
               InStr(vla, "(sub on:sheet-change ") > 0, _
               "compiled VLA did not declare on:sheet-change: " & vla
    End If

    Dim d2 As String, frame As Object
    On Error Resume Next
    Set frame = VLA_Interpreter.VlaInterpret(vla, wbA)
    If Err.Number <> 0 Then d2 = Err.Description
    On Error GoTo 0
    If Len(d2) > 0 Then
        Report "in.7: a declared handler stays dormant during a normal run", False, d2
    Else
        Report "in.7: a declared handler stays dormant during a normal run", _
               VLA_Runtime.VlaDictGet(frame, "total") = 1, _
               "total was " & VLA_Runtime.VlaDictGet(frame, "total") & " after the run - the handler must not have executed"
    End If
    Report "in.7: VlaHasProc sees the just-run declaration", _
           VLA_Interpreter.VlaHasProc("on:sheet-change"), _
           "VlaHasProc(""on:sheet-change"") was False right after a run that declared it"

    ' VlaInterpretEntry: call the handler directly, from OUTSIDE any
    ' running program - the seam VLA_Events' own dispatch needs.
    ' thisworkbook must resolve to the PASSED host, not the add-in -
    ' the same IN.6 discipline TestInterpreterHostWorkbook proves,
    ' checked again here since VlaInterpretEntry is a second, real
    ' route into PrepareInterpret, not a thin pass-through trusted on
    ' its own word. Hand-written VLA, not English - dotted member
    ' access (thisworkbook.name) is an interpreter-internal/VLA-level
    ' concept; English's own tokenizer has no word-char role for '.'
    ' (it is sentence punctuation there).
    ' Checked via the effect log with debug-print, NOT set!/the
    ' returned frame - two separate reasons, found live rather than
    ' guessed: (1) the handler runs through CallUserProc, same as any
    ' user-defined action - a local (dim probe)/(set! probe ...) lives
    ' in CallUserProc's own ISOLATED call frame, discarded when the
    ' call returns, never the outer frame VlaInterpretEntry hands back
    ' (IN.10's own documented non-goal: a callee's local write never
    ' reaches the caller's frame) - reading the returned frame for
    ' "probe" raises "there is nothing stored at key 'probe'", the same
    ' isolation every other CallUserProc caller already relies on. (2)
    ' even reading the RIGHT frame, a PLAIN variable set! (ExecSet's
    ' "VLA_Runtime.VlaDictSet frame, plainName, v" path, no dot in the
    ' place) never calls LogEffect at all - only a dotted-global-
    ' receiver place, a '.' member set, or a computed place does. Only
    ' debug-print logs unconditionally, so it is the effect this test
    ' (and the dispatch one below) actually watches for.
    Dim d3 As String, vla2 As String
    vla2 = "(sub on:sheet-change () (debug-print thisworkbook.name))"
    On Error Resume Next
    VLA_Interpreter.VlaInterpretEntry vla2, "on:sheet-change", wbB
    If Err.Number <> 0 Then d3 = Err.Description
    On Error GoTo 0
    If Len(d3) > 0 Then
        Report "in.7: VlaInterpretEntry runs a named handler against the passed host", False, d3
    Else
        Report "in.7: VlaInterpretEntry runs a named handler against the passed host", _
               InStr(VLA_Interpreter.VlaInterpreterEffectLog(), "debug-print: " & wbB.Name) > 0, _
               "effect log did not show the handler's own effect: " & VLA_Interpreter.VlaInterpreterEffectLog()
    End If

    ' The registry: register/dispatch/unregister by real object
    ' identity. The handler's only effect is a debug-print (see the
    ' note above for why, not a cell write or set!) - the effect log
    ' is the interpreter's own record of what the LAST call did, read
    ' right after VlaDispatchSheetChange returns, which is also what
    ' proves EnableEvents's own save/restore bracket around the call
    ' didn't swallow the effect.
    Dim rawVla As String
    rawVla = "(sub on:sheet-change () (debug-print ""handler-ran""))"
    VLA_Events.VlaRegisterSheetChangeHandler wbA, rawVla
    Report "in.7: VlaRegisterSheetChangeHandler registers the named workbook", _
           VLA_Events.VlaHasSheetChangeHandler(wbA), "wbA was not registered"
    Report "in.7: an unregistered workbook has no handler", _
           Not VLA_Events.VlaHasSheetChangeHandler(wbB), "wbB reported a handler it was never given"

    Dim d4 As String
    On Error Resume Next
    VLA_Events.VlaDispatchSheetChange wbA
    If Err.Number <> 0 Then d4 = Err.Description
    On Error GoTo 0
    If Len(d4) > 0 Then
        Report "in.7: VlaDispatchSheetChange runs the registered handler", False, d4
    Else
        Report "in.7: VlaDispatchSheetChange runs the registered handler", _
               InStr(VLA_Interpreter.VlaInterpreterEffectLog(), "debug-print: handler-ran") > 0, _
               "effect log did not show the handler's own effect: " & VLA_Interpreter.VlaInterpreterEffectLog()
    End If

    Dim d5 As String
    On Error Resume Next
    VLA_Events.VlaDispatchSheetChange wbB   ' never registered - must be a silent no-op
    If Err.Number <> 0 Then d5 = Err.Description
    On Error GoTo 0
    Report "in.7: dispatching an unregistered workbook is a silent no-op", Len(d5) = 0, d5

    VLA_Events.VlaUnregisterHandlers wbA
    Report "in.7: VlaUnregisterHandlers clears the registration", _
           Not VLA_Events.VlaHasSheetChangeHandler(wbA), "wbA still reported a handler after VlaUnregisterHandlers"

    ' A second handler in the same program is refused loudly, not
    ' silently dropped - RegisterProc's own redefinition-replaces
    ' semantics would otherwise let the SECOND block silently win.
    Dim d6 As String
    On Error Resume Next
    EnglishStepTracking False
    EnglishToVla "When the sheet changes:" & vbCrLf & "Set total to 1." & vbCrLf & vbCrLf & _
                 "When the sheet changes:" & vbCrLf & "Set total to 2."
    EnglishStepTracking True
    If Err.Number <> 0 Then d6 = Err.Description
    On Error GoTo 0
    Report "in.7: a second 'When the sheet changes:' is refused, not silently dropped", _
           InStr(d6, "already a 'When the sheet changes:'") > 0, _
           "expected the duplicate-handler refusal, got: '" & d6 & "'"

    Application.DisplayAlerts = False
    wbA.Close SaveChanges:=False
    wbB.Close SaveChanges:=False
    Application.DisplayAlerts = True
    priorWb.Activate
End Sub

' ---------------------------------------------------------------------
'  IN.7 (button-click half): "When "<caption>" is clicked:" - the same
'  shape TestInterpreterSheetChangeEvent just proved for the sheet-
'  change half, generalized two ways: (1) multiple named handlers per
'  program, each its own caption-derived internal proc ("on:click:" &
'  ClickHandlerSlug(caption), VLA.bas - compiled parity's own follow-
'  up needs the target name recomputable from the caption alone)
'  rather than one fixed name, so the registry is keyed by (workbook,
'  button name) instead of workbook alone; (2) a
'  real MakeVlaButton smoke test, which the sheet-change half never
'  needed (nothing there draws onto a sheet) - the one genuinely new
'  piece of Excel object-model code this half adds, so it is the part
'  most worth proving live rather than trusting by analogy to the
'  registry code (which otherwise mirrors VlaRegisterSheetChangeHandler/
'  VlaDispatchSheetChange closely enough that a second full narrative
'  would only restate that file's own header). Same untestable-from-
'  the-dev-rig boundary as sheet-change: a real button CLICK routes
'  through VlaButtonClickDispatch's own Application.Caller/ActiveWorkbook
'  read, which only a live click ever populates - left to a manual
'  check, as TestInterpreterSheetChangeEvent's own header already
'  accepts for WithEvents_Application_SheetChange.
' ---------------------------------------------------------------------
Private Sub TestInterpreterButtonClickEvent()
    Dim priorWb As Workbook
    Set priorWb = ActiveWorkbook
    Dim wbA As Workbook, wbB As Workbook
    Set wbA = Workbooks.Add
    Set wbB = Workbooks.Add

    ' English grammar: two named handlers in one program, each its own
    ' declaration - not folded into main, so a normal run leaves both
    ' dormant, and EnglishClickHandlerNames/Procs must report both,
    ' caption bound to the matching numbered proc, in declaration order.
    Dim d As String, vla As String
    On Error Resume Next
    EnglishStepTracking False
    vla = EnglishToVla("Create a number called total." & vbCrLf & _
                        "Set total to 1." & vbCrLf & _
                        "When ""Go"" is clicked:" & vbCrLf & _
                        "Set total to 97." & vbCrLf & vbCrLf & _
                        "When ""Stop"" is clicked:" & vbCrLf & _
                        "Set total to 98.")
    EnglishStepTracking True
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    If Len(d) > 0 Then
        Report "in.7: 'When ""...."" is clicked:' translates", False, d
    Else
        Report "in.7: 'When ""...."" is clicked:' translates", _
               InStr(vla, "(sub on:click:Go ") > 0 And InStr(vla, "(sub on:click:Stop ") > 0, _
               "compiled VLA did not declare both on:click:Go and on:click:Stop: " & vla
    End If

    Dim clkNames As Collection, clkProcs As Collection
    Set clkNames = EnglishClickHandlerNames()
    Set clkProcs = EnglishClickHandlerProcs()
    Dim clkDebug As String
    Dim ck As Variant
    For Each ck In clkNames
        clkDebug = clkDebug & "'" & ck & "' "
    Next
    Report "in.7: EnglishClickHandlerNames/Procs report both handlers, in order", _
           clkNames.Count = 2 And clkProcs.Count = 2 And _
           CStr(clkNames.Item(1)) = "Go" And CStr(clkProcs.Item(1)) = "on:click:Go" And _
           CStr(clkNames.Item(2)) = "Stop" And CStr(clkProcs.Item(2)) = "on:click:Stop", _
           "got names: " & clkDebug & "(count " & clkNames.Count & "/" & clkProcs.Count & ")"

    Dim d2 As String, frame As Object
    On Error Resume Next
    Set frame = VLA_Interpreter.VlaInterpret(vla, wbA)
    If Err.Number <> 0 Then d2 = Err.Description
    On Error GoTo 0
    If Len(d2) > 0 Then
        Report "in.7: declared button handlers stay dormant during a normal run", False, d2
    Else
        Report "in.7: declared button handlers stay dormant during a normal run", _
               VLA_Runtime.VlaDictGet(frame, "total") = 1, _
               "total was " & VLA_Runtime.VlaDictGet(frame, "total") & " after the run - neither handler must have executed"
    End If
    Report "in.7: VlaHasProc sees both just-run declarations", _
           VLA_Interpreter.VlaHasProc("on:click:Go") And VLA_Interpreter.VlaHasProc("on:click:Stop"), _
           "VlaHasProc was False for on:click:Go or on:click:Stop right after a run that declared both"

    ' The registry: register/dispatch/unregister by (workbook, name),
    ' proven against TWO buttons in the SAME workbook so a dispatch by
    ' name, not just by workbook identity, is what actually happened -
    ' the one real generalization over VlaRegisterSheetChangeHandler's
    ' own single-per-workbook shape. Effect log, not set!/the returned
    ' frame, for the exact reasons TestInterpreterSheetChangeEvent's own
    ' header already gives (CallUserProc's isolated call frame; plain
    ' set! never calls LogEffect).
    Dim rawVlaGo As String, rawVlaStop As String
    rawVlaGo = "(sub on:click:1 () (debug-print ""go-ran""))"
    rawVlaStop = "(sub on:click:1 () (debug-print ""stop-ran""))"
    VLA_Events.VlaRegisterButtonClickHandler wbA, "Go", rawVlaGo, "on:click:1"
    VLA_Events.VlaRegisterButtonClickHandler wbA, "Stop", rawVlaStop, "on:click:1"
    Report "in.7: VlaRegisterButtonClickHandler registers each named button", _
           VLA_Events.VlaHasButtonClickHandler(wbA, "Go") And VLA_Events.VlaHasButtonClickHandler(wbA, "Stop"), _
           "wbA did not report both 'Go' and 'Stop' as registered"
    Report "in.7: an unregistered (workbook, name) pair has no handler", _
           Not VLA_Events.VlaHasButtonClickHandler(wbB, "Go") And Not VLA_Events.VlaHasButtonClickHandler(wbA, "Nope"), _
           "a handler was reported for a pair that was never registered"

    Dim d4 As String
    On Error Resume Next
    VLA_Events.VlaDispatchButtonClick wbA, "Stop"
    If Err.Number <> 0 Then d4 = Err.Description
    On Error GoTo 0
    If Len(d4) > 0 Then
        Report "in.7: VlaDispatchButtonClick runs the handler for the matching name only", False, d4
    Else
        Report "in.7: VlaDispatchButtonClick runs the handler for the matching name only", _
               InStr(VLA_Interpreter.VlaInterpreterEffectLog(), "debug-print: stop-ran") > 0, _
               "effect log did not show 'Stop''s own effect: " & VLA_Interpreter.VlaInterpreterEffectLog()
    End If

    Dim d5 As String
    On Error Resume Next
    VLA_Events.VlaDispatchButtonClick wbB, "Go"   ' never registered on wbB - must be a silent no-op
    If Err.Number <> 0 Then d5 = Err.Description
    On Error GoTo 0
    Report "in.7: dispatching an unregistered (workbook, name) pair is a silent no-op", Len(d5) = 0, d5

    VLA_Events.VlaUnregisterHandlers wbA
    Report "in.7: VlaUnregisterHandlers clears every button this workbook registered", _
           Not VLA_Events.VlaHasButtonClickHandler(wbA, "Go") And Not VLA_Events.VlaHasButtonClickHandler(wbA, "Stop"), _
           "wbA still reported a handler after VlaUnregisterHandlers"

    ' A repeated caption within one program is refused loudly - CheckName/
    ' CheckDupAction's own "one name, one definition" reflex, IN.7's own
    ' precedent for a second 'When the sheet changes:' generalized to
    ' per-caption instead of per-program.
    Dim d6 As String
    On Error Resume Next
    EnglishStepTracking False
    EnglishToVla "When ""Go"" is clicked:" & vbCrLf & "Set total to 1." & vbCrLf & vbCrLf & _
                 "When ""Go"" is clicked:" & vbCrLf & "Set total to 2."
    EnglishStepTracking True
    If Err.Number <> 0 Then d6 = Err.Description
    On Error GoTo 0
    Report "in.7: a repeated 'When ""Go"" is clicked:' caption is refused, not silently dropped", _
           InStr(d6, "already a 'When ""Go"" is clicked:'") > 0, _
           "expected the duplicate-handler refusal, got: '" & d6 & "'"

    ' Two DIFFERENT captions whose slug collides (punctuation aside) are
    ' refused too, with words and a line number, at PARSE time - before
    ' this check existed, this would have silently ridden into VBA's own
    ' "duplicate declaration" compile error the first time Compile ran
    ' (compiled parity's own comment on the make-button EmitStmt case
    ' has the full reasoning).
    Dim d6b As String
    On Error Resume Next
    EnglishStepTracking False
    EnglishToVla "When ""Go!"" is clicked:" & vbCrLf & "Set total to 1." & vbCrLf & vbCrLf & _
                 "When ""Go?"" is clicked:" & vbCrLf & "Set total to 2."
    EnglishStepTracking True
    If Err.Number <> 0 Then d6b = Err.Description
    On Error GoTo 0
    Report "in.7: two captions colliding on the same slug are refused with a line number", _
           InStr(d6b, "Go!") > 0 And InStr(d6b, "Go?") > 0 And InStr(d6b, "line") > 0, _
           "expected a slug-collision refusal naming both captions and a line, got: '" & d6b & "'"

    ' Owner-found live gap: 'Create a button called "..." at cell ...:'
    ' reads like every other 'Create a <kind> called <name>.' sentence
    ' (number/text/value/list/lookup all use that exact shape), but
    ' "create"/"called" is CORE grammar, dispatched unconditionally on
    ' the sentence's first word - no phrasebook rule (make-button's
    ' own included) could ever intercept it. Before this redirect
    ' existed, the quoted caption tripped a confusing "expected a name
    ' after 'called'" error, with no hint that Make a button ... at
    ' cell ...' was the sentence that actually works.
    Dim d6c As String
    On Error Resume Next
    EnglishStepTracking False
    EnglishToVla "Create a button called ""Clicky"" at cell C3."
    EnglishStepTracking True
    If Err.Number <> 0 Then d6c = Err.Description
    On Error GoTo 0
    Report "in.7: 'Create a button called ...' redirects to 'Make a button ...'", _
           InStr(d6c, "isn't created with 'Create") > 0 And InStr(d6c, "Make a button") > 0, _
           "expected the button-redirect refusal, got: '" & d6c & "'"

    ' G-PIVOT: the same collision, hit live by pivot-create's own
    ' pareto.txt surface text ("Create a pivot table from ... called
    ' ...") - "table" arrives where this handler's Create-declaration
    ' shape expects "called", same mismatch class as button above.
    Dim d6p As String
    On Error Resume Next
    EnglishStepTracking False
    EnglishToVla "Create a pivot table from A1:C10 at F1 called SalesPivot."
    EnglishStepTracking True
    If Err.Number <> 0 Then d6p = Err.Description
    On Error GoTo 0
    Report "in.7: 'Create a pivot table ...' redirects to 'Make a pivot table ...'", _
           InStr(d6p, "isn't created with 'Create") > 0 And InStr(d6p, "Make a pivot table") > 0, _
           "expected the pivot-redirect refusal, got: '" & d6p & "'"

    ' Owner-found live gap, the second half of the same story: 'When
    ' Clicky is clicked:' (forgotten quotes around the button name)
    ' does not refuse at all - AtButtonClickEvent's own 5-token
    ' lookahead requires a QUOTED name, misses on the bare word, and
    ' the sentence silently falls through to the generic 'When <value>
    ' is <case>:' choices form instead, emitting a real
    ' "Select Case clicky" against a variable that was never declared -
    ' an untrappable VBA "Variable not defined" compile error, caught
    ' live. ParseStmt's own "when" case now redirects before that
    ' happens.
    Dim d6d As String
    On Error Resume Next
    EnglishStepTracking False
    EnglishToVla "When Clicky is clicked:" & vbCrLf & "Show ""Clicky clicked.""."
    EnglishStepTracking True
    If Err.Number <> 0 Then d6d = Err.Description
    On Error GoTo 0
    Report "in.7: 'When Clicky is clicked:' (unquoted) redirects instead of compiling a phantom variable", _
           InStr(d6d, "needs quotes") > 0 And InStr(d6d, """clicky""") > 0, _
           "expected the unquoted-click-name redirect, got: '" & d6d & "'"

    ' LE.2 follow-up, owner-found live gap: the other half of the same
    ' story - 'When "Clicky" is clocked:' (QUOTED name, but "clicked"
    ' itself misspelled). AtButtonClickEvent's exact 4-token shape
    ' misses on word 3 this time, and the sentence again falls through
    ' to the generic choices form: a literal string scrutinee (always
    ' the same case, never sensible on its own) paired with a bare,
    ' undeclared case value - the same untrappable VBA "Variable not
    ' defined" compile error, just from the other missing word.
    Dim d6e As String
    On Error Resume Next
    EnglishStepTracking False
    EnglishToVla "When ""Clicky"" is clocked:" & vbCrLf & "Show ""Clicky clicked.""."
    EnglishStepTracking True
    If Err.Number <> 0 Then d6e = Err.Description
    On Error GoTo 0
    Report "in.7: 'When ""Clicky"" is clocked:' (misspelled) redirects instead of compiling a phantom variable", _
           InStr(d6e, "misspelled") > 0 And InStr(d6e, "clocked") > 0, _
           "expected the misspelled-clicked redirect, got: '" & d6e & "'"

    ' MakeVlaButton: the one genuinely new piece of Excel object-model
    ' code this half adds. "range" resolves bare against ActiveSheet
    ' (EvalDynamicHead's own documented shape, unrelated to hostWb), so
    ' wbA must be the active workbook for its cell A1 to be where the
    ' button actually lands - VlaTimeIt's own "aim by activating"
    ' discipline, not a new rule invented for this test.
    wbA.Activate
    Dim d7 As String
    On Error Resume Next
    VLA_Interpreter.VlaInterpret "(sub main () (make-button ""Test Button"" (range ""a1"")))", wbA
    If Err.Number <> 0 Then d7 = Err.Description
    On Error GoTo 0
    If Len(d7) > 0 Then
        Report "in.7: make-button draws a real Form Control button", False, d7
    Else
        Dim btn As Button
        Set btn = Nothing
        On Error Resume Next
        Set btn = wbA.ActiveSheet.Buttons("Test Button")
        On Error GoTo 0
        If btn Is Nothing Then
            Report "in.7: make-button draws a real Form Control button", False, _
                   "no Button named 'Test Button' was found on wbA's active sheet"
        Else
            Report "in.7: make-button draws a real Form Control button", _
                   btn.Caption = "Test Button" And InStr(btn.OnAction, "VlaButtonClickDispatch") > 0, _
                   "caption='" & btn.Caption & "' onaction='" & btn.OnAction & "'"
        End If
    End If

    Application.DisplayAlerts = False
    wbA.Close SaveChanges:=False
    wbB.Close SaveChanges:=False
    Application.DisplayAlerts = True
    priorWb.Activate
End Sub

' ---------------------------------------------------------------------
'  IN.7 (button-click half), compiled parity: EmitStmt's "make-button"
'  case and EmitProc's narrowed colon-guard, proven from VlaTranspile's
'  own emitted VBA TEXT - unlike TestInterpreterButtonClickEvent just
'  above, none of this needs a live workbook (VlaTranspile is a pure
'  text transform), so it stays a text-shape check: caption ->
'  'on:click:' & ClickHandlerSlug wiring, a decorative button with no
'  matching handler leaving OnAction unset instead of referencing a
'  Sub that was never declared, a punctuated caption's slug folding
'  end to end, the shared helper Sub appearing only when make-button
'  was actually used, and 'When the sheet changes:' still refusing
'  Compile now that click-handler names no longer do - the one
'  regression the narrowed guard could most easily have introduced.
' ---------------------------------------------------------------------
Private Sub TestCompiledButtonClickParity()
    ' Local source-text variables are named "vlaSrc*", never "vla" -
    ' VBA is case-insensitive, and a local var named "vla" silently
    ' SHADOWS the "VLA" module itself: "VLA.VlaTranspile(...)" inside
    ' such a scope resolves "VLA" to the local STRING variable, not
    ' the module, and fails to compile with "Invalid qualifier" (an
    ' owner-found live compile error on this test's own first pass -
    ' local scope wins over module scope for a case-insensitive name).
    ' "Make a button ... at cell ...:" is a english.vla PHRASEBOOK
    ' macro, not core grammar (unlike "Set"/"Show"/"When ... is
    ' clicked:", all built in) - EnglishToVla only understands it once
    ' the real vocab file is loaded, exactly like VerifyReportInterpreter's
    ' own vocabPath/EnglishLoadVocabulary pair above. Nothing later in
    ' this suite depends on staying at the built-in-only 11 phrases (no
    ' test asserts a phrase count), so this is left loaded rather than
    ' reset back afterward.
    Dim vocabPath As String
    vocabPath = VLA_Tests.FindDevFile("english.vla")
    EnglishResetGrammar
    EnglishLoadVocabulary vocabPath

    Dim d As String, vlaSrc As String, vba As String
    On Error Resume Next
    EnglishStepTracking False
    vlaSrc = EnglishToVla("When ""Go"" is clicked:" & vbCrLf & "Show ""hi""." & vbCrLf & vbCrLf & _
                        "Make a button ""Go"" at cell a1.")
    EnglishStepTracking True
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    If Len(d) > 0 Then
        Report "in.7 (compiled): 'When ""Go"" is clicked:' + make-button translates", False, d
    Else
        Dim d1 As String
        On Error Resume Next
        Err.Clear
        vba = VLA.VlaTranspile(vlaSrc)
        If Err.Number <> 0 Then d1 = Err.Description
        On Error GoTo 0
        If Len(d1) > 0 Then
            Report "in.7 (compiled): a matched caption compiles clean", False, d1
        Else
            Report "in.7 (compiled): a matched caption compiles clean", _
                   InStr(vba, "Sub on_click_Go(") > 0 And _
                   InStr(vba, "Call VlaCompiledMakeButton(""Go"", range(") > 0 And _
                   InStr(vba, ", ""on_click_Go"")") > 0 And _
                   InStr(vba, "Private Sub VlaCompiledMakeButton(") > 0, _
                   "compiled VBA: " & vba
        End If
    End If

    Dim d2 As String, vlaSrc2 As String, vba2 As String
    On Error Resume Next
    EnglishStepTracking False
    vlaSrc2 = EnglishToVla("Make a button ""Solo"" at cell b2.")
    EnglishStepTracking True
    If Err.Number <> 0 Then d2 = Err.Description
    On Error GoTo 0
    If Len(d2) > 0 Then
        Report "in.7 (compiled): an unhandled caption still compiles, OnAction unset", False, d2
    Else
        Dim d3 As String
        On Error Resume Next
        Err.Clear
        vba2 = VLA.VlaTranspile(vlaSrc2)
        If Err.Number <> 0 Then d3 = Err.Description
        On Error GoTo 0
        If Len(d3) > 0 Then
            Report "in.7 (compiled): an unhandled caption still compiles, OnAction unset", False, d3
        Else
            Report "in.7 (compiled): an unhandled caption still compiles, OnAction unset", _
                   InStr(vba2, "Call VlaCompiledMakeButton(""Solo"", range(") > 0 And _
                   InStr(vba2, ", """")") > 0, _
                   "compiled VBA: " & vba2
        End If
    End If

    Dim d4 As String, vlaSrc4 As String, vba4 As String
    On Error Resume Next
    EnglishStepTracking False
    vlaSrc4 = EnglishToVla("When ""Recalculate Totals!"" is clicked:" & vbCrLf & "Show ""hi""." & vbCrLf & vbCrLf & _
                         "Make a button ""Recalculate Totals!"" at cell c3.")
    EnglishStepTracking True
    If Err.Number <> 0 Then d4 = Err.Description
    On Error GoTo 0
    If Len(d4) > 0 Then
        Report "in.7 (compiled): a punctuated caption's slug wires through", False, d4
    Else
        Dim d5 As String
        On Error Resume Next
        Err.Clear
        vba4 = VLA.VlaTranspile(vlaSrc4)
        If Err.Number <> 0 Then d5 = Err.Description
        On Error GoTo 0
        If Len(d5) > 0 Then
            Report "in.7 (compiled): a punctuated caption's slug wires through", False, d5
        Else
            Report "in.7 (compiled): a punctuated caption's slug wires through", _
                   InStr(vba4, "Sub on_click_Recalculate_Totals_(") > 0 And _
                   InStr(vba4, """on_click_Recalculate_Totals_""") > 0, _
                   "compiled VBA: " & vba4
        End If
    End If

    Dim d6 As String, vlaSrc6 As String, vba6 As String
    On Error Resume Next
    Err.Clear
    vlaSrc6 = EnglishToVla("Show ""no buttons here"".")
    vba6 = VLA.VlaTranspile(vlaSrc6)
    If Err.Number <> 0 Then d6 = Err.Description
    On Error GoTo 0
    If Len(d6) > 0 Then
        Report "in.7 (compiled): the make-button helper is not emitted when unused", False, d6
    Else
        Report "in.7 (compiled): the make-button helper is not emitted when unused", _
               InStr(vba6, "VlaCompiledMakeButton") = 0, "compiled VBA: " & vba6
    End If

    Dim vlaSrc7 As String
    On Error Resume Next
    EnglishStepTracking False
    vlaSrc7 = EnglishToVla("When the sheet changes:" & vbCrLf & "Show ""hi"".")
    EnglishStepTracking True
    On Error GoTo 0
    Dim errNum8 As Long, d8 As String
    On Error Resume Next
    Err.Clear
    VLA.VlaTranspile vlaSrc7
    errNum8 = Err.Number
    d8 = Err.Description
    On Error GoTo 0
    Report "in.7 (compiled): 'When the sheet changes:' still refuses Compile", _
           errNum8 = VLA.VLA_ERR_INTERPRETER_ONLY, _
           "expected the interpreter-only refusal (" & VLA.VLA_ERR_INTERPRETER_ONLY & "), got err " & errNum8 & ": '" & d8 & "'"
End Sub

' ---------------------------------------------------------------------
'  IN.4 (part two, walking skeleton): proves the export MECHANISM -
'  SaveCopyAs a real file, reopen it, inject the runtime and a compiled
'  module into THAT copy, run it, read the result back from the copy's
'  own cells. What this does NOT prove, and cannot from inside a
'  running VBA host: that the exported file runs with the add-in
'  actually CLOSED - this test's own host workbook is still open and
'  on the call stack the whole time, so a bug where the injected code
'  accidentally still reached back into this project's own modules
'  (rather than its own injected Frazaro_EN_Runtime copy) would not be
'  caught here. That claim needs the manual check
'  EnglishIdeExportSkeleton's own confirmation message asks for. Raw
'  VLA, no macro dependency (vlaensuresheet is a bare runtime helper;
'  "(. (worksheets ...) activate)" is activate-sheet's own macro body,
'  written directly rather than depending on english.vla being
'  loaded into a fresh compile's macro table) - same discipline as
'  every other hand-built pin in this file.
' ---------------------------------------------------------------------
Private Sub TestExportSkeleton()
    Dim exportPath As String
    exportPath = Environ$("TEMP") & Application.PathSeparator & "VlaExportSkeletonTest.xlsm"
    On Error Resume Next
    If Len(Dir$(exportPath)) > 0 Then Kill exportPath
    On Error GoTo 0

    Dim d As String
    Dim copyWb As Workbook
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.SaveCopyAs exportPath
    Application.DisplayAlerts = True
    If Err.Number <> 0 Then d = "SaveCopyAs: " & Err.Description
    On Error GoTo 0

    If Len(d) = 0 Then
        On Error Resume Next
        Set copyWb = Workbooks.Open(exportPath)
        If Err.Number <> 0 Then d = "reopen: " & Err.Description
        On Error GoTo 0
    End If

    If Len(d) = 0 Then
        On Error Resume Next
        VlaInjectRuntime copyWb
        ' VBA has no loose top-level executable statements - the body
        ' needs a real (sub main () ...) wrapper (English's own
        ' translation always supplies one; this raw VLA has to do it
        ' by hand), with each statement as its own trailing element,
        ' not a single (begin ...) bundle - EmitProc's own body-forms
        ' shape, matching (sub NAME (params) stmt stmt stmt...).
        VLA.VlaCompileToModule _
            "(sub main () " & _
            "(vlaensuresheet ""VlaExportSkeletonTest"") " & _
            "(. (worksheets ""VlaExportSkeletonTest"") activate) " & _
            "(set! (range ""A1"") 4242))", _
            "VlaExportSkeletonMod", copyWb
        If Err.Number <> 0 Then d = "inject/compile: " & Err.Description
        On Error GoTo 0
    End If

    Dim gotVal As Variant
    If Len(d) = 0 Then
        On Error Resume Next
        Application.Run "'" & copyWb.Name & "'!VlaExportSkeletonMod.main"
        If Err.Number <> 0 Then
            d = "run: " & Err.Description
        Else
            gotVal = copyWb.Worksheets("VlaExportSkeletonTest").Range("A1").Value
        End If
        On Error GoTo 0
    End If

    If Len(d) > 0 Then
        Report "in.4: exported copy compiles and runs from its own injected runtime", False, d
    Else
        Report "in.4: exported copy compiles and runs from its own injected runtime", _
               gotVal = 4242, "got '" & gotVal & "', wanted 4242"
    End If

    On Error Resume Next
    If Not copyWb Is Nothing Then copyWb.Close SaveChanges:=False
    If Len(Dir$(exportPath)) > 0 Then Kill exportPath
    On Error GoTo 0
End Sub

' ---------------------------------------------------------------------
'  IN2.3: for-each over a real Range, live - matching how the corpus
'  actually uses it ("Remember range B2:B4 as results. For each r in
'  results, log r." - the range IS the list; there is no separate
'  Collection construction step). A pure pin would need
'  VLA_Runtime.VlaDictNew() reached through Application.Run with an
'  OBJECT return, round-tripped as an argument to a second such call
'  (VlaDictSet) - a real gap (this interpreter has no (new Collection)
'  support at all, a separate, pre-existing hole IN2.3 does not touch)
'  that this pin sidesteps entirely by testing for-each the way it is
'  actually written today, over an already-proven place helper
'  ((range addr), IN2.0) instead. AS.6-guarded, same discipline as
'  TestInterpreterObjectDispatch just above.
' ---------------------------------------------------------------------
Private Sub TestInterpreterForEach()
    Dim prior As Worksheet
    Set prior = ActiveSheet
    VlaEnsureSheet "VlaInterpForEachSheet"
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets("VlaInterpForEachSheet")
    ws.Activate
    ws.Range("A1").Value = 10
    ws.Range("A2").Value = 20
    ws.Range("A3").Value = 30

    Dim d As String
    Dim frame As Object
    On Error Resume Next
    Set frame = VLA_Interpreter.VlaInterpret( _
        "(begin (dim total Double) (set! total 0)" & _
        " (for-each (c (range ""A1:A3""))" & _
        "   (set! total (+ total (. c value)))))")
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0

    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
    prior.Activate

    If Len(d) > 0 Then
        Report "in2.3: for-each over a real range", False, d
        Exit Sub
    End If
    CheckV "in2.3: for-each over a real range sums each cell's value", _
           VLA_Runtime.VlaDictGet(frame, "total"), 60
End Sub

' ---------------------------------------------------------------------
'  IN2.5: named-argument dispatch, live - four of the eleven real
'  named-arg calls (VLA_Interpreter.bas's own IN2.5 header note has the
'  full census and reasoning), a representative sample rather than
'  exhaustive coverage of all eleven: a Worksheet receiver (protect/
'  unprotect, paired), a Range receiver with a Range-valued keyword
'  argument (copy :destination), and a Range receiver with two
'  keyword arguments including an Excel enum constant
'  (removeduplicates :columns :header). close/printout/
'  exportasfixedformat are deliberately NOT exercised here, for the
'  same reason msgbox/inputbox never are - closing the workbook,
'  opening a print preview, or writing a file mid-VlaSelfTest would
'  disrupt or hang the run, not just this one check. sort is
'  implemented (matches Range.Sort's real signature) but not pinned
'  either, purely for scope - four representative checks was judged
'  enough to trust the mechanism, the same "a few checks, not
'  exhaustive" precedent this session's other new pins already set.
'  A separate, pre-existing gap found while writing this, not
'  introduced by it: Excel's named enum constants (xlYes, xlAscending,
'  ...) are not resolved anywhere in the interpreter - EvalExpr's atom
'  fallback just does a frame lookup, and nothing binds these names
'  specially the way ResolveGlobalReceiver does for application/
'  activesheet/etc. Only the EMITTER's generated VBA resolves them,
'  at compile time, against the Excel type library. So this test uses
'  their literal numeric values (xlNo=2 for Header, matching what
'  removeduplicates actually needs here) rather than the symbolic
'  names english.vla's own macros write - named remainder, not
'  hidden, same as with/on-error/goto/label/exit-sub/exit-function.
' ---------------------------------------------------------------------
Private Sub TestInterpreterNamedArgs()
    Dim prior As Worksheet
    Set prior = ActiveSheet
    VlaEnsureSheet "VlaNamedArgSheet"
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets("VlaNamedArgSheet")
    ws.Activate

    Dim d As String

    ' protect/unprotect - Worksheet, via (worksheets "...") since a
    ' bare 'activesheet' receiver does not resolve through EvalExpr's
    ' atom path today (ResolveGlobalReceiver only fires for a DOTTED
    ' HEAD like activesheet.name, not a bare receiver expression - a
    ' separate, pre-existing gap this item does not touch either).
    d = ""
    On Error Resume Next
    VLA_Interpreter.VlaInterpret _
        "(begin (. (worksheets ""VlaNamedArgSheet"") protect :password ""pw123""))"
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    If Len(d) > 0 Then
        Report "in2.5: protect with a named password", False, d
    Else
        CheckV "in2.5: protect with a named password", ws.ProtectContents, True
    End If

    d = ""
    On Error Resume Next
    VLA_Interpreter.VlaInterpret _
        "(begin (. (worksheets ""VlaNamedArgSheet"") unprotect :password ""pw123""))"
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    If Len(d) > 0 Then
        Report "in2.5: unprotect with a named password", False, d
    Else
        CheckV "in2.5: unprotect with a named password", ws.ProtectContents, False
    End If

    ' copy :destination - Range, a Range-VALUED keyword argument, not
    ' just a scalar.
    ws.Range("A1").Value = "hello"
    ws.Range("B1").ClearContents
    d = ""
    On Error Resume Next
    VLA_Interpreter.VlaInterpret _
        "(begin (. (range ""A1"") copy :destination (range ""B1"")))"
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    If Len(d) > 0 Then
        Report "in2.5: copy with a named, range-valued destination", False, d
    Else
        CheckV "in2.5: copy with a named, range-valued destination", ws.Range("B1").Value, "hello"
    End If

    ' removeduplicates :columns :header - Range, two keyword arguments,
    ' one an Excel enum constant (xlNo = 2, no header row here).
    ws.Range("D1").Value = 1
    ws.Range("D2").Value = 2
    ws.Range("D3").Value = 2
    ws.Range("D4").Value = 3
    d = ""
    On Error Resume Next
    VLA_Interpreter.VlaInterpret _
        "(begin (. (range ""D1:D4"") removeduplicates :columns 1 :header 2))"
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    If Len(d) > 0 Then
        Report "in2.5: removeduplicates with named columns/header", False, d
    Else
        Report "in2.5: removeduplicates with named columns/header collapses the duplicate row", _
               ws.Range("D1").Value = 1 And ws.Range("D2").Value = 2 And ws.Range("D3").Value = 3, _
               "D1=" & ws.Range("D1").Value & " D2=" & ws.Range("D2").Value & " D3=" & ws.Range("D3").Value
    End If

    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
    prior.Activate
End Sub

' ---------------------------------------------------------------------
'  VerifyReport: the machine-checked end state of instructions.txt.
' ---------------------------------------------------------------------
Public Function VerifyReport() As Boolean
    mPass = 0
    mFail = 0
    Set mFailedNames = New Collection
    Debug.Print "===== VERIFY REPORT (instructions.txt end state) ====="
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets("Output")
    On Error GoTo 0
    If ws Is Nothing Then
        Report "emitter: Output sheet exists", False, "no Output sheet - Run the program first"
        Debug.Print "===== VERIFY: " & mPass & " passed, " & mFail & " failed ====="
        VerifyReport = False
        Exit Function
    End If
    ' VerifyReports stale-read fix (VLA_Runtime's own VlaStampRunProvenance
    ' header note has the full incident): this function has never run its
    ' own compile+Run - it only ever reads whatever is CURRENTLY on the
    ' Output/GStruct sheets, which VerifyReportInterpreter's own live run
    ' (or a manually-run Interpret and Trace) silently overwrites with
    ' INTERPRETER output. A whole debugging session was lost to reports
    ' that looked like real per-rule bugs but were actually just reading
    ' the wrong backend's leftover state, mislabeled "emitter." Refuse
    ' loudly instead of guessing which backend actually wrote this sheet.
    On Error Resume Next
    Dim provenance As String
    provenance = VLA_Runtime.VlaReadRunProvenance()
    On Error GoTo 0
    If provenance <> "emitter" Then
        Report "emitter: Output sheet was produced by a real Compile-and-Trace Run", False, _
               "last write was tagged '" & provenance & "' (empty means untagged/older workbook) - " & _
               "run Compile and Trace now, THEN call VerifyReport/VerifyReports with nothing else run in between"
        Debug.Print "===== VERIFY: " & mPass & " passed, " & mFail & " failed ====="
        VerifyReport = False
        Exit Function
    End If

    VerifyReportChecks ws
    Debug.Print "===== VERIFY: " & mPass & " passed, " & mFail & " failed ====="
    VerifyReport = (mFail = 0)
End Function

' IN.3: VerifyReport's own cell-by-cell checks, factored out unchanged so
' VerifyReportInterpreter (below) - the interpreter-side counterpart this
' item's own text always meant to build ("instructions.txt produces identical
' Output sheets under both backends") - can run the exact same checks
' against the exact same expected values, IN3.2's own two-sided precedent
' extended from four hand-picked statement forms to the whole corpus:
' each backend independently against a hand-computed expected value is
' what proves agreement, not a direct A-vs-B cell diff. mPass/mFail reset
' is the CALLER's job (both entry points already did this before the
' split), not this function's, so a caller could in principle chain
' checks across more than one sheet in one running total - not used that
' way today, but no reason to foreclose it for a `~one-line` cost.
Private Sub VerifyReportChecks(ws As Worksheet)
    CheckV "B2", ws.Range("B2").Value, 15
    CheckV "B3 (formula result)", ws.Range("B3").Value, 30
    CheckV "B4", ws.Range("B4").Value, 45
    CheckV "C4 text", ws.Range("C4").Value, "PASS"
    Report "C4 bold", ws.Range("C4").Font.Bold = True, "not bold"
    Report "C4 yellow", ws.Range("C4").Interior.Color = vbYellow, "wrong fill"
    CheckV "E1", ws.Range("E1").Value, 10
    CheckV "E2", ws.Range("E2").Value, 20
    CheckV "E3", ws.Range("E3").Value, 30
    CheckV "F1", ws.Range("F1").Value, "ok"
    Report "F1 bold (contains-test hit)", ws.Range("F1").Font.Bold = True, "not bold"
    CheckV "F2", ws.Range("F2").Value, "beta"
    Report "F2 not bold (contains-test miss)", ws.Range("F2").Font.Bold = False, "bolded wrongly"

    CheckV "A1 title", ws.Range("A1").Value, "Test Report"
    Report "A1 bold", ws.Range("A1").Font.Bold = True, "not bold"
    CheckV "A1 size", ws.Range("A1").Font.Size, 14
    Report "A1 hot-pink (Define alias + BGR)", ws.Range("A1").Font.Color = RGB(255, 105, 180), _
           "got " & ws.Range("A1").Font.Color
    Report "A1 centered", ws.Range("A1").HorizontalAlignment = xlCenter, "not centered"
    ' Column D: an explicit width the program never autofits away.
    ' (The original probe set column A, which Tidy-up's "Fit column
    ' A." then overrode - VerifyReport's first catch, on its first
    ' run: a set-then-autofit contradiction eyeballing never saw.)
    CheckV "column D width", ws.Columns(4).ColumnWidth, 24
    Report "A6 hot-pink", ws.Range("A6").Font.Color = RGB(255, 105, 180), "wrong color"
    Report "A6 nonempty", Len(Trim$(CStr(ws.Range("A6").Value))) > 0, "empty"

    ' The H probes: values the program computes and writes for us.
    CheckV "H1 grand (sum)", ws.Range("H1").Value, 45
    CheckV "H2 biggest (largest of)", ws.Range("H2").Value, 45
    Report "H3 rounded 3.14", Abs(ws.Range("H3").Value - 3.14) < 0.0001, "got " & ws.Range("H3").Value
    CheckV "H4 thousands 1500", ws.Range("H4").Value, 1500
    CheckV "H5 stopped at 3", ws.Range("H5").Value, 3
    CheckV "H6 last filled F row", ws.Range("H6").Value, 2
    CheckV "H7 list count (B1)", ws.Range("H7").Value, 2
    CheckV "H8 chain verdict (B2)", ws.Range("H8").Value, "solid"
    CheckV "H9 when-chain label (B2)", ws.Range("H9").Value, "warm"
    CheckV "H10 repeat-until passes (B2)", ws.Range("H10").Value, 3
    CheckV "H11 try rescued (B3)", ws.Range("H11").Value, "rescued"
    CheckV "H12 try success path (B3)", ws.Range("H12").Value, 7
    CheckV "H13 tax of 100 (B4)", ws.Range("H13").Value, 8
    CheckV "H14 condition used tax (B4)", ws.Range("H14").Value, "big"
    CheckV "H15 column-number prim (B5)", ws.Range("H15").Value, "yes"
    CheckV "H16 sheet-qualified bang (B5.2)", ws.Range("H16").Value, "bang"
    CheckV "H17 value-in synonym (B5.2)", ws.Range("H17").Value, "ok"
    CheckV "H18 spaced-sheet roundtrip (B5.3)", ws.Range("H18").Value, "spaced"
    CheckV "H19 using both named args (B6)", ws.Range("H19").Value, 200
    CheckV "H20 using a default (B6)", ws.Range("H20").Value, 30
    CheckV "H21 grow/increase-share chain (B7.4)", ws.Range("H21").Value, 165
    CheckV "H22 item n of a list (B7)", ws.Range("H22").Value, 400
    CheckV "H23 doubled quotes in text (B7)", ws.Range("H23").Value, "He said ""ok"""
    CheckV "H24 nullary value action with percent (B7.2)", ws.Range("H24").Value, 0.2
    CheckV "H25 raw VLA row wrote its probe (L0)", ws.Range("H25").Value, "vla-row"
    ' V3: keyed memory end-to-end - replace keeps the value fresh and
    ' the key's first spelling, keys are case-insensitive, walks are
    ' insertion-ordered, and reads compose in arithmetic and
    ' conditions.
    CheckV "H26 lookup read after replace + case-insensitive key (V3)", ws.Range("H26").Value, 120
    CheckV "H27 key count stays 2 after replace (V3)", ws.Range("H27").Value, 2
    CheckV "H28 keys walk in first-stored spelling and order (V3)", ws.Range("H28").Value, "ax-7bx-2"
    CheckV "H29 keyed reads chain in arithmetic (V3)", ws.Range("H29").Value, 370
    CheckV "H30 keyed read decides a condition (V3)", ws.Range("H30").Value, "steep"
    CheckV "H31 pair walk filters by value, collects keys (V3.1)", ws.Range("H31").Value, "bx-2"

    ' G-TABLES: real ListObject state, not just "the script didn't
    ' crash" - style compared case-insensitively (StrComp vbTextCompare)
    ' since English text slots lowercase in transit and it's untested
    ' whether Excel's own TableStyle getter normalizes the case back out.
    Report "SalesTable exists", Not (ws.Range("J1").ListObject Is Nothing), "no ListObject at J1"
    If Not (ws.Range("J1").ListObject Is Nothing) Then
        Report "SalesTable style", StrComp(CStr(ws.Range("J1").ListObject.TableStyle), "TableStyleMedium9", vbTextCompare) = 0, _
               "got " & ws.Range("J1").ListObject.TableStyle
        Report "SalesTable totals shown", ws.Range("J1").ListObject.ShowTotals = True, "totals not shown"
    End If
    If Not (ws.Range("J5").ListObject Is Nothing) Then
        Report "QuietTable totals hidden after show-then-hide", ws.Range("J5").ListObject.ShowTotals = False, "totals still shown"
    End If
    Report "TempTable converted back to a plain range", ws.Range("J8").ListObject Is Nothing, "still a ListObject"

    ' G-TABLES: table-add-row/table-delete-row/table-column, real state -
    ' EditTable starts with 2 data rows (Bolt/5, Nut/8); add makes 3,
    ' delete row 1 (Bolt) removes the FIRST row specifically, not just
    ' the count, so Nut/8 shifts up and the added blank row lands last.
    If Not (ws.Range("J11").ListObject Is Nothing) Then
        Dim editTbl As ListObject
        Set editTbl = ws.Range("J11").ListObject
        Report "EditTable row count after add-then-delete nets to 2", editTbl.ListRows.Count = 2, _
               "got " & editTbl.ListRows.Count
        Report "EditTable row 1 is Nut (Bolt row was the one actually deleted)", _
               CStr(editTbl.DataBodyRange.Cells(1, 1).Value) = "Nut", _
               "got " & editTbl.DataBodyRange.Cells(1, 1).Value
        Report "EditTable row 1 Qty still 8 after the shift", editTbl.DataBodyRange.Cells(1, 2).Value = 8, _
               "got " & editTbl.DataBodyRange.Cells(1, 2).Value
        Report "EditTable row 2 is the blank row table-add-row appended", _
               Len(CStr(editTbl.DataBodyRange.Cells(2, 1).Value)) = 0, _
               "got " & editTbl.DataBodyRange.Cells(2, 1).Value
    Else
        Report "EditTable exists", False, "no ListObject at J11"
    End If

    ' G-PIVOT: real PivotTable state - PivotTables(name) raises rather
    ' than returning Nothing for a miss, unlike Range.ListObject above,
    ' so the lookup itself needs the guard.
    Dim pvt As PivotTable
    On Error Resume Next
    Set pvt = ws.PivotTables("SalesPivot")
    On Error GoTo 0
    Report "SalesPivot exists", Not (pvt Is Nothing), "no PivotTable named SalesPivot"

    ' G-PIVOT (rows/columns/filters): real field-ORIENTATION state, not
    ' just "the sentence didn't crash" - proves VlaPivotSetOrientation
    ' (VLA_Runtime.bas) actually moved the right fields to the right
    ' pivot areas. PivotFields(name) lookup is case-insensitive against
    ' the source header text (Excel's own object-model convention, the
    ' same one Worksheets(name)/Range(name) already rely on elsewhere
    ' in this file) - instructions.txt's own sentences pass the field
    ' names as bare, tokenizer-lowercased words ("region", not
    ' "Region"), so this is the live proof that assumption holds, not
    ' just a traced one.
    If Not pvt Is Nothing Then
        Report "SalesPivot: Region is a row field", pvt.PivotFields("Region").Orientation = xlRowField, _
               "got " & pvt.PivotFields("Region").Orientation
        Report "SalesPivot: Product is a row field", pvt.PivotFields("Product").Orientation = xlRowField, _
               "got " & pvt.PivotFields("Product").Orientation
        Report "SalesPivot: Segment is a column field", pvt.PivotFields("Segment").Orientation = xlColumnField, _
               "got " & pvt.PivotFields("Segment").Orientation
        Report "SalesPivot: Channel is a filter field", pvt.PivotFields("Channel").Orientation = xlPageField, _
               "got " & pvt.PivotFields("Channel").Orientation

        ' G-PIVOT (pivot-values-sum/-count/-avg): real per-field
        ' aggregation state - Revenue deliberately appears in TWO data
        ' fields (summed AND counted), the live proof that
        ' VlaPivotAddValues's own no-guard-against-re-adding decision
        ' (owner's own explicit call) is genuinely safe, not just
        ' untested. Checked by (SourceName, Function) pair, since
        ' SourceName alone can't distinguish Revenue's two entries.
        Dim wantSumRevenue As Boolean, wantCountRevenue As Boolean
        Dim wantCountUnits As Boolean, wantAvgUnits As Boolean
        Dim dfVal As PivotField
        For Each dfVal In pvt.DataFields
            If dfVal.SourceName = "Revenue" And dfVal.Function = xlSum Then wantSumRevenue = True
            If dfVal.SourceName = "Revenue" And dfVal.Function = xlCount Then wantCountRevenue = True
            If dfVal.SourceName = "Units" And dfVal.Function = xlCount Then wantCountUnits = True
            If dfVal.SourceName = "Units" And dfVal.Function = xlAverage Then wantAvgUnits = True
        Next
        Report "SalesPivot: values are Sum of Revenue, Count of Revenue, Count of Units, Average of Units", _
               pvt.DataFields.Count = 4 And wantSumRevenue And wantCountRevenue And wantCountUnits And wantAvgUnits, _
               "got " & pvt.DataFields.Count & " data field(s)"
    End If

    ' G-PIVOT (pivot-full): the composite sentence's own second,
    ' independent pivot - real row/column/data-field state, including
    ' the genuine three-item Oxford-comma row list (Region, Product,
    ' Channel) instructions.txt now proves against a real sentence, not
    ' just TestG6's synthetic one.
    Dim pvtFull As PivotTable
    On Error Resume Next
    Set pvtFull = ws.PivotTables("FullPivot")
    On Error GoTo 0
    Report "FullPivot exists (pivot-full's own composite sentence)", Not (pvtFull Is Nothing), "no PivotTable named FullPivot"
    If Not pvtFull Is Nothing Then
        Report "FullPivot: Region is a row field", pvtFull.PivotFields("Region").Orientation = xlRowField, _
               "got " & pvtFull.PivotFields("Region").Orientation
        Report "FullPivot: Product is a row field", pvtFull.PivotFields("Product").Orientation = xlRowField, _
               "got " & pvtFull.PivotFields("Product").Orientation
        Report "FullPivot: Channel is a row field (three-item Oxford-comma list, not just two)", _
               pvtFull.PivotFields("Channel").Orientation = xlRowField, _
               "got " & pvtFull.PivotFields("Channel").Orientation
        Report "FullPivot: Segment is a column field", pvtFull.PivotFields("Segment").Orientation = xlColumnField, _
               "got " & pvtFull.PivotFields("Segment").Orientation

        ' VlaPivotAddValues's own v1 cap: sum-only, checked by
        ' SourceName (the underlying field) rather than the
        ' auto-generated ".Name" caption ("Sum of Revenue"), which is
        ' rendering, not identity.
        Dim dataOk As Boolean, sawRevenue As Boolean, sawUnits As Boolean
        dataOk = (pvtFull.DataFields.Count = 2)
        If dataOk Then
            Dim df As PivotField
            For Each df In pvtFull.DataFields
                If df.SourceName = "Revenue" Then
                    sawRevenue = True
                    If df.Function <> xlSum Then dataOk = False
                ElseIf df.SourceName = "Units" Then
                    sawUnits = True
                    If df.Function <> xlSum Then dataOk = False
                End If
            Next
            dataOk = dataOk And sawRevenue And sawUnits
        End If
        Report "FullPivot: values are Revenue and Units, both summed (two-item list)", dataOk, _
               "got " & pvtFull.DataFields.Count & " data field(s)"

        ' G-PIVOT (pivot-collapse/pivot-expand): real ShowDetail state,
        ' checked per PIVOT ITEM, not per field - PivotField.ShowDetail
        ' itself raises 1004 under Compact Form (the default row layout
        ' CreatePivotTable leaves in place), the exact live crash
        ' VlaPivotSetShowDetail's own header note now documents; every
        ' PivotItem's own ShowDetail is what the real UI's "+"/"-"
        ' control actually drives, and reads reliably in any layout.
        ' SalesPivot's own Region field ends COLLAPSED (every item) - a
        ' real, distinguishing change from Excel's own True default, not
        ' just "the sentence didn't crash". FullPivot's own Product
        ' field is collapsed THEN expanded in instructions.txt, ending
        ' True - proves the round trip runs end to end, though (named
        ' honestly, not overclaimed) the final value alone can't by
        ' itself distinguish "collapsed then correctly reversed" from
        ' "never touched", a real limit of checking end-state rather
        ' than a trace.
        Dim piChk As PivotItem
        Dim regionCollapsed As Boolean
        regionCollapsed = True
        For Each piChk In pvt.PivotFields("Region").PivotItems
            If piChk.ShowDetail <> False Then regionCollapsed = False
        Next
        Report "SalesPivot: Region is collapsed (every item)", regionCollapsed, _
               "at least one Region item is not collapsed"

        Dim productExpanded As Boolean
        productExpanded = True
        For Each piChk In pvtFull.PivotFields("Product").PivotItems
            If piChk.ShowDetail <> True Then productExpanded = False
        Next
        Report "FullPivot: Product is expanded after a collapse-then-expand round trip (every item)", _
               productExpanded, "at least one Product item is not expanded"

        ' G-PIVOT close-out (pivot-layout): FIRST VERSION checked
        ' PivotField.LayoutForm expecting it to mirror RowAxisLayout's
        ' own compact/tabular/outline state - wrong, caught live (both
        ' checks read back the same value regardless of which layout
        ' had actually been applied). Traced to the real cause after
        ' the live failure, not before it: LayoutForm is type
        ' XlLayoutFormType (Microsoft's own docs: values xlOutline/
        ' xlTabular ONLY, no compact option, defaults to xlTabular) - a
        ' different, older per-field property entirely, unrelated to
        ' XlLayoutRowType (RowAxisLayout's own parameter type). Per
        ' Microsoft's own RowAxisLayout reference, there is no
        ' object-model readback for XlLayoutRowType state at all - so
        ' pivot-layout's own live proof is necessarily thinner than
        ' every other pivot verb's, the same honestly-named limit
        ' pivot-refresh's own comment already accepted: the call
        ' succeeds without error, not a visible, checkable state
        ' change. What IS checkable and real: RowAxisLayout re-forms
        ' every row field in the table at once, so SalesPivot's own
        ' already-collapsed Region field has its ShowDetail state
        ' re-read fresh after two layout switches, rather than assumed
        ' to have survived untouched - the exact class of Compact-
        ' Form/per-item interaction that broke VlaPivotSetShowDetail's
        ' first version.
        Dim regionStillCollapsed As Boolean
        regionStillCollapsed = True
        For Each piChk In pvt.PivotFields("Region").PivotItems
            If piChk.ShowDetail <> False Then regionStillCollapsed = False
        Next
        Report "SalesPivot: Region ShowDetail survives a row-layout switch", regionStillCollapsed, _
               "at least one Region item is no longer collapsed after the layout switch"

        ' G-PIVOT close-out (pivot-subtotals-hide/-show): real
        ' PivotField.Subtotals state. Index 1 is "Automatic" (Sum) -
        ' the only index any verb here ever sets, matching Excel's own
        ' default (True) - so "hidden" means index 1 is False and
        ' "shown" means it's True again after a hide-then-show round
        ' trip on a different pivot/field, the same
        ' distinguishing-change-plus-round-trip pair pivot-collapse/
        ' -expand already established above.
        Report "SalesPivot: Region subtotals hidden", _
               pvt.PivotFields("Region").Subtotals(1) = False, "Region subtotal 1 is still True"
        Report "FullPivot: Product subtotals shown after hide-then-show round trip", _
               pvtFull.PivotFields("Product").Subtotals(1) = True, "Product subtotal 1 is False"

        ' G-PIVOT close-out (pivot-blank-line-add/-remove): real
        ' PivotField.LayoutBlankLine state, same distinguishing-change-
        ' plus-round-trip pair.
        Report "SalesPivot: Region has a blank row after each item", _
               pvt.PivotFields("Region").LayoutBlankLine = True, "LayoutBlankLine is False"
        Report "FullPivot: Product blank row removed after add-then-remove round trip", _
               pvtFull.PivotFields("Product").LayoutBlankLine = False, "LayoutBlankLine is still True"

        ' G-PIVOT round 2 (pivot-sort/pivot-sort-by-value): real
        ' PivotItem ORDER state, not just "the sentence didn't crash".
        ' instructions.txt sorts each field descending then ascending,
        ' so the FINAL state is always the ascending order - a genuine
        ' transition from the descending sort immediately before it,
        ' not "untouched" (the same discipline collapse/expand's own
        ' round trip already established). SalesPivot's Region:
        ' East/North/South ascending alphabetically puts East first.
        ' FullPivot's Product: Gadget's own Revenue total (200+300+250=
        ' 750) is less than Widget's (500+900+600=2000), so ascending
        ' by Revenue puts Gadget first - PivotItems(1) proves the VALUE
        ' actually drove the order, not alphabetical accident (Gadget
        ' also happens to sort alphabetically before Widget, but the
        ' MAGNITUDE relationship is the real thing being proved here).
        Report "SalesPivot: Region is ascending-sorted after descending-then-ascending (East first)", _
               pvt.PivotFields("Region").PivotItems(1).Name = "East", _
               "got '" & pvt.PivotFields("Region").PivotItems(1).Name & "'"
        Report "FullPivot: Product is ascending-sorted by Revenue after descending-then-ascending (Gadget first)", _
               pvtFull.PivotFields("Product").PivotItems(1).Name = "Gadget", _
               "got '" & pvtFull.PivotFields("Product").PivotItems(1).Name & "'"
    End If

    ' G-PIVOT round 2 (pivot-rename): RenameMePivot no longer resolves
    ' under its old name; RenamedPivot resolves under the new one - a
    ' real, positive pair of checks, not just "the call didn't crash".
    Dim pvtOldName As PivotTable, pvtNewName As PivotTable
    On Error Resume Next
    Set pvtOldName = ws.PivotTables("RenameMePivot")
    Set pvtNewName = ws.PivotTables("RenamedPivot")
    On Error GoTo 0
    Report "RenameMePivot no longer exists under its old name", pvtOldName Is Nothing, "old name still resolves"
    Report "RenamedPivot exists under its new name", Not (pvtNewName Is Nothing), "no PivotTable named RenamedPivot"

    ' G-PIVOT round 2 (pivot-clear): ClearMePivot itself survives
    ' (ClearTable empties a pivot, it doesn't delete it - pivot-delete's
    ' own job) but its one row field and one value field are both gone.
    Dim pvtClear As PivotTable
    On Error Resume Next
    Set pvtClear = ws.PivotTables("ClearMePivot")
    On Error GoTo 0
    Report "ClearMePivot still exists (ClearTable empties, it doesn't delete)", Not (pvtClear Is Nothing), _
           "no PivotTable named ClearMePivot"
    If Not pvtClear Is Nothing Then
        Report "ClearMePivot: Region is no longer a row field after Clear", _
               pvtClear.PivotFields("Region").Orientation = xlHidden, _
               "got " & pvtClear.PivotFields("Region").Orientation
        Report "ClearMePivot: no data fields remain after Clear", pvtClear.DataFields.Count = 0, _
               "got " & pvtClear.DataFields.Count & " data field(s)"
    End If

    ' G-PIVOT round 2 (pivot-remove-field): Product is hidden (the
    ' missing inverse of pivot-rows/-columns/-filters actually took
    ' effect) while Region, untouched, proves the removal was selective
    ' - not a wholesale reset (that's pivot-clear's own job, checked
    ' separately above).
    Dim pvtRemoveField As PivotTable
    On Error Resume Next
    Set pvtRemoveField = ws.PivotTables("RemoveFieldMePivot")
    On Error GoTo 0
    Report "RemoveFieldMePivot exists", Not (pvtRemoveField Is Nothing), "no PivotTable named RemoveFieldMePivot"
    If Not pvtRemoveField Is Nothing Then
        Report "RemoveFieldMePivot: Product removed (hidden) from the layout", _
               pvtRemoveField.PivotFields("Product").Orientation = xlHidden, _
               "got " & pvtRemoveField.PivotFields("Product").Orientation
        Report "RemoveFieldMePivot: Region untouched by removing Product (still a row field)", _
               pvtRemoveField.PivotFields("Region").Orientation = xlRowField, _
               "got " & pvtRemoveField.PivotFields("Region").Orientation
    End If

    ' G-PIVOT round 2 (pivot-source): SourceTestPivot was built from
    ' N1:S7 (missing the West row on purpose) then repointed to N1:S8 -
    ' West actually showing up as a Region PivotItem is the real proof
    ' that ChangePivotCache-plus-refresh pulled in genuinely new data,
    ' not just re-read the same range.
    Dim pvtSource As PivotTable
    On Error Resume Next
    Set pvtSource = ws.PivotTables("SourceTestPivot")
    On Error GoTo 0
    Report "SourceTestPivot exists", Not (pvtSource Is Nothing), "no PivotTable named SourceTestPivot"
    If Not pvtSource Is Nothing Then
        Dim sawWest As Boolean, piSrc As PivotItem
        For Each piSrc In pvtSource.PivotFields("Region").PivotItems
            If piSrc.Name = "West" Then sawWest = True
        Next
        Report "SourceTestPivot: Region includes West after changing the source to N1:S8", sawWest, _
               "West not found among Region's PivotItems"
    End If

    ' G-PIVOT (pivot-delete): TempPivot is a throwaway pivot built and
    ' immediately deleted (instructions.txt), the same "disposable
    ' object proves real removal" shape TempTable/table-to-range
    ' already uses above - PivotTables(name) raises rather than
    ' returning Nothing for a miss, same guard as every other pivot
    ' lookup in this file.
    Dim pvtTemp As PivotTable
    On Error Resume Next
    Set pvtTemp = ws.PivotTables("TempPivot")
    On Error GoTo 0
    Report "TempPivot deleted (pivot-delete removes the PivotTable, not just its cells)", _
           pvtTemp Is Nothing, "TempPivot still exists"

    ' G-STRUCT: pareto.txt sections 2-5, real end state on instructions.txt's
    ' own dedicated GStruct sheet (not Output - row/column insert-delete-
    ' group-move operations shift EVERYTHING below them on a sheet, and
    ' Output already carries the H-probe/G-TABLES/G-PIVOT absolute
    ' references above). A missing sheet (Run hasn't happened yet) fails
    ' one named check and skips the rest, same guard shape as SalesTable/
    ' SalesPivot above.
    Dim wsG As Worksheet
    On Error Resume Next
    Set wsG = ActiveWorkbook.Worksheets("GStruct")
    On Error GoTo 0
    Report "GStruct sheet exists", Not (wsG Is Nothing), "no GStruct sheet - Run the program first"
    If Not wsG Is Nothing Then
        Report "unhide-all: row 3 no longer hidden", wsG.Rows(3).Hidden = False, "still hidden"
        Report "unhide-all: column B no longer hidden", wsG.Columns(2).Hidden = False, "still hidden"

        Report "fit-row: row 6 grew past the default height for 36pt text", _
               wsG.Rows(6).RowHeight > 20, "got " & wsG.Rows(6).RowHeight

        Report "group-rows: rows 10-12 are one outline level deep", wsG.Rows(10).OutlineLevel = 2, _
               "got " & wsG.Rows(10).OutlineLevel
        Report "ungroup-rows: rows 14-16 are back to no outline level (a real group-then-ungroup, not never-grouped)", _
               wsG.Rows(14).OutlineLevel = 1, "got " & wsG.Rows(14).OutlineLevel

        Report "insert-n-rows-at: 3 blank rows pushed the old row 20 down to row 23", _
               CStr(wsG.Range("A23").Value) = "before-insert", "got " & wsG.Range("A23").Value
        Report "insert-n-rows-at: the new row 20 is blank", Len(CStr(wsG.Range("A20").Value)) = 0, _
               "got " & wsG.Range("A20").Value

        Report "delete-rows-range: deleting rows 38-39 pulled row 40 up to row 38", _
               CStr(wsG.Range("A38").Value) = "before-delete", "got " & wsG.Range("A38").Value

        ' move-column: "Move column B before column D" cuts B, then closes
        ' the gap (C shifts left into B) and inserts B's own content
        ' immediately before D's ORIGINAL position - D itself never moves.
        ' Final order by content: A, [old C], [old B], [old D] - i.e. old
        ' B's value lands in column C, old C's value lands in column B,
        ' and old D's value stays in column D. Reasoned through by hand,
        ' not live-verified before this pass - the one check in this
        ' whole section worth double-checking first if it fails.
        Report "move-column: old column C's value shifted left into B", _
               CStr(wsG.Range("B50").Value) = "marker-c", "got " & wsG.Range("B50").Value
        Report "move-column: old column B's value landed in C, immediately before D", _
               CStr(wsG.Range("C50").Value) = "marker-b", "got " & wsG.Range("C50").Value
        Report "move-column: column D's own value did not move", _
               CStr(wsG.Range("D50").Value) = "marker-d", "got " & wsG.Range("D50").Value

        ' freeze-panes: a window property, not a range/sheet property - it
        ' only reflects whichever sheet is CURRENTLY displayed, and by now
        ' Output is active again (instructions.txt's own final "Go to
        ' sheet Output."), so GStruct has to be reactivated to read its
        ' own frozen-pane state, then handed back the way it was found.
        Dim priorActive As Worksheet
        Set priorActive = ActiveSheet
        wsG.Activate
        Report "freeze-panes: the window is split", ActiveWindow.FreezePanes = True, "not frozen"
        Report "freeze-panes: split after row 2 (""the first 2 rows"")", ActiveWindow.SplitRow = 2, _
               "got " & ActiveWindow.SplitRow
        priorActive.Activate

        Report "copy-formats: bold copied onto C60", wsG.Range("C60").Font.Bold = True, "not bold"
        Report "copy-formats: fill color copied onto C60", wsG.Range("C60").Interior.Color = vbRed, _
               "got " & wsG.Range("C60").Interior.Color
        Report "copy-style (""look like""): bold copied onto C62", wsG.Range("C62").Font.Bold = True, "not bold"
        Report "copy-style (""look like""): fill color copied onto C62", wsG.Range("C62").Interior.Color = vbRed, _
               "got " & wsG.Range("C62").Interior.Color

        Report "copy-formulas: D64 evaluates the copied formula", wsG.Range("D64").Value = 10, _
               "got " & wsG.Range("D64").Value

        CheckV "copy-widths: column H matches column F's width", wsG.Columns(8).ColumnWidth, wsG.Columns(6).ColumnWidth

        Report "paste-transposed: column became a row (C68)", wsG.Range("C68").Value = 1, "got " & wsG.Range("C68").Value
        Report "paste-transposed: column became a row (D68)", wsG.Range("D68").Value = 2, "got " & wsG.Range("D68").Value
        Report "paste-transposed: column became a row (E68)", wsG.Range("E68").Value = 3, "got " & wsG.Range("E68").Value

        Report "cut-range: value arrived at the destination", CStr(wsG.Range("C72").Value) = "cutme", _
               "got " & wsG.Range("C72").Value
        Report "cut-range: source is empty (a real move, not a copy)", Len(CStr(wsG.Range("A72").Value)) = 0, _
               "got " & wsG.Range("A72").Value

        Report "duplicate-row: row 76 got row 74's value", CStr(wsG.Range("A76").Value) = "rowdata", _
               "got " & wsG.Range("A76").Value

        Report "clear-all: value cleared", Len(CStr(wsG.Range("A80").Value)) = 0, "got " & wsG.Range("A80").Value
        Report "clear-all: formatting cleared too (not just contents)", wsG.Range("A80").Font.Bold = False, _
               "still bold"

        Report "delete-shift-up: row 87's value shifted up to row 84", CStr(wsG.Range("A84").Value) = "m4", _
               "got " & wsG.Range("A84").Value
        Report "delete-shift-up: row 88's value shifted up to row 85", CStr(wsG.Range("A85").Value) = "m5", _
               "got " & wsG.Range("A85").Value

        Report "delete-shift-left: column E's value shifted left to column B", _
               CStr(wsG.Range("B90").Value) = "rightdata", "got " & wsG.Range("B90").Value

        Report "delete-blank-rows: row 94 (no blanks) survived untouched", wsG.Range("A94").Value = 1, _
               "got " & wsG.Range("A94").Value
        Report "delete-blank-rows: the blank row was removed, pulling row 96 up to row 95", _
               wsG.Range("A95").Value = 5, "got " & wsG.Range("A95").Value

        Report "banded-rows: the range's own 2nd row is shaded", wsG.Range("A151").Interior.Color = RGB(217, 217, 217), _
               "got " & wsG.Range("A151").Interior.Color
        Report "banded-rows: the range's own 1st row is NOT shaded", wsG.Range("A150").Interior.Color <> RGB(217, 217, 217), _
               "row 1 got shaded too"
        Report "banded-rows: the range's own 3rd row is NOT shaded (every OTHER row)", _
               wsG.Range("A152").Interior.Color <> RGB(217, 217, 217), "row 3 got shaded too"

        Report "header-style: bold", wsG.Range("A165").Font.Bold = True, "not bold"
        Report "header-style: light fill", wsG.Range("A165").Interior.Color = RGB(217, 217, 217), _
               "got " & wsG.Range("A165").Interior.Color

        CheckV "fill-series: linear start", wsG.Range("A170").Value, 1
        CheckV "fill-series: linear end (10 cells, step 1)", wsG.Range("A179").Value, 10
        CheckV "fill-series: linear-with-step start", wsG.Range("A180").Value, 1
        CheckV "fill-series: linear-with-step end (5 cells, step 2)", wsG.Range("A184").Value, 9
        CheckV "fill-series: growth start", wsG.Range("A190").Value, 2
        CheckV "fill-series: growth end (5 cells, doubling)", wsG.Range("A194").Value, 32
        CheckV "fill-series: growth-with-step start", wsG.Range("A200").Value, 2
        CheckV "fill-series: growth-with-step end (4 cells, x3 each)", wsG.Range("A203").Value, 54

        ' trim-sheet: Z500 was touched (bold) then emptied, exactly the
        ' "formatted, then cleared" case UsedRange itself gets wrong -
        ' bounds checked loosely (well under the real Z500/row-500 extent)
        ' rather than pinned to an exact row/column, since the true last
        ' cell depends on which of this section's own ranges is rightmost/
        ' lowest, not on anything trim-sheet itself should be judged by.
        Report "trim-sheet: used range no longer reaches row 500", _
               wsG.UsedRange.Row + wsG.UsedRange.Rows.Count - 1 < 300, _
               "got last row " & (wsG.UsedRange.Row + wsG.UsedRange.Rows.Count - 1)
        Report "trim-sheet: used range no longer reaches column Z", _
               wsG.UsedRange.Column + wsG.UsedRange.Columns.Count - 1 < 15, _
               "got last column " & (wsG.UsedRange.Column + wsG.UsedRange.Columns.Count - 1)
    End If
End Sub

' ---------------------------------------------------------------------
'  IN.3: VerifyReportInterpreter - the interpreter-side counterpart to
'  VerifyReport, and the instrument IN.3's own opening line always meant:
'  "instructions.txt produces identical Output sheets under both backends."
'  Unlike VerifyReport, this does not need a prior manual Run - VlaInterpret
'  IS the run (IN2.0's real Range/member dispatch, IN.10's real call
'  stack, and this pass's own module-scope fix all feed a live Output
'  sheet the same way a real emitter Run would), so this function compiles
'  and executes instructions.txt itself before checking it. NOT dispatched
'  from VlaSelfTestHost, on purpose, mirroring VerifyReport's own
'  reasoning for staying manual: both mutate the real "Output" sheet in
'  whatever workbook is active, and folding this into the automatic host
'  suite would make an ordinary VlaSelfTestHost run silently clobber
'  whatever a manual emitter Run had just written there, breaking the
'  documented "Run the program, then VerifyReport" workflow out from
'  under it. AS.6-guarded throughout (TestExprParity's own precedent):
'  a raise reports FAIL under this pin's own name, with the interpreter's
'  own effect-log tail attached (EffectLogTail, below) so a failure names
'  where it happened, not just what happened - not crash the caller.
' ---------------------------------------------------------------------

' IN.11: the last N lines of VlaInterpreterEffectLog's own text, newest
' last - a raise-time diagnostic, not a golden: VlaWriteInterpreterGoldens
' (VLA_Tests.bas) already owns the full, exact-text, committed record of
' this log; this is a throwaway tail read back out of a run that DIED,
' so there is nothing to commit, only something to read once, right
' after the failure that needed it.
Private Function EffectLogTail(ByVal fullLog As String, ByVal n As Long) As String
    Dim allLines() As String
    allLines = Split(fullLog, vbCrLf)
    Dim lo As Long
    lo = UBound(allLines) - n + 1
    If lo < LBound(allLines) Then lo = LBound(allLines)
    Dim i As Long
    Dim r As String
    For i = lo To UBound(allLines)
        If Len(allLines(i)) > 0 Then
            If Len(r) > 0 Then r = r & " | "
            r = r & allLines(i)
        End If
    Next
    EffectLogTail = r
End Function

Public Function VerifyReportInterpreter() As Boolean
    mPass = 0
    mFail = 0
    Set mFailedNames = New Collection
    Debug.Print "===== VERIFY REPORT, INTERPRETER BACKEND (instructions.txt end state) ====="

    Dim vocabPath As String, progPath As String
    vocabPath = VLA_Tests.FindDevFile("english.vla")
    progPath = VLA_Tests.FindDevFile("instructions.txt")

    Dim d As String
    On Error Resume Next
    Dim englishText As String
    englishText = ReadTextFileUtf8(progPath)
    EnglishResetGrammar
    EnglishLoadVocabulary vocabPath
    ' IN.3: every interpreter-targeted compile so far sidesteps the
    ' vla-step/vlatraceon/at-line harness BuildSub wraps every procedure
    ' in when step tracking is on (ParseTracked's own gate - a pure
    ' passthrough when mStepTracking is False) - same convention, applied
    ' here to instructions.txt's own compile for the first time.
    EnglishStepTracking False
    Dim vla As String
    vla = EnglishToVla(englishText)
    EnglishStepTracking True
    If Err.Number <> 0 Then d = "compile: " & Err.Description
    On Error GoTo 0
    If Len(d) > 0 Then
        Report "in.3: instructions.txt compiles for the interpreter (step tracking off)", False, d
        Debug.Print "===== VERIFY: " & mPass & " passed, " & mFail & " failed ====="
        VerifyReportInterpreter = False
        Exit Function
    End If

    ' Never coast on a PRIOR run's leftover cells (IN3.2's own "cleared
    ' before each run" discipline, widened from one probe cell to the
    ' whole sheet): delete any existing Output sheet first, so every
    ' checked cell below was written by THIS interpreter run or not at
    ' all - "Work on sheet Output." (instructions.txt's own first real
    ' sentence) recreates it, the same create-if-missing path a genuinely
    ' fresh workbook already exercises.
    Dim prior As Worksheet
    Set prior = ActiveSheet
    Dim existing As Worksheet
    On Error Resume Next
    Set existing = ActiveWorkbook.Worksheets("Output")
    On Error GoTo 0
    If Not existing Is Nothing Then
        Application.DisplayAlerts = False
        existing.Delete
        Application.DisplayAlerts = True
    End If

    d = ""
    On Error Resume Next
    VLA_Interpreter.VlaInterpret vla
    If Err.Number <> 0 Then
        ' IN.11: a raise here used to name only the error, not where it
        ' happened - two straight guessing rounds against the breadth
        ' pass (docs/BETA_ROADMAP.md's own IN.11 entry has the history) cost
        ' real turnaround time that the interpreter's own effect log
        ' (LogEffect, VLA_Interpreter.bas) could have closed immediately:
        ' it already records every real effect in order, so its last few
        ' entries name the exact statement that was running when this
        ' run died, without needing instructions.txt's own line numbers at
        ' all. Appended here instead of guessed at from outside.
        d = "run: " & Err.Description & " -- last effects: " & _
            EffectLogTail(VLA_Interpreter.VlaInterpreterEffectLog(), 5)
    End If
    On Error GoTo 0
    If Len(d) > 0 Then
        Report "in.3: instructions.txt runs to completion under the interpreter", False, d
        On Error Resume Next
        prior.Activate
        On Error GoTo 0
        Debug.Print "===== VERIFY: " & mPass & " passed, " & mFail & " failed ====="
        VerifyReportInterpreter = False
        Exit Function
    End If
    ' VerifyReports stale-read fix (VLA_Runtime's own header note on
    ' VlaStampRunProvenance has the incident) - stamped as soon as the
    ' interpret itself succeeded, independent of whether the checks
    ' below pass, since it reflects who wrote the sheet, not who
    ' verified it.
    On Error Resume Next
    VLA_Runtime.VlaStampRunProvenance "interpreter"
    On Error GoTo 0

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets("Output")
    prior.Activate
    On Error GoTo 0
    If ws Is Nothing Then
        Report "in.3: the interpreter run left an Output sheet behind", False, "no Output sheet after VlaInterpret"
        Debug.Print "===== VERIFY: " & mPass & " passed, " & mFail & " failed ====="
        VerifyReportInterpreter = False
        Exit Function
    End If

    VerifyReportChecks ws
    Debug.Print "===== VERIFY: " & mPass & " passed, " & mFail & " failed ====="
    VerifyReportInterpreter = (mFail = 0)
End Function

' IN.3: convenience, VlaSelfTests's own shape (VLA_Tests_Host.bas:220) -
' both backends' end-state checks in one call, each printing its own
' summary plus one combined verdict line. ORDER STILL MATTERS - VerifyReport
' must run FIRST, while the Output sheet still holds whatever a prior
' manual emitter Run left behind, because VerifyReportInterpreter deletes
' and rebuilds that same sheet as part of its own live run - but a
' violation is no longer SILENT: VlaReadRunProvenance (VLA_Runtime.bas)
' now makes VerifyReport refuse loudly if the sheet wasn't actually
' produced by a real emitter Run, instead of quietly checking the
' interpreter's leftover cells and reporting them as emitter failures
' (the incident VlaStampRunProvenance's own header note documents - a
' full debugging session spent chasing per-rule bugs that were actually
' one stale-read bug in this harness). Does not skip VerifyReportInterpreter
' when VerifyReport fails (e.g. no prior Run at all, so no Output sheet
' yet) - the two are otherwise independent, same reasoning VlaSelfTests
' already applies to its own pure/host pair.
Public Function VerifyReports() As Boolean
    Dim emitterOk As Boolean, interpreterOk As Boolean
    Dim emitterPass As Long, emitterFail As Long, emitterFailed As Collection
    Dim interpPass As Long, interpFail As Long, interpFailed As Collection

    emitterOk = VerifyReport()
    emitterPass = mPass : emitterFail = mFail : Set emitterFailed = mFailedNames

    interpreterOk = VerifyReportInterpreter()
    interpPass = mPass : interpFail = mFail : Set interpFailed = mFailedNames

    Debug.Print "===== VERIFY REPORTS: " & IIf(emitterOk, "emitter PASS", "emitter FAIL") & _
                " (" & emitterPass & "/" & (emitterPass + emitterFail) & ")" & _
                ", " & IIf(interpreterOk, "interpreter PASS", "interpreter FAIL") & _
                " (" & interpPass & "/" & (interpPass + interpFail) & ") ====="
    Dim fn As Variant
    If emitterFail > 0 Then
        Debug.Print "----- emitter failures (" & emitterFailed.Count & ") -----"
        For Each fn In emitterFailed
            Debug.Print "  FAIL  " & fn
        Next fn
    End If
    If interpFail > 0 Then
        Debug.Print "----- interpreter failures (" & interpFailed.Count & ") -----"
        For Each fn In interpFailed
            Debug.Print "  FAIL  " & fn
        Next fn
    End If
    VerifyReports = emitterOk And interpreterOk
End Function

' Rule-12 duplicate of VLA_Tests's own ReadTextFileUtf8 (itself a
' Rule-12 duplicate of VLA_English's private VocabReadFile) - modules
' stay self-contained, same reasoning, same precedent, not called
' cross-module for a Private helper.
Private Function ReadTextFileUtf8(ByVal filePath As String) As String
    On Error GoTo ansiFallback
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2                              ' adTypeText
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile filePath
    ReadTextFileUtf8 = st.ReadText(-1)
    st.Close
    Exit Function

ansiFallback:
    On Error GoTo 0
    Dim f As Integer
    Dim b() As Byte
    f = FreeFile
    Open filePath For Binary Access Read As #f
    If LOF(f) = 0 Then
        Close #f
        ReadTextFileUtf8 = ""
        Exit Function
    End If
    ReDim b(0 To LOF(f) - 1)
    Get #f, , b
    Close #f
    ReadTextFileUtf8 = StrConv(b, vbUnicode)
End Function

' ---------------------------------------------------------------------
'  Assertion plumbing (local copy - R7: this module's whole reason for
'  existing is to be runnable, and reportable, independently of
'  VLA_Tests.bas).
' ---------------------------------------------------------------------
Private Sub CheckV(ByVal name As String, ByVal got As Variant, ByVal want As Variant)
    Report name, CStr(got) = CStr(want), "got '" & got & "', wanted '" & want & "'"
End Sub

Private Sub Report(ByVal name As String, ByVal ok As Boolean, ByVal detail As String)
    If ok Then
        mPass = mPass + 1
        Debug.Print "  PASS  " & name
    Else
        mFail = mFail + 1
        Debug.Print "  FAIL  " & name & " - " & detail
        If Not mFailedNames Is Nothing Then mFailedNames.Add name & " - " & detail
    End If
End Sub
