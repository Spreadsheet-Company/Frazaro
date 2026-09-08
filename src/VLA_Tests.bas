Attribute VB_Name = "VLA_Tests"
Option Explicit
Public Const VLA_TESTS_VERSION As String = "AS7.1"
' AS7.1: VlaSelfTestLastResult - a Public accessor exposing this
' module's own Private mPass/mFail/mFailedNames to VlaSelfTests
' (VLA_Tests_Host.bas), which cannot read them directly (cross-module
' Private). Read immediately after VlaSelfTest() returns - a snapshot
' of that one call, not live. Built so VlaSelfTests' own combined
' summary can print full counts and failure text for BOTH halves
' together at the very end (previously: pure's own failure list only
' ever appeared once, at VlaSelfTest's own tail, long since scrolled
' past the Immediate pane's buffer by the time the combined verdict
' line printed) - the owner's own request, prompted while scoping AS.7.
' COND.0: TestCond - pins the tenth engine primitive (VLA.bas's own
' COND.0, VLA_CORE_VERSION) - expand-time folding with the untaken
' branch never touched, the standalone compound-test case that broke
' the original naive defmacro attempt, runtime deferral into a native
' If/Else, a later folded-true clause hoisted into else rather than
' re-wrapped, the no-match-no-else no-op, a misplaced else, a malformed
' clause, reserved-name refusal, and nesting inside an ordinary
' defmacro. Wired into VlaSelfTest right after TestListopsExpand.
' LISTOPSEXPAND.0: TestListopsExpand - pins the seven new expand-time
' arithmetic/comparison primitives (VLA.bas's own LISTOPSEXPAND.0,
' VLA_CORE_VERSION) - arity, the strict non-numeric-operand gate,
' IsNumericLiteralText's own narrower scope vs. VBA's IsNumeric (a
' thousands-separator literal explicitly rejected), reserved-name
' refusal, and both recognition sites (standalone and nested inside a
' defmacro). Wired into VlaSelfTest right after TestListops.
' G6.0: TestArrayPrimitive - pins the new `array` runtime primitive
' (VLA.bas's EmitExpr, VLA_Interpreter.bas's EvalExpr) directly against
' both backends, structurally on the interpreter side (CStr on a VBA
' array raises Type mismatch, so CheckV can't be used there). TestG6
' itself (list-valued slots, the grammar-tier half) lives in
' VLA_Tests_Grammar.bas, dispatched right after TestG11r.
' TABLESPECSCALE.0: TestListopsBudget/TestListopsDepthSafety/
' TestTablespecDepthSafety extended with 2000/5500-row cases to prove
' TABLESPEC-SCALE's ExpandMacros trampoline (VLA.bas) at genuinely large
' scale - 250/900-row cases that used to fail past the old depth>200
' ceiling now expand cleanly; new 5500-row cases prove the trampoline's
' own new, purely-logical 5000 chainLen cap still fires loudly. New
' mRunScaleTests flag (off by default) and VlaSelfTestScale wrapper gate
' the cdr/ListTail- and GetMacro-registry-cost cases out of the ordinary
' VlaSelfTest run - a separate, pre-existing O(n^2) wall-clock cost the
' trampoline doesn't touch (it only removes VBA-stack cost), confirmed
' live at several minutes for a 5500-row walk. Run VlaSelfTestScale
' instead of VlaSelfTest when a change touches ExpandMacros/cdr/
' ListTail/GetMacro/tablespec's own walk and needs re-proving at scale.
' IN2.7: TestInterpreterOperators - pure pins for VLA_Interpreter.bas's
' own new EvalOpChain (18 operators AS.8's coverage scan found had no
' interpreter case at all - see VLA_Interpreter.bas's own IN2.7 header
' note for the fold-vs-chaining semantics this pins). Dispatched right
' after TestInterpreterExcelConstants.
' IN2.6: TestInterpreterExcelConstants - pure pins for
' VLA_Interpreter.bas's own new ResolveExcelConstant (one representative
' name per enum family the live corpus measured, plus a case-fold check
' and an ordinary-variable regression check). Dispatched right after
' TestInterpreterDynamicHead.
' IN2.3: TestInterpreterControlFlow - pure pins for IN.2's for-each/
' do-until/select/exit-for/exit-do (VLA_Interpreter.bas's own IN2.3
' header note has the full reasoning for what shipped and what was
' deliberately left out). do-until basic, exit-do fired from inside a
' nested if (proving mLoopBreak propagates through ExecBody/ExecIf,
' not just a loop's own direct child), select matching a multi-value
' case, select falling to case-else. Dispatched right after
' TestInterpreterDynamicHead.
' GROWLOOP.0: dispatch TestGRowLoop (VLA_Tests_Grammar.bas) right after
' TestG8 - G-ROWLOOP's step clause on the existing "Count" loop.
' IN2.0: TestInterpreterDynamicHead - pins IN.2's bounded-builtin tier
' (len/trim/ucase/left/round/instr - real VBA global functions with no
' object receiver) and its VLA_Runtime-helper fallback tier (vlacolor,
' via Application.Run against the module by name) plus obj-set! to a
' plain variable. None of these touch worksheet/range state, so this
' pin stays in the pure suite - VLA_Interpreter.bas's own IN2.0 header
' note has the full four-tier design. msgbox/inputbox are deliberately
' NOT exercised here (a modal mid-VlaSelfTest would hang the run
' untrappably - see TryEvalBuiltin's own note); the object-place/
' dot-dispatch/computed-set! tiers need a real worksheet and are
' pinned in VLA_Tests_Host.bas's TestInterpreterObjectDispatch instead.
' IN3_5.1: VlaGoldens - a convenience wrapper (owner request, same
' spirit as VLA_Tests_Host.bas's VlaSelfTests) that runs VlaWriteGoldens
' then VlaWriteInterpreterGoldens in one call. Both writers still run
' (and print) even if one fails, so a broken golden never hides the
' other's result.
' IN3_5.0: VlaWriteInterpreterGoldens - the interpreter's own
' VlaWriteGoldens (IN.3.5), scoped to what VLA_Interpreter.bas can run
' today (the walking skeleton's own demo program via the new
' VLA_Interpreter.VlaSkeletonDemoVla, not instructions.txt - IN.2 is what
' widens the corpus this can cover). Writes scripts/interpreter_golden.txt,
' same "commit deliberately, empty git diff is the pass" discipline as
' the emitter's own goldens. New pin TestInterpreterEffectLog, dispatched
' right after TestInterpreterSkeleton, proves the log's content on a
' small hand-built program independent of the demo/golden.
' N1.0: TestHelpers - VlaCheckSheetName / VlaCheckRangeName pins
' (accepts an ordinary name, rejects empty/forbidden-char/too-long/
' reserved-word for sheets, rejects hyphen/cell-shaped/leading-digit
' for range names, and checks the hyphen refusal's message actually
' offers "demo_table" as the fix).
' G8.0: dispatch TestG8 (VLA_Tests_Grammar.bas) right after TestG7 -
' number words and ordinals.
' G7.0: dispatch TestG7 (VLA_Tests_Grammar.bas) right after TestG11r -
' slot defaults, {n:expr=1}.
' G11r.0: dispatch TestG11r (VLA_Tests_Grammar.bas) right after TestG11 -
' the template-face remainder (wrap-text/unwrap-text, hide-row/unhide-row,
' hide-column/unhide-column, format-as-currency/percent/date) G11 itself
' left as an open remainder.
' IN5.0: TestHeadTableExportOnly - pins IN.5's adjudicated export-only
' set (raw, deflambda, lambda) by exact symbol list, and the nine rows
' corrected from the draft's True (sub/function/type/enum/public/
' private/include/at-line/doc) by name, not just by count. Dispatched
' right after TestHeadTable.
' F5.0: TestContextPushPop + TestContextUnderflow - VLA.bas's new
' VlaPushContext/VlaPopContext (F.5) round-trip real compile state
' (observed through VlaMacroDoc, not just the depth counter) and the
' underflow guard raises instead of silently no-opping. Both dispatched
' from VlaSelfTest, registered at the end of the list.
' IN1.0: TestHeadTable - sanity checks on the new VLA_HeadTable catalog
' (IN.1). Not behavior pins; VLA.bas's emitter is unchanged.
' LX6.0: TestNonAsciiIdentifiers - an accented sub name transliterates
' and emits correctly; an entirely unmappable name (a different script)
' refuses in words instead of reaching generated VBA source untouched.
' F1.0: VlaDotCount + TestDotCount - the published scoreboard for
' english.vla's dot count, pinned at 0.
' F1.1: TestMacroCallKeepsSourceTag, reverted in F1.3 - it pinned a
' VLA.bas change (tagging a macro call's single-statement expansion)
' that broke two pre-existing pins asserting the opposite, deliberate
' behavior. See VLA_CORE_VERSION's F1.1/F1.2 notes for the full account.
' F1.2: VlaSelfTest now collects every FAIL's name+detail (mFailedNames)
' and reprints them as a block after the summary line, so a failure
' near the top of the suite survives the Immediate Window's finite
' scrollback instead of scrolling off before the run finishes. Kept -
' unaffected by F1.3's revert, and how the two broken pins below were
' found in the first place.
' F1.3: TestMacroCallKeepsSourceTag removed - it pinned incorrect
' behavior (see F1.1 above). Nothing replaces it; the two pre-existing
' pins it should have been checked against already cover this.
' F1.4: F.1's residual closed - the seven application.worksheetfunction.*
' calls (sum/average/max/min/vlookup/sumif/countif) left out of the F1.0
' pass are now named macros too. VlaDotCount extended to also count
' "application.worksheetfunction." outside a macro: line, so this
' category is guarded the same way "(. " already is; the pinned count
' stays 0, now over both categories.
' IO6.0: TestIdeNaming's two module-name pins updated for the
' generated-module rename (EN_Sheet/EN_<tag> -> Frazaro_EN_Sheet/
' Frazaro_EN_<tag>) - a declared contract change, not a surprise diff.
' IN0.5: TestInterpreterSkeleton - one deterministic pin for the new
' walking-skeleton interpreter (VLA_Interpreter.bas), exercising all
' ten forms it supports in a single hand-written VLA snippet.
' LX4.0: TestHeadTable extended with VlaHeadTableRow's alias-scan path;
' TestHeadAlias (new) proves the actual chokepoint at the emitter -
' VlaTranspile of set!/if/debug-print and their new aliases
' (fijar!/si/depurar) emit byte-identical VBA.
' F9.0: VlaWriteGoldens now splices stable section markers into both
' goldens (InsertSectionMarkers, EnglishParagraphs and small helpers) -
' one (raw "' ---- instructions.txt:N label ----") line per blank-line-
' delimited paragraph, so a diff's surrounding context names the
' section instead of a bare vla:N tag. TestSectionMarkers (new) pins
' the boundary detection and escaping directly, on a hand-built
' sample, not the real corpus - this test's expectations do not drift
' when a corpus sentence is added or reworded.
' F8.0: this file crossed 4,945 lines - REBUILD.md R4's measured
' trigger for this exact split (F.8). The grammar/phrasebook feature
' pins (the G*/L*/V*/PL* series plus TestAlonzoLib, ~2,500 lines, the
' largest and fastest-growing concern in the file) moved to a new
' module, VLA_Tests_Grammar.bas. VlaSelfTest's dispatch list is
' unchanged - VBA calls a Public Sub in another standard module by
' bare name, no qualification needed - but every moved sub changed
' Private->Public (VBA cannot call a Private Sub across modules), and
' twelve shared assertion/report helpers below (Report, CheckFrags,
' CheckV, Norm, AssertVla, AssertEnglish, TryTranspile, TryEnglish,
' AssertErrLine, AssertClaim, AssertClaimRefusal, CountOcc) plus
' FindDevFile changed Private->Public for the same reason, since the
' new module calls them. Nothing else moved or changed; this pass is
' a pure file split, verified by an unchanged VlaSelfTest pass/fail
' count and an empty golden diff.
' F11.0: split again, this time by what a test NEEDS rather than by
' concern - F.11, "promoted to first on Theory-of-Constraints grounds"
' (the roadmap's own words: a human at a Windows box is the velocity
' ceiling and the bus factor). Everything needing a live workbook
' moved to a new module, VLA_Tests_Host.bas: VerifyReport (whole
' function, unchanged in behavior); TestRuntimeModule + DevModuleText
' (whole subs - VBProject/CodeModule reads); and two extracted
' fragments that used to be the host-dependent TAIL of an otherwise-
' pure test - TestHelpers lost its scratch-sheet section (kept its
' VlaColor/VlaDict-family assertions, which need no host), and
' TestRuntimeTrace lost its workbook-Name-persistence section (kept
' its in-memory VlaTrace/VlaTraceStep/VlaTraceReport assertions).
' VlaSelfTest's dispatch list dropped TestRuntimeModule; TestHelpers
' and TestRuntimeTrace stayed on it under their original names, now
' pure. The new module gets its own dispatcher (VlaSelfTestHost) and
' its own local Report/CheckV (R7 - a real architectural boundary,
' not a size-only split, so counters do not cross it). Nothing pure
' changed behavior; verified by an unchanged VlaSelfTest pass/fail
' count and an empty golden diff, same as F8.0.
' AS6.0: SD-6 ("the corpus family is a fixture... already in force")
' turned out to hold for none of the three files under a genuine
' absence. english.vla's absence crashed VlaSelfTest mid-run
' (TestDotCount's bare VlaDotCount call, unguarded); alonzo.vla's
' absence ALSO crashed it (TestAlonzoLib, VLA_Tests_Grammar.bas -
' its FindDevFile call was equally bare; the Dir$ recheck right
' after it was dead code, since FindDevFile raises on a miss and
' never returns "" - caught only by this pass's adversarial
' verification, not by the first read of that Sub, which is why the
' first draft of this note wrongly called it "already fine"); and
' instructions.txt was never checked by VlaSelfTest at all - only by the
' separately-invoked VlaWriteGoldens. Both bare calls now capture
' Err.Description under Resume Next and report FAIL instead of
' dying, same hazard class as L2.1/L2.2. New: VlaCorpusFamilyOk (one
' function, checks all three via FindDevFile, names every missing
' file) and TestCorpusFamily (dispatched FIRST, so a missing corpus
' file is named immediately instead of surfacing as an unrelated
' failure - or a crash - further down the run).

' =====================================================================
'  VLA_Tests - the developer rig's proof half: the harness (assertion
'  and report primitives), the goldens machinery, and the pure
'  core-mechanics pins (emitter, loader, IDE naming, head table,
'  interpreter skeleton, section markers) - nothing here touches a
'  live workbook. The grammar/phrasebook feature pins live in the
'  sibling module VLA_Tests_Grammar.bas (F.8), also pure. Everything
'  needing a live Excel - VerifyReport, TestRuntimeModule, and the
'  host-dependent tails TestHelpers and TestRuntimeTrace used to
'  carry - lives in VLA_Tests_Host.bas (F.11). All three are dev-only:
'  NOT shipped in the add-in build.
'
'  VlaSelfTest    golden tests for the transpiler (which previously
'                 had NO direct tests), tokenizer regressions through
'                 EnglishToVla, loader behavior (two-line gathering,
'                 failing tests must fail loads), and pure helper
'                 units (VlaColor's BGR swap, the VlaDict family).
'                 Prints PASS/FAIL per test to the Immediate window.
'                 Dispatches into VLA_Tests_Grammar's pins too - a
'                 Public Sub in another standard module is callable
'                 by bare name, so the single dispatch list here still
'                 covers both modules' tests. Needs nothing but VBA
'                 itself - no live workbook state, so it is the run
'                 every pass can iterate against without a Windows box
'                 (F.11).
'                 NOTE: resets the grammar - Reload vocab afterwards.
'                 NOTE: host-required tests are a separate run -
'                 VlaSelfTestHost, then VerifyReport after a Run - see
'                 VLA_Tests_Host.bas.
'
'  VlaWriteGoldens  S1: regenerates the whole-corpus golden files
'                 (instructions_golden.vla / instructions_golden.vba) beside
'                 instructions.txt. git diff on them is the total-pipeline
'                 witness - run in every verification loop.
'
'  VlaWriteInterpreterGoldens  IN.3.5: the interpreter's own version -
'                 regenerates scripts/interpreter_golden.txt, the
'                 walking skeleton's ordered runtime-effect record.
'
'  VlaGoldens     convenience: VlaWriteGoldens then
'                 VlaWriteInterpreterGoldens, one call.
'
'  LAYER:     Harness (dev-only; never ships - see REBUILD.md SS3
'             Layer 5)
'  MAY CALL:  VLA (transpiler), VLA_English (grammar engine),
'             VLA_Loader, VLA_HeadTable, VLA_Runtime, VLA_IDE,
'             VLA_Build, VLA_Interpreter, VLA_Tests_Grammar (dispatch
'             only, from VlaSelfTest) - none of these calls touches a
'             live workbook; VLA_Runtime's VlaColor/VlaDict family and
'             VLA_IDE's naming/scan helpers are pure functions
'  SHIPS:     nowhere. Dev-rig only, like VLA_DevRig.
'  PAYS INTO: F.8 (the concern split), F.11 (the pure/host split - this
'             is now the pure half, callable without a live workbook),
'             R4 (the 1,200-line budget REBUILD.md names - not yet
'             met; this is a first cut, not the full Layer-5
'             topology), every roadmap item these pins hold the
'             machinery for.
'  REASON:    see the F8.0 history note above.
' =====================================================================

Private mPass As Long
Private mFail As Long
Private mFailedNames As Collection
' TABLESPEC-SCALE / TABLESPECSCALE.0: off by default so an ordinary
' VlaSelfTest run stays fast during daily development. Gates the
' 2000/5500-row cdr-based scale cases in TestListopsDepthSafety/
' TestTablespecDepthSafety - NOT the trampoline fix's own stack-depth
' cost (that's O(1) per row now, cheap at any N tested), but a
' pre-existing, separate O(n^2) cost in cdr/ListTail (VLA.bas): every
' cdr call copies the ENTIRE remaining tail into a new Collection, so a
' full walk-rows/tablespec walk of length N does ~N^2/2 element copies
' regardless of the trampoline - invisible below ~1000 rows, several
' minutes at 5500, confirmed live. Fixing cdr itself is real, separate
' engine work (same family as the already-backlogged P-NTH, gated on
' P-PROF) - not something to do under time pressure here. Set True and
' call VlaSelfTest (or use VlaSelfTestScale below) only when a change
' actually touches ExpandMacros/cdr/ListTail/tablespec's own walk and
' needs re-proving at genuinely large scale.
Public mRunScaleTests As Boolean

Public Function VlaSelfTest() As Boolean
    mPass = 0
    mFail = 0
    Set mFailedNames = New Collection
    Debug.Print "===== VLA SELF-TEST ====="

    TestCorpusFamily
    TestEmitter
    TestTier1
    TestPreludeLibrary
    TestVlaExpand
    TestEnglishCore
    TestRawVlaRows
    TestExplainTrace
    TestLoader
    TestHelpers
    TestIdeNaming
    TestBuildRibbon
    TestUndoScan
    TestGoldens
    TestRuleUsage
    TestMessageSeam
    TestSec8Provenance
    TestResolveCheck
    TestRuntimeTrace
    TestVlaTry
    TestPreludeGrowth
    TestTemplateLimits
    TestG1
    TestG2
    TestLE2
    TestG4
    TestG3
    TestL4
    TestVocabDiff
    TestMetaVocab
    TestVocabMacroProbe
    TestTableFamily
    TestAtRow
    TestGExpander
    TestRuleCoverage
    TestRawConsentDeviceScope
    TestRawConsentWorkbookScope
    TestRawConsentTextPathUngated
    TestPhrasebookPersistence
    TestPhrasebookReplayAddsNotReplaces
    TestGenRow
    TestListopsBudget
    TestListopsConfluence
    TestQuasiquote
    TestListops
    TestListopsExpand
    TestCond
    TestListopsDepthSafety
    TestAntonymSweep
    TestListopsStdlib
    TestTablespecDepthSafety
    TestTablespec
    TestG5
    TestG10
    TestG11
    TestG11r
    TestG6
    TestG7
    TestG8
    TestGRowLoop
    TestV7
    TestV8
    TestL6
    TestEvalDisplay
    TestL16
    TestL7
    TestL17
    TestL11
    TestL11_1
    TestAproposPlus
    TestL8
    TestL82
    TestL12
    TestL15
    TestPL1
    TestPL2
    TestPL3
    TestPL4
    TestPL5
    TestPL6
    TestPL7
    TestL13
    TestG12
    TestL14
    TestL18
    TestAlonzoLib
    TestLx10NonEnglishFixture
    TestF4NoiseWordBeforeSlot
    TestF4CrossRuleShadow
    TestGPath
    TestHeadTable
    TestHeadTableExportOnly
    TestHeadAlias
    TestNonAsciiIdentifiers
    TestDotCount
    TestInterpreterSkeleton
    TestInterpreterEffectLog
    TestInterpreterDynamicHead
    TestInterpreterCsvInjectionGuard
    TestInterpreterExcelConstants
    TestInterpreterVbConstants
    TestInterpreterOperators
    TestInterpreterQuote
    TestArrayPrimitive
    TestLInterpolate
    TestInterpreterControlFlow
    TestInterpreterCallReturn
    TestInterpreterModuleScope
    TestInterpreterNewCollection
    TestInterpreterSec1DynamicMemberRefused
    TestInterpreterOnError
    TestSectionMarkers
    TestContextPushPop
    TestContextUnderflow
    TestCo4VersionSemver
    TestF10Requires
    TestVlaLint

    Debug.Print "===== SELF-TEST: " & mPass & " passed, " & mFail & " failed ====="
    If mFail > 0 Then
        Debug.Print "===== FAILED TESTS (" & mFailedNames.Count & ") ====="
        Dim fn As Variant
        For Each fn In mFailedNames
            Debug.Print "  FAIL  " & fn
        Next fn
        Debug.Print "===== (paste the block above into a bug report) ====="
    End If
    If mFail = 0 Then Debug.Print "(grammar was reset - Reload vocabulary before the next Run)"
    Debug.Print "(host-required tests are separate - VlaSelfTestHost, then VerifyReport after a Run)"
    Debug.Print "(QUERY AND LOGIC's DSL tests are separate too, and never run from here - TestDSLs)"
    VlaSelfTest = (mFail = 0)
End Function

' TABLESPEC-SCALE / TABLESPECSCALE.0: convenience wrapper, VlaSelfTests's
' own shape (VLA_Tests_Host.bas) - runs the full suite INCLUDING the
' 2000/5500-row cdr-based scale cases mRunScaleTests normally skips (see
' its own declaration above for why they're gated). Several minutes
' slower than a plain VlaSelfTest run; use only when a change actually
' touches ExpandMacros/cdr/ListTail/tablespec's own walk.
Public Function VlaSelfTestScale() As Boolean
    mRunScaleTests = True
    On Error GoTo cleanup
    VlaSelfTestScale = VlaSelfTest()
    mRunScaleTests = False
    Exit Function
cleanup:
    Dim n As Long, s As String, d As String
    n = Err.Number: s = Err.Source: d = Err.Description
    mRunScaleTests = False
    Err.Raise n, s, d
End Function

' ---------------------------------------------------------------------
'  AS.6: SD-6 says the corpus family is a fixture - absence of
'  instructions.txt / english.vla / alonzo.vla fails the suite - and
'  called that "already in force." Auditing every corpus-file touch in
'  the dispatch chain found it was NOT actually true for any of the
'  three: english.vla's absence crashed the suite (TestDotCount's
'  bare VlaDotCount call), alonzo.vla's absence ALSO crashed it
'  (TestAlonzoLib's FindDevFile call was equally bare - its Dir$
'  recheck right after was dead code, since FindDevFile raises on a
'  miss and never returns "" - both fixed alongside this test), and
'  instructions.txt was never checked by VlaSelfTest at all - only by the
'  separately-invoked VlaWriteGoldens. Dispatched first, so a missing
'  corpus file is named immediately instead of surfacing as an
'  unrelated-looking failure (or a crash) further down the run.
' ---------------------------------------------------------------------
Private Sub TestCorpusFamily()
    Dim detail As String
    Report "corpus: instructions.txt / english.vla / alonzo.vla are all present", _
           VlaCorpusFamilyOk(detail), "missing:" & detail
End Sub

' ---------------------------------------------------------------------
'  Emitter goldens: VLA -> VBA fragments (case-insensitive contains).
' ---------------------------------------------------------------------
Private Sub TestEmitter()
    AssertVla "dim", "(sub t () (dim x Long))", "Dim x As Long"
    AssertVla "assignment+operators", "(sub t () (set! x (+ 1 2)))", "x = (1 + 2)"
    AssertVla "obj-set!", "(sub t () (obj-set! r (range ""a1"")))", "Set r = range(""a1"")"
    AssertVla "dot chain property set", "(sub t () (set! (. (range ""a1"") font.bold) true))", _
              "range(""a1"").font.bold = True"
    AssertVla "keyword args", "(sub t () (. ws exportasfixedformat :type xltypepdf))", "type:=xltypepdf"
    AssertVla "hyphen names", "(sub t () (set! grand-total 5))", "grand_total = 5"
    AssertVla "if/else", "(sub t () (if (> x 1) (then (debug-print 1)) (else (debug-print 2))))", _
              "If (x > 1) Then", "Else", "End If"
    AssertVla "macro: when", "(sub t () (when (> x 1) (debug-print x)))", "If (x > 1) Then"
    AssertVla "macro: dotimes", "(sub t () (dotimes counter 3 (debug-print counter)))", _
              "For counter = 1 To 3"
    AssertVla "string literal", "(sub t () (debug-print ""hi""))", "Debug.Print ""hi"""
    AssertVla "top-level const", "(const pink ""#FF69B4"")", "Const pink = ""#FF69B4"""
    AssertVla "optional param default", "(sub t ((optional p Variant 1)) (debug-print p))", _
              "Optional p As Variant = 1"
    ' (function name (params) ReturnType body...) - params third,
    ' return type fourth; the original golden had them swapped and
    ' the emitter's own error corrected it. First self-test, 2026.
    AssertVla "function return", "(function f () Long (return 5))", "As Long", "f = 5", "Exit Function"
    AssertVla "for step", "(sub t () (for (k 9 2 -1) (debug-print k)))", "Step -1"
    ' B1: the (new ...) expression - present in the emitter but never
    ' pinned until lists made it a shipped path.
    AssertVla "new expression", "(sub t () (obj-set! c (new Collection)))", _
              "Set c = New Collection"
End Sub

' ---------------------------------------------------------------------
'  C2 Tier 1 goldens: visibility, arrays/ReDim, Type/Enum, and the
'  ' vla:N source map. Map pins use Norm's newline collapse the same
'  way C1's ordering pins do: "For k = 1 To 3 Debug.Print k ' vla:3"
'  pins that the macro-template line is clean AND the user's argument
'  kept its line, in one fragment.
' ---------------------------------------------------------------------
Private Sub TestTier1()
    AssertVla "tier1: private sub", _
              "(private (sub helper () (set! x 1)))", _
              "Private Sub helper"
    AssertVla "tier1: visibility on dim, const, function", _
              "(begin (private (dim counter-base Long)) (private (const k 5)) (public (function f () Long (return k))))", _
              "Private counter_base As Long", _
              "Private Const k = 5", _
              "Public Function f"
    AssertVla "tier1: array declarations", _
              "(sub t () (dim a (array Long 10)) (dim b (array Double (to 1 5))) (dim c (array)))", _
              "Dim a(10) As Long", _
              "Dim b(1 To 5) As Double", _
              "Dim c() As Variant"
    AssertVla "tier1: redim and redim preserve", _
              "(sub t () (dim a (array Long)) (redim a n) (redim preserve a (to 1 n) 4))", _
              "ReDim a(n)", _
              "ReDim Preserve a(1 To n, 4)"
    AssertVla "tier1: type definition", _
              "(type Point (x Double) (y Double) (tags (array String 3)))", _
              "Type Point", _
              "x As Double", _
              "tags(3) As String", _
              "End Type"
    AssertVla "tier1: private enum with values", _
              "(private (enum Color red (green 5) blue))", _
              "Private Enum Color", _
              "green = 5", _
              "End Enum"
    ' The source map: a statement is tagged with the line of its
    ' opening paren in the USER source (the prelude never maps).
    AssertVla "map: statements carry their vla line", _
              "(sub t ()" & vbCrLf & "(set! x 1)" & vbCrLf & "(set! y 2))", _
              "x = 1 ' vla:2", _
              "y = 2 ' vla:3"
    ' Through a macro: the template's For line is clean (no user line
    ' to name), the user's body statement keeps its own.
    AssertVla "map: macro args keep lines, templates stay clean", _
              "(sub t ()" & vbCrLf & "(dotimes k 3" & vbCrLf & "(debug-print k)))", _
              "For k = 1 To 3 Debug.Print k ' vla:3"

    ' V4: the three-layer join. (at-line N ...) is a zero-runtime
    ' annotation: everything emitted inside carries src:N beside its
    ' vla line. The first pin fixes the full dual-tag format AND both
    ' values ((set! x 1) opens on user line 3 of this source).
    AssertVla "map: at-line adds the source coordinate", _
              "(sub t ()" & vbCrLf & "(at-line 7" & vbCrLf & "(set! x 1)))", _
              "x = 1 ' vla:3 src:7"
    ' Several sibling forms share one wrapper (English's Try emits a
    ' train of forms from one sentence) - EACH must carry the tag.
    Dim vb As String
    vb = ""
    On Error Resume Next
    vb = VlaTranspile("(sub t () (at-line 3 (set! x 1) (set! y 2)))")
    On Error GoTo 0
    Report "map: at-line covers every sibling it wraps", CountOcc(vb, "src:3") = 2, _
           "wanted 2 x 'src:3', got " & CountOcc(vb, "src:3")
    ' Through (begin ...): begin only delegates and never tags itself,
    ' so the wrapper's dynamic extent is what carries the sentence
    ' line onto begin's inner statements - the vocabulary multi-step
    ' sentence path.
    vb = ""
    On Error Resume Next
    vb = VlaTranspile("(sub t () (at-line 4 (begin (set! x 1) (set! y 2))))")
    On Error GoTo 0
    Report "map: at-line reaches through begin", CountOcc(vb, "src:4") = 2, _
           "wanted 2 x 'src:4', got " & CountOcc(vb, "src:4")
    ' Nesting: an inner wrapper overrides within its own extent (block
    ' bodies re-wrap per sentence); the If opener keeps the outer line.
    AssertVla "map: an inner at-line overrides within its extent", _
              "(sub t ()" & vbCrLf & "(at-line 2" & vbCrLf & "(if x" & vbCrLf & _
              "(then" & vbCrLf & "(at-line 5" & vbCrLf & "(set! y 1))))))", _
              "If x Then ' vla:3 src:2", _
              "y = 1 ' vla:6 src:5"
    ' The extent ENDS with the wrapper: a following statement must not
    ' inherit the tag (exactly one src: in this whole module).
    vb = ""
    On Error Resume Next
    vb = VlaTranspile("(sub t () (at-line 2 (set! x 1)) (set! z 9))")
    On Error GoTo 0
    Report "map: at-line scope ends with its wrapper", CountOcc(vb, "src:") = 1, _
           "wanted exactly 1 'src:', got " & CountOcc(vb, "src:")
    ' The refusal exists (proven before pinning): a non-number where
    ' the line belongs is a loud emitter error, not a silent 0.
    Dim atErr As String
    atErr = ""
    On Error Resume Next
    vb = VlaTranspile("(sub t () (at-line banana (set! x 1)))")
    If Err.Number <> 0 Then atErr = Err.Description
    On Error GoTo 0
    Report "map: at-line refuses a non-number line", InStr(1, atErr, "line number", vbTextCompare) > 0, _
           "wanted a 'line number' refusal, got: " & Left$(atErr, 120)
End Sub

' Occurrences of a fragment in a text (used by the V4 map pins).
Public Function CountOcc(ByVal text As String, ByVal frag As String) As Long
    Dim i As Long, k As Long
    i = 1
    Do
        i = InStr(i, text, frag, vbTextCompare)
        If i = 0 Then Exit Do
        k = k + 1
        i = i + Len(frag)
    Loop
    CountOcc = k
End Function

' ---------------------------------------------------------------------
'  C1 prelude library goldens: the with- macros. Norm collapses
'  newlines to single spaces, so a fragment spanning two emitted lines
'  ("fastexit: Application.ScreenUpdating = True") pins ORDERING, not
'  just presence - the restore must sit after the label, the exit
'  before it. That trick is what makes soft pins strong enough here.
' ---------------------------------------------------------------------
Private Sub TestPreludeLibrary()
    ' with-fast-excel: both switches off, the trap armed at the
    ' caller-supplied label, the re-raise guard present.
    AssertVla "prelude: with-fast-excel arms and guards", _
              "(sub t () (with-fast-excel fastexit (set! (range ""a1"") 5)))", _
              "Application.ScreenUpdating = False", _
              "Application.Calculation = xlcalculationmanual", _
              "On Error GoTo fastexit", _
              "Call err.raise(err.number, err.source, err.description)", _
              "On Error GoTo 0"
    ' ...and the restore lands AFTER the label (ordering pin), so the
    ' error path cannot skip it.
    AssertVla "prelude: with-fast-excel restores after the label", _
              "(sub t () (with-fast-excel fastexit (set! (range ""a1"") 5)))", _
              "fastexit: Application.ScreenUpdating = True", _
              "Application.Calculation = xlcalculationautomatic"
    ' with-error-handler in a Sub: the classic tail idiom - Exit Sub
    ' immediately before the label (ordering pin), handler after it.
    AssertVla "prelude: with-error-handler (sub)", _
              "(sub t () (with-error-handler oops (debug-print err.description) (set! x 1)))", _
              "On Error GoTo oops", _
              "Exit Sub oops:", _
              "oops: Debug.Print Err.Description"
    ' ...and in a Function the same macro exits with Exit Function -
    ' the (return) polymorphism is the reason the macro uses it.
    AssertVla "prelude: with-error-handler (function)", _
              "(function f () Long (with-error-handler oops (return 0) (return (* x 2))))", _
              "On Error GoTo oops", _
              "Exit Function oops:"
    ' with-sheet: save the departure point into the caller-named
    ' temporary, activate the target, work, hand the prior sheet back.
    AssertVla "prelude: with-sheet saves and activates", _
              "(sub t () (with-sheet ""Data"" prev (set! (range ""a1"") 5)))", _
              "Dim prev As Object", _
              "Set prev = activesheet", _
              "Call worksheets(""Data"").activate", _
              "Call prev.activate"
    ' ...restore strictly after the body (ordering pin). Blessed at
    ' C2: the body statement is the USER'S form, so it now carries its
    ' source-map tag - the pin carries it too, and so double-pins the
    ' C2 rule that macro arguments keep their lines.
    AssertVla "prelude: with-sheet restores after the body", _
              "(sub t () (with-sheet ""Data"" prev (set! (range ""a1"") 5)))", _
              "range(""a1"") = 5 ' vla:1 Call prev.activate"
End Sub

' ---------------------------------------------------------------------
'  L1 goldens: VlaExpand's programmatic half. ONE PASS shows the
'  template substitution untouched (nested macro calls survive as
'  calls); FIXPOINT shows what the emitter sees; firedCount reports
'  applications. Fragment checks ride Norm, so the pretty-printer's
'  line breaks never enter the contract - only the datums do.
' ---------------------------------------------------------------------
Private Sub TestVlaExpand()
    Dim t As String
    Dim fired As Long

    ' One pass: the template substitution, nothing more.
    t = VlaExpandText("(inc! x)", False, fired)
    CheckFrags "expand: one pass shows the template", t, Array("(set! x (+ x 1))")
    CheckV "expand: one pass fired count", fired, 1

    ' One pass stops at the outermost call: the inner macro survives
    ' spelled as a call - the defining contrast with fixpoint.
    t = VlaExpandText("(when (> x 1) (inc! x))", False, fired)
    CheckFrags "expand: one pass keeps inner macro calls", t, _
               Array("(if (> x 1) (then (inc! x)))")

    ' Fixpoint expands the nest fully and leaves no macro heads.
    t = VlaExpandText("(when (> x 1) (inc! x))", True, fired)
    CheckFrags "expand: fixpoint expands the nest", t, _
               Array("(if (> x 1)", "(set! x (+ x 1))")
    Report "expand: fixpoint leaves no macro calls", _
           InStr(1, t, "(inc!", vbTextCompare) = 0, _
           "found '(inc!' in: " & Left$(Norm(t), 160)
    CheckV "expand: fixpoint fired count", fired, 2

    ' Core VLA passes through untouched, count zero.
    t = VlaExpandText("(set! x 1)", True, fired)
    CheckFrags "expand: core form prints normalized", t, Array("(set! x 1)")
    CheckV "expand: core form fires nothing", fired, 0

    ' String literals round-trip with their escapes (backslashes
    ' re-escaped before quotes - the printer's ordering rule).
    t = VlaExpandText("(debug-print ""a\""b\\c"")", True, fired)
    CheckFrags "expand: string literal re-escapes", t, Array("""a\""b\\c""")

    ' A user defmacro drives expansion and is not echoed - only the
    ' expanded body forms print, same two-pass shape as a transpile.
    t = VlaExpandText("(defmacro (twice s) (begin s s))" & vbCrLf & _
                      "(twice (set! x 1))", True, fired)
    CheckFrags "expand: user defmacro expands", t, _
               Array("(begin (set! x 1) (set! x 1))")
    Report "expand: defmacro definition not echoed", _
           InStr(1, t, "defmacro", vbTextCompare) = 0, _
           "found 'defmacro' in: " & Left$(Norm(t), 160)

    ' The C1 flagship: with-fast-excel's whole bracket, body spliced,
    ' restore after the label - the expansion a macro author reads
    ' here is the one the emitter receives.
    t = VlaExpandText("(with-fast-excel done (set! x 1))", True, fired)
    CheckFrags "expand: with-fast-excel golden", t, Array( _
        "(set! application.screenupdating false)", _
        "(on-error goto done)", _
        "(set! x 1)", _
        "(label done)", _
        "(set! application.screenupdating true)")

    ' Several top-level forms all print, in order.
    t = VlaExpandText("(inc! a)" & vbCrLf & "(dec! b)", True, fired)
    CheckFrags "expand: multi-form source prints all forms", t, _
               Array("(set! a (+ a 1))", "(set! b (- b 1))")

    ' Reader errors still raise from the programmatic half (the
    ' printing Sub catches and prints them; Explain's panel skips).
    Dim d As String
    On Error Resume Next
    t = VlaExpandText("(set! x", True, fired)
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    Report "expand: reader error raises", _
           InStr(1, d, "unbalanced", vbTextCompare) > 0, "got: " & d
End Sub

' ---------------------------------------------------------------------
'  English core: structural forms and tokenizer regressions, prelude
'  only (no vocabulary file needed).
' ---------------------------------------------------------------------
Private Sub TestEnglishCore()
    EnglishResetGrammar
    AssertEnglish "thousands separator", "Set budget to 1,000,000.", "(set! budget 1000000)"
    AssertEnglish "decimal point", "Set pi-ish to 3.14 plus 1.", "(+ 3.14 1)"
    AssertEnglish "mixed 1,234.56", "Set price to 1,234.56.", "1234.56"
    AssertEnglish "count loop", "Count k from 2 to 9, log k.", "(for (k 2 9)"
    AssertEnglish "count down", "Count k down from 9 to 2, log k.", "(for (k 9 2 -1)"
    AssertEnglish "stop the loop picks Exit For", "Repeat 2 times, stop the loop.", "(exit-for)"
    AssertEnglish "stop the loop picks Exit Do", "While x is greater than 0, stop the loop.", "(exit-do)"
    AssertEnglish "text comparator", "If code contains ""x"", log ""hit"".", "vbtextcompare"
    ' DF1: the is-empty condition now reads through blank? - the
    ' frag follows the machine (the old plumbing spelling lives one
    ' expansion below, where TestPL1 pins it byte-exact).
    AssertEnglish "is empty", "If x is empty, log ""blank"".", "(blank? x)"
    AssertEnglish "cells as values", "Set v to cell B2 plus 1.", "(+ (range ""b2"") 1)"
    AssertEnglish "function word", "Set n to length of code.", "(len code)"
    AssertEnglish "nullary today", "Set d to month of today.", "(month (date))"
    AssertEnglish "show is a box", "Show total.", "(msgbox total)"
    AssertEnglish "log is the log", "Log total.", "(debug-print total)"

    ' B1: lists are core language - no vocabulary file needed.
    AssertEnglish "create a list declares", "Create a list called found-items.", _
                  "(dim found-items Collection)"
    AssertEnglish "create a list initializes", "Create a list called found-items.", _
                  "(new Collection)"
    AssertEnglish "append is Collection.Add", "Append 5 to found-items.", _
                  "(. found-items add 5)"
    AssertEnglish "count of a list", "Set n to count of found-items.", _
                  "(vlacount found-items)"

    ' B2: branching completeness. Block forms carry step markers, so
    ' these pin fragments rather than whole translations.
    Dim chainP As String
    chainP = "If x is 1:" & vbLf & "  Log ""one""." & vbLf & vbLf & _
             "Otherwise, if x is 2:" & vbLf & "  Log ""two""." & vbLf & vbLf & _
             "Otherwise:" & vbLf & "  Log ""many""."
    AssertEnglish "otherwise-if becomes elseif", chainP, "(elseif (= x 2)"
    AssertEnglish "chain keeps its final else", chainP, "(else"
    AssertEnglish "one-line otherwise-if branch", _
                  "If x is 1:" & vbLf & "  Log ""one""." & vbLf & vbLf & _
                  "Otherwise, if x is 2, log ""two"".", "(elseif (= x 2)"
    Dim whenP As String
    whenP = "When region is ""North"":" & vbLf & "  Log ""cold""." & vbLf & vbLf & _
            "When it is ""South"" or ""East"":" & vbLf & "  Log ""warm""." & vbLf & vbLf & _
            "Otherwise:" & vbLf & "  Log ""where?""."
    AssertEnglish "when chain is a select", whenP, "(select region"
    AssertEnglish "when values join with or", whenP, "(case (""South"" ""East"")"
    AssertEnglish "when chain catch-all", whenP, "(case-else"
    AssertEnglish "repeat until is do-until", _
                  "Repeat until fuel is 0, decrease fuel by 1.", "(do-until (= fuel 0)"
    AssertEnglish "stop inside repeat-until picks Exit Do", _
                  "Repeat until fuel is 0, stop the loop.", "(exit-do)"

    ' B3: Try blocks. Label numbers are per-translation, so fragments
    ' pin the label FAMILY, never the number.
    Dim tryP As String
    tryP = "Try:" & vbLf & "  Set x to 1." & vbLf & vbLf & _
           "If that fails:" & vbLf & "  Log ""saved""."
    AssertEnglish "try arms a local handler", tryP, "(on-error goto vla-tryf-"
    AssertEnglish "try failure resumes out of the handler", tryP, "(resume vla-tryr-"
    AssertEnglish "try restores the step handler", tryP, "(on-error goto vla-fail)"
    AssertEnglish "try success skips the recovery", tryP, "(goto vla-tryd-"
    AssertEnglish "try without recovery is legal", _
                  "Try:" & vbLf & "  Set x to 1.", "(label vla-tryr-"

    ' A stray recovery paragraph is refused, naming its line.
    AssertErrLine "stray 'If that fails' names its line", _
        "Set x to 1." & vbLf & vbLf & "If that fails:" & vbLf & "  Log ""x"".", 3

    ' The Known Sentences listing must cover the structural forms -
    ' they live in the parser, not the rule list, so this gate is what
    ' keeps the listing honest as the grammar grows. Add a structural
    ' form? Add its fragment here AND its line in EnglishListPhrases.
    Dim structFrags As Variant
    structFrags = Array("create a number", "if <condition>", "otherwise, if", _
                        "when <value> is", "when it is", "try:", "if that fails", _
                        "repeat <n> times", "repeat until", "while <condition>", _
                        "count <name> from", "for each", "stop the loop", _
                        "to <name>:", "to <name> of <param>", "to [get] <name> using", _
                        "give back", "define <name> as")
    CheckFrags "listing covers structural forms", EnglishListPhrases(), structFrags

    ' LE.1 (thin slice): EnglishPhraseRows is EnglishListPhrases's own
    ' walk handed back structured (Template/Example/isHeader parallel
    ' Collections) instead of preformatted text - the "What can I say?"
    ' button's real two-column table data source (VLA_IDE.bas). Proven
    ' here: the three Collections stay in lockstep, the built-in
    ' section header comes back marked as a header with a blank
    ' example, and - the part that actually matters for the button - a
    ' freshly loaded rule's own PASSING test: sentence becomes its
    ' worked example in the Example column, not a synthesized one.
    Dim rowTemplates As Collection, rowExamples As Collection, rowIsHeader As Collection
    EnglishPhraseRows rowTemplates, rowExamples, rowIsHeader
    Report "phrase rows: template/example/isHeader stay the same length", _
           rowTemplates.Count = rowExamples.Count And rowTemplates.Count = rowIsHeader.Count, _
           "templates " & rowTemplates.Count & " examples " & rowExamples.Count & " isHeader " & rowIsHeader.Count
    Report "phrase rows: the built-in section header is a header row with a blank example", _
           CStr(rowTemplates.Item(1)) = "built into the language (every vocabulary shares these)" And _
           CBool(rowIsHeader.Item(1)) = True And CStr(rowExamples.Item(1)) = "", _
           "template '" & CStr(rowTemplates.Item(1)) & "' isHeader " & CBool(rowIsHeader.Item(1)) & _
           " example '" & CStr(rowExamples.Item(1)) & "'"

    EnglishResetGrammar
    EnglishLoadVocabularyText _
        "(english-vla ""praise cell {r:cell}"" (debug-print {r}))" & vbCrLf & _
        "(test-success ""Praise cell B2."" (debug-print ""b2""))", "phrase-rows-vocab"
    EnglishPhraseRows rowTemplates, rowExamples, rowIsHeader
    Dim prI As Long, foundRow As Boolean, foundExample As String
    foundRow = False
    For prI = 1 To rowTemplates.Count
        If InStr(1, CStr(rowTemplates.Item(prI)), "praise cell", vbTextCompare) > 0 Then
            foundRow = True
            foundExample = CStr(rowExamples.Item(prI))
        End If
    Next
    Report "phrase rows: a loaded rule's own passing test becomes its worked example", _
           foundRow And InStr(1, foundExample, "Praise cell B2", vbTextCompare) > 0, _
           "found " & foundRow & " example '" & foundExample & "'"
    EnglishResetGrammar

    ' B4: value-returning actions. One program covers the definition,
    ' the Function shape, composition into Set, and - the phase's
    ' done-when - a user-defined calculation inside a condition.
    Dim fnP As String
    fnP = "To tax of amount:" & vbLf & "  Give back amount times 0.08." & vbLf & vbLf & _
          "Set fee to tax of 100." & vbLf & _
          "If tax of 50 is greater than 3, log ""big""."
    AssertEnglish "to-of defines a Function", fnP, "(function tax ((amount Variant))"
    AssertEnglish "give back is return", fnP, "(return (* amount 0.08))"
    AssertEnglish "the new word composes in Set", fnP, "(set! fee (tax 100))"
    AssertEnglish "the new word composes in a condition", fnP, "(if (> (tax 50) 3)"
    AssertEnglish "function instrumentation exits as Function", fnP, "(exit-function)"
    AssertEnglish "bare Stop inside a function exits the Function", _
        "To probe of x:" & vbLf & "  Stop." & vbLf & "  Give back x.", "(exit-function)"

    ' The guards, each pinned to its line.
    AssertErrLine "Give back outside a function names its line", "Give back 5.", 1
    AssertErrLine "shadowing a built-in word is refused", _
        "To length of x:" & vbLf & "  Give back x.", 1
    AssertErrLine "duplicate action names are refused", _
        "To greet:" & vbLf & "  Log 1." & vbLf & vbLf & "To greet:" & vbLf & "  Log 2.", 4

    ' B5: the prim's "number" branch - computed columns as VALUES.
    AssertEnglish "column-number cell in a condition", _
        "If cell in column number 3 row k is 5, log ""hit"".", "(= (cells k 3) 5)"
    ' Blessed after this golden's maiden failure: the row expression
    ' is greedy - "row 2 plus 1" IS row 3 (the lettered form always
    ' bound the same way). Arithmetic on the VALUE names it first.
    AssertEnglish "row expression is greedy (row 2 plus 1 = row 3)", _
        "Set x to cell in column number c-pick row 2 plus 1.", "(cells (+ 2 1) c-pick)"

    ' B5.2: "value in ..." - the honest synonym for reading cells.
    AssertEnglish "value in column composes in a condition", _
        "If value in column number 3 row k is 5, log ""hit"".", "(= (cells k 3) 5)"
    AssertEnglish "value in cell reads the cell", _
        "Set v to value in cell B2.", "(set! v (range ""b2""))"

    ' B5.2, owner decision: word-flanked ! is a sheet qualifier;
    ' free-standing ! is still a full stop.
    AssertEnglish "sheet-qualified bang is one reference", _
        "Set x to cell Data!B2.", "(range ""data!b2"")"
    AssertEnglish "qualified bang chains with the range colon", _
        "Set x to cell Data!B2 plus 1.", "(+ (range ""data!b2"") 1)"
    AssertEnglish "free-standing bang still ends the sentence", _
        "Log total!", "(debug-print total)"

    ' B5.3, owner decision: Excel's spaced-sheet syntax works bare.
    AssertEnglish "single-quoted spaced sheet binds bare", _
        "Set x to cell 'Q1 Data'!B2.", "(range ""'Q1 Data'!B2"")"
    AssertEnglish "spaced sheet chains with the range colon", _
        "Set x to cell 'Q1 Data'!A1:B10 plus 0.", "(range ""'Q1 Data'!A1:B10"")"
    AssertErrLine "a contraction apostrophe is still refused", _
        "Set label to don't.", 1

    ' B6: using-style value actions - named parameters and defaults.
    Dim useP As String
    useP = "To get tax using income of 1000 and rate of 15:" & vbLf & _
           "  Give back income times rate divided by 100." & vbLf & vbLf & _
           "Set fee to tax using income of 50000 and rate of 20." & vbLf & _
           "Set base-fee to get tax using income of 800." & vbLf & _
           "If tax using income of 5000 and rate of 10 is greater than 400 and fee is greater than 0, log ""rich""."
    AssertEnglish "using-def is a Function with optionals", useP, _
        "(function tax ((optional income Variant 1000) (optional rate Variant 15))"
    AssertEnglish "using-call emits named arguments", useP, "(tax :income 50000 :rate 20)"
    AssertEnglish "omitted parameters take defaults (and get is call sugar)", useP, "(tax :income 800)"
    AssertEnglish "the and-collision resolves by parameter lookahead", useP, _
        "(> (tax :income 5000 :rate 10) 400)"
    AssertEnglish "required parameters emit plain", _
        "To ship using weight and speed of 9:" & vbLf & "  Give back weight times speed." & vbLf & vbLf & _
        "Set x to ship using weight of 2.", "(function ship ((weight Variant) (optional speed Variant 9))"

    ' The guards, line-pinned.
    AssertErrLine "unknown call parameter names its line", _
        "To tx using a of 1:" & vbLf & "  Give back a." & vbLf & vbLf & _
        "Set x to tx using b of 2.", 4
    AssertErrLine "missing required parameter names the call line", _
        "To ship2 using weight and speed of 9:" & vbLf & "  Give back weight." & vbLf & vbLf & _
        "Set x to ship2 using speed of 1.", 4
    AssertErrLine "required-after-optional refused at definition", _
        "To bad using a of 1 and b:" & vbLf & "  Give back a.", 1
    AssertErrLine "with-clause ordering hole is closed too", _
        "To stampx, with a of 1 and b:" & vbLf & "  Log a.", 1
    AssertErrLine "parameter shadowing a function word is refused", _
        "To f using length of 2:" & vbLf & "  Give back length.", 1

    ' B6.3, owner decision: % is a postfix divide-by-100, Excel's own
    ' reading, binding at the primary level.
    AssertEnglish "percent on a number", "Set x to 15%.", "(/ 15 100)"
    AssertEnglish "percent binds tighter than times", _
        "Set x to 2 times 50%.", "(* 2 (/ 50 100))"
    AssertEnglish "percent on a name", "Set x to rate%.", "(/ rate 100)"
    AssertEnglish "percent as a using default", _
        "To tv using rate of 15%:" & vbLf & "  Give back rate.", _
        "(optional rate Variant (/ 15 100))"

    ' B7: the percent hazard is loud, and the multiplicative verbs
    ' carry the meaning typists intend.
    ' B7.4 (owner revision): a SIMPLE share - one number or name
    ' wearing % or percent - after Increase/Decrease/Add is the
    ' idiomatic growth spelling and now MEANS it, emitting the exact
    ' Grow/Shrink template.
    AssertEnglish "increase by percent grows", "Increase total by 10%.", "(set! total (* total (+ 1 (/ 10 100))))"
    AssertEnglish "decrease by percent-name shrinks", "Decrease x by rate%.", "(set! x (* x (- 1 (/ rate 100))))"
    AssertEnglish "add percent grows too", "Add 10% to total.", "(set! total (* total (+ 1 (/ 10 100))))"
    AssertEnglish "the word percent grows too", "Increase total by 10 percent.", "(+ 1 (/ 10 100))"
    ' ...and % blended into longer arithmetic still refuses - HARDER
    ' than B7, which refused "5 plus 5%" but let "10% times 2" fall
    ' through to a silent additive read of 0.2.
    AssertErrLine "percent blended after arithmetic is refused", "Increase total by 5 plus 5%.", 1
    AssertErrLine "percent blended before arithmetic is refused", "Increase total by 10% times 2.", 1
    AssertErrLine "add with a blended percent is refused", "Add 10% times 2 to total.", 1
    ' A variable merely named percent stays usable as a target.
    AssertEnglish "a target named percent survives the guard", "Increase percent by 5.", "(add! percent 5)"
    AssertEnglish "plain increase untouched", "Increase total by 5.", "(add! total 5)"
    AssertEnglish "plain add untouched", "Add 5 to total.", "(add! total 5)"
    AssertEnglish "grow is multiplicative", "Grow total by 10%.", "(* total (+ 1 (/ 10 100)))"
    AssertEnglish "shrink is multiplicative", "Shrink total by 25%.", "(* total (- 1 (/ 25 100)))"

    ' B7: doubled quotes inside text (VBA's own escape).
    AssertEnglish "doubled quotes read as one", _
        "Set label to " & Chr$(34) & "He said " & Chr$(34) & Chr$(34) & "hi" & Chr$(34) & Chr$(34) & Chr$(34) & ".", _
        "(set! label " & Chr$(34) & "He said \" & Chr$(34) & "hi\" & Chr$(34) & Chr$(34) & ")"

    ' B7: element access over lists.
    AssertEnglish "item n of a list", "Set x to item 3 of found-items.", "(vlaitem found-items 3)"
    AssertEnglish "first of a list", "Set x to first of found-items.", "(vlafirst found-items)"
    AssertEnglish "last of a list", "Set x to last of found-items.", "(vlalast found-items)"

    ' B7: recovery can say what went wrong.
    Dim probP As String
    probP = "Try:" & vbLf & "  Set x to 1." & vbLf & vbLf & _
            "If that fails:" & vbLf & "  Log the problem."
    AssertEnglish "the problem is declared", probP, "(dim vla-problem String)"
    AssertEnglish "the problem is captured before Resume", probP, "(set! vla-problem err.description)"
    AssertEnglish "the problem reads in recovery", probP, "(debug-print vla-problem)"

    ' B7: a bare using-call is refused with a pointer.
    AssertErrLine "standalone Get is refused with a pointer", _
        "To tx2 using a of 1:" & vbLf & "  Give back a." & vbLf & vbLf & _
        "Get tx2 using a of 5.", 4

    ' B7.2: the word percent is the % postfix. (Its refusal pin for
    ' "Increase ... by 10 percent." was superseded by B7.4 - that
    ' sentence now GROWS, pinned above as "the word percent grows
    ' too" - so this slot pins the word-form BLEND refusal instead,
    ' keeping the guard symmetric across both spellings.)
    AssertEnglish "the word percent is the postfix", "Set x to 10 percent.", "(/ 10 100)"
    AssertEnglish "grow by percent-the-word", "Grow total by 10 percent.", "(* total (+ 1 (/ 10 100)))"
    AssertErrLine "percent-the-word blended is refused too", "Increase total by 5 plus 5 percent.", 1

    ' B7.2: nullary value actions - To get <name>: ... Give back ...
    Dim nulP As String
    nulP = "To get answer:" & vbLf & "  Give back 42." & vbLf & vbLf & _
           "Set x to answer." & vbLf & _
           "If answer is 42, log ""yes""."
    AssertEnglish "get-nullary defines a Function", nulP, "(function answer ()"
    AssertEnglish "nullary name is a value in Set", nulP, "(set! x (answer))"
    AssertEnglish "nullary name composes in a condition", nulP, "(= (answer) 42)"
    ' Blessed after this pin's maiden failure: Get is VBA-reserved,
    ' so "To get:" was NEVER definable - CheckName refuses it with
    ' the hyphenate-it guidance, correctly. What the sugar gate must
    ' actually protect is sentences merely STARTING with get reaching
    ' the phrase rules, pinned directly:
    EnglishAddPhrase "get busy", "(debug-print 1)"
    AssertEnglish "a get-rule still reaches the phrase rules", "Get busy.", "(debug-print 1)"


    ' B5: stray characters refuse loudly, on their line, with a hint.
    AssertErrLine "a formula plus sign is refused with words", "Set x to 5 + 3.", 1
    AssertErrLine "stray char names its line mid-program", _
        "Set x to 1." & vbLf & "Set y to 2 * x.", 2

    ' B5: curly double quotes tokenize as straight ones.
    AssertEnglish "curly quotes are accepted", _
        "Set label to " & ChrW$(8220) & "hi" & ChrW$(8221) & ".", "(set! label ""hi"")"

    ' A1: errors carry their source line.
    AssertErrLine "unmatched sentence names its line", _
        "Set x to 1." & vbLf & vbLf & "Blorp the fizzles.", 3
    AssertErrLine "call-check names the CALL's line", _
        "To stamp, with row-number of 1:" & vbLf & "  Log row-number." & vbLf & vbLf & _
        "Stamp with row of 2.", 4

    ' V1: the alias-assignment refusal is raised during sub assembly,
    ' AFTER parsing - the one user-reachable Check error with no token
    ' position in hand at raise time. It now carries the assigning
    ' sentence's line (recorded at MarkAssigned time), so the IDE's
    ' translate-once Check - which retired the progressive-prefix
    ' fallback that used to locate this row by re-translation - still
    ' lands it red on the exact row. First assignment wins when the
    ' alias is assigned more than once.
    AssertErrLine "assigning a Defined alias names the assigning line", _
        "Define pi as 3." & vbLf & "Set total to 1." & vbLf & "Set pi to 4.", 3
    AssertErrLine "alias line: first assignment wins", _
        "Define pi as 3." & vbLf & "Set pi to 4." & vbLf & "Set pi to 5.", 2

    ' V4: every tracked sentence rides in an (at-line N ...) wrapper
    ' carrying its own line - the English half of the three-layer map
    ' (the emitter half is pinned in TestTier1). Block bodies re-wrap
    ' per sentence, so a body line carries ITS line, not the header's.
    ' Vocab test: lines and Explain go through ParseStmt directly and
    ' stay wrapper-free - the 179 vocabulary proofs compare exact VLA
    ' and must never see instrumentation.
    AssertEnglish "V4: a tracked sentence is wrapped with its line", _
        "Set total to 5.", "(at-line 1"
    AssertEnglish "V4: a block-body sentence carries its own line", _
        "Repeat 2 times:" & vbCrLf & "Log counter.", "(at-line 2"

    ' V3: keyed memory. "Create a lookup" declares, initializes, and
    ' registers the name; registration is what turns on the keyed
    ' read, "<lookup> for <key>" - a value anywhere, key binding
    ' tightly like "length of" so reads chain in arithmetic. Without
    ' the declaration "for" stays inert and the sentence refuses -
    ' pinning that no rule's greediness changed for ordinary
    ' programs. Walking a lookup directly is refused (the two
    ' representations disagree on what a walk yields); "keys of" is
    ' the taught spelling. The store rules themselves (at key /
    ' under, one template - V3.2's owner-revised verb) are
    ' vocabulary and carry their proofs in english.vla.
    AssertEnglish "V3: Create a lookup declares and initializes", _
        "Create a lookup called prices.", _
        "(obj-set! prices (vladictnew))"
    AssertEnglish "V3: the keyed read composes in Set", _
        "Create a lookup called prices." & vbLf & "Set p to prices for ""ax-7"".", _
        "(vladictget prices ""ax-7"")"
    AssertEnglish "V3: the keyed read composes in a condition", _
        "Create a lookup called prices." & vbLf & _
        "If prices for ""x"" is greater than 5, log ""big"".", _
        "(> (vladictget prices ""x"") 5)"
    AssertEnglish "V3: the key binds tightly, reads chain in arithmetic", _
        "Create a lookup called prices." & vbLf & "Set p to prices for a plus prices for b.", _
        "(+ (vladictget prices a) (vladictget prices b))"
    AssertErrLine "V3: 'for' stays inert without a lookup declaration", _
        "Set p to prices for 3.", 1
    ' V3.1 (owner revision): the direct walk is no longer refused -
    ' it walks ENTRIES as (key, value) pairs via VlaDictPairs, which
    ' answers identically under both representations, DISSOLVING the
    ' ambiguity V3 had guarded rather than guarding it. key of /
    ' value of are engine-seeded (a structural feature must not
    ' depend on a vocabulary file for its reading), so these pins
    ' run on the prelude-only grammar.
    AssertEnglish "V3.1: For each over a lookup walks its pairs", _
        "Create a lookup called prices." & vbLf & "For each pair in prices:" & vbLf & "Log key of pair.", _
        "(for-each (pair (vladictpairs prices))"
    AssertEnglish "V3.1: value of decides a condition on the pair", _
        "Create a lookup called prices." & vbLf & _
        "For each pair in prices:" & vbLf & _
        "If value of pair is greater than 200, log key of pair.", _
        "(> (vlapairvalue pair) 200)"
    AssertEnglish "V3.1: pair accessors compose in arithmetic (one-liner walk)", _
        "Create a lookup called prices." & vbLf & _
        "For each pair in prices, log value of pair times 10.", _
        "(* (vlapairvalue pair) 10)"
End Sub

Public Sub AssertErrLine(ByVal name As String, ByVal program As String, ByVal wantLine As Long)
    Dim failed As Boolean
    On Error Resume Next
    Err.Clear
    EnglishToVla program
    failed = (Err.Number <> 0)
    On Error GoTo 0
    If Not failed Then
        Report "line: " & name, False, "no error was raised"
    Else
        Report "line: " & name, EnglishLastErrorLine() = wantLine, _
               "got line " & EnglishLastErrorLine() & ", wanted " & wantLine
    End If
End Sub

' ---------------------------------------------------------------------
'  Loader behavior: two-line gathering, function: lines, and the rule
'  that a failing test: must fail the load.
' ---------------------------------------------------------------------
' ---------------------------------------------------------------------
'  A4 claim goldens: the Explain trace's attribution state, pinned
'  through EnglishLastClaim(). Claims are per top-level sentence and
'  first-wins within one, so each program's LAST top-level sentence
'  is the one whose claim these read. Grammar state: runs right after
'  TestEnglishCore, prelude only - core forms and built-in rules.
' ---------------------------------------------------------------------
' L0 pins: raw VLA rows. A row beginning with "(" is a VLA statement
' wherever a sentence can stand; Check validates it; steps cover it.
Private Sub TestRawVlaRows()
    AssertEnglish "vla row: a raw form passes through verbatim", _
        "(set! (range ""h25"") ""vla-row"")", _
        "(set! (range ""h25"") ""vla-row"")"
    AssertEnglish "vla row: composes inside a block", _
        "Repeat 2 times:" & vbCrLf & "(debug-print counter)", _
        "(debug-print counter)"
    AssertErrLine "vla row: unbalanced form names its row", _
        "(set! x 1", 1
    ' (First spelling of this pin used "(if x)" and assumed the
    ' emitter would refuse it - it does not, correctly: a clause-less
    ' If emits as a vacuous "If x Then / End If", legal VBA, per the
    ' 1:1 doctrine. The pin now uses a form that IS malformed - a
    ' set! missing its value - so the emitter's own "expects" error
    ' rides the probe into a red, line-attributed Check row.)
    AssertErrLine "vla row: a malformed form goes red at Check", _
        "(set! x)", 1
    AssertErrLine "vla row: top-level heads are refused with directions", _
        "(sub evil () (set! x 1))", 1
    ' L0.2: a form may now span multiple rows - the old one-row rule
    ' existed only for the progressive-prefix Check's false-red interior
    ' rows, and V1's translate-once Check has no such interior state to
    ' be wrong about (EnTokenize's own header comment, VLA_English.bas).
    AssertEnglish "vla row: a form now spans multiple rows", _
        "(set!" & vbCrLf & "x" & vbCrLf & "(+ 1 2))", _
        "(set! x (+ 1 2))"
    ' The real regression risk isn't the spanning form itself - it's
    ' every row AFTER it. lineNo must advance once per swallowed
    ' newline inside the form, or every later error attributes to the
    ' wrong row (this pin's form spans rows 1-3; row 4 is deliberately
    ' unparseable, and the refusal must still name row 4, not row 2).
    AssertErrLine "vla row: line numbers stay correct after a multi-row form", _
        "(set!" & vbCrLf & "x" & vbCrLf & "(+ 1 2))" & vbCrLf & "Xyzzy plugh nonsense.", 4
    AssertErrLine "vla row: a form spanning rows still refuses if never closed", _
        "(set!" & vbCrLf & "x" & vbCrLf & "(+ 1 2)", 1
End Sub

Private Sub TestExplainTrace()
    ' A When paragraph is claimed by the choices form - the trace no
    ' longer answers with the pre-B generic list.
    AssertClaim "when chain names its case", _
        "Set region to 1." & vbCrLf & "When region is 1: Set region-label to 2.", _
        "the When choices form"
    ' One-line If: the OUTERMOST form is the headline even though the
    ' body fired a phrase rule (shown by Explain as the inner rule).
    AssertClaim "one-line If headlines over its inner rule", _
        "If total is 1, set x to 2.", _
        "the If form"
    AssertClaim "Try block claims", _
        "Try: Set x to 1.", _
        "the Try block"
    ' An action call is a claim too - Explain used to show the stale
    ' structural list for these.
    AssertClaim "action call claims by name", _
        "To tidy-up:" & vbCrLf & "Set x to 1." & vbCrLf & vbCrLf & "Tidy-up.", _
        "a call to the action 'tidy-up'"
    ' The A4 acceptance case: a failed using-call is claimed by the
    ' value call, and the refusal names the parameter check.
    AssertClaimRefusal "failed using-call names the parameter check", _
        "To get commission using sale of 1000 and rate of 5:" & vbCrLf & _
        "Give back sale times rate." & vbCrLf & vbCrLf & _
        "Set fee to commission using bogus of 3.", _
        "a value call: 'commission using ...'", _
        "no parameter called 'bogus'"
    ' B7.4: the simple share is now a success owned by the percent
    ' form; only a % blended into arithmetic still meets the guard,
    ' and the trace names whichever one spoke.
    AssertClaim "a raw VLA row claims as itself", _
        "(set! x 1)", _
        "a raw VLA form"
    AssertClaim "percent form claims its growth", _
        "Increase total by 10%.", _
        "the percent form of Increase/Decrease/Add"
    AssertClaimRefusal "percent guard claims the blend refusal", _
        "Increase total by 10% times 2.", _
        "the percent guard on Increase/Decrease/Add", _
        "mixes % into a longer amount"
End Sub

Public Sub AssertClaim(ByVal name As String, ByVal program As String, ByVal wantClaim As String)
    Dim t As String, d As String
    On Error Resume Next
    t = EnglishToVla(program)
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    If Len(d) > 0 Then
        Report "claim: " & name, False, "unexpected refusal: " & d
    Else
        Report "claim: " & name, EnglishLastClaim() = wantClaim, _
               "got '" & EnglishLastClaim() & "', wanted '" & wantClaim & "'"
    End If
End Sub

Public Sub AssertClaimRefusal(ByVal name As String, ByVal program As String, _
                               ByVal wantClaim As String, ByVal errFrag As String)
    Dim t As String, d As String
    On Error Resume Next
    t = EnglishToVla(program)
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    If Len(d) = 0 Then
        Report "claim: " & name, False, "expected a refusal mentioning '" & errFrag & "' but it translated"
    ElseIf InStr(1, d, errFrag, vbTextCompare) = 0 Then
        Report "claim: " & name, False, "refusal missing '" & errFrag & "': " & d
    Else
        Report "claim: " & name, EnglishLastClaim() = wantClaim, _
               "got claim '" & EnglishLastClaim() & "', wanted '" & wantClaim & "'"
    End If
End Sub

' F.13: rewritten for the vocabulary-grammar migration - every fixture
' below used to be old-format text (pattern => template, test:,
' function:), which the loader no longer understands at all (forms are
' the only path now - see BETA_ROADMAP.md's F.13 entry). Same
' intent, same assertions, new-grammar fixtures.
Private Sub TestLoader()
    Dim v As String
    EnglishResetGrammar
    v = "(english-vla ""make cell {r:text} glow""" & vbLf & _
        "    (set! (. (range {r}) interior.color)" & vbLf & _
        "          vbyellow))" & vbLf & _
        "(test-success ""Make cell B2 glow.""" & vbLf & _
        "    (set! (. (range ""b2"") interior.color) vbyellow))" & vbLf & _
        "(english-function ""doubled of"" vladoubletest)"
    On Error Resume Next
    Err.Clear
    EnglishLoadVocabularyText v, "selftest-vocab"
    Report "multi-line load with passing test", Err.Number = 0, Err.Description
    On Error GoTo 0
    AssertEnglish "loaded rule matches (bare range token)", "Make cell C4 glow.", "(range ""c4"")"
    AssertEnglish "function-word directive registered", "Set d to doubled of x.", "(vladoubletest x)"

    ' G0: the old question-mark slot spelling refuses with directions
    ' (braces are the only spelling; a stale phrasebook fails loud).
    ' Unrelated to file format - IsSlotTok's own refusal, exercised the
    ' same way regardless of which directive carries the pattern text.
    EnglishResetGrammar
    v = "(english-vla ""make cell ?r:text glow"" (set! (range {r}) 1))"
    On Error Resume Next
    Err.Clear
    EnglishLoadVocabularyText v, "selftest-vocab"
    Report "old ?slot spelling refuses with directions", _
           Err.Number <> 0 And InStr(1, Err.Description, "braces", vbTextCompare) > 0, _
           "wanted a braces-teaching refusal, got: " & Err.Description
    On Error GoTo 0

    ' A vocabulary whose test fails must refuse to load.
    EnglishResetGrammar
    v = "(test-success ""Log total."" (this-is-wrong))"
    On Error Resume Next
    Err.Clear
    EnglishLoadVocabularyText v, "selftest-vocab"
    Report "failing test refuses the load", Err.Number <> 0, "load succeeded but should not have"
    On Error GoTo 0

    ' The audit: a duplicate-shaped rule is reported; a clean
    ' phrasebook returns "".
    Dim aud As String
    aud = EnglishAuditText( _
        "(english-vla ""warm cell {r:text}"" (set! (. (range {r}) interior.color) vbyellow))" & vbLf & _
        "(english-vla ""warm cell {x:text}"" (set! (. (range {x}) interior.color) vbyellow))" & vbLf & _
        "(test-success ""Warm cell B2."" (set! (. (range ""b2"") interior.color) vbyellow))")
    Report "audit flags duplicate signatures", InStr(1, aud, "duplicates", vbTextCompare) > 0, _
           "report was: " & Left$(aud, 120)
    aud = EnglishAuditText( _
        "(english-vla ""warm cell {r:text}"" (set! (. (range {r}) interior.color) vbyellow))" & vbLf & _
        "(test-success ""Warm cell B2."" (set! (. (range ""b2"") interior.color) vbyellow))")
    Report "audit passes a clean phrasebook", Len(aud) = 0, "unexpected: " & Left$(aud, 120)

    ' An unclosed form must say so - now VLA.bas's own reader's job
    ' (VlaReadFormsWithLines), not a hand-rolled paren scanner's.
    EnglishResetGrammar
    v = "(english-vla ""broken rule {r:text}"" (never closes"
    On Error Resume Next
    Err.Clear
    EnglishLoadVocabularyText v, "selftest-vocab"
    Report "unclosed form is refused", Err.Number <> 0, "load succeeded but should not have"
    On Error GoTo 0

    ' G0: a bare atom sitting at the top level (not a parenthesized
    ' directive) must refuse cleanly, not crash with a raw VBA "Object
    ' required" - exactly what old-format text fed to this reader by
    ' mistake produces (a bare word sequence, not a syntax error), live-
    ' caught during the F.13 migration and worth pinning permanently.
    EnglishResetGrammar
    v = "stray-atom"
    On Error Resume Next
    Err.Clear
    EnglishLoadVocabularyText v, "selftest-vocab"
    Report "a bare top-level atom refuses cleanly", _
           Err.Number <> 0 And InStr(1, Err.Description, "bare word", vbTextCompare) > 0, _
           "wanted a bare-word teaching refusal, got: " & Err.Description
    On Error GoTo 0

    EnglishResetGrammar
End Sub

' ---------------------------------------------------------------------
'  Helper units.
' ---------------------------------------------------------------------
' =====================================================================
'  SEC.8 - the provenance gate's PURE half.
'
'  What this can and cannot cover, stated plainly because SEC.8's own
'  roadmap entry promises the same honesty SEC.13 and F.10 gave their
'  coverage gaps: the pure suite does no COM and cannot fabricate a
'  Mark-of-the-Web, so NOTHING here proves that a real internet-marked
'  workbook is actually refused - that is a live test the owner runs,
'  plus check_sec8_provenance_gate.ps1's static pin on the call sites.
'  What IS covered here is the whole of the DECISION: the stream parse,
'  the path classification, the policy that combines them, and the
'  guard's own refusal, each exercised through the memo seam so no file
'  is needed. Every assertion below fails if the line it targets is
'  removed - checked by writing each one to distinguish the real
'  behaviour from the most plausible wrong one, not merely to pass.
' =====================================================================
Private Sub TestSec8Provenance()
    ' --- the stream parse ---
    Report "SEC.8 parses ZoneId=3 (CRLF, the real stream's own form)", _
           VlaParseZoneIdentifier("[ZoneTransfer]" & vbCrLf & "ZoneId=3" & vbCrLf) = VlaZoneInternet, _
           "got " & VlaParseZoneIdentifier("[ZoneTransfer]" & vbCrLf & "ZoneId=3" & vbCrLf)
    Report "SEC.8 parses ZoneId=3 with bare LF too", _
           VlaParseZoneIdentifier("[ZoneTransfer]" & vbLf & "ZoneId=3") = VlaZoneInternet, _
           "bare-LF stream did not parse"
    Report "SEC.8 parses ZoneId=0 as local machine, not as 'no mark'", _
           VlaParseZoneIdentifier("ZoneId=0") = VlaZoneLocalMachine, _
           "got " & VlaParseZoneIdentifier("ZoneId=0")
    Report "SEC.8 ignores ReferrerUrl (attacker-supplied text, never a decision input)", _
           VlaParseZoneIdentifier("[ZoneTransfer]" & vbCrLf & _
                                  "ReferrerUrl=https://evil.invalid/ZoneId=0" & vbCrLf & _
                                  "ZoneId=3") = VlaZoneInternet, _
           "a ReferrerUrl containing 'ZoneId=0' changed the verdict"
    Report "SEC.8 reads an absent ZoneId line as unreadable", _
           VlaParseZoneIdentifier("[ZoneTransfer]" & vbCrLf & "HostUrl=x") = VlaZoneUnreadable, _
           "a stream with no ZoneId did not read as unreadable"
    Report "SEC.8 reads an empty stream as unreadable", _
           VlaParseZoneIdentifier("") = VlaZoneUnreadable, "empty stream was not unreadable"
    ' The rounding trap: IsNumeric+CLng alone would turn "3.7" into 4
    ' (VlaZoneRestricted) - a decimal must never land on a zone id it
    ' does not name. Fails if AllDigits is dropped for a bare IsNumeric.
    Report "SEC.8 refuses a decimal ZoneId rather than rounding it onto a real zone", _
           VlaParseZoneIdentifier("ZoneId=3.7") = VlaZoneUnreadable, _
           "got " & VlaParseZoneIdentifier("ZoneId=3.7") & " - a decimal rounded onto a zone id"
    ' The stream's bytes travel WITH the file, so its content is
    ' attacker-authored. An all-digit value too big for a Long would
    ' overflow CLng and raise error 6 out of a function this module
    ' calls pure and total. Fails if the length cap is dropped.
    Dim overflowed As Boolean
    On Error Resume Next
    Err.Clear
    Report "SEC.8 reads an oversized ZoneId as unreadable", _
           VlaParseZoneIdentifier("ZoneId=99999999999999") = VlaZoneUnreadable, _
           "an oversized ZoneId did not read as unreadable"
    overflowed = (Err.Number <> 0)
    On Error GoTo 0
    Report "SEC.8 does not raise on an oversized ZoneId (no CLng overflow escapes)", _
           Not overflowed, "parsing an oversized ZoneId raised an error"
    ' A number that is not one of the six real zones is not a zone.
    Report "SEC.8 reads a ZoneId outside the real zone range as unreadable", _
           VlaParseZoneIdentifier("ZoneId=7") = VlaZoneUnreadable, _
           "got " & VlaParseZoneIdentifier("ZoneId=7") & " - an unknown zone number was accepted"
    Report "SEC.8 refuses a negative ZoneId", _
           VlaParseZoneIdentifier("ZoneId=-1") = VlaZoneUnreadable, _
           "a negative ZoneId was accepted"

    ' --- the path classification ---
    Report "SEC.8 counts a drive-letter path as demonstrably local", _
           VlaPathIsDemonstrablyLocal("C:\Users\someone\book.xlsx"), "C:\ path was not local"
    Report "SEC.8 does NOT count a UNC share as demonstrably local", _
           Not VlaPathIsDemonstrablyLocal("\\server\share\book.xlsx"), "UNC path counted as local"
    Report "SEC.8 does NOT count an https:// (WebDAV/SharePoint) path as local", _
           Not VlaPathIsDemonstrablyLocal("https://contoso.sharepoint.com/x/book.xlsx"), _
           "https path counted as local"
    Report "SEC.8 does NOT count an http:// path as local", _
           Not VlaPathIsDemonstrablyLocal("http://host/book.xlsx"), "http path counted as local"
    Report "SEC.8 counts a never-saved workbook (empty path) as local", _
           VlaPathIsDemonstrablyLocal(""), "unsaved workbook was not treated as local"

    ' --- the policy ---
    Report "SEC.8 refuses a workbook positively marked internet-zone", _
           VlaProvenanceRefuses(VlaZoneInternet, True), "ZoneId=3 on a local path was allowed"
    Report "SEC.8 refuses a workbook marked restricted-zone", _
           VlaProvenanceRefuses(VlaZoneRestricted, True), "ZoneId=4 was allowed"
    Report "SEC.8 allows a workbook positively marked local-machine", _
           Not VlaProvenanceRefuses(VlaZoneLocalMachine, True), "ZoneId=0 was refused"
    Report "SEC.8 allows an intranet-marked workbook", _
           Not VlaProvenanceRefuses(VlaZoneIntranet, True), "ZoneId=1 was refused"
    ' The two halves of the forced positive-only answer. An unreadable
    ' mark is safe ONLY where a mark could have been stored and was not.
    Report "SEC.8 allows an unreadable mark on a demonstrably local path", _
           Not VlaProvenanceRefuses(VlaZoneUnreadable, True), _
           "an ordinary local file with no mark was refused - this would break every normal workbook"
    Report "SEC.8 refuses an unreadable mark where no mark could have been stored", _
           VlaProvenanceRefuses(VlaZoneUnreadable, False), _
           "unknown provenance on a UNC/URL path was allowed"

    ' --- the guard itself, through the memo seam (no file, no COM) ---
    Dim refused As Boolean
    VlaProvenanceResetMemo
    VlaProvenanceSeedMemoForTest "C:\Users\someone\mailed.xlsx", VlaZoneInternet
    On Error Resume Next
    Err.Clear
    VlaProvenanceGuardCaptured "email something out of Excel"
    refused = (Err.Number <> 0)
    Dim msg As String
    msg = Err.Description
    On Error GoTo 0
    Report "SEC.8 guard refuses an external effect on an internet-marked workbook", _
           refused, "the guard allowed the call"
    Report "SEC.8 refusal names the verb the program actually tried", _
           InStr(msg, "email something out of Excel") > 0, "verb missing from: " & msg
    Report "SEC.8 refusal tells the person how to grant it (Windows' own Unblock)", _
           InStr(msg, "Unblock") > 0, "no Unblock instruction in: " & msg

    VlaProvenanceResetMemo
    VlaProvenanceSeedMemoForTest "C:\Users\someone\mine.xlsx", VlaZoneLocalMachine
    On Error Resume Next
    Err.Clear
    VlaProvenanceGuardCaptured "email something out of Excel"
    refused = (Err.Number <> 0)
    On Error GoTo 0
    Report "SEC.8 guard allows the same effect on a local-marked workbook", _
           Not refused, "an ordinary local workbook was refused"
    VlaProvenanceResetMemo
End Sub

Private Sub TestHelpers()
    ' The BGR swap: naive CLng("&HFF69B4") would be 16738740; the
    ' correct VBA color for hot pink is RGB(255,105,180).
    Report "VlaColor #FF69B4 is RGB(255,105,180)", VlaColor("#FF69B4") = RGB(255, 105, 180), _
           "got " & VlaColor("#FF69B4")
    Report "VlaColor bare hex", VlaColor("FF69B4") = RGB(255, 105, 180), "bare hex failed"
    Report "VlaColor numeric passthrough", VlaColor(255) = 255, "got " & VlaColor(255)
    Dim bad As Boolean
    On Error Resume Next
    VlaColor "#GGGGGG"
    bad = (Err.Number <> 0)
    On Error GoTo 0
    Report "VlaColor rejects non-hex", bad, "accepted #GGGGGG"

    ' N1: VlaCheckSheetName / VlaCheckRangeName - name-range and
    ' add-sheet-called's shared refusal, checked directly (no live
    ' Excel object needed - pure string rules).
    Dim ok As Boolean
    On Error Resume Next
    Err.Clear
    VlaCheckSheetName "Demo"
    ok = (Err.Number = 0)
    On Error GoTo 0
    Report "VlaCheckSheetName accepts an ordinary name", ok, "refused a legal name"

    On Error Resume Next
    Err.Clear
    VlaCheckSheetName ""
    bad = (Err.Number <> 0)
    On Error GoTo 0
    Report "VlaCheckSheetName rejects an empty name", bad, "accepted an empty sheet name"

    On Error Resume Next
    Err.Clear
    VlaCheckSheetName "Q1/Report"
    bad = (Err.Number <> 0)
    On Error GoTo 0
    Report "VlaCheckSheetName rejects a forbidden character", bad, "accepted a slash"

    On Error Resume Next
    Err.Clear
    VlaCheckSheetName String(32, "x")
    bad = (Err.Number <> 0)
    On Error GoTo 0
    Report "VlaCheckSheetName rejects over 31 characters", bad, "accepted a 32-char name"

    On Error Resume Next
    Err.Clear
    VlaCheckSheetName "History"
    bad = (Err.Number <> 0)
    On Error GoTo 0
    Report "VlaCheckSheetName rejects the reserved word History", bad, "accepted History"

    On Error Resume Next
    Err.Clear
    VlaCheckRangeName "prices"
    ok = (Err.Number = 0)
    On Error GoTo 0
    Report "VlaCheckRangeName accepts an ordinary name", ok, "refused a legal name"

    ' The bug report this whole helper exists for: a hyphenated name,
    ' the same shape VLA's own identifiers use freely, rejected here
    ' because Range.Name is a different, stricter namespace.
    On Error Resume Next
    Err.Clear
    VlaCheckRangeName "demo-table"
    bad = (Err.Number <> 0)
    On Error GoTo 0
    Report "VlaCheckRangeName rejects a hyphen", bad, "accepted demo-table"

    On Error Resume Next
    Err.Clear
    Dim d As String
    VlaCheckRangeName "demo-table"
    d = Err.Description
    On Error GoTo 0
    Report "VlaCheckRangeName's message offers a legal alternative", _
           InStr(1, d, "demo_table", vbTextCompare) > 0, "no suggestion in: " & d

    On Error Resume Next
    Err.Clear
    VlaCheckRangeName "A1"
    bad = (Err.Number <> 0)
    On Error GoTo 0
    Report "VlaCheckRangeName rejects a cell-shaped name", bad, "accepted A1"

    On Error Resume Next
    Err.Clear
    VlaCheckRangeName "9prices"
    bad = (Err.Number <> 0)
    On Error GoTo 0
    Report "VlaCheckRangeName rejects a leading digit", bad, "accepted 9prices"

    ' V3: the VlaDict family. First against whatever representation
    ' this machine gives (Scripting.Dictionary on Windows), then
    ' explicitly against the pair-Collection fallback - the branch a
    ' Windows dev machine would otherwise never execute. The contract
    ' both must honor: add-or-replace, first key spelling survives
    ' replacement, case-insensitive lookup, insertion-ordered keys,
    ' one .Count, loud miss naming the key.
    Dim dd As Object
    Set dd = VlaDictNew()
    VlaDictSet dd, "ax-7", 100
    VlaDictSet dd, "bx-2", 250
    Report "VlaDict: set/get roundtrip", VlaDictGet(dd, "ax-7") = 100, _
           "got " & VlaDictGet(dd, "ax-7")
    VlaDictSet dd, "AX-7", 120
    Report "VlaDict: store-again replaces", VlaDictGet(dd, "ax-7") = 120, _
           "expected 120 after replace"
    Report "VlaDict: keys are case-insensitive", VlaDictGet(dd, "Bx-2") = 250, _
           "case-varied key missed"
    Report "VlaDict: count stays 2 after replace", VlaCount(dd) = 2, _
           "got " & VlaCount(dd)
    Dim kk As Collection
    Dim entriesOk As Boolean
    Set kk = VlaDictKeys(dd)
    ' Guarded shape (V3.1 hardening): And evaluates BOTH sides, so a
    ' wrong count must not reach .Item(1) - a regression should FAIL,
    ' not crash the suite.
    entriesOk = False
    If kk.Count = 2 Then
        entriesOk = CStr(kk.Item(1)) = "ax-7" And CStr(kk.Item(2)) = "bx-2"
    End If
    Report "VlaDict: keys walk in first-stored order", entriesOk, _
           "got " & kk.Count & " keys"
    Dim missed As Boolean
    Dim missMsg As String
    On Error Resume Next
    VlaDictGet dd, "zz-9"
    missed = (Err.Number <> 0)
    missMsg = Err.Description
    On Error GoTo 0
    Report "VlaDict: a miss is loud and names the key", _
           missed And InStr(1, missMsg, "zz-9", vbTextCompare) > 0, _
           "miss message: " & missMsg
    ' The fallback branch, forced: a plain Collection through the
    ' same helpers, same contract.
    Dim fb As Object
    Set fb = New Collection
    VlaDictSet fb, "ax-7", 100
    VlaDictSet fb, "AX-7", 120
    VlaDictSet fb, "bx-2", 250
    Report "VlaDict fallback: replace + case-insensitive get", _
           VlaDictGet(fb, "Ax-7") = 120, "fallback get failed"
    Set kk = VlaDictKeys(fb)
    entriesOk = False
    If kk.Count = 2 Then
        entriesOk = CStr(kk.Item(1)) = "ax-7" And CStr(kk.Item(2)) = "bx-2"
    End If
    Report "VlaDict fallback: first spelling survives, order kept", entriesOk, _
           "got " & kk.Count & " keys"
    Report "VlaDict fallback: VlaCount agrees", VlaCount(fb) = 2, _
           "got " & VlaCount(fb)
    ' V3.1: pair walks - the entries as (key, value) pairs, same
    ' order and content under both representations. dd holds
    ' ax-7 -> 120 (after replace) and bx-2 -> 250; so does fb.
    Dim pp As Collection
    Set pp = VlaDictPairs(dd)
    entriesOk = False
    If pp.Count = 2 Then
        entriesOk = CStr(VlaPairKey(pp.Item(1))) = "ax-7" And VlaPairValue(pp.Item(1)) = 120 _
                    And CStr(VlaPairKey(pp.Item(2))) = "bx-2" And VlaPairValue(pp.Item(2)) = 250
    End If
    Report "VlaDict pairs: entries in order, key and value readable", entriesOk, _
           "got " & pp.Count & " pairs"
    Set pp = VlaDictPairs(fb)
    entriesOk = False
    If pp.Count = 2 Then
        entriesOk = CStr(VlaPairKey(pp.Item(1))) = "ax-7" And VlaPairValue(pp.Item(1)) = 120 _
                    And CStr(VlaPairKey(pp.Item(2))) = "bx-2" And VlaPairValue(pp.Item(2)) = 250
    End If
    Report "VlaDict pairs: fallback deals identical pairs", entriesOk, _
           "got " & pp.Count & " pairs"
End Sub

' ---------------------------------------------------------------------
'  Assertion plumbing.
' ---------------------------------------------------------------------
' ---------------------------------------------------------------------
'  V2: program identity - the pure name-derivation core of the multi-
'  program IDE, pinned without a workbook in hand. The contract:
'  named workspaces "Frazaro (<name>)" derive an alnum-only tag (max
'  10) that scopes their module (Frazaro_EN_<tag>) and snapshot
'  prefixes (VLAu_<tag>_...); the default workspace keeps tag "Main"
'  and module Frazaro_EN_Sheet (IO.6: was EN_Sheet/EN_<tag> - renamed
'  to prevent, not just catch, a collision with a foreign module of
'  the same name). No-underscore-in-tags is what keeps snapshot names
'  parseable, so the sanitizer is pinned hard.
' ---------------------------------------------------------------------
Private Sub TestIdeNaming()
    CheckV "ide: default sheet is a workspace", VlaIdeIsWorkspaceName("Frazaro"), True
    CheckV "ide: named sheet is a workspace", VlaIdeIsWorkspaceName("Frazaro (Invoices)"), True
    CheckV "ide: a plain sheet is not", VlaIdeIsWorkspaceName("Invoices"), False
    CheckV "ide: default tag is Main", VlaIdeProgramTag("Frazaro"), "Main"
    CheckV "ide: named tag is the name", VlaIdeProgramTag("Frazaro (Invoices)"), "Invoices"
    CheckV "ide: tag keeps letters and digits only", VlaIdeProgramTag("Frazaro (Q1 Data!)"), "Q1Data"
    CheckV "ide: tag never carries an underscore", VlaIdeProgramTag("Frazaro (My_Copy)"), "MyCopy"
    CheckV "ide: tag caps at 10", VlaIdeProgramTag("Frazaro (ABCDEFGHIJKLMNOP)"), "ABCDEFGHIJ"
    CheckV "ide: default module stays Frazaro_EN_Sheet", VlaIdeModuleFor("Frazaro"), "Frazaro_EN_Sheet"
    CheckV "ide: named module is Frazaro_EN_tag", VlaIdeModuleFor("Frazaro (Invoices)"), "Frazaro_EN_Invoices"
End Sub

' ---------------------------------------------------------------------
'  V5: the ribbon contract. VlaRibbonXml (VLA_Build) and the
'  VlaRibbonAction dispatcher (VLA_IDE) meet only inside Excel's
'  ribbon loader, which fails SILENTLY on a mismatch - so the pins
'  hold the two ends together here instead: every button id the XML
'  declares must be one the dispatcher's Select Case knows, and
'  every button (fifteen now, with IN.6's Interpret/Interpret and
'  Trace, IN.4's Show Compiled VBA, DI.2 Pass 3's Register for
'  Auto-Load, and DI.2 Pass 4's Uninstall Frazaro) must ride the one
'  callback. Twenty as of this session's Lint VLA button (owner
'  request - VLA_Lint.VlaLintFormat, on demand, over any .vla file).
'  Twenty-two with GO.6's own Load Phrasebook button.
'  (Requires VLA_Build in
'  the dev project, which the build flow already assumes.)
' ---------------------------------------------------------------------
' ---------------------------------------------------------------------
'  V5.2: the Undo target scan, pinned against the first field report
'  - a screenshot of the Undo dialog naming phantom sheets minted
'  from comment prose ("and", "already", "the") and from a raw VLA
'  row's set! (the ! scanner checked only its left flank). Every
'  phantom class from that screenshot is a pin here; the legitimate
'  reaches are pinned beside them so the fix cannot overcorrect.
' ---------------------------------------------------------------------
Private Sub TestUndoScan()
    CheckV "scan: a sentence target is found", VlaIdeScanTargets("Go to sheet Data."), "data"
    CheckV "scan: comment prose mints no targets", _
           VlaIdeScanTargets("# fails at the missing sheet and runs" & vbLf & "# the sheet already exists" & vbLf & "# removes a sheet the run created"), ""
    CheckV "scan: a commented-out Go to is inactive", VlaIdeScanTargets("# Go to sheet Ghost."), ""
    CheckV "scan: a raw VLA set! is punctuation, not a sheet", VlaIdeScanTargets("(set! (range " & Chr$(34) & "h25" & Chr$(34) & ") 5)"), ""
    CheckV "scan: qualified reference still found", VlaIdeScanTargets("Put 1 into cell Output!H16."), "output"
    CheckV "scan: spaced sheet qualifier keeps its case", VlaIdeScanTargets("Put 1 into cell 'Q1 Data'!A1."), "Q1 Data"
    CheckV "scan: Add sheet called chains to the quoted name", VlaIdeScanTargets("Add sheet called " & Chr$(34) & "Report" & Chr$(34) & "."), "Report"
End Sub

Private Sub TestBuildRibbon()
    Dim x As String
    x = VlaRibbonXml()
    CheckV "ribbon: one well-formed customUI document", _
           (InStr(x, "<customUI ") > 0 And InStr(x, "</customUI>") > 0), True
    Dim ids As Variant
    ids = Array("VlaSetup", "VlaRegister", "VlaAddProgram", "VlaImport", "VlaReload", "VlaUninstall", _
                "VlaCheck", "VlaInterpret", "VlaInterpretTrace", "VlaRun", "VlaRunTrace", "VlaShowVba", _
                "VlaTranslateVla", "VlaTranslateVba", "VlaUndo", "VlaPhrases", "VlaLoadPhrasebook", _
                "VlaExportExpanded", "VlaRuleCoverage", "VlaLintVla", "VlaFeedback", "VlaOpenCli")
    Dim i As Long
    Dim missing As String
    For i = LBound(ids) To UBound(ids)
        If InStr(x, "id=" & Chr$(34) & ids(i) & Chr$(34)) = 0 Then missing = missing & " " & ids(i)
    Next
    Report "ribbon: all twenty-two command ids present", Len(missing) = 0, "missing:" & missing
    Report "ribbon: every button rides the one callback", _
           CountOcc(x, "onAction=" & Chr$(34) & "VlaRibbonAction" & Chr$(34)) = 22, _
           "got " & CountOcc(x, "onAction=" & Chr$(34) & "VlaRibbonAction" & Chr$(34))
End Sub

' =====================================================================
'  S1: the whole-corpus goldens. Translates instructions.txt through the
'  Run path's own machinery - vocabulary loaded fresh from disk, step
'  tracking on (its default and the Run state) - and writes TWO
'  golden files beside the program: the complete generated VLA
'  (instructions_golden.vla) and the transpiled VBA (instructions_golden.vba).
'
'  git diff on the goldens is the total-pipeline witness: a pass that
'  claims to preserve behavior regenerates them and shows an EMPTY
'  diff; a pass that changes translation shows exactly what changed,
'  everywhere - not just where a fragment pin happened to look. Run
'  in every verification loop; commit golden changes DELIBERATELY
'  (a golden change is a contract change - say so in the commit).
'
'  Dev-only like everything in this module. Paths default beside the
'  dev workbook (the git working tree); pass others to golden a
'  different corpus. Leaves the loaded vocabulary in place, like a
'  Check would.
'
'  F.9: stable section markers. Every ' vla:N tag already names the
'  ONE line a statement came from; nothing named which instructions.txt
'  PARAGRAPH a reader is inside without decoding a tag by hand.
'  InsertSectionMarkers below splices one (raw "' ---- ...") line
'  into `vla` before the first statement of each blank-line-delimited
'  paragraph (instructions.txt's own documented structural unit - "blank
'  line ends a block"), keyed to that paragraph's own starting line
'  and first sentence - so a git diff's surrounding context, or a
'  human scrolling either golden, names the section without leaving
'  the file. Entirely post-processing: EnglishToVla/VlaTranspile run
'  unmodified (raw is an existing, well-supported core form; markers
'  carry no at-line tag of their own, the same "not user-authored"
'  treatment raw/begin/at-line already get - see EmitStmt's own
'  exclusion list). Confined to this dev-only golden writer; nothing
'  a real user's program compiles through changes.
' =====================================================================
Public Function VlaWriteGoldens(Optional ByVal programPath As String = "", _
                                Optional ByVal vocabPath As String = "") As Boolean
    On Error GoTo failed
    ' S1.1 (maiden-run incident): the first draft assumed the corpus
    ' beside the workbook; the real layout keeps it in a scripts
    ' subfolder. Defaults now PROBE the known layouts instead of
    ' assuming one, and the refusal names every location tried plus
    ' the escape hatch. The deeper fix - one workbook-scoped path
    ' seam instead of per-tool guesses - is proposed item S6 on the
    ' Alpha 3 roadmap.
    If Len(programPath) = 0 Then programPath = FindDevFile("instructions.txt")
    If Len(vocabPath) = 0 Then vocabPath = FindDevFile("english.vla")
    If Len(Dir$(programPath)) = 0 Then Err.Raise 53, "VLA_Tests", "program file not found: " & programPath
    If Len(Dir$(vocabPath)) = 0 Then Err.Raise 53, "VLA_Tests", "vocabulary file not found: " & vocabPath

    Debug.Print "===== WRITE GOLDENS ====="
    EnglishResetGrammar
    Dim rules As Long
    rules = EnglishLoadVocabulary(vocabPath)
    Debug.Print "  vocabulary: " & rules & " rules from " & vocabPath

    ' The Run state: step tracking on. It is the module default and
    ' there is no state getter to save/restore; on IS the resting
    ' state, so the generator asserts it rather than assuming it.
    EnglishStepTracking True

    Dim program As String
    program = ReadTextFileUtf8(programPath)
    Dim vla As String
    vla = EnglishToVla(program)
    vla = InsertSectionMarkers(vla, program)   ' F.9
    Dim vbaText As String
    vbaText = VlaTranspile(vla)

    Dim vlaPath As String, vbaPath As String
    vlaPath = GoldenPathFor(programPath, ".vla")
    vbaPath = GoldenPathFor(programPath, ".vba")
    WriteTextFile vlaPath, vla
    WriteTextFile vbaPath, vbaText
    Debug.Print "  wrote " & Len(vla) & " chars -> " & vlaPath
    Debug.Print "  wrote " & Len(vbaText) & " chars -> " & vbaPath
    Debug.Print "  (git diff is the witness: empty = behavior preserved;"
    Debug.Print "   a golden change is a contract change - commit it deliberately)"
    Debug.Print "===== GOLDENS WRITTEN ====="
    VlaWriteGoldens = True
    Exit Function
failed:
    Debug.Print "  GOLDENS FAILED: " & Err.Description
End Function

' =====================================================================
'  IN.3.5: the interpreter's own VlaWriteGoldens. instructions_golden.vla/
'  .vba witness the WHOLE-CORPUS emitter pipeline; nothing played that
'  role for the interpreter, and IN.9 makes the interpreter the
'  runtime, not a side experiment - "making the interpreter the
'  runtime demotes the primary golden with it," per the roadmap entry
'  this implements. VlaWriteInterpreterGoldens is that replacement,
'  scoped honestly to what VLA_Interpreter.bas can run TODAY: the
'  walking skeleton's own ten-form demo program (VlaSkeletonDemoVla),
'  not instructions.txt - the interpreter cannot run the real corpus until
'  IN.2 exists, and this golden's whole point is to be a behavior
'  witness for code that actually runs, not an aspiration. Same
'  discipline as VlaWriteGoldens: run by hand from the Immediate
'  window, commit a changed golden deliberately, empty git diff is
'  the pass.
' =====================================================================
Public Function VlaWriteInterpreterGoldens() As Boolean
    On Error GoTo failed
    Debug.Print "===== WRITE INTERPRETER GOLDEN ====="
    Dim english As String
    Dim vla As String
    vla = VLA_Interpreter.VlaSkeletonDemoVla(english)
    VLA_Interpreter.VlaInterpret vla
    Dim effectLog As String
    effectLog = VLA_Interpreter.VlaInterpreterEffectLog()

    Dim goldenPath As String
    goldenPath = ThisWorkbook.Path & Application.PathSeparator & "scripts" & _
                 Application.PathSeparator & "interpreter_golden.txt"
    WriteTextFile goldenPath, effectLog
    Debug.Print "  wrote " & Len(effectLog) & " chars -> " & goldenPath
    Debug.Print "  (git diff is the witness: empty = interpreter behavior preserved;"
    Debug.Print "   a golden change is a contract change - commit it deliberately)"
    Debug.Print "===== INTERPRETER GOLDEN WRITTEN ====="
    VlaWriteInterpreterGoldens = True
    Exit Function
failed:
    Debug.Print "  INTERPRETER GOLDEN FAILED: " & Err.Description
End Function

' Convenience: both golden writers in one call (owner request, same
' spirit as VLA_Tests_Host.bas's VlaSelfTests). Lives here rather than
' split across a wrapper module - both VlaWriteGoldens and
' VlaWriteInterpreterGoldens already live in this module, so this is a
' same-module convenience, not a new cross-module dependency. Each
' writer keeps running (and printing) even if the other fails, so one
' broken golden never hides the other's result; the grammar reset both
' already do individually happens twice here, harmlessly redundant -
' same "reload vocab afterwards" caveat VlaSelfTest already carries.
Public Function VlaGoldens() As Boolean
    Dim corpusOk As Boolean, interpOk As Boolean
    corpusOk = VlaWriteGoldens()
    interpOk = VlaWriteInterpreterGoldens()
    Debug.Print "===== VLA GOLDENS: " & IIf(corpusOk, "corpus PASS", "corpus FAILED") & _
                ", " & IIf(interpOk, "interpreter PASS", "interpreter FAILED") & " ====="
    Debug.Print "(grammar was reset twice - Reload vocabulary before the next Run)"
    VlaGoldens = corpusOk And interpOk
End Function

' VLALINT.0: unlike VlaWriteGoldens (which REGENERATES a golden and
' leaves git diff as the check), scripts/english.vla/prelude.vla are
' hand-edited directly - nothing in the runtime writes them
' programmatically - so the only "is this still canonically linted"
' answer has to come from reformatting each in memory (VLA_Lint.
' VlaLintFormat) and comparing against what's actually on disk.
' Dev-only, like everything else in this module; run by hand from the
' Immediate window (?VlaLintCheck(d): ?d) after any hand-edit to either
' corpus file - same discipline VlaGoldens already asks for.
' Deliberately NOT wired into VlaSelfTest's own dispatch list: nothing
' there does real disk I/O beyond FindDevFile's own cheap Dir$
' existence probe (VlaWriteGoldens/VlaGoldens are themselves excluded
' for the identical reason) - a full read+reformat+compare of two
' multi-hundred-line files would be the first entry to break that norm.
' Deliberately scoped to exactly these two files - NOT
' scripts/english_expanded.vla, G-EXPANDER's own artifact: that file is
' a flat, one-form-per-line machine-scan artifact for AS.1's own future
' line-based scan, and running the full pretty-printer over it would
' split every english-vla/test-success pair across two lines, breaking
' that design. A future "lint every .vla file" pass must not silently
' sweep it in.
Public Function VlaLintCheck(ByRef detail As String) As Boolean
    Dim names As Variant
    names = Array("english.vla", "prelude.vla")
    Dim bad As String
    Dim i As Long
    For i = LBound(names) To UBound(names)
        Dim p As String
        p = ""
        On Error Resume Next
        p = FindDevFile(CStr(names(i)))
        On Error GoTo 0
        If Len(p) = 0 Then
            bad = bad & " " & names(i) & " (not found)"
        Else
            Dim orig As String, formatted As String
            orig = ReadTextFileUtf8(p)
            formatted = VLA_Lint.VlaLintFormat(orig)
            Dim origLines() As String, fmtLines() As String
            origLines = Split(Replace(orig, vbCrLf, vbLf), vbLf)
            fmtLines = Split(Replace(formatted, vbCrLf, vbLf), vbLf)
            Dim maxJ As Long
            maxJ = UBound(origLines)
            If UBound(fmtLines) > maxJ Then maxJ = UBound(fmtLines)
            Dim firstDiff As Long
            firstDiff = -1
            Dim j As Long
            For j = 0 To maxJ
                Dim oLine As String, fLine As String
                oLine = "": fLine = ""
                If j <= UBound(origLines) Then oLine = origLines(j)
                If j <= UBound(fmtLines) Then fLine = fmtLines(j)
                If oLine <> fLine Then
                    firstDiff = j + 1
                    Exit For
                End If
            Next
            If firstDiff > 0 Then
                bad = bad & " " & names(i) & " (not canonically linted - first differing line " & firstDiff & ")"
            End If
        End If
    Next i
    detail = bad
    VlaLintCheck = (Len(bad) = 0)
End Function

' =====================================================================
'  VLALINT.0 self-test - cheap, in-memory only (no file I/O; the real
'  corpus-file check is VlaLintCheck above, deliberately NOT dispatched
'  from here). Each pin's expected string is computed by hand, exactly,
'  so a mismatch names precisely what diverged rather than a fuzzy
'  substring guess.
' =====================================================================
Private Sub TestVlaLint()
    Dim s1 As String
    s1 = "(vladictset prices ""ax-7"" 100)"
    Report "vlalint: a short form stays flat regardless of its 3-argument count", _
           VLA_Lint.VlaLintFormat(s1) = s1 & vbCrLf & vbCrLf & vbCrLf, _
           "got: [" & VLA_Lint.VlaLintFormat(s1) & "]"

    Dim functor As String, a1 As String, a2 As String, a3 As String
    functor = "plain-call"
    a1 = String$(40, "a")
    a2 = String$(40, "b")
    a3 = String$(40, "c")
    Dim s2 As String
    s2 = "(" & functor & " " & a1 & " " & a2 & " " & a3 & ")"
    ' Owner correction (third pass, settling where the second pass
    ' overreached): "plain-call" (10 chars) sits at or under
    ' MAX_INLINE_FUNCTOR, so it keeps PpVerticalizeInline's original
    ' shape - functor and first argument share the opening line, the
    ' rest align beneath it - same as VLA's own short control-flow
    ' keywords (if, begin, quote-if, ...) now do. Only a functor LONGER
    ' than the threshold (expected2b below) falls back to PpVerticalize's
    ' full split.
    Dim expected2 As String
    expected2 = "(" & functor & " " & a1 & vbCrLf & _
                Space$(12) & a2 & vbCrLf & _
                Space$(12) & a3 & ")" & vbCrLf & vbCrLf & vbCrLf
    Report "vlalint: a short-functor form too wide to fit flat keeps PpVerticalizeInline's shape - functor and first argument share the opening line", _
           VLA_Lint.VlaLintFormat(s2) = expected2, _
           "got: [" & VLA_Lint.VlaLintFormat(s2) & "]"

    ' The actual motivating case: a functor longer than
    ' MAX_INLINE_FUNCTOR (table-property-family's own real length is 22;
    ' this pin uses it directly) falls back to PpVerticalize's full
    ' split even though the arguments themselves are short - length of
    ' the FUNCTOR decides, not overall line width alone.
    Dim longFunctor As String
    longFunctor = "table-property-family"
    Dim s2b As String
    s2b = "(" & longFunctor & " " & a1 & " " & a2 & " " & a3 & ")"
    Dim expected2b As String
    expected2b = "(" & longFunctor & vbCrLf & _
                 Space$(4) & a1 & vbCrLf & _
                 Space$(4) & a2 & vbCrLf & _
                 Space$(4) & a3 & ")" & vbCrLf & vbCrLf & vbCrLf
    Report "vlalint: a functor longer than MAX_INLINE_FUNCTOR falls back to PpVerticalize's full split", _
           VLA_Lint.VlaLintFormat(s2b) = expected2b, _
           "got: [" & VLA_Lint.VlaLintFormat(s2b) & "]"

    ' The regression this whole correction exists for, caught live: a
    ' short control-flow-shaped functor (if - 2 chars) must NOT get
    ' PpVerticalize's full split just because it doesn't fit flat -
    ' "(if" alone on its own line, with the condition and both branches
    ' each on their own line below, reads as alien Lisp, not a fix.
    Dim s2c As String
    s2c = "(if " & a1 & " " & a2 & " " & a3 & ")"
    Dim expected2c As String
    expected2c = "(if " & a1 & vbCrLf & _
                 Space$(4) & a2 & vbCrLf & _
                 Space$(4) & a3 & ")" & vbCrLf & vbCrLf & vbCrLf
    Report "vlalint: a short control-flow functor (if) keeps the inline shape even when it doesn't fit flat - the regression this correction fixes", _
           VLA_Lint.VlaLintFormat(s2c) = expected2c, _
           "got: [" & VLA_Lint.VlaLintFormat(s2c) & "]"

    Dim expected3 As String
    expected3 = "(defmacro" & vbCrLf & _
                "    (f x)" & vbCrLf & _
                "    " & Chr$(34) & "doc" & Chr$(34) & vbCrLf & _
                "    (stop))" & vbCrLf & vbCrLf & vbCrLf
    Report "vlalint: defmacro always gets the fixed multi-line layout, even when short", _
           VLA_Lint.VlaLintFormat("(defmacro (f x) ""doc"" (stop))") = expected3, _
           "got: [" & VLA_Lint.VlaLintFormat("(defmacro (f x) ""doc"" (stop))") & "]"

    ' Owner correction: a directive-shaped node's OWN two arguments now
    ' both drop below the functor (PpVerticalize), rather than the
    ' first staying inline with it - "both arguments as one aligned
    ' list beneath the functor," matching defmacro's own layout.
    Dim expected4 As String
    expected4 = "(english-vla" & vbCrLf & _
                Space$(4) & Chr$(34) & "hi." & Chr$(34) & vbCrLf & _
                Space$(4) & "(debug-print 1))" & vbCrLf & vbCrLf & vbCrLf
    Report "vlalint: english-vla always splits EVERY argument onto its own line below the functor, even when short", _
           VLA_Lint.VlaLintFormat("(english-vla ""hi."" (debug-print 1))") = expected4, _
           "got: [" & VLA_Lint.VlaLintFormat("(english-vla ""hi."" (debug-print 1))") & "]"

    ' begin (5 chars, well under MAX_INLINE_FUNCTOR) is NOT directive-
    ' shaped itself, but containing a directive-shaped argument
    ' (test-success) still forces IT vertical - keeping
    ' PpVerticalizeInline's own shape ("x" stays inline with "begin");
    ' the nested test-success is directive-shaped, so IT still gets
    ' PpVerticalize's unconditional full split regardless of its own
    ' (12-char) functor length.
    Dim expected5 As String
    expected5 = "(begin x" & vbCrLf & _
                Space$(7) & "(test-success" & vbCrLf & _
                Space$(11) & Chr$(34) & "a" & Chr$(34) & vbCrLf & _
                Space$(11) & "b))" & vbCrLf & vbCrLf & vbCrLf
    Report "vlalint: a non-first directive-shaped argument forces the parent vertical too - the parent (short functor) stays inline, the directive itself still fully splits", _
           VLA_Lint.VlaLintFormat("(begin x (test-success ""a"" b))") = expected5, _
           "got: [" & VLA_Lint.VlaLintFormat("(begin x (test-success ""a"" b))") & "]"

    Dim s6 As String
    s6 = "; a comment" & vbCrLf & vbCrLf & "(stop)"
    Dim expected6 As String
    expected6 = "; a comment" & vbCrLf & vbCrLf & "(stop)" & vbCrLf & vbCrLf & vbCrLf
    Report "vlalint: comments and blank lines pass through byte-for-byte at their original position", _
           VLA_Lint.VlaLintFormat(s6) = expected6, _
           "got: [" & VLA_Lint.VlaLintFormat(s6) & "]"

    Dim s7 As String
    s7 = "(begin" & vbCrLf & "  ; inner comment, dropped on purpose" & vbCrLf & "  (stop))"
    Dim out7 As String
    out7 = VLA_Lint.VlaLintFormat(s7)
    Report "vlalint: a comment INSIDE a form's own parens is dropped, matching tools/vla_lint.pl's own already-installed behavior (not a regression)", _
           InStr(1, out7, "inner comment", vbTextCompare) = 0, _
           "got: [" & out7 & "]"

    ' Owner QoL house rule (VLA_Lint.bas's own addition, no Perl-tool
    ' ancestor): a bare top-level defmacro is always followed by a
    ' blank line before whatever comes next - inserted when missing,
    ' not duplicated when already present. Both inputs below canonicalize
    ' to the identical expected8 text.
    Dim defmacroBlock As String
    defmacroBlock = "(defmacro" & vbCrLf & _
                     "    (f x)" & vbCrLf & _
                     "    " & Chr$(34) & "doc" & Chr$(34) & vbCrLf & _
                     "    (stop))"
    Dim expected8 As String
    expected8 = defmacroBlock & vbCrLf & vbCrLf & "(g x)" & vbCrLf & vbCrLf & vbCrLf

    Dim s8 As String
    s8 = "(defmacro (f x) ""doc"" (stop))" & vbCrLf & "(g x)"
    Report "vlalint: a bare top-level defmacro with no blank line after it gets one inserted before the next form", _
           VLA_Lint.VlaLintFormat(s8) = expected8, _
           "got: [" & VLA_Lint.VlaLintFormat(s8) & "]"

    Dim s9 As String
    s9 = "(defmacro (f x) ""doc"" (stop))" & vbCrLf & vbCrLf & "(g x)"
    Report "vlalint: a blank line already present after a defmacro is not duplicated", _
           VLA_Lint.VlaLintFormat(s9) = expected8, _
           "got: [" & VLA_Lint.VlaLintFormat(s9) & "]"

    ' Regression pin, caught live: a real corpus file with its OWN
    ' trailing blank lines came back with six, not two, because the
    ' first cut of the two-blank-lines-at-EOF rule appended three more
    ' unconditionally instead of normalizing what the source already
    ' had. Both checks below must hold for the rule to be trustworthy:
    ' excess blanks collapse to exactly two, and an already-canonical
    ' file (also VlaLintCheck's whole comparison depends on this) is
    ' left byte-for-byte alone.
    Dim expected10 As String
    expected10 = "(stop)" & vbCrLf & vbCrLf & vbCrLf

    Dim s10 As String
    s10 = "(stop)" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf
    Report "vlalint: excess trailing blank lines already in the source collapse to the canonical two, not accumulate", _
           VLA_Lint.VlaLintFormat(s10) = expected10, _
           "got: [" & VLA_Lint.VlaLintFormat(s10) & "]"

    Report "vlalint: formatting an already-canonical file is idempotent (a no-op)", _
           VLA_Lint.VlaLintFormat(expected10) = expected10, _
           "got: [" & VLA_Lint.VlaLintFormat(expected10) & "]"

    ' Owner-found regression, live: alonzo.vla's own register-bricks (a
    ' sub whose body is five back-to-back deflambdas) had no blank line
    ' between any of them - the top-level "defmacro gets a blank line
    ' after it" rule never reached a NESTED body position at all.
    ' IntroducesDefinition (defmacro OR deflambda) now fires inside
    ' every body-rendering loop (PpDefmacro, PpVerticalize,
    ' PpVerticalizeInline's both the first-argument slot and its own
    ' m=3+ loop), not just at VlaLintFormat's own top-level segment
    ' boundary. Exact-match pin below traces a nested DEFMACRO (via
    ' PpVerticalizeInline's first-argument slot, table-property-family's
    ' own real (begin (defmacro ...) ...) shape) end to end; the InStr
    ' pin after it confirms DEFLAMBDA specifically triggers the same
    ' rule, without re-tracing a second full layout by hand.
    Dim expected11 As String
    expected11 = "(begin (defmacro" & vbCrLf & _
                 Space$(11) & "(f x)" & vbCrLf & _
                 Space$(11) & Chr$(34) & "doc" & Chr$(34) & vbCrLf & _
                 Space$(11) & "(stop))" & vbCrLf & _
                 vbCrLf & _
                 Space$(7) & "(g x))" & vbCrLf & vbCrLf & vbCrLf
    Report "vlalint: a nested defmacro (first argument of an enclosing begin) also gets a blank line after it, not just a top-level one", _
           VLA_Lint.VlaLintFormat("(begin (defmacro (f x) ""doc"" (stop)) (g x))") = expected11, _
           "got: [" & VLA_Lint.VlaLintFormat("(begin (defmacro (f x) ""doc"" (stop)) (g x))") & "]"

    Dim padDoc As String
    padDoc = "this docstring is intentionally long enough to push the whole enclosing form past one hundred columns wide"
    Dim s12 As String
    s12 = "(sub reg () ""d"" (deflambda f1 (x) """ & padDoc & """ (+ x 1)) (deflambda f2 (y) ""short"" (+ y 2)))"
    Dim out12 As String
    out12 = VLA_Lint.VlaLintFormat(s12)
    Report "vlalint: a deflambda body form also gets a blank line after it, same as defmacro - the alonzo.vla regression itself", _
           InStr(1, out12, "(+ x 1))" & vbCrLf & vbCrLf, vbBinaryCompare) > 0, _
           "got: [" & out12 & "]"
End Sub

' F.9: instructions.txt's own paragraphs (its documented structural unit -
' "blank line ends a block"), as (startLine, endLine, label) triples,
' one Collection per paragraph, 1-indexed lines to match the at-line
' tags InsertSectionMarkers below correlates them against. label is
' the paragraph's first non-blank, non-#-comment line, truncated for
' a marker that stays a one-liner.
Private Function EnglishParagraphs(ByVal englishText As String) As Collection
    Dim lines() As String
    lines = Split(Replace(englishText, vbCrLf, vbLf), vbLf)
    Dim result As New Collection
    Dim inPara As Boolean, paraStart As Long, paraLabel As String
    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        Dim lineNo As Long
        lineNo = i + 1
        Dim trimmed As String
        trimmed = Trim$(lines(i))
        If Len(trimmed) = 0 Then
            If inPara Then
                result.Add ParagraphRec(paraStart, lineNo - 1, paraLabel)
                inPara = False
            End If
        Else
            If Not inPara Then
                inPara = True
                paraStart = lineNo
                paraLabel = ""
            End If
            If Len(paraLabel) = 0 And Left$(trimmed, 1) <> "#" Then
                paraLabel = trimmed
                If Len(paraLabel) > 70 Then paraLabel = Left$(paraLabel, 67) & "..."
            End If
        End If
    Next i
    If inPara Then result.Add ParagraphRec(paraStart, UBound(lines) + 1, paraLabel)
    Set EnglishParagraphs = result
End Function

Private Function ParagraphRec(ByVal startLine As Long, ByVal endLine As Long, _
                               ByVal label As String) As Collection
    Dim rec As New Collection
    rec.Add startLine
    rec.Add endLine
    rec.Add label
    Set ParagraphRec = rec
End Function

' Which paragraph (1-based index into paragraphs) a source line falls
' in, or 0 if none does (defensive - should not happen for a real
' at-line value, since every non-blank source line belongs to exactly
' one paragraph by construction).
Private Function ParagraphIndexForLine(paragraphs As Collection, ByVal n As Long) As Long
    Dim i As Long
    For i = 1 To paragraphs.Count
        Dim rec As Collection
        Set rec = paragraphs.Item(i)
        If n >= CLng(rec.Item(1)) And n <= CLng(rec.Item(2)) Then
            ParagraphIndexForLine = i
            Exit Function
        End If
    Next i
End Function

Private Function ParseLeadingNumber(ByVal s As String) As Long
    Dim j As Long, numTxt As String
    j = 1
    Do While j <= Len(s)
        Dim c As String
        c = Mid$(s, j, 1)
        If c >= "0" And c <= "9" Then
            numTxt = numTxt & c
            j = j + 1
        Else
            Exit Do
        End If
    Loop
    If Len(numTxt) > 0 Then ParseLeadingNumber = CLng(numTxt)
End Function

' A VLA string-literal token (leading-quote tag, per VLA.bas's own
' reader/StrLitContent contract) from raw text - escapes \ and " the
' way Tokenize's string-literal reader un-escapes \" and \\, so a
' label containing a quoted phrase (most of instructions.txt's sentences
' do) round-trips correctly through (raw "..."). R7 duplicate: an
' identical VlaStringLit already lives Private in VLA_English.bas,
' unreachable from here for the same reason every other cross-module
' Private call in this codebase gets its own small copy instead.
Private Function VlaStringLit(ByVal raw As String) As String
    Dim esc As String
    esc = Replace(raw, "\", "\\")
    esc = Replace(esc, Chr$(34), "\" & Chr$(34))
    VlaStringLit = Chr$(34) & esc & Chr$(34)
End Function

' F.9: splice one (raw "' ---- instructions.txt:N label ----") line into
' vlaText before the first statement of each instructions.txt paragraph -
' a single forward pass, anchored on "(at-line N" (present exactly
' once per statement, indentation-independent since the marker copies
' whatever leading whitespace the at-line line already has). No
' backward insertion, no splicing an already-built Collection: the
' marker lands right before the at-line line itself, which reads fine
' even with the (set! vla-step ...)/(if (vlatraceon) ...) step-
' tracking preamble sitting just above it - that scaffolding has no
' user-visible meaning to protect, and anchoring on it too would trade
' this function's only real complexity for none of the benefit.
Private Function InsertSectionMarkers(ByVal vlaText As String, ByVal englishText As String) As String
    Dim paragraphs As Collection
    Set paragraphs = EnglishParagraphs(englishText)

    Dim lines() As String
    lines = Split(Replace(vlaText, vbCrLf, vbLf), vbLf)

    ' Plain concatenation, not VLA.bas's SbAdd/SbText buffer (Private
    ' to that module - R7 territory, but this runs once on a ~300-line
    ' corpus at golden-generation time, not once per statement of
    ' every compile, so the O(n^2) a plain & would cost on a HOT path
    ' is not a cost worth avoiding here.
    Dim outText As String
    Dim currentPara As Long
    currentPara = -1
    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        Dim ln As String
        ln = lines(i)
        Dim trimmed As String
        trimmed = LTrim$(ln)
        If Left$(trimmed, 9) = "(at-line " Then
            Dim srcLn As Long
            srcLn = ParseLeadingNumber(Mid$(trimmed, 10))
            If srcLn > 0 Then
                Dim pIdx As Long
                pIdx = ParagraphIndexForLine(paragraphs, srcLn)
                If pIdx > 0 And pIdx <> currentPara Then
                    currentPara = pIdx
                    Dim rec As Collection
                    Set rec = paragraphs.Item(pIdx)
                    Dim pad As String
                    pad = Left$(ln, Len(ln) - Len(trimmed))
                    Dim marker As String
                    marker = "' ---- instructions.txt:" & CLng(rec.Item(1)) & " " & CStr(rec.Item(3)) & " ----"
                    outText = outText & pad & "(raw " & VlaStringLit(marker) & ")" & vbCrLf
                End If
            End If
        End If
        outText = outText & ln
        If i < UBound(lines) Then outText = outText & vbCrLf
    Next i
    InsertSectionMarkers = outText
End Function

' F.9: pins EnglishParagraphs' block-boundary detection and
' InsertSectionMarkers' placement/escaping directly - a correctness
' guarantee independent of reading through the (much larger) real
' corpus goldens by eye. A hand-built three-paragraph sample, not
' instructions.txt itself, so this test's expectations do not drift every
' time a corpus sentence is added or reworded.
Private Sub TestSectionMarkers()
    Dim eng As String
    eng = "# a leading comment, skipped for the label" & vbCrLf & _
          "Set a to 1." & vbCrLf & _
          vbCrLf & _
          "Log ""two"" now." & vbCrLf & _
          vbCrLf & _
          "Set c to 3."
    Dim paras As Collection
    Set paras = EnglishParagraphs(eng)
    CheckV "f9: three blank-line-delimited paragraphs found", paras.Count, 3
    If paras.Count = 3 Then
        Dim p1 As Collection, p2 As Collection, p3 As Collection
        Set p1 = paras.Item(1)
        Set p2 = paras.Item(2)
        Set p3 = paras.Item(3)
        CheckV "f9: paragraph 1 starts at its leading comment's line", CLng(p1.Item(1)), 1
        CheckV "f9: paragraph 1's label skips the leading # comment", CStr(p1.Item(3)), "Set a to 1."
        CheckV "f9: paragraph 2 starts where its own text begins", CLng(p2.Item(1)), 4
        CheckV "f9: paragraph 2's label keeps its embedded quotes (pre-VlaStringLit)", CStr(p2.Item(3)), "Log ""two"" now."
        CheckV "f9: paragraph 3 starts at the file's last block", CLng(p3.Item(1)), 6
    End If

    Dim vla As String
    vla = "(sub main ()" & vbCrLf & _
          "    (at-line 2" & vbCrLf & _
          "    (set! a 1)))" & vbCrLf & _
          "    (at-line 4" & vbCrLf & _
          "    (debug-print ""two""))" & vbCrLf & _
          "    (at-line 6" & vbCrLf & _
          "    (set! c 3)))"
    Dim marked As String
    marked = InsertSectionMarkers(vla, eng)
    Report "f9: a marker precedes each new paragraph's first at-line, quotes escaped", _
           InStr(1, marked, "(raw ""' ---- instructions.txt:1 Set a to 1. ----"")") > 0 And _
           InStr(1, marked, "(raw ""' ---- instructions.txt:4 Log \""two\"" now. ----"")") > 0 And _
           InStr(1, marked, "(raw ""' ---- instructions.txt:6 Set c to 3. ----"")") > 0, _
           Left$(marked, 400)
    Report "f9: exactly three markers - one per paragraph, none repeated mid-paragraph", _
           CountOcc(marked, "(raw ""' ----") = 3, "got: " & CountOcc(marked, "(raw ""' ----")
End Sub

' F.1: the dot count - rules whose template reaches the VBA object
' model directly instead of through a named macro or runtime helper.
' Publish it; watch it fall; fail a test when it rises (BETA_ROADMAP.md's
' own words for this item). Counts every top-level form containing
' "(. " that is NOT itself a (defmacro ...) form - macro BODIES are
' where the dots are supposed to live, per F.1's own finish line ("the
' dots live in prelude.vla," today the phrasebook's own defmacro
' forms, since prelude.vla (F.3) does not exist yet). Also counts
' "application.worksheetfunction." outside a defmacro form - F.1's
' other bare-dotted-call category (sum/average/max/min/vlookup/sumif/
' countif).
' Excludes the two core-language "Append to <list>" proofs by name
' (they test Collection.Add, a language feature, not a spreadsheet
' rule) rather than silently, so a real regression is never absorbed
' into that exception unnoticed.
' F.13: rewritten for the vocabulary-grammar migration - the old
' version recognized a macro body by its own "macro:" line prefix,
' which no longer exists (defmacro forms are bare now, no wrapper).
' The new version reads the whole file as real forms (VLA.VlaReadForms
' - the same reader every other caller in this codebase now uses) and
' asks each top-level form's OWN head symbol whether it is "defmacro"
' - structural, not a line-prefix guess, and correct regardless of how
' a defmacro form happens to be formatted across lines.
Public Function VlaDotCount(Optional ByVal vocabPath As String = "") As Long
    If Len(vocabPath) = 0 Then vocabPath = FindDevFile("english.vla")
    Dim src As String
    src = ReadTextFileUtf8(vocabPath)
    Dim forms As Collection
    Set forms = VLA.VlaReadForms(src)
    Dim n As Long
    Dim f As Variant
    For Each f In forms
        Dim isMacro As Boolean
        isMacro = False
        If IsObject(f) Then
            Dim fl As Collection
            Set fl = f
            If fl.Count > 0 Then
                If Not IsObject(fl.Item(1)) Then
                    isMacro = (VLA_Identity.Fold(CStr(fl.Item(1))) = "defmacro")
                End If
            End If
        End If
        If Not isMacro Then
            Dim ftxt As String
            ftxt = VLA.VlaWriteForm(f)
            If InStr(1, ftxt, "found-items", vbBinaryCompare) = 0 Then
                If InStr(1, ftxt, "(. ", vbBinaryCompare) > 0 Then
                    n = n + 1
                ElseIf InStr(1, ftxt, "application.worksheetfunction.", vbBinaryCompare) > 0 Then
                    n = n + 1
                End If
            End If
        End If
    Next
    VlaDotCount = n
End Function

' AS.6: SD-6's contract, as one callable function instead of three
' ad-hoc checks - instructions.txt, english.vla, and alonzo.vla (the
' function loader, renamed from hello.vla) all resolve through
' FindDevFile's own probe (beside the workbook, then its scripts
' folder, then scripts/polyglotta), so this can never drift onto a
' different assumption about where the corpus lives. detail comes
' back empty when all three are present; otherwise it names every
' missing file (not just the first), so one report explains the
' whole gap.
Public Function VlaCorpusFamilyOk(ByRef detail As String) As Boolean
    Dim names As Variant
    names = Array("instructions.txt", "english.vla", "alonzo.vla")
    Dim missing As String
    Dim i As Long
    For i = LBound(names) To UBound(names)
        Dim p As String
        p = ""
        On Error Resume Next
        p = FindDevFile(CStr(names(i)))
        On Error GoTo 0
        If Len(p) = 0 Then missing = missing & " " & names(i)
    Next i
    detail = missing
    VlaCorpusFamilyOk = (Len(missing) = 0)
End Function

' S1.1: resolve a dev corpus file by probing the known layouts -
' beside the workbook, then in its scripts subfolder, then in
' scripts/polyglotta (the non-English dialect files' own home, so the
' corpus can grow to ten or twenty languages without scripts/ itself
' getting crowded) - taking the first hit. A miss refuses loudly,
' naming every probed location and the explicit-path escape hatch
' (error voice: say what was tried, teach the way out). One resolver,
' so the corpus files cannot drift onto different assumptions.
Public Function FindDevFile(ByVal fileName As String) As String
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
        Err.Raise 53, "VLA_Tests", fileName & " not found - I looked beside the workbook (" & _
                  c1 & "), in its scripts folder (" & c2 & "), and in scripts/polyglotta (" & c3 & _
                  "); pass the full path to VlaWriteGoldens if it lives elsewhere"
    End If
End Function

' <program folder>\<base>_golden<ext> - instructions.txt -> instructions_golden.vla
Private Function GoldenPathFor(ByVal programPath As String, ByVal ext As String) As String
    Dim sep As String
    sep = Application.PathSeparator
    Dim i As Long
    i = InStrRev(programPath, sep)
    Dim folder As String, base As String
    folder = Left$(programPath, i)           ' keeps the separator; "" if none
    base = Mid$(programPath, i + 1)
    Dim d As Long
    d = InStrRev(base, ".")
    If d > 0 Then base = Left$(base, d - 1)
    GoldenPathFor = folder & base & "_golden" & ext
End Function

' Rule-12 duplicate of VLA_English's private VocabReadFile (modules
' stay self-contained): utf-8 via ADODB.Stream with an ANSI binary
' fallback, so the golden translation reads the program byte-for-byte
' the way a Run's file import would.
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

' Plain-text writer for the goldens. Print # writes CRLF line endings
' and appends one trailing break - both deterministic, which is all a
' diff baseline asks; the corpus and everything generated from it is
' plain ASCII.
Private Sub WriteTextFile(ByVal filePath As String, ByVal text As String)
    Dim f As Integer
    f = FreeFile
    Open filePath For Output As #f
    Print #f, text
    Close #f
End Sub

' ---------------------------------------------------------------------
'  S1 pins: the property the goldens rely on - translation and
'  transpilation are DETERMINISTIC, so a golden diff is always a
'  behavior diff, never noise. (The golden files themselves are
'  proven by git diff in the verification loop, which the self-test
'  cannot reach; these pins prove the property that makes that proof
'  sound.)
' ---------------------------------------------------------------------
Private Sub TestGoldens()
    EnglishResetGrammar
    Dim prog As String
    prog = "Create a number called total." & vbCrLf & _
           "Repeat 3 times:" & vbCrLf & _
           "  Increase total by counter." & vbCrLf & vbCrLf & _
           "Log total."
    Dim a As String, b As String
    a = EnglishToVla(prog)
    b = EnglishToVla(prog)
    Report "golden: translation is deterministic", a = b And Len(a) > 0, _
           "two translations of one program differ"
    Dim va As String, vbOut As String
    va = VlaTranspile(a)
    vbOut = VlaTranspile(b)
    Report "golden: transpilation is deterministic", va = vbOut And Len(va) > 0, _
           "two transpiles of one translation differ"

    ' S3.1 regression: the pin above caught heap-dependent PHANTOM map
    ' tags - template-copied lists were keyless in the ObjPtr map, so
    ' one born at a dead tagged form's address read the dead form's
    ' line. Template statements must be deliberately, EXPLICITLY
    ' untagged: present in the output, never wearing a tag.
    Dim tv As String
    tv = VlaTranspile("(sub t () (inc! x))")
    Report "map: template statements are deliberately untagged", _
           InStr(1, tv, "x = (x + 1)", vbTextCompare) > 0 And _
           InStr(1, tv, "x = (x + 1) ' vla:", vbTextCompare) = 0, _
           Left$(Norm(tv), 200)
End Sub

' ---------------------------------------------------------------------
'  LISTOPS-PROVENANCE: (gen-row label stmt...), a zero-runtime emitter
'  annotation, same shape as at-line but a DIFFERENT provenance - which
'  row of a GENERATOR's own source table a macro's body came from, not
'  which line of the calling program. Deliberately does NOT touch
'  TagLine/FormLine/ObjPtr at all (label is ordinary DATA copied by
'  Substitute like any other literal) - the previous pin's own
'  regression (template statements deliberately untagged) must hold
'  completely unchanged when gen-row is never used; these pins prove
'  the tag DOES appear when it is, and restores cleanly afterward.
' ---------------------------------------------------------------------
Private Sub TestGenRow()
    Dim tv As String

    tv = VlaTranspile("(sub t () (gen-row ""37"" (set! x 1)))")
    Report "gen-row: a wrapped statement's emitted comment carries vla-row", _
           InStr(1, tv, "vla-row:37", vbTextCompare) > 0, Left$(Norm(tv), 200)

    ' The wrapper's own dynamic extent ends where it says it does - a
    ' sibling statement AFTER the wrapper is an ordinary, directly-
    ' parsed statement (a real source line, unlike a macro's own
    ' Substitute-copied body), so it keeps its NORMAL vla:N tag - it
    ' just must not ALSO pick up vla-row. Exactly one "vla-row" in the
    ' whole output, not two, proves both that the sibling didn't
    ' inherit it AND that the wrapper's own line was never additionally
    ' tagged (at-line's own "tagging the wrapper would inject a
    ' duplicate" note applies here identically).
    tv = VlaTranspile("(sub t () (gen-row ""37"" (set! x 1)) (set! y 2))")
    Dim rowCount As Long
    rowCount = (Len(tv) - Len(Replace(tv, "vla-row", ""))) \ Len("vla-row")
    Report "gen-row: its own extent ends where it says - a later sibling keeps its normal tag but no row-tag, and there is no duplicate", _
           rowCount = 1 And InStr(1, tv, "y = 2", vbTextCompare) > 0, Left$(Norm(tv), 250)
End Sub

' ---------------------------------------------------------------------
'  LISTOPS-BUDGET: verify, don't assume, that the engine fails LOUDLY
'  under realistic LISTOPS-scale expansion volume rather than silently
'  handing back a partial, truncated-but-plausible-looking rule set.
'  LISTOPS itself doesn't exist yet (no car/cdr to walk a table with),
'  but the actual risk doesn't need it to exist: a table-walker written
'  the natural recursive way (each step consumes one row, calls itself
'  on the rest) is a CHAIN of macro re-expansions at one position -
'  exactly what ExpandMacros's own `depth` counts, already, today, with
'  ordinary substitution-only defmacro. Built here with a VBA loop
'  (never hand-typed), not because LISTOPS is being simulated, but
'  because the depth-guard's OWN behavior at scale is what needs
'  proving, independent of how the chain gets generated.
'
'  Finding, stated so it isn't rediscovered by surprise later: the
'  guard (`ExpandMacros`'s own `depth > 200`) DOES fail loudly, never
'  silently - the good half of the news. The other half: 200 sequential
'  steps is not a large number for a real spreadsheet table (a 250-row
'  pricing sheet is entirely ordinary), so a naive recursive LISTOPS
'  table-walker would hit this ceiling on a table smaller than "large"
'  by any normal measure. Two options for whoever scopes LISTOPS, not
'  decided here: raise the constant, or steer LISTOPS's own walking
'  convention away from one long recursive chain. Either way, the
'  failure mode today is the safe one (loud, not silent) - this pin is
'  the proof, not a fix.
'
'  TABLESPEC-SCALE / TABLESPECSCALE.0: the fix, eventually - neither of
'  the two options above turned out to be it. Raising the constant was
'  tried (LISTOPS.0) and disproven live (native stack, not the soft
'  guard, was the real ceiling). Steering the walking convention away
'  from a chain (TABLESPEC's own first cut) didn't help either, once
'  traced: ANY self-recursive macro walk pays the identical per-row
'  ExpandMacros `depth` cost regardless of accumulator shape, because
'  the cost comes from chained re-expansion, not from what shape the
'  result takes. The actual fix was hand-trampolining `ExpandMacros`'s
'  own head-resolution chain (VLA.bas) - VBA has no built-in tail-call
'  optimization, but the chain was always a tail call in substance - so
'  it no longer costs a VBA stack frame per row at all. The 250-deep case
'  below now expands cleanly instead of failing; a new 2000-deep case
'  and a new 5500-deep (past the trampoline's own new, purely-logical
'  5000 cap) case were added alongside it.
' ---------------------------------------------------------------------
Private Sub TestListopsBudget()
    Dim i As Long, chainLen As Long

    ' A legitimate, finite chain safely under the guard expands
    ' COMPLETELY - the fixpoint is reached, nothing silently stops
    ' partway and hands back a plausible-looking but incomplete result.
    chainLen = 150
    Dim src150 As String
    src150 = ""
    For i = 0 To chainLen - 1
        src150 = src150 & "(defmacro (step" & i & ") (step" & (i + 1) & "))" & vbCrLf
    Next
    src150 = src150 & "(defmacro (step" & chainLen & ") (debug-print ""done""))" & vbCrLf
    src150 = src150 & "(sub t () (step0))"
    Dim vba150 As String
    On Error Resume Next
    Err.Clear
    vba150 = VlaTranspile(src150)
    Dim n150 As Long, d150 As String
    n150 = Err.Number
    d150 = Err.Description
    On Error GoTo 0
    ' G0 gotcha: On Error GoTo 0 itself resets Err.Number to 0 - Err.Number
    ' must be captured to a local (n150) BEFORE it executes, the same
    ' discipline Err.Description already gets everywhere else in this
    ' suite, never read live after the fact.
    Report "listops-budget: a 150-deep legitimate macro chain expands completely, not truncated", _
           n150 = 0 And InStr(1, vba150, "done", vbTextCompare) > 0, _
           "err: " & d150 & " / vba: " & Left$(vba150, 200)

    ' The stepper's own budget counters (mExpandBudget/mExpandFired,
    ' LISTOPS-BUDGET's own literal namesakes) stay correct at scale too -
    ' asking for step 100 of the 150-chain (chainLen is still 150 here,
    ' read before the next block reassigns it) returns exactly the
    ' state after 100 applications (parked on step100, not one off in
    ' either direction, and nowhere near the fixpoint's own "done").
    Dim expectTotal As Long
    expectTotal = chainLen + 1
    Dim total As Long, step100 As String
    step100 = VlaExpandStepText(src150, 100, total)
    Report "listops-budget: the stepper's own count stays correct across a large chain", _
           total = expectTotal And _
           InStr(1, step100, "step100", vbTextCompare) > 0 And _
           InStr(1, step100, "step101", vbTextCompare) = 0 And _
           InStr(1, step100, "done", vbTextCompare) = 0, _
           "total: " & total & " (expected " & expectTotal & ") / step100: " & Left$(step100, 150)

    ' TABLESPEC-SCALE / TABLESPECSCALE.0: this used to be "a 250-deep
    ' chain fails LOUDLY" - the whole reason LISTOPS-BUDGET's depth-chain
    ' problem stayed open. Trampolining ExpandMacros's own head-chain
    ' (VLA.bas) uncoupled chain length from native VBA stack usage
    ' entirely, so 250 - "a 250-row pricing sheet is entirely ordinary,"
    ' this item's own original illustrative case - now EXPANDS CLEANLY
    ' instead of failing. Renamed from src250/n250/etc. to reflect what
    ' it now proves.
    ' TABLESPEC-SCALE / TABLESPECSCALE.0: built via a pre-sized array +
    ' Join, not `src250 = src250 & ...` in the loop - naive string
    ' concatenation reallocates and copies the WHOLE result on every
    ' append (VLA.bas's own SbAdd comment: "invisible at 100... dominant
    ' at 5,000"), and this file has no SbAdd of its own (Private,
    ' invisible cross-module) - found live, Excel visibly hanging on
    ' this test's own 2000/5500-row siblings, not caught by tracing.
    chainLen = 250
    Dim parts250() As String
    ReDim parts250(0 To chainLen - 1)
    For i = 0 To chainLen - 1
        parts250(i) = "(defmacro (rung" & i & ") (rung" & (i + 1) & "))"
    Next
    Dim src250 As String
    src250 = Join(parts250, vbCrLf) & vbCrLf & _
             "(defmacro (rung" & chainLen & ") (debug-print ""done""))" & vbCrLf & _
             "(sub t () (rung0))"
    Dim vba250 As String
    On Error Resume Next
    Err.Clear
    vba250 = VlaTranspile(src250)
    Dim n250 As Long, d250 As String
    n250 = Err.Number
    d250 = Err.Description
    On Error GoTo 0
    Report "listops-budget: a 250-deep chain now expands completely - TABLESPEC-SCALE's own motivating case, no longer failing", _
           n250 = 0 And InStr(1, vba250, "done", vbTextCompare) > 0, _
           "err: " & d250 & " / vba: " & Left$(vba250, 200)

    ' TABLESPEC-SCALE / TABLESPECSCALE.0: 2000/5500-deep chains - the
    ' former genuinely large (20x the "250-row pricing sheet" case), the
    ' latter past the trampoline's OWN new ceiling (chainLen > 5000,
    ' ExpandMacros, VLA.bas), a purely logical guard now rather than a
    ' stack-safety one, still proven to fire loudly. Gated on
    ' mRunScaleTests, same reasoning as TestListopsDepthSafety/
    ' TestTablespecDepthSafety's own gated cases: `GetMacro` (VLA.bas)
    ' looks up `mMacros` by key on every chain step, and a chain this
    ' long resolving through a growing macro table is the same cost
    ' family flagged by the roadmap's own P-DICT ("registry lookups off
    ' Collection error-traps") - noticeably slow at 2000, confirmed live,
    ' not the O(1)-per-step cost this test's own header comment assumed
    ' when only the cdr-based tests were gated. Worth a P-DICT/P-PROF
    ' note for whoever picks that up; not something to fix here.
    If mRunScaleTests Then
        chainLen = 2000
        Dim parts2000() As String
        ReDim parts2000(0 To chainLen - 1)
        For i = 0 To chainLen - 1
            parts2000(i) = "(defmacro (link" & i & ") (link" & (i + 1) & "))"
        Next
        Dim src2000 As String
        src2000 = Join(parts2000, vbCrLf) & vbCrLf & _
                  "(defmacro (link" & chainLen & ") (debug-print ""done""))" & vbCrLf & _
                  "(sub t () (link0))"
        Dim vba2000 As String
        On Error Resume Next
        Err.Clear
        vba2000 = VlaTranspile(src2000)
        Dim n2000 As Long, d2000 As String
        n2000 = Err.Number
        d2000 = Err.Description
        On Error GoTo 0
        Report "listops-budget: a 2000-deep chain expands completely - realistic table scale, well past the old native-stack ceiling", _
               n2000 = 0 And InStr(1, vba2000, "done", vbTextCompare) > 0, _
               "err: " & d2000 & " / vba: " & Left$(vba2000, 200)

        chainLen = 5500
        Dim parts5500() As String
        ReDim parts5500(0 To chainLen - 1)
        For i = 0 To chainLen - 1
            parts5500(i) = "(defmacro (far" & i & ") (far" & (i + 1) & "))"
        Next
        Dim src5500 As String
        src5500 = Join(parts5500, vbCrLf) & vbCrLf & _
                  "(defmacro (far" & chainLen & ") (debug-print ""done""))" & vbCrLf & _
                  "(sub t () (far0))"
        Dim vba5500 As String
        On Error Resume Next
        Err.Clear
        vba5500 = VlaTranspile(src5500)
        Dim n5500 As Long, d5500 As String
        n5500 = Err.Number
        d5500 = Err.Description
        On Error GoTo 0
        Report "listops-budget: a 5500-deep chain - past the trampoline's own 5000 cap - fails loudly, never a silent partial expansion", _
               n5500 <> 0 And Len(vba5500) = 0 And InStr(1, d5500, "5000", vbTextCompare) > 0, _
               "err: " & d5500 & " / vba: " & Left$(vba5500, 200)
    End If

    ' Volume alone - many INDEPENDENT (never chained) macro calls, the
    ' shape (begin call1 call2 ... callN) splicing already produces
    ' today - carries no depth risk at all, at any N: depth only grows
    ' through CHAINED re-expansion at one position, never across
    ' independent siblings. The realistic wide-table shape is entirely
    ' safe regardless of row count; only the naive-recursive shape
    ' above needs the ceiling raised.
    Dim wideSrc As String, wi As Long
    wideSrc = "(defmacro (leaf n) (debug-print n))" & vbCrLf & "(sub w ()" & vbCrLf
    For wi = 1 To 1000
        wideSrc = wideSrc & "(leaf " & wi & ")" & vbCrLf
    Next
    wideSrc = wideSrc & ")"
    Dim wideVba As String
    On Error Resume Next
    Err.Clear
    wideVba = VlaTranspile(wideSrc)
    Dim nWide As Long, dWide As String
    nWide = Err.Number
    dWide = Err.Description
    On Error GoTo 0
    Report "listops-budget: 1000 independent (non-chained) macro calls carry no depth risk at all", _
           nWide = 0 And InStr(1, wideVba, "Debug.Print 1000", vbTextCompare) > 0, _
           "err: " & dWide & " / len: " & Len(wideVba)
End Sub

' ---------------------------------------------------------------------
'  LISTOPS-CONFLUENCE: prove ExpandMacros's walk produces the same
'  result regardless of sibling-expansion order for recursive-over-
'  recursive expansion specifically - nothing before LISTOPS forced
'  two arbitrary recursive expansions to interact, so this property
'  was never exercised.
'
'  Closed the same way LISTOPS-PURITY was - by reading, backed here by
'  a concrete test, not asserted from the read alone. The structural
'  argument: `ExpandMacros`'s own `depth` is `ByVal` (no cross-call
'  mutation - VBA value semantics, not a convention); `mMacros` is
'  fully populated in pass 1, before any expansion begins, and never
'  written again during the walk (`VlaTranspile`/`VlaExpandText`'s own
'  two-pass shape - collect every defmacro, THEN expand body forms);
'  and `Substitute` (re-read for this pass) touches ONLY its own four
'  parameters plus pure helpers - no module-level state at all, not
'  `mExpandFired`, not `mAtLine`, nothing. A macro's own expansion
'  result is therefore a pure function of its own template and the
'  arguments explicitly bound to it - there is no channel through
'  which one sibling's expansion could observably affect another's,
'  regardless of which one a walk touches first. Order is already
'  fixed today (VBA Collections iterate in insertion order, no
'  randomization possible) - the test below cannot literally vary
'  order, so it proves the CONSEQUENCE order-independence predicts
'  instead: two unrelated, independently-recursive macro chains,
'  brought together as siblings by a third macro, both resolve to
'  their own correct fixpoint with no cross-contamination, in either
'  argument order.
' ---------------------------------------------------------------------
Private Sub TestListopsConfluence()
    Dim i As Long, chainLen As Long
    chainLen = 50

    ' Two INDEPENDENT recursive chains (never calling each other),
    ' brought together as siblings by one combining macro - the actual
    ' new shape LISTOPS-CONFLUENCE names: two arbitrary recursive
    ' expansions interacting, not just coexisting in one file.
    Dim chains As String
    chains = "(defmacro (combine-two a b) (begin (debug-print a) (debug-print b)))" & vbCrLf
    For i = 0 To chainLen - 2
        chains = chains & "(defmacro (stepC" & i & ") (stepC" & (i + 1) & "))" & vbCrLf
        chains = chains & "(defmacro (stepD" & i & ") (stepD" & (i + 1) & "))" & vbCrLf
    Next
    chains = chains & "(defmacro (stepC" & (chainLen - 1) & ") ""TAGC"")" & vbCrLf
    chains = chains & "(defmacro (stepD" & (chainLen - 1) & ") ""TAGD"")" & vbCrLf

    Dim vbaFwd As String, nFwd As Long, dFwd As String
    On Error Resume Next
    Err.Clear
    vbaFwd = VlaTranspile(chains & "(sub t () (combine-two (stepC0) (stepD0)))")
    nFwd = Err.Number
    dFwd = Err.Description
    On Error GoTo 0
    Report "listops-confluence: two independent 50-deep sibling chains both resolve correctly, in order", _
           nFwd = 0 And InStr(1, vbaFwd, "TAGC", vbTextCompare) > 0 And _
           InStr(1, vbaFwd, "TAGD", vbTextCompare) > InStr(1, vbaFwd, "TAGC", vbTextCompare), _
           "err: " & dFwd & " / vba: " & Left$(vbaFwd, 200)

    ' The same two chains, arguments swapped at the call site - proves
    ' the result tracks the CALL's own argument order, not some
    ' incidental artifact of which chain happened to be defined first
    ' in the source text.
    Dim vbaRev As String, nRev As Long, dRev As String
    On Error Resume Next
    Err.Clear
    vbaRev = VlaTranspile(chains & "(sub t () (combine-two (stepD0) (stepC0)))")
    nRev = Err.Number
    dRev = Err.Description
    On Error GoTo 0
    Report "listops-confluence: swapping the sibling call order swaps the output order, cleanly", _
           nRev = 0 And InStr(1, vbaRev, "TAGD", vbTextCompare) > 0 And _
           InStr(1, vbaRev, "TAGC", vbTextCompare) > InStr(1, vbaRev, "TAGD", vbTextCompare), _
           "err: " & dRev & " / vba: " & Left$(vbaRev, 200)

    ' Recursive-INTO-recursive, not just recursive-BESIDE-recursive: a
    ' value threaded through one recursive macro FAMILY that hands off,
    ' mid-chain, into a SECOND independent recursive family, arrives at
    ' the far terminal uncorrupted - proving argument bindings stay
    ' isolated across a nested recursive handoff, not just across
    ' sibling recursions.
    Dim halfLen As Long
    halfLen = 25
    Dim nestedSrc As String
    nestedSrc = ""
    For i = 0 To halfLen - 2
        nestedSrc = nestedSrc & "(defmacro (stepA" & i & " tag) (stepA" & (i + 1) & " tag))" & vbCrLf
    Next
    nestedSrc = nestedSrc & "(defmacro (stepA" & (halfLen - 1) & " tag) (stepB0 tag))" & vbCrLf   ' the handoff
    For i = 0 To halfLen - 2
        nestedSrc = nestedSrc & "(defmacro (stepB" & i & " tag) (stepB" & (i + 1) & " tag))" & vbCrLf
    Next
    nestedSrc = nestedSrc & "(defmacro (stepB" & (halfLen - 1) & " tag) (debug-print tag))" & vbCrLf
    nestedSrc = nestedSrc & "(sub t () (stepA0 ""PAYLOAD""))"

    Dim vbaNested As String, nNested As Long, dNested As String
    On Error Resume Next
    Err.Clear
    vbaNested = VlaTranspile(nestedSrc)
    nNested = Err.Number
    dNested = Err.Description
    On Error GoTo 0
    Report "listops-confluence: a value threaded through a nested two-family recursive handoff arrives uncorrupted", _
           nNested = 0 And InStr(1, vbaNested, "PAYLOAD", vbTextCompare) > 0, _
           "err: " & dNested & " / vba: " & Left$(vbaNested, 200)
End Sub

' ---------------------------------------------------------------------
'  QUASIQUOTE: the full bundle - (quasiquote ...), (unquote ...),
'  (unquote-splicing ...), (symbol ...). Every assertion here mirrors
'  a live-verified Immediate Window check from the scoping/build
'  session (VlaExpandStep), now pinned as a permanent regression.
'  Grouped in build order: symbol (both call sites), the four reserved
'  names, quasiquote/unquote's core shielding behavior, then every
'  guard rail (nesting, wrong-context, wrong-arity, non-list splice
'  operand) that exists specifically to fail loud instead of silently
'  misbehaving - each one is a distinct code path, not a duplicate.
' ---------------------------------------------------------------------
Private Sub TestQuasiquote()
    Dim t As String
    Dim fired As Long

    ' --- symbol: standalone (ExpandMacros) and inside a template
    '     (Substitute) - the one primitive of the four recognized
    '     generally, since it needs no ambient bindings environment.
    t = VlaExpandText("(symbol ""make-"" ""bold"")", True, fired)
    CheckFrags "quasiquote: symbol fuses standalone", t, Array("make-bold")

    t = VlaExpandText("(symbol ""bold"")", True, fired)
    CheckFrags "quasiquote: symbol single-argument is legal (string-literal coercion, not a no-op error)", t, Array("bold")

    CheckExpandErr "quasiquote: symbol refuses zero arguments", _
        "(symbol)", "symbol expects at least 1 argument, got 0"

    t = VlaExpandText("(defmacro (make-getter suffix) (function (symbol ""get-"" suffix) () String (return ""hi""))) (make-getter widget)", True, fired)
    CheckFrags "quasiquote: symbol fuses inside a macro template - closes METAVOCAB's stated limit", t, _
               Array("(function get-widget () String (return ""hi""))")

    ' --- reserved names: none of the four can be shadowed as a macro
    '     name - each would be permanently unreachable if defined.
    CheckExpandErr "quasiquote: 'symbol' refused as a macro name", _
        "(defmacro (symbol x) x)", "'symbol' is the identifier-fusion primitive and cannot be a macro name"
    CheckExpandErr "quasiquote: 'quasiquote' refused as a macro name", _
        "(defmacro (quasiquote x) x)", "'quasiquote' is the template-shielding form and cannot be a macro name"
    CheckExpandErr "quasiquote: 'unquote' refused as a macro name", _
        "(defmacro (unquote x) x)", "'unquote' is the quasiquote escape form and cannot be a macro name"
    CheckExpandErr "quasiquote: 'unquote-splicing' refused as a macro name", _
        "(defmacro (unquote-splicing x) x)", "'unquote-splicing' is the quasiquote splice form and cannot be a macro name"

    ' --- core behavior: quasiquote shields, unquote punches through -
    '     the same param name in both roles in one template, so a
    '     regression in either direction (shield leaks, or unquote
    '     stays shielded) shows up as the wrong output, not a crash.
    t = VlaExpandText("(defmacro (demo x) (quasiquote (x (unquote x)))) (demo hello)", True, fired)
    CheckFrags "quasiquote: shields a literal symbol while unquote punches through it, same name both roles", t, _
               Array("(x hello)")

    ' unquote-splicing generalizes the existing bare-restName splice
    ' beyond the implicit case, explicit inside a shield.
    t = VlaExpandText("(defmacro (demo & rest) (quasiquote (literal (unquote-splicing rest) end))) (demo a b c)", True, fired)
    CheckFrags "quasiquote: unquote-splicing splices a rest param explicitly inside a shield", t, _
               Array("(literal a b c end)")

    ' Ordinary (non-quasiquoted) template substitution is untouched -
    ' the backward-compatibility guarantee inQuasi's False default
    ' exists for, pinned directly rather than just asserted.
    t = VlaExpandText("(defmacro (make-adder n) (begin (dim total Long) (set! total (+ total n)))) (make-adder 5)", True, fired)
    CheckFrags "quasiquote: ordinary template substitution is unaffected by inQuasi's existence", t, _
               Array("(set! total (+ total 5))")

    ' --- guard rails: every one fails loud, never silently, by design.
    CheckExpandErr "quasiquote: nested quasiquote refused (data-structure mismatch, not a stack risk)", _
        "(defmacro (demo x) (quasiquote (a (quasiquote (unquote x))))) (demo hello)", _
        "nested quasiquote is not supported"

    CheckExpandErr "quasiquote: unquote with no enclosing macro at all (ExpandMacros's own guard)", _
        "(unquote x)", "'unquote' only has meaning inside a defmacro's own template"

    CheckExpandErr "quasiquote: unquote inside a macro that never opens a quasiquote (Substitute's own guard)", _
        "(defmacro (demo x) (unquote x)) (demo hello)", "unquote used outside quasiquote"

    CheckExpandErr "quasiquote: unquote-splicing inside a macro that never opens a quasiquote", _
        "(defmacro (demo x) (unquote-splicing x)) (demo hello)", "unquote-splicing used outside quasiquote"

    CheckExpandErr "quasiquote: unquote-splicing used as a bare value, not a list element", _
        "(defmacro (demo x) (quasiquote (unquote-splicing x))) (demo hello)", _
        "unquote-splicing must appear as a list element, not as a value"

    CheckExpandErr "quasiquote: unquote-splicing operand that doesn't resolve to a list", _
        "(defmacro (demo x) (quasiquote (a (unquote-splicing x) b))) (demo hello)", _
        "unquote-splicing: operand did not resolve to a list"

    CheckExpandErr "quasiquote: wrong arity refused explicitly", _
        "(defmacro (demo x) (quasiquote a b)) (demo hello)", _
        "quasiquote expects exactly 1 argument, got 2"

    CheckExpandErr "quasiquote: unquote wrong arity refused explicitly", _
        "(defmacro (demo x) (quasiquote (unquote x y))) (demo hello)", _
        "unquote expects exactly 1 argument, got 2"

    CheckExpandErr "quasiquote: unquote-splicing wrong arity refused explicitly", _
        "(defmacro (demo x) (quasiquote (a (unquote-splicing x y) b))) (demo hello)", _
        "unquote-splicing expects exactly 1 argument, got 2"
End Sub

' ---------------------------------------------------------------------
'  LISTOPS.0: the complete expand-time data bundle - car/cdr/cddr/cons/
'  list/null?/eq?/equal?/quote-if. Reopens L-TIER3's declined
'  "procedural macros" question, owner-gated the same way QUASIQUOTE's
'  bundle was, in two rounds: first the general L-TIER3 reading (a
'  recursive expand-time walker over quoted data), then the 9th
'  primitive (quote-if) specifically, once no expand-time conditional
'  anywhere in the codebase turned out to make the named 8-primitive
'  bundle alone insufficient for genuine termination. Grouped in build
'  order: each eager primitive standalone (including the quote-
'  unwrapping LISTOPS-PURITY's own data boundary requires), reserved
'  names, a bound-template-parameter case, the quasiquote-shielding
'  interaction (decided differently from `symbol` - see EvalListopsPrim's
'  own header comment), then the acceptance test - a self-terminating
'  recursive walker composing all nine, proving TABLESPEC's own
'  predicted shape actually works, not just that the primitives exist
'  in isolation.
' ---------------------------------------------------------------------
Private Sub TestListops()
    Dim t As String
    Dim fired As Long

    ' --- car/cdr/cddr: standalone, unwrapping (quote ...) to the
    '     literal data underneath - the data boundary LISTOPS-PURITY
    '     names, not something Substitute does for ordinary templates.
    t = VlaExpandText("(car (quote (1 2 3)))", True, fired)
    CheckFrags "listops: car returns the first element, unwrapping quote", t, Array("1")

    t = VlaExpandText("(cdr (quote (1 2 3)))", True, fired)
    CheckFrags "listops: cdr returns a new list of the remaining elements", t, Array("(2 3)")

    t = VlaExpandText("(cddr (quote (1 2 3 4)))", True, fired)
    CheckFrags "listops: cddr skips two elements", t, Array("(3 4)")

    CheckExpandErr "listops: car refuses an empty list", _
        "(car (quote ()))", "car: expected a non-empty list"

    CheckExpandErr "listops: cdr refuses an empty list", _
        "(cdr (quote ()))", "cdr: expected a list of at least 1 element(s), got 0"

    CheckExpandErr "listops: cddr refuses a 1-element list", _
        "(cddr (quote (1)))", "cddr: expected a list of at least 2 element(s), got 1"

    CheckExpandErr "listops: car refuses a non-list argument", _
        "(car 5)", "car: expected a list, got '5'"

    CheckExpandErr "listops: car refuses the wrong arity", _
        "(car (quote (1 2)) (quote (3 4)))", "car expects exactly 1 argument, got 2"

    ' --- cons/list: build NEW lists, unwrapping quote on every argument.
    t = VlaExpandText("(cons 1 (quote (2 3)))", True, fired)
    CheckFrags "listops: cons prepends an item to a list", t, Array("(1 2 3)")

    CheckExpandErr "listops: cons refuses a non-list second argument", _
        "(cons 1 2)", "cons: second argument must be a list, got '2'"

    t = VlaExpandText("(list 1 2 3)", True, fired)
    CheckFrags "listops: list builds a list from N arguments", t, Array("(1 2 3)")

    t = VlaExpandText("(list)", True, fired)
    CheckFrags "listops: list with zero arguments builds the empty list", t, Array("()")

    ' --- null?/eq?/equal?: return the SYMBOL true/false, never a raw
    '     VBA Boolean - see EvalListopsPrim's own header comment in
    '     VLA.bas for why a raw Boolean would be unsafe here.
    t = VlaExpandText("(null? (quote ()))", True, fired)
    CheckFrags "listops: null? is true for the empty list", t, Array("true")

    t = VlaExpandText("(null? (quote (1)))", True, fired)
    CheckFrags "listops: null? is false for a non-empty list", t, Array("false")

    t = VlaExpandText("(null? 5)", True, fired)
    CheckFrags "listops: null? is false for a non-list atom, never an error", t, Array("false")

    t = VlaExpandText("(eq? (quote a) (quote a))", True, fired)
    CheckFrags "listops: eq? is true for two equal symbols", t, Array("true")

    t = VlaExpandText("(eq? (quote a) (quote b))", True, fired)
    CheckFrags "listops: eq? is false for two different symbols", t, Array("false")

    CheckExpandErr "listops: eq? refuses a list operand (use equal? instead)", _
        "(eq? (quote (1)) (quote (1)))", "eq?: expected an atom, got a list"

    t = VlaExpandText("(equal? (quote (1 2 (3 4))) (quote (1 2 (3 4))))", True, fired)
    CheckFrags "listops: equal? is true for deeply-equal structures", t, Array("true")

    t = VlaExpandText("(equal? (quote (1 2 (3 4))) (quote (1 2 (3 5))))", True, fired)
    CheckFrags "listops: equal? is false when a nested element differs", t, Array("false")

    ' --- quote-if: the one special form - only the SELECTED branch is
    '     ever resolved, proven by an untaken branch that would raise
    '     if it were ever touched.
    t = VlaExpandText("(quote-if (null? (quote ())) (debug-print ""base"") (car (quote ())))", True, fired)
    CheckFrags "listops: quote-if never resolves the untaken branch (would raise if it did)", t, Array("base")

    t = VlaExpandText("(quote-if (null? (quote (1))) (car (quote ())) (debug-print ""recurse""))", True, fired)
    CheckFrags "listops: quote-if resolves the taken branch when the test is false", t, Array("recurse")

    CheckExpandErr "listops: quote-if refuses a non-true/false test", _
        "(quote-if 5 (debug-print ""a"") (debug-print ""b""))", "quote-if: test must resolve to true or false"

    CheckExpandErr "listops: quote-if refuses the wrong arity", _
        "(quote-if (null? (quote ())) (debug-print ""a""))", "quote-if expects exactly 3 arguments"

    ' --- reserved names: none of the nine can be shadowed as a macro
    '     name - each would be permanently unreachable if defined.
    CheckExpandErr "listops: 'car' refused as a macro name", _
        "(defmacro (car x) x)", "'car' is the list-head primitive and cannot be a macro name"
    CheckExpandErr "listops: 'cdr' refused as a macro name", _
        "(defmacro (cdr x) x)", "'cdr' is the list-tail primitive and cannot be a macro name"
    CheckExpandErr "listops: 'cddr' refused as a macro name", _
        "(defmacro (cddr x) x)", "'cddr' is the list-tail primitive and cannot be a macro name"
    CheckExpandErr "listops: 'cons' refused as a macro name", _
        "(defmacro (cons x) x)", "'cons' is the list-construction primitive and cannot be a macro name"
    CheckExpandErr "listops: 'list' refused as a macro name", _
        "(defmacro (list x) x)", "'list' is the list-construction primitive and cannot be a macro name"
    CheckExpandErr "listops: 'null?' refused as a macro name", _
        "(defmacro (null? x) x)", "'null?' is the empty-list primitive and cannot be a macro name"
    CheckExpandErr "listops: 'eq?' refused as a macro name", _
        "(defmacro (eq? x) x)", "'eq?' is the atom-comparison primitive and cannot be a macro name"
    CheckExpandErr "listops: 'equal?' refused as a macro name", _
        "(defmacro (equal? x) x)", "'equal?' is the structural-comparison primitive and cannot be a macro name"
    CheckExpandErr "listops: 'quote-if' refused as a macro name", _
        "(defmacro (quote-if x) x)", "'quote-if' is the expand-time conditional and cannot be a macro name"

    ' --- inside a defmacro template, using the enclosing macro's own
    '     bindings/restName/restItems - the same machinery car/cdr/cons
    '     resolve their own arguments through, no new evaluation model.
    t = VlaExpandText("(defmacro (first-of xs) (car xs)) (first-of (quote (7 8 9)))", True, fired)
    CheckFrags "listops: car resolves a bound template parameter, not just a literal", t, Array("7")

    ' --- quasiquote interaction: an explicit unquote is the real,
    '     intended way to combine LISTOPS with quasiquote - computing a
    '     primitive's value INSIDE the shield, not leaving the bare form
    '     for later (see VLA.bas's own LISTOPS.0 header comment for why
    '     a shielded-but-never-unquoted call was always a dead end
    '     either way, not something this pass needs to make bulletproof).
    t = VlaExpandText("(defmacro (demo xs) (quasiquote (unquote (car xs)))) (demo (quote (1 2 3)))", True, fired)
    CheckFrags "listops: an explicit unquote inside a shield evaluates car normally", t, Array("1")

    ' --- TABLESPEC-shaped acceptance test: a self-terminating recursive
    '     macro walking a 5-element quoted list via quote-if/null?/car/
    '     cdr, emitting one statement per row - proof the bundle
    '     composes the way the roadmap's own TABLESPEC entry predicts,
    '     not a TABLESPEC ship. Base case is bare (begin), NOT
    '     (quote (begin)) - quote-if's branches are CODE (resolved by
    '     plain Substitute), not data, and EmitStmt's own "quote is an
    '     expression, not a statement" refusal would fire if a quoted
    '     form ever reached statement position here.
    Dim walker As String
    walker = "(defmacro (walk-rows lst) " & _
             "(quote-if (null? lst) " & _
             "(begin) " & _
             "(begin (debug-print (car lst)) (walk-rows (cdr lst)))))"
    Dim vbaWalk As String, nWalk As Long, dWalk As String
    On Error Resume Next
    Err.Clear
    vbaWalk = VlaTranspile(walker & vbCrLf & "(sub t () (walk-rows (quote (10 20 30 40 50))))")
    nWalk = Err.Number
    dWalk = Err.Description
    On Error GoTo 0
    Report "listops: a self-terminating recursive walker emits one statement per row, in order", _
           nWalk = 0 And _
           InStr(1, vbaWalk, "Debug.Print 10", vbTextCompare) > 0 And _
           InStr(1, vbaWalk, "Debug.Print 50", vbTextCompare) > InStr(1, vbaWalk, "Debug.Print 10", vbTextCompare), _
           "err: " & dWalk & " / vba: " & Left$(vbaWalk, 300)
End Sub

' ---------------------------------------------------------------------
'  LISTOPS-EXPAND: the seven expand-time arithmetic/comparison
'  primitives (+expand/=expand/<>expand/>expand/<expand/>=expand/
'  <=expand) - constant folding over LITERAL numeric operands only,
'  scoped and adjudicated in BETA_ROADMAP.md before a line of this was
'  written (see VLA_CORE_VERSION's own LISTOPSEXPAND.0 comment,
'  VLA.bas). Non-numeric-operand cases prove the strict gate (guardrail
'  2) rather than a silent coercion; the "not a canonical literal"
'  cases prove IsNumericLiteralText's own scope, not VBA's IsNumeric,
'  is what decides - the locale-invariance guarantee itself needs
'  Application.International actually switched to observe directly,
'  which a single-locale CI run cannot do (stated honestly in the
'  roadmap entry, not silently skipped here).
' ---------------------------------------------------------------------
Private Sub TestListopsExpand()
    Dim t As String
    Dim fired As Long

    ' --- +expand: folds two literal numbers to a new literal number.
    t = VlaExpandText("(+expand 2 3)", True, fired)
    CheckFrags "listops-expand: +expand folds two literals", t, Array("5")

    t = VlaExpandText("(+expand -2 3)", True, fired)
    CheckFrags "listops-expand: +expand accepts a negative literal", t, Array("1")

    t = VlaExpandText("(+expand 2.5 1.5)", True, fired)
    CheckFrags "listops-expand: +expand folds fractional literals", t, Array("4")

    ' --- the six comparisons: return the SYMBOL true/false, same
    '     convention null?/eq?/equal? already use.
    t = VlaExpandText("(=expand 5 5)", True, fired)
    CheckFrags "listops-expand: =expand is true for equal numbers", t, Array("true")

    t = VlaExpandText("(=expand 5 6)", True, fired)
    CheckFrags "listops-expand: =expand is false for unequal numbers", t, Array("false")

    t = VlaExpandText("(<>expand 5 6)", True, fired)
    CheckFrags "listops-expand: <>expand is true for unequal numbers", t, Array("true")

    t = VlaExpandText("(>expand 5 3)", True, fired)
    CheckFrags "listops-expand: >expand is true when the first is greater", t, Array("true")

    t = VlaExpandText("(>expand 3 5)", True, fired)
    CheckFrags "listops-expand: >expand is false when the first is not greater", t, Array("false")

    t = VlaExpandText("(<expand 3 5)", True, fired)
    CheckFrags "listops-expand: <expand is true when the first is smaller", t, Array("true")

    t = VlaExpandText("(>=expand 5 5)", True, fired)
    CheckFrags "listops-expand: >=expand is true on equality", t, Array("true")

    t = VlaExpandText("(<=expand 5 5)", True, fired)
    CheckFrags "listops-expand: <=expand is true on equality", t, Array("true")

    ' --- strict, not heuristic: a non-numeric operand raises, never
    '     coerces (guardrail 2, BETA_ROADMAP.md's own LISTOPS-EXPAND
    '     entry) - same discipline eq?'s atom-only requirement holds.
    CheckExpandErr "listops-expand: +expand refuses a non-numeric operand", _
        "(+expand 2 (quote x))", "+expand: expected a number, got 'x'"

    CheckExpandErr "listops-expand: >expand refuses a list operand", _
        "(>expand (quote (1)) 2)", ">expand: expected a number, got a list"

    ' --- IsNumericLiteralText's own strict scope - no thousands
    '     separator, no exponent notation - deliberately NOT VBA's own
    '     IsNumeric, which would accept locale-specific variants of
    '     these and break "same source text, same result everywhere".
    CheckExpandErr "listops-expand: +expand refuses a thousands separator", _
        "(+expand (quote 1,000) 1)", "+expand: expected a number, got '1,000'"

    ' --- arity: exactly 2 arguments, same convention every LISTOPS
    '     primitive already holds.
    CheckExpandErr "listops-expand: +expand refuses the wrong arity", _
        "(+expand 2)", "+expand expects exactly 2 arguments, got 1"

    ' --- reserved names: none of the seven can be shadowed as a macro
    '     name, same treatment the original nine already get.
    CheckExpandErr "listops-expand: '+expand' refused as a macro name", _
        "(defmacro (+expand x) x)", "'+expand' is the expand-time addition primitive and cannot be a macro name"
    CheckExpandErr "listops-expand: '>expand' refused as a macro name", _
        "(defmacro (>expand x) x)", "'>expand' is the expand-time comparison primitive and cannot be a macro name"

    ' --- standalone AND nested, same dual-site treatment LISTOPS.0's
    '     own nine already get - proof the two new recognition-site
    '     Select Case additions (ExpandMacros/Substitute, VLA.bas) both
    '     actually fired, not just one of them.
    t = VlaExpandText("(quote-if (>expand 5 3) (debug-print ""yes"") (debug-print ""no""))", True, fired)
    CheckFrags "listops-expand: >expand composes as quote-if's own test, standalone", t, Array("yes")

    t = VlaExpandText("(defmacro (bigger x) (quote-if (>expand x 10) (quote yes) (quote no))) (bigger 20)", True, fired)
    CheckFrags "listops-expand: >expand resolves a bound template parameter, nested inside a macro", t, Array("yes")
End Sub

' ---------------------------------------------------------------------
'  COND.0: the tenth engine primitive, unifying the expand-time and
'  runtime halves of `cond` under one name (BETA_ROADMAP.md's own COND
'  entry - a first pass split them into two separate roadmap items that
'  both wanted the name `cond`, caught before either was built).
' ---------------------------------------------------------------------
Private Sub TestCond()
    Dim t As String
    Dim fired As Long

    ' --- expand-time folding: a test that resolves to true/false is
    '     folded immediately, the untaken branch never touched (proven
    '     the same way quote-if's own test already proves it - a branch
    '     that would raise if it were ever resolved).
    t = VlaExpandText("(cond (true (debug-print ""a"")) (false (car (quote ()))))", True, fired)
    CheckFrags "cond: a true clause folds immediately, the untaken clause never touched", t, Array("a")

    t = VlaExpandText("(cond (false (car (quote ()))) (true (debug-print ""b"")))", True, fired)
    CheckFrags "cond: a false clause is skipped, never resolved, the next clause tried", t, Array("b")

    ' --- the standalone case that broke the original naive defmacro-
    '     based attempt (BETA_ROADMAP.md's own COND entry history) - a
    '     compound LISTOPS test, called at the TOP LEVEL, no enclosing
    '     defmacro. This is the entire reason cond became an engine
    '     primitive instead of staying a prelude.vla macro.
    t = VlaExpandText("(cond ((null? (quote ())) (debug-print ""yes"")) (else (debug-print ""no"")))", True, fired)
    CheckFrags "cond: a compound LISTOPS test folds correctly when called standalone", t, Array("yes")

    t = VlaExpandText("(cond ((null? (quote (1))) (debug-print ""wrong"")) (else (debug-print ""right"")))", True, fired)
    CheckFrags "cond: falls through to else when no test matches", t, Array("right")

    ' --- runtime deferral: a genuine runtime expression (a real
    '     variable, not a literal) is spliced unresolved into a native
    '     if/then/else, evaluated at RUNTIME exactly like a hand-written
    '     `if` - real Lisp cond semantics for the traditional use.
    Dim vba1 As String, n1 As Long, d1 As String
    On Error Resume Next
    Err.Clear
    vba1 = VlaTranspile("(sub t () (dim x) (set! x 5) (cond ((> x 0) (debug-print ""pos"")) (else (debug-print ""nonpos""))))")
    n1 = Err.Number
    d1 = Err.Description
    On Error GoTo 0
    Report "cond: a runtime test defers to a native If/Else, never folded", _
           n1 = 0 And _
           InStr(1, vba1, "If (x > 0) Then", vbTextCompare) > 0 And _
           InStr(1, vba1, "Debug.Print ""pos""", vbTextCompare) > 0 And _
           InStr(1, vba1, "Else", vbTextCompare) > 0 And _
           InStr(1, vba1, "Debug.Print ""nonpos""", vbTextCompare) > 0, _
           "err: " & d1 & " / vba: " & Left$(vba1, 300)

    ' --- a later clause that DOES fold true gets hoisted straight into
    '     the enclosing else slot, not wrapped in another redundant
    '     `if` - no "If True Then" anywhere in the emitted VBA.
    Dim vba2 As String, n2 As Long, d2 As String
    On Error Resume Next
    Err.Clear
    vba2 = VlaTranspile("(sub t () (dim x) (set! x 5) (cond ((> x 0) (debug-print ""pos"")) (true (debug-print ""fallback""))))")
    n2 = Err.Number
    d2 = Err.Description
    On Error GoTo 0
    Report "cond: a later folded-true clause is hoisted into else, not re-wrapped", _
           n2 = 0 And _
           InStr(1, vba2, "If (x > 0) Then", vbTextCompare) > 0 And _
           InStr(1, vba2, "Debug.Print ""fallback""", vbTextCompare) > 0 And _
           InStr(1, vba2, "If True Then", vbTextCompare) = 0, _
           "err: " & d2 & " / vba: " & Left$(vba2, 300)

    ' --- no match and no else: expands to (begin), a deliberate no-op,
    '     not an error - matches walk-rows's own base-case convention.
    Dim vba3 As String, n3 As Long, d3 As String
    On Error Resume Next
    Err.Clear
    vba3 = VlaTranspile("(sub t () (cond (false (debug-print ""a""))))")
    n3 = Err.Number
    d3 = Err.Description
    On Error GoTo 0
    Report "cond: no match and no else is a silent no-op, never an error", _
           n3 = 0 And InStr(1, vba3, "Debug.Print", vbTextCompare) = 0, _
           "err: " & d3 & " / vba: " & Left$(vba3, 300)

    t = VlaExpandText("(cond)", True, fired)
    CheckFrags "cond: zero clauses also expands to (begin), not an error", t, Array("begin")

    ' --- strictness: a misplaced else, and a malformed clause, both
    '     raise loudly - the strictness-over-smoothing value this
    '     project holds everywhere else.
    CheckExpandErr "cond: a misplaced 'else' (not last) raises", _
        "(cond (else (debug-print ""a"")) (true (debug-print ""b"")))", "cond: 'else' must be the last clause"

    CheckExpandErr "cond: a malformed clause (wrong element count) raises", _
        "(cond ((> 1 0)))", "cond: each clause must be (test form) or (else form), got 1 element(s)"

    ' --- reserved name: 'cond' cannot be shadowed as a macro name.
    CheckExpandErr "cond: 'cond' refused as a macro name", _
        "(defmacro (cond x) x)", "'cond' is the multi-clause conditional"

    ' --- nested inside an ordinary defmacro - the case that already
    '     worked for quote-if, must keep working identically for cond.
    t = VlaExpandText("(defmacro (branch x) (cond ((null? x) (quote empty)) (else (quote nonempty)))) (branch (quote ()))", True, fired)
    CheckFrags "cond: composes inside an ordinary defmacro, resolving a bound parameter's own argument", t, Array("empty")
End Sub

' ---------------------------------------------------------------------
'  LISTOPS.0 / LISTOPS-BUDGET follow-up. FIRST version of this test
'  guessed the depth guard's raise to 1000 was safe at 900 rows -
'  disproven by a live host run: VBA's real native stack exhausted with
'  "Out of stack space" at 900, AND at the pre-existing 250-deep
'  TestListopsBudget chain, which used to pass cleanly at the guard's
'  prior value of 200. The guard is reverted to 200 (VLA.bas); this
'  test no longer asserts that any SPECIFIC large row-count is
'  physically safe for the heavier walk-rows shape - that claim is
'  exactly what already went wrong once.
'
'  TABLESPEC-SCALE / TABLESPECSCALE.0: the claim that went wrong is now
'  actually true, because the mechanism it was wrong about (native VBA
'  stack growing per chain link) no longer exists - `ExpandMacros`'s own
'  head-resolution chain is hand-trampolined (VLA.bas), uncoupling
'  chain length from stack usage entirely, regardless of how much extra
'  work (like walk-rows's own quote-if/car/cdr/begin, heavier than a
'  bare substitution chain) each link does. The 900-row case below now
'  expands cleanly; a new 5500-row case (past the trampoline's own new,
'  purely-logical 5000 cap) proves the heavier shape ALSO hits that cap
'  loudly, not just the lighter chain TestListopsBudget already covers.
' ---------------------------------------------------------------------
Private Sub TestListopsDepthSafety()
    Dim walker As String
    walker = "(defmacro (walk-rows lst) " & _
             "(quote-if (null? lst) " & _
             "(begin) " & _
             "(begin (debug-print (car lst)) (walk-rows (cdr lst)))))"

    ' 20 rows: a modest step up from TestListops's own already-passing
    ' 5-row acceptance test, still far below any boundary this session
    ' has live evidence for in either direction.
    Dim items As String, i As Long
    items = ""
    For i = 1 To 20
        If Len(items) > 0 Then items = items & " "
        items = items & i
    Next
    Dim src20 As String
    src20 = walker & vbCrLf & "(sub t () (walk-rows (quote (" & items & "))))"
    Dim vba20 As String, n20 As Long, d20 As String
    On Error Resume Next
    Err.Clear
    vba20 = VlaTranspile(src20)
    n20 = Err.Number
    d20 = Err.Description
    On Error GoTo 0
    Report "listops-depth: a 20-row walk-rows recursion expands completely", _
           n20 = 0 And InStr(1, vba20, "Debug.Print 20", vbTextCompare) > 0, _
           "err: " & d20 & " / len: " & Len(vba20)

    ' TABLESPEC-SCALE / TABLESPECSCALE.0: 900 rows used to fail past the
    ' (reverted) 200 ceiling either way - now expands cleanly, the
    ' trampolined head-chain no longer costing a VBA stack frame per row
    ' regardless of walk-rows's own heavier per-link work.
    ' TABLESPEC-SCALE / TABLESPECSCALE.0: pre-sized array + Join, not
    ' `items2 = items2 & ...` in the loop - see the matching note on
    ' TestListopsBudget's own src250/src2000/src5500 for why (VLA.bas's
    ' SbAdd comment: naive concatenation is O(n^2), "dominant at 5,000" -
    ' confirmed live, Excel visibly hanging on this test's own 5500-row
    ' sibling below before this fix).
    Dim parts900() As String
    ReDim parts900(1 To 900)
    For i = 1 To 900
        parts900(i) = CStr(i)
    Next
    Dim items2 As String
    items2 = Join(parts900, " ")
    Dim src900 As String
    src900 = walker & vbCrLf & "(sub t () (walk-rows (quote (" & items2 & "))))"
    Dim vba900 As String, n900 As Long, d900 As String
    On Error Resume Next
    Err.Clear
    vba900 = VlaTranspile(src900)
    n900 = Err.Number
    d900 = Err.Description
    On Error GoTo 0
    Report "listops-depth: a 900-row walk-rows recursion now expands completely - the trampolined chain has no native-stack cost per row", _
           n900 = 0 And InStr(1, vba900, "Debug.Print 900", vbTextCompare) > 0, _
           "err: " & d900 & " / len: " & Len(vba900)

    ' TABLESPEC-SCALE / TABLESPECSCALE.0: past the trampoline's own new,
    ' purely-logical 5000 cap (ExpandMacros, VLA.bas) - fails loudly for
    ' the HEAVIER walk-rows shape too, not just the bare substitution
    ' chain TestListopsBudget's own 5500-row case already covers
    ' (gated there too - see its own comment, corrected: GetMacro's
    ' registry lookup turned out to share this same cost family, not
    ' cheap the way this comment originally assumed).
    ' Gated on mRunScaleTests: cdr/ListTail (VLA.bas) copies the ENTIRE
    ' remaining tail on every call, so walking 5000+ of a 5500-element
    ' list to reach the cap costs ~15M element copies - several minutes,
    ' confirmed live - a pre-existing O(n^2) cost the trampoline never
    ' touched, not something worth paying on every ordinary VlaSelfTest
    ' run. See mRunScaleTests's own declaration above for the full
    ' reasoning; run VlaSelfTestScale instead of VlaSelfTest to include
    ' this when a change actually touches ExpandMacros/cdr/ListTail.
    If mRunScaleTests Then
        Dim parts5500w() As String
        ReDim parts5500w(1 To 5500)
        For i = 1 To 5500
            parts5500w(i) = CStr(i)
        Next
        Dim items3 As String
        items3 = Join(parts5500w, " ")
        Dim src5500w As String
        src5500w = walker & vbCrLf & "(sub t () (walk-rows (quote (" & items3 & "))))"
        Dim vba5500w As String, n5500w As Long, d5500w As String
        On Error Resume Next
        Err.Clear
        vba5500w = VlaTranspile(src5500w)
        n5500w = Err.Number
        d5500w = Err.Description
        On Error GoTo 0
        Report "listops-depth: a 5500-row walk-rows recursion - past the trampoline's own 5000 cap - fails loudly, never a silent partial expansion", _
               n5500w <> 0 And Len(vba5500w) = 0 And InStr(1, d5500w, "5000", vbTextCompare) > 0, _
               "err: " & d5500w & " / vba: " & Left$(vba5500w, 200)
    End If
End Sub

' ---------------------------------------------------------------------
'  TABLESPEC (BETA_ROADMAP.md's own entry) depth-safety proof - the
'  same TestListopsDepthSafety discipline applied to tablespec/
'  tablespec-row (scripts/english.vla) instead of the acceptance test's
'  own walk-rows: tablespec is ALSO a self-recursive macro, so it
'  inherits the identical ExpandMacros depth>200 guard, unchanged.
'  tablespec/tablespec-row are redefined inline here rather than loaded
'  from scripts/english.vla - this test proves the WALK MECHANISM in
'  isolation via bare VlaExpandText (no vocabulary loader needed, same
'  reason TestListopsStdlib tests prelude.vla functions directly) -
'  TestTablespec (VLA_Tests_Grammar.bas) separately proves the REAL
'  copy composes with the vocabulary loader end to end.
'  20 rows, not a larger guessed number: tablespec does MORE work per
'  row than walk-rows (seven field accessors plus a tablespec-row call,
'  not one debug-print), so 20 - walk-rows' own already-proven-safe
'  number - is the honest, evidence-grounded claim; LISTOPS-BUDGET's
'  own first-round mistake was guessing a larger number was safe
'  without a live test, disproven by a real host run.
'
'  TABLESPEC-SCALE / TABLESPECSCALE.0: 250 rows - this item's own
'  original documented ceiling - now expands cleanly instead of failing,
'  `ExpandMacros`'s trampolined head-chain (VLA.bas) having uncoupled
'  row count from native VBA stack usage entirely, regardless of
'  tablespec's own heavier per-row work. A new 2000-row case proves this
'  at genuinely large, realistic table scale (TABLESPEC-SCALE's own
'  motivating case: "a 250-row pricing sheet is entirely ordinary" no
'  longer needs hedging); a new 5500-row case proves the trampoline's
'  own new, purely-logical 5000 cap still fires loudly for TABLESPEC's
'  own heavier shape, not just the lighter ones TestListopsBudget/
'  TestListopsDepthSafety already cover.
' ---------------------------------------------------------------------
Private Sub TestTablespecDepthSafety()
    Dim t As String
    Dim fired As Long
    Dim gen As String
    gen = "(defmacro (tablespec-row name doc body pattern call testsentence testcall) " & _
          "(begin (defmacro (name n) doc body) (english-vla pattern call) (test-success testsentence testcall))) " & _
          "(defmacro (tablespec spec) (quote-if (null? spec) (begin) " & _
          "(begin (tablespec-row (car (car spec)) (cadr (car spec)) (car (cddr (car spec))) " & _
          "(cadr (cddr (car spec))) (caddr (cddr (car spec))) " & _
          "(cadr (cddr (cddr (car spec)))) (caddr (cddr (cddr (car spec))))) " & _
          "(tablespec (cdr spec)))))"

    Dim items As String, i As Long
    items = ""
    For i = 1 To 20
        If Len(items) > 0 Then items = items & " "
        items = items & "(row" & i & " ""d"" (debug-print " & i & ") ""p"" (c) ""s"" (c))"
    Next
    Dim n20 As Long, d20 As String
    On Error Resume Next
    Err.Clear
    t = VlaExpandText(gen & " (tablespec (quote (" & items & ")))", True, fired)
    n20 = Err.Number
    d20 = Err.Description
    On Error GoTo 0
    ' First live run found this substring check itself broken, not the
    ' walk: VlaExpandText renders through WritePretty (VLA.bas), which
    ' line-wraps any list whose flat form exceeds 90 columns (indent
    ' included) - a 20-row nested-begin tree trips that well before row
    ' 20, splitting "debug-print" and "20" across a line break, so the
    ' contiguous substring was never going to be there regardless of
    ' whether the walk itself was correct. Stripping whitespace first
    ' makes the check indifferent to how the pretty-printer wraps.
    Dim tFlat As String
    tFlat = Replace(Replace(t, vbCrLf, ""), " ", "")
    Report "tablespec-depth: a 20-row walk expands completely, reaching the last row", _
           n20 = 0 And InStr(1, tFlat, "debug-print20", vbTextCompare) > 0, _
           "err: " & d20 & " / len: " & Len(t)

    ' TABLESPEC-SCALE / TABLESPECSCALE.0: pre-sized array + Join, not
    ' `items2 = items2 & ...` in the loop - naive concatenation is
    ' O(n^2) in the result's length (VLA.bas's own SbAdd comment,
    ' "dominant at 5,000"), confirmed live - Excel visibly hanging on
    ' this test's own 2000/5500-row siblings below before this fix, this
    ' test's own rows being wider strings per item than the plain-number
    ' chains elsewhere, making the blowup worse, not better.
    Dim parts250t() As String
    ReDim parts250t(1 To 250)
    For i = 1 To 250
        parts250t(i) = "(row" & i & " ""d"" (debug-print " & i & ") ""p"" (c) ""s"" (c))"
    Next
    Dim items2 As String
    items2 = Join(parts250t, " ")
    Dim n250 As Long, d250 As String
    ' Second bug this test's own first live run found: `t` reused across
    ' calls without resetting - fixed the same way below (each block now
    ' clears `t` before its own VlaExpandText call, honest about what
    ' THIS call actually produced, not a stale carry-over).
    t = ""
    On Error Resume Next
    Err.Clear
    t = VlaExpandText(gen & " (tablespec (quote (" & items2 & ")))", True, fired)
    n250 = Err.Number
    d250 = Err.Description
    On Error GoTo 0
    Dim t250Flat As String
    t250Flat = Replace(Replace(t, vbCrLf, ""), " ", "")
    Report "tablespec-depth: a 250-row walk now expands completely - the trampolined chain has no native-stack cost per row", _
           n250 = 0 And InStr(1, t250Flat, "debug-print250", vbTextCompare) > 0, _
           "err: " & d250 & " / len: " & Len(t)

    ' TABLESPEC-SCALE / TABLESPECSCALE.0: 2000/5500 rows - genuinely
    ' large, realistic table scale, not just a token step past the old
    ' ceiling. Gated on mRunScaleTests, same reasoning as
    ' TestListopsDepthSafety's own 5500-row case above: tablespec walks
    ' via the same cdr/ListTail (VLA.bas), which copies the ENTIRE
    ' remaining tail on every call - O(n^2) total, several minutes at
    ' 5500 rows (worse than walk-rows's own equivalent case, since
    ' tablespec does more work per row too), a pre-existing cost the
    ' trampoline never touched. See mRunScaleTests's own declaration
    ' above; run VlaSelfTestScale instead of VlaSelfTest to include
    ' this when a change actually touches ExpandMacros/cdr/ListTail/
    ' tablespec's own walk.
    If mRunScaleTests Then
        Dim parts2000t() As String
        ReDim parts2000t(1 To 2000)
        For i = 1 To 2000
            parts2000t(i) = "(row" & i & " ""d"" (debug-print " & i & ") ""p"" (c) ""s"" (c))"
        Next
        Dim items2000 As String
        items2000 = Join(parts2000t, " ")
        Dim n2000 As Long, d2000 As String
        t = ""
        On Error Resume Next
        Err.Clear
        t = VlaExpandText(gen & " (tablespec (quote (" & items2000 & ")))", True, fired)
        n2000 = Err.Number
        d2000 = Err.Description
        On Error GoTo 0
        Dim t2000Flat As String
        t2000Flat = Replace(Replace(t, vbCrLf, ""), " ", "")
        Report "tablespec-depth: a 2000-row walk expands completely - realistic table scale, well past the old native-stack ceiling", _
               n2000 = 0 And InStr(1, t2000Flat, "debug-print2000", vbTextCompare) > 0, _
               "err: " & d2000 & " / len: " & Len(t)

        ' TABLESPEC-SCALE / TABLESPECSCALE.0: past the trampoline's own
        ' new, purely-logical 5000 cap (ExpandMacros, VLA.bas) - fails
        ' loudly for TABLESPEC's own heavier per-row shape too.
        Dim parts5500t() As String
        ReDim parts5500t(1 To 5500)
        For i = 1 To 5500
            parts5500t(i) = "(row" & i & " ""d"" (debug-print " & i & ") ""p"" (c) ""s"" (c))"
        Next
        Dim items5500 As String
        items5500 = Join(parts5500t, " ")
        Dim n5500 As Long, d5500 As String
        t = ""
        On Error Resume Next
        Err.Clear
        t = VlaExpandText(gen & " (tablespec (quote (" & items5500 & ")))", True, fired)
        n5500 = Err.Number
        d5500 = Err.Description
        On Error GoTo 0
        Report "tablespec-depth: a 5500-row walk - past the trampoline's own 5000 cap - fails loudly, never a silent partial expansion", _
               n5500 <> 0 And Len(t) = 0 And InStr(1, d5500, "5000", vbTextCompare) > 0, _
               "err: " & d5500 & " / text: " & Left$(t, 200)
    End If
End Sub

' ---------------------------------------------------------------------
'  LISTOPS.0 follow-up: identity/cadr/caddr/length/reverse/append/
'  quote-map/quote-reduce, prelude.vla's own new small standard library
'  over the nine LISTOPS primitives. Always loaded (prelude.vla is
'  prepended to every compile), so tested directly via VlaExpandText,
'  same as TestListops's own primitives - no vocabulary/English layer
'  involved at all. filter is deliberately absent (see prelude.vla's
'  own comment) - not pinned here because it does not exist.
'  First live run (768/3... no, 781/2) found TWO real bugs, both fixed
'  in prelude.vla, neither caught by tracing beforehand: (1) append's
'  first version produced garbage - (append (quote (1 2)) (quote (3
'  4))) came back "((1 2) reverse 3 4)", not "(1 2 3 4)" - fixed by a
'  continuation-passing redesign (append-reverse-onto), see prelude.vla
'  for the full trace. (2) "map"/"reduce" were already Excel's own
'  native dynamic-array function names, already in real use inside
'  alonzo.vla's own deflambda bodies - a global macro of the same name
'  silently intercepted them and broke TestAlonzoLib. Renamed to
'  quote-map/quote-reduce; the assertions below were updated to match,
'  not just the append fix.
' ---------------------------------------------------------------------
Private Sub TestListopsStdlib()
    Dim t As String
    Dim fired As Long

    t = VlaExpandText("(identity 5)", True, fired)
    CheckFrags "listops-stdlib: identity returns its argument unchanged", t, Array("5")

    t = VlaExpandText("(cadr (quote (1 2 3)))", True, fired)
    CheckFrags "listops-stdlib: cadr is the second element", t, Array("2")

    t = VlaExpandText("(caddr (quote (1 2 3)))", True, fired)
    CheckFrags "listops-stdlib: caddr is the third element", t, Array("3")

    ' COND.0/LISTOPS-EXPAND follow-up: length now folds to a genuine
    ' compile-time literal via +expand (was a runtime sum of 1s before
    ' - see prelude.vla's own comment at length's definition for the
    ' full accumulator-safety trace). Expands all the way to the bare
    ' atom "3", no wrapping form left at all - same shape cadr/caddr's
    ' own pins above already check for.
    t = VlaExpandText("(length (quote (a b c)))", True, fired)
    CheckFrags "listops-stdlib: length folds to a compile-time literal via +expand", t, Array("3")

    CheckExpandErr "listops-stdlib: length still refuses the wrong arity", _
        "(length (quote (a b c)) (quote (d e)))", "macro 'length' expects 1 argument(s), got 2"

    t = VlaExpandText("(reverse (quote (1 2 3)))", True, fired)
    CheckFrags "listops-stdlib: reverse, accumulator-style, reverses a quoted list", t, Array("(3 2 1)")

    t = VlaExpandText("(append (quote (1 2)) (quote (3 4)))", True, fired)
    CheckFrags "listops-stdlib: append concatenates two quoted lists in order (regression pin - this exact call produced garbage before the fix)", t, Array("(1 2 3 4)")

    t = VlaExpandText("(defmacro (double x) (+ x x)) (quote-map double (quote (1 2 3)))", True, fired)
    CheckFrags "listops-stdlib: quote-map applies an ordinary macro to every element, in order", t, _
               Array("(+ 1 1)", "(+ 2 2)", "(+ 3 3)")

    t = VlaExpandText("(defmacro (add a b) (+ a b)) (quote-reduce add 0 (quote (1 2 3)))", True, fired)
    CheckFrags "listops-stdlib: quote-reduce folds an ordinary macro left to right, starting from init", t, _
               Array("(+ (+ (+ 0 1) 2) 3)")
End Sub

' ---------------------------------------------------------------------
'  S2 pins: rule-usage profiling. Firings count once each and
'  accumulate; the claiming construct counts once per top-level
'  sentence; vocab test: proofs deliberately do NOT count; counters
'  are text-keyed, so they outlive grammar resets; the report lists
'  loaded-but-never-fired rules (the dead-rule view G1 will read).
' ---------------------------------------------------------------------
Private Sub TestRuleUsage()
    EnglishRuleUsageReset
    EnglishResetGrammar
    Dim junk As String
    junk = EnglishToVla("Create a number called t." & vbCrLf & "Set t to 5.")
    CheckV "usage: a fired rule is counted once", _
           EnglishRuleUsageCount("rule: set {v:var} to {e:expr}"), 1
    CheckV "usage: the claiming construct is counted", _
           EnglishRuleUsageCount("form: the Create declaration"), 1
    junk = EnglishToVla("Set t to 6.")
    CheckV "usage: counts accumulate across translations", _
           EnglishRuleUsageCount("rule: set {v:var} to {e:expr}"), 2

    EnglishRuleUsageReset
    EnglishResetGrammar
    Dim v As String
    v = "(english-vla ""warm cell {r:text}""" & vbLf & _
        "    (set! (. (range {r}) interior.color) vbyellow))" & vbLf & _
        "(test-success ""Warm cell B2.""" & vbLf & _
        "    (set! (. (range ""b2"") interior.color) vbyellow))"
    EnglishLoadVocabularyText v, "usage-vocab"
    CheckV "usage: vocab proofs are not usage", _
           EnglishRuleUsageCount("rule: warm cell {r:text}"), 0
    junk = EnglishToVla("Warm cell C3.")
    CheckV "usage: a loaded rule's real firing counts", _
           EnglishRuleUsageCount("rule: warm cell {r:text}"), 1
    EnglishResetGrammar
    CheckV "usage: counters survive a grammar reset", _
           EnglishRuleUsageCount("rule: warm cell {r:text}"), 1
    Dim rep As String
    rep = EnglishRuleUsage()
    Report "usage: never-fired rules are listed", _
           InStr(1, rep, "grow {v:var} by {e:expr}", vbTextCompare) > 0, _
           Left$(rep, 200)
    EnglishRuleUsageReset
    CheckV "usage: reset clears counts", _
           EnglishRuleUsageCount("rule: warm cell {r:text}"), 0
End Sub

' ---------------------------------------------------------------------
'  S3 pins: the message seam. Dialog CONTENT is now a pinnable
'  artifact - the V5.2 class of defect (a dialog talking nonsense,
'  invisible to every test) gets its regression home here. The
'  capture flag is restored OFF at section end no matter what, so a
'  dev session after a self-test still gets real dialogs.
' ---------------------------------------------------------------------
Private Sub TestMessageSeam()
    VlaMessageCapture True
    VlaShowError "boom happened"
    Dim cap As String
    cap = VlaCapturedMessages()
    Report "seam: captured message carries text and default title", _
           InStr(1, cap, "Frazaro: boom happened", vbTextCompare) > 0, cap
    VlaShowError "second thing"
    cap = VlaCapturedMessages()
    Report "seam: messages accumulate in order", _
           InStr(1, cap, "boom happened", vbTextCompare) > 0 And _
           InStr(1, cap, "second thing", vbTextCompare) > InStr(1, cap, "boom happened", vbTextCompare), cap
    VlaShowError "titled thing", "My Title"
    cap = VlaCapturedMessages()
    Report "seam: a custom title is honored", _
           InStr(1, cap, "My Title: titled thing", vbTextCompare) > 0, cap
    VlaMessageCapture True
    CheckV "seam: re-enabling capture clears the log", VlaCapturedMessages(), ""

    ' The generated step-failure dialog routes through the seam; the
    ' user's own Show dialog does NOT - both pinned on one program.
    EnglishResetGrammar
    Dim vla As String
    vla = EnglishToVla("Create a number called t." & vbCrLf & "Show t.")
    CheckFrags "seam: step infra routes through the seam", vla, _
               Array("(vlashowerror (& ""Something went wrong at step """)
    Report "seam: the user's Show keeps its own dialog", _
           InStr(1, vla, "(msgbox t)", vbTextCompare) > 0, Left$(Norm(vla), 200)
    Report "seam: no raw msgbox remains in the step infra", _
           InStr(1, Replace(vla, "(msgbox t)", ""), "(msgbox", vbTextCompare) = 0, _
           Left$(Norm(vla), 200)

    ' S3.2: the informational sibling rides the same capture.
    VlaMessageCapture True
    VlaShowInfo "gentle news"
    Report "seam: info messages ride the same capture", _
           InStr(1, VlaCapturedMessages(), "Frazaro: gentle news", vbTextCompare) > 0, _
           VlaCapturedMessages()

    ' S3.2 (owner catch): the step table quotes the user's ORIGINAL
    ' sentence - binary compare, because case is the whole point
    ' (the parsed rendition would read "create a number called Total"
    ' in lowercase).
    Dim vla2 As String
    EnglishStepTracking True   ' the step table only exists when
                               ' tracking is on (its resting default)
    vla2 = EnglishToVla("Create a number called Total.")
    Report "seam: step table quotes the original sentence", _
           InStr(1, vla2, "Create a number called Total. [line 1]", vbBinaryCompare) > 0, _
           Left$(Norm(vla2), 240)
    VlaMessageCapture False
End Sub


' ---------------------------------------------------------------------
'  S4.2 pins: the resolve check - the probe's replacement, and unlike
'  the probe it is PURE STRING WORK, so the self-test can hold the
'  whole contract: clean programs pass (live, against the real
'  manifest), a typo'd vla-helper is named, in-program definitions
'  count, hyphen/underscore spellings unify, string literals never
'  trip it, and the at-line marker threads the source line to the
'  refusal. The sheet-target gate (S4.3) is sheet surgery the
'  self-test cannot reach - the verification loop and D3 own it.
' ---------------------------------------------------------------------
Private Sub TestResolveCheck()
    Dim missing As String
    Dim ml As Long

    ' Live end-to-end: a real tracked translation resolves cleanly -
    ' the step infra's own (vlashowerror ...) call proves the manifest
    ' path against the actual runtime module.
    EnglishResetGrammar
    EnglishStepTracking True
    Dim vla As String
    vla = EnglishToVla("Create a number called t.")
    missing = EnglishResolveCheck(vla, ml)
    CheckV "resolve: a real tracked program resolves cleanly", missing, ""

    missing = EnglishResolveCheck("(sub main () (vlacolr ""#fff""))", ml)
    CheckV "resolve: a typo'd helper is named", missing, "vlacolr"

    missing = EnglishResolveCheck("(sub vla-mine () (set! x 1))" & vbCrLf & _
                                  "(sub main () (vla-mine))", ml)
    CheckV "resolve: in-program definitions count", missing, ""

    missing = EnglishResolveCheck("(sub main () (vla-show-error ""x""))", ml)
    CheckV "resolve: hyphen and underscore spellings unify", missing, ""

    missing = EnglishResolveCheck("(sub main () (debug-print ""(vlacolr)""))", ml)
    CheckV "resolve: string literals never trip it", missing, ""

    missing = EnglishResolveCheck("(sub main () (at-line 7 (vlacolr 1)))", ml)
    Report "resolve: the at-line marker threads the line", _
           missing = "vlacolr" And ml = 7, "got '" & missing & "' at line " & ml

    Dim man As String
    man = VlaHelperManifest()
    Report "resolve: the manifest is live and populated", _
           InStr(1, man, " vlacolor ", vbTextCompare) > 0 And _
           InStr(1, man, " vlashowerror ", vbTextCompare) > 0, Left$(man, 160)
End Sub

' ---------------------------------------------------------------------
'  S5 pins: the runtime trace. The emission contract (every tracked
'  step carries the guarded call, and it survives to VBA under the
'  names the runtime provides) plus the seam's own semantics (order,
'  clearing, off-means-silent). The full Run-with-trace is the
'  verification loop's smoke; F1's parity harness is the consumer.
'  The trace flag is restored OFF at section end unconditionally.
' ---------------------------------------------------------------------
Private Sub TestRuntimeTrace()
    EnglishResetGrammar
    EnglishStepTracking True
    ' L2.2 (owner-directed hardening): the two engine calls here were
    ' bare - the exact hazard L2.1 hit one section down - so a
    ' translation or transpile raise killed the suite mid-run. Both
    ' now ride the guards; each pin still reports exactly once on
    ' every path (a skipped dependent pin reports FAIL explicitly,
    ' never silently vanishing from the printout).
    Dim vla As String
    vla = TryEnglish("trace: every tracked step carries the guarded call", _
                     "Create a number called t." & vbCrLf & "Set t to 5.")
    Dim vbaOut As String
    If Len(vla) > 0 Then
        CheckFrags "trace: every tracked step carries the guarded call", vla, _
                   Array("(if (vlatraceon) (then (vlatracestep 1 (vla-step-text 1))))", _
                         "(if (vlatraceon) (then (vlatracestep 2 (vla-step-text 2))))")
        vbaOut = TryTranspile("trace: the calls survive to VBA under runtime names", vla)
        If Len(vbaOut) > 0 Then
            Report "trace: the calls survive to VBA under runtime names", _
                   InStr(1, vbaOut, "vlatraceon", vbTextCompare) > 0 And _
                   InStr(1, vbaOut, "vlatracestep", vbTextCompare) > 0, _
                   Left$(Norm(vbaOut), 200)
        End If
    Else
        Report "trace: the calls survive to VBA under runtime names", False, _
               "skipped - the translation above failed"
    End If

    VlaTrace True
    VlaTraceStep 1, "Alpha. [line 1]"
    VlaTraceStep 2, "Beta. [line 2]"
    Dim rep As String
    rep = VlaTraceReport()
    Report "trace: steps log in execution order with their text", _
           InStr(1, rep, "Alpha. [line 1]", vbTextCompare) > 0 And _
           InStr(1, rep, "Beta. [line 2]", vbTextCompare) > InStr(1, rep, "Alpha. [line 1]", vbTextCompare), _
           rep
    VlaTrace True
    Report "trace: re-enabling clears the log", _
           InStr(1, VlaTraceReport(), "Alpha", vbTextCompare) = 0, VlaTraceReport()
    VlaTrace False
    VlaTraceStep 3, "Gamma. [line 3]"
    Report "trace: off means silent", _
           InStr(1, VlaTraceReport(), "Gamma", vbTextCompare) = 0, VlaTraceReport()
    CheckV "trace: the guard reads the flag", VlaTraceOn(), False
End Sub

' ---------------------------------------------------------------------
'  L2: VlaTry's pure half. VlaTryBuild (VLA_DevRig - dev-only calling
'  dev-only, both absent from the add-in by the same build list) is
'  plain string work, so the self-test holds its whole contract:
'  splitting, hoisting, string/comment paren-blindness, the refusal.
'  The inject/run/delete half mutates the live project (module-level
'  state loss included), so it belongs to the verification loop's
'  Immediate-window smoke - the trace's precedent (S5).
'
'  L2.1 (maiden-run incident, owner-caught): the defmacro pin was
'  written in the WRONG DIALECT - "(defmacro twice (x) ...)" where
'  this engine's shape is "(defmacro (twice x) ...)", name and
'  params in one signature list (the prelude and the L1 pins both
'  say so) - and the raise it drew (NthList, "expected a list at
'  position 2") escaped a BARE VlaTranspile call and killed the
'  whole suite mid-run. Two fixes: the pin now speaks the engine's
'  dialect, and every transpile in this section rides TryTranspile,
'  which reports FAIL under the pin's own name instead of crashing -
'  a self-test must never be able to die on the code it judges.
'  (TestRuntimeTrace's bare VlaTranspile calls carried the same
'  latent hazard on engine-generated input; hardened at the owner's
'  word as L2.2 - see that section and TryEnglish below.)
' ---------------------------------------------------------------------
Private Sub TestVlaTry()
    Dim b As String, t As String
    b = VlaTryBuild("(debug-print 1)")
    CheckFrags "vlatry: a bare statement wraps in the scratch sub", b, _
               Array("(sub vla-scratch ()", "(debug-print 1)")
    t = TryTranspile("vlatry: the scratch transpiles runnable", b)
    If Len(t) > 0 Then
        CheckFrags "vlatry: the scratch transpiles runnable", t, _
                   Array("Sub vla_scratch", "Debug.Print 1", "End Sub")
    End If

    b = VlaTryBuild("(defmacro (twice x) (begin x x))" & vbCrLf & "(twice (debug-print 7))")
    t = TryTranspile("vlatry: a defmacro hoists top-level and its body call expands", b)
    If Len(t) > 0 Then
        Dim p1 As Long
        p1 = InStr(1, t, "Debug.Print 7", vbTextCompare)
        Report "vlatry: a defmacro hoists top-level and its body call expands", _
               p1 > 0 And InStr(p1 + 1, t, "Debug.Print 7", vbTextCompare) > p1, _
               Left$(Norm(t), 160)
    End If

    b = VlaTryBuild("(debug-print 9) (sub helper () (debug-print 2)) (helper)")
    Dim ps As Long
    ps = InStr(1, b, "(sub vla-scratch", vbTextCompare)
    Report "vlatry: definitions hoist above the scratch; calls ride the body in order", _
           ps > 0 And InStr(1, b, "(sub helper", vbTextCompare) > 0 And _
           InStr(1, b, "(sub helper", vbTextCompare) < ps And _
           InStr(ps, b, "(debug-print 9)", vbTextCompare) > ps And _
           InStr(ps, b, "(helper)", vbTextCompare) > InStr(ps, b, "(debug-print 9)", vbTextCompare), _
           Left$(Norm(b), 160)

    b = VlaTryBuild("(debug-print "");smile("")" & vbCrLf & _
                    "; a comment (with parens" & vbCrLf & "(debug-print 3)")
    Report "vlatry: strings and comments never miscount parens", _
           InStr(1, b, "(debug-print "");smile("")", vbTextCompare) > 0 And _
           InStr(1, b, "(debug-print 3)", vbTextCompare) > 0 And _
           InStr(1, b, "a comment", vbTextCompare) = 0, _
           Left$(Norm(b), 160)

    b = VlaTryBuild("(dim g Long) (set! g 3)")
    t = TryTranspile("vlatry: dim stays a procedure local in the scratch body", b)
    If Len(t) > 0 Then
        Dim pm As Long
        pm = InStr(1, t, "Sub vla_scratch", vbTextCompare)
        Report "vlatry: dim stays a procedure local in the scratch body", _
               pm > 0 And InStr(pm, t, "Dim g As Long", vbTextCompare) > pm And _
               InStr(pm, t, "g = 3", vbTextCompare) > pm, _
               Left$(Norm(t), 160)
    End If

    t = TryTranspile("vlatry: empty input still builds a runnable scratch", VlaTryBuild(""))
    If Len(t) > 0 Then
        CheckFrags "vlatry: empty input still builds a runnable scratch", t, _
                   Array("Sub vla_scratch")
    End If

    ' Rule 9's caveat honored: this refusal is VlaTryBuild's own
    ' (the raise exists by construction in the splitter), so the pin
    ' proves the shipped message, not an assumed strictness.
    Dim errd As String
    On Error Resume Next
    b = VlaTryBuild("(debug-print 1")
    errd = Err.Description
    On Error GoTo 0
    Report "vlatry: an unclosed '(' refuses at build", _
           InStr(1, errd, "unclosed", vbTextCompare) > 0, "got: " & errd
End Sub

' ---------------------------------------------------------------------
'  L3 goldens: the prelude growth - binding forms and control
'  spellings, templates only. Same conventions as the C1 prelude
'  pins (AssertVla fragments) plus one expansion golden through
'  L1's instrument; every manual engine call rides a guard, the
'  L2.1 rule - a section must not be able to crash the suite.
' ---------------------------------------------------------------------
Private Sub TestPreludeGrowth()
    AssertVla "macro: let-one binds, assigns, and runs the body", _
              "(sub t () (let-one x 5 (debug-print x)))", _
              "x As Variant", "x = 5", "Debug.Print x"
    AssertVla "macro: let-one splices a multi-statement body", _
              "(sub t () (let-one x 5 (debug-print x) (set! x 6)))", _
              "x = 6"
    AssertVla "macro: let-one-as carries the type", _
              "(sub t () (let-one-as n Long 5 (debug-print n)))", _
              "n As Long", "n = 5"
    AssertVla "macro: if-not is unless's twin spelling", _
              "(sub t () (if-not (> x 1) (debug-print x)))", _
              "If (Not (x > 1)) Then"
    AssertVla "macro: while-not compiles to Do While Not", _
              "(sub t () (while-not (= x 3) (inc! x)))", _
              "Do While (Not (x = 3))", "Loop"

    ' comment: the positive and the negative in one pin - the
    ' neighbor statement survives, the disabled code truly vanishes.
    Dim t As String
    t = TryTranspile("macro: comment vanishes, neighbors survive", _
                     "(sub t () (comment (set! x 999)) (set! y 1))")
    If Len(t) > 0 Then
        Report "macro: comment vanishes, neighbors survive", _
               InStr(1, t, "y = 1", vbTextCompare) > 0 And _
               InStr(1, t, "999", vbTextCompare) = 0, Left$(Norm(t), 160)
    End If
    t = TryTranspile("macro: empty comment is legal", "(sub t () (comment))")
    If Len(t) > 0 Then
        CheckFrags "macro: empty comment is legal", t, Array("Sub t", "End Sub")
    End If

    ' The expansion golden through L1's instrument: the roadmap's
    ' let-one expansion, verbatim (Norm rides over the printer's
    ' line breaks - only the datums are the contract).
    Dim fired As Long
    Dim d As String
    On Error Resume Next
    t = VlaExpandText("(let-one x 5 (debug-print x))", True, fired)
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    If Len(d) > 0 Then
        Report "expand: let-one golden", False, "expand error: " & d
    Else
        CheckFrags "expand: let-one golden", t, _
                   Array("(begin (dim x) (set! x 5) (debug-print x))")
    End If
End Sub

' ---------------------------------------------------------------------
'  L5: the guide's quoted refusals, kept honest. VLA_GUIDE's new
'  template-limits section quotes two loader messages verbatim;
'  rule 9 says a documented refusal must be proven to exist, so
'  these pins lock doc and code together. The second is doubly a
'  SENTINEL: the keyword-delimited split ("& body :rescue & handler")
'  is RECORDED in the guide as the approved future shape but not
'  implemented - it must refuse today, and whoever implements it
'  will fail this pin and be sent to update the guide's paragraph.
' ---------------------------------------------------------------------
Private Sub TestTemplateLimits()
    Dim d As String
    Dim junk As String
    On Error Resume Next
    junk = VlaTranspile("(defmacro (bad-mac x &) (begin x))")
    d = Err.Description
    On Error GoTo 0
    Report "limits: '&' without a rest name refuses as the guide quotes", _
           InStr(1, d, "'&' must be followed by a rest parameter", vbTextCompare) > 0, _
           "got: " & d
    d = ""
    On Error Resume Next
    junk = VlaTranspile("(defmacro (try-mac & body :rescue & handler) (begin body))")
    d = Err.Description
    On Error GoTo 0
    Report "limits: the guide's future shape still refuses today (sentinel)", _
           InStr(1, d, "rest parameter must be last", vbTextCompare) > 0, _
           "got: " & d
End Sub

' IN.1: the head table is new machinery with no caller yet (VLA.bas's
' Select Case arms are still authoritative), so these are sanity
' checks on the catalog itself, not behavior pins - the count is a
' printed baseline (LESSONS.md: "baselines must print themselves"),
' not a magic number to defend.
Private Sub TestHeadTable()
    Report "headtable: builds without a duplicate-key error", _
           VlaHeadTableCount() > 0, "count: " & VlaHeadTableCount()

    Report "headtable: row count is the printed baseline (65)", _
           VlaHeadTableCount() = 65, "got: " & VlaHeadTableCount()

    Dim ifRow As Collection
    Set ifRow = VlaHeadTableRow("if")
    Report "headtable: 'if' resolves to EmitIf, formula-supported", _
           Not ifRow Is Nothing, "if: not found"
    If Not ifRow Is Nothing Then
        Report "headtable: 'if' row detail", _
               CStr(ifRow.Item(HT_VBA)) = "EmitIf" And CStr(ifRow.Item(HT_FORMULA)) = "yes", _
               "vba=[" & CStr(ifRow.Item(HT_VBA)) & "] formula=[" & CStr(ifRow.Item(HT_FORMULA)) & "]"
        Report "headtable: the interpreter column is empty for every row (IN.2 not built)", _
               Len(CStr(ifRow.Item(HT_INTERP))) = 0, "if row's interpreter cell: [" & CStr(ifRow.Item(HT_INTERP)) & "]"
    End If

    Dim rawRow As Collection
    Set rawRow = VlaHeadTableRow("raw")
    Report "headtable: 'raw' is flagged export-only", _
           Not rawRow Is Nothing, "raw: not found"
    If Not rawRow Is Nothing Then
        Report "headtable: 'raw' export-only flag", _
               CBool(rawRow.Item(HT_EXPORTONLY)) = True, "got: " & CBool(rawRow.Item(HT_EXPORTONLY))
    End If

    Report "headtable: lookup folds case invariantly (SET! finds set!)", _
           Not VlaHeadTableRow("SET!") Is Nothing, "SET!: not found"

    Report "headtable: an unknown symbol returns Nothing, not an error", _
           VlaHeadTableRow("not-a-real-form") Is Nothing, "expected Nothing"

    Report "headtable: the summary line names the row count", _
           InStr(1, VlaHeadTableSummary(), "65 forms", vbBinaryCompare) > 0, _
           "got: " & VlaHeadTableSummary()

    ' LX.4: VlaHeadTableRow's own alias scan (the catalog-lookup path,
    ' not VlaHeadTableAliasMap - the emitter's own O(1) path is proven
    ' separately by TestHeadAlias below).
    Dim aliasRow As Collection
    Set aliasRow = VlaHeadTableRow("fijar!")
    Report "headtable: alias 'fijar!' resolves to the 'set!' row", _
           (Not aliasRow Is Nothing), "fijar!: not found via VlaHeadTableRow"
    If Not aliasRow Is Nothing Then
        CheckV "headtable: 'fijar!' row's own symbol reads 'set!'", _
               CStr(aliasRow.Item(HT_SYMBOL)), "set!"
    End If
    Report "headtable: alias lookup also folds case (FIJAR! finds set!)", _
           Not VlaHeadTableRow("FIJAR!") Is Nothing, "FIJAR!: not found"
End Sub

' IN5.0: pins the export-only column IN.5 adjudicated - the exact set
' of symbols, not just "some are true," so a future accidental edit to
' BuildRows fails loudly instead of drifting unnoticed. Also proves the
' negative for the nine corrected rows (sub/function/type/enum/public/
' private/include/at-line/doc) by name, not just by the total count -
' two wrong rows swapping True for False on DIFFERENT symbols could
' still sum to the right total. The finding this item made: nine forms
' were marked export-only by rote in an earlier pass, not by reason
' checked against the actual emitter code, and turned out not to be.
Private Sub TestHeadTableExportOnly()
    Dim rows As Collection
    Set rows = VlaHeadTableRows()
    Dim gotList As String
    Dim row As Variant
    For Each row In rows
        If CBool(row.Item(HT_EXPORTONLY)) Then
            If Len(gotList) > 0 Then gotList = gotList & ","
            gotList = gotList & CStr(row.Item(HT_SYMBOL))
        End If
    Next
    CheckV "in5: the adjudicated export-only set is exactly raw, deflambda, lambda", _
           gotList, "raw,deflambda,lambda"

    Dim corrected As Variant
    corrected = Array("sub", "function", "type", "enum", "public", "private", "include", "at-line", "doc")
    Dim i As Long
    Dim allFalse As Boolean
    allFalse = True
    Dim detail As String
    For i = LBound(corrected) To UBound(corrected)
        Dim r As Collection
        Set r = VlaHeadTableRow(CStr(corrected(i)))
        If r Is Nothing Then
            allFalse = False
            detail = detail & CStr(corrected(i)) & "=not-found "
        ElseIf CBool(r.Item(HT_EXPORTONLY)) Then
            allFalse = False
            detail = detail & CStr(corrected(i)) & "=still-True "
        End If
    Next
    Report "in5: the nine corrected rows (sub/function/type/enum/public/private/include/at-line/doc) are all False", _
           allFalse, IIf(Len(detail) > 0, detail, "all correct")
End Sub

' LX.4: the actual chokepoint, proven at the emitter, not just the
' catalog - VlaTranspile of an aliased head must produce BYTE-IDENTICAL
' VBA to its canonical spelling, since ResolveHeadAlias runs before the
' Select Case that does the real emitting, not a second emitter beside
' it. Three separate Emit* dispatchers exercised: EmitStmt (set!/
' fijar!, debug-print/depurar) and EmitStmt's own EmitIf (if/si).
Private Sub TestHeadAlias()
    Dim canon As String, aliased As String
    canon = VlaTranspile("(sub t () (set! x 5) (if (> x 1) (then (debug-print x))))")
    aliased = VlaTranspile("(sub t () (fijar! x 5) (si (> x 1) (then (depurar x))))")
    CheckV "lx4: an aliased program emits byte-identical VBA to its canonical spelling", _
           aliased, canon
End Sub

' LX.6: non-ASCII identifiers transliterate rather than reaching
' generated VBA source untouched. Built via ChrW$ code points, never a
' literal accented character in this file's own text - the same
' encoding-safety reason TransliterateChar's own table uses hex case
' labels instead (LESSONS.md's string-literal hazard: the .bas
' import/export round trip is not guaranteed UTF-8).
Private Sub TestNonAsciiIdentifiers()
    Dim accented As String
    accented = "caf" & ChrW$(&HE9)                ' "cafe" + U+00E9 (e-acute) = "cafe with an accent"
    Dim t As String
    t = TryTranspile("lx6: an accented sub name transliterates", _
                     "(sub " & accented & " () (debug-print 1))")
    If Len(t) > 0 Then
        CheckFrags "lx6: an accented sub name transliterates", t, Array("Sub cafe(")
    End If

    Dim unmappable As String
    unmappable = ChrW$(&H4E2D) & ChrW$(&H6587)    ' two CJK characters - no table entry, on purpose
    Dim d As String
    On Error Resume Next
    VlaTranspile "(sub " & unmappable & " () (debug-print 1))"
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    Report "lx6: an entirely unmappable name refuses in words, not silently", _
           InStr(1, d, "no plain letters", vbTextCompare) > 0, _
           "expected the LX.6 refusal, got: [" & d & "]"
End Sub

' F.1: the published scoreboard, pinned. Zero was reached by converting
' every rule template that reached the object model directly - both
' "(. " forms and bare-dotted application.worksheetfunction.* calls -
' into a call on a named macro; this assertion is what keeps that
' number from drifting back up unnoticed the next time someone adds a
' rule under deadline. Two known exceptions (Collection.Add in the
' "Append to <list>" core-language proofs) are excluded BY NAME in
' VlaDotCount itself, not absorbed into a bigger tolerance here.
Private Sub TestDotCount()
    ' AS.6 (owner-caught class, same shape as L2.1/L2.2): this call
    ' was bare - VlaDotCount's FindDevFile raises loudly when
    ' english.vla is missing, and an unguarded raise here would
    ' escape TestDotCount and kill VlaSelfTest mid-run, the exact
    ' hazard TryTranspile/TryEnglish exist to prevent elsewhere. A
    ' self-test must never be able to die on the corpus it judges;
    ' TestCorpusFamily (dispatched first) already names a missing
    ' file explicitly, but this guard holds regardless of dispatch
    ' order, on its own.
    Dim n As Long
    Dim d As String
    On Error Resume Next
    n = VlaDotCount()
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    If Len(d) > 0 Then
        Report "f1: the phrasebook's dot count is zero", False, "VlaDotCount error: " & d
        Exit Sub
    End If
    Report "f1: the phrasebook's dot count is zero", _
           n = 0, "got: " & n & " (see VlaDotCount's own header for what counts)"
End Sub

' IN0.5: a correctness pin for the walking-skeleton interpreter
' (VLA_Interpreter.bas) - dev-rig only, no error handling, no parity
' with the emitter, but its ten forms still get one deterministic
' proof rather than eyeballing VlaInterpretDemo's Immediate-window
' output. Hand-written VLA, not English (VLA_Interpreter's own demo
' covers the English-in half); exercises all ten forms this pass
' supports (dim, set!, if, for, while, begin, debug-print, +, -, >)
' in one trace: dim x=0, set! x=10, set! x=(10-3)=7, if(7>5) true so
' set! x=(7+1)=8, for i=1..3 accumulating x=9,11,14, while(x>12)
' subtracting 1 each pass: 13, then 12 (stops, 12 is not > 12).
Private Sub TestInterpreterSkeleton()
    Dim frame As Object
    Set frame = VLA_Interpreter.VlaInterpret( _
        "(begin" & _
        "  (dim x Double)" & _
        "  (set! x 10)" & _
        "  (set! x (- x 3))" & _
        "  (if (> x 5) (then (set! x (+ x 1))) (else (set! x 0)))" & _
        "  (for (i 1 3) (set! x (+ x i)))" & _
        "  (while (> x 12) (set! x (- x 1)))" & _
        "  (debug-print x))")
    CheckV "in0.5: the walking skeleton's ten forms (dim/set!/if/for/while/begin/debug-print/+/-/>) land on the right value", _
           VLA_Runtime.VlaDictGet(frame, "x"), 12
End Sub

' IN3_5.0: proves the effect log's CONTENT, not just that VlaInterpret
' still returns the right frame - a small, hand-built program (three
' debug-prints: a literal, an arithmetic result, a string) whose
' expected log text is fixed and does not drift as the walking
' skeleton's own demo program changes, the same "small, hand-built
' sample" discipline F.9's TestSectionMarkers already uses for the
' emitter-side goldens.
Private Sub TestInterpreterEffectLog()
    VLA_Interpreter.VlaInterpret _
        "(begin (debug-print 1) (debug-print (+ 2 3)) (debug-print ""hi""))"
    CheckV "in3.5: effect log records debug-print effects in order", _
           VLA_Interpreter.VlaInterpreterEffectLog(), _
           "debug-print: 1" & vbCrLf & "debug-print: 5" & vbCrLf & "debug-print: hi" & vbCrLf
End Sub

' IN2.0: the bounded-builtin tier (TryEvalBuiltin) and the
' VLA_Runtime-helper fallback tier (TryRuntimeHelper), pure - see this
' file's own IN2.0 history note for why msgbox/inputbox are absent on
' purpose. obj-set! to a plain variable is pinned identical to set!,
' proving ExecSet's own claim that the two are the SAME operation for
' this interpreter's dict-backed frame, not merely similar.
Private Sub TestInterpreterDynamicHead()
    CheckV "in2: len via the bounded builtin tier", _
           VLA_Interpreter.VlaEvalExpression("(len ""hello"")"), 5
    CheckV "in2: trim via the bounded builtin tier", _
           VLA_Interpreter.VlaEvalExpression("(trim ""  hi  "")"), "hi"
    CheckV "in2: ucase via the bounded builtin tier", _
           VLA_Interpreter.VlaEvalExpression("(ucase ""ok"")"), "OK"
    CheckV "in2: left via the bounded builtin tier", _
           VLA_Interpreter.VlaEvalExpression("(left ""hello"" 3)"), "hel"
    CheckV "in2: round (2-arg) via the bounded builtin tier", _
           VLA_Interpreter.VlaEvalExpression("(round 3.14159 2)"), 3.14
    CheckV "in2: instr via the bounded builtin tier", _
           VLA_Interpreter.VlaEvalExpression("(instr ""hello world"" ""world"")"), 7

    ' VLA_Runtime.VlaColor, reached through Application.Run against the
    ' module by name (tier 4) - not a hand-written branch for this
    ' specific helper. VlaColor's own contract (VLA_Runtime.bas): a
    ' "#RRGGBB" string, hex-decoded through RGB() - #FF0000 is red.
    CheckV "in2: VLA_Runtime helper dispatch (vlacolor)", _
           VLA_Interpreter.VlaEvalExpression("(vlacolor ""#FF0000"")"), RGB(255, 0, 0)

    ' obj-set! to a plain variable, aliased straight to ExecSet.
    Dim frame As Object
    Set frame = VLA_Interpreter.VlaInterpret("(begin (dim x Double) (obj-set! x 42))")
    CheckV "in2: obj-set! to a plain variable acts exactly like set!", _
           VLA_Runtime.VlaDictGet(frame, "x"), 42
End Sub

' SEC.4: VLA_Interpreter.NeutralizeFormulaInjection, pure - DynamicSet's
' own "value" case only ever fires against a real Range-like object (a
' pure test has no way to construct one), so this pins the guard's own
' decision logic directly instead, made Public for exactly this reason.
' Covers: all four OWASP-named trigger characters (=/+/-/@), each
' neutralized with a leading apostrophe; an ordinary string with none of
' them, untouched; an empty string, untouched (no Left$ crash on a
' zero-length string); a trigger character appearing INSIDE a string but
' not at its own start, untouched (position-specific, not a substring
' search); every non-string VarType (a real number, a Boolean) left
' completely alone, confirming the guard never touches genuine Excel
' numeric/boolean types; and an object input passed through unchanged
' via Set, not a bare assignment (this Sub's own regression pin for the
' exact "bare Variant assignment invokes an object's own default member"
' trap this function's own header already documents guarding against).
Private Sub TestInterpreterCsvInjectionGuard()
    CheckV "sec.4: an ordinary string is left untouched", _
           VLA_Interpreter.NeutralizeFormulaInjection("hello"), "hello"
    CheckV "sec.4: a leading = is neutralized with a leading apostrophe", _
           VLA_Interpreter.NeutralizeFormulaInjection("=SUM(A1:A2)"), "'=SUM(A1:A2)"
    CheckV "sec.4: a leading + is neutralized", _
           VLA_Interpreter.NeutralizeFormulaInjection("+1234"), "'+1234"
    CheckV "sec.4: a leading - is neutralized", _
           VLA_Interpreter.NeutralizeFormulaInjection("-5 days"), "'-5 days"
    CheckV "sec.4: a leading @ is neutralized", _
           VLA_Interpreter.NeutralizeFormulaInjection("@SUM"), "'@SUM"
    CheckV "sec.4: an empty string is left untouched, not crashed on", _
           VLA_Interpreter.NeutralizeFormulaInjection(""), ""
    CheckV "sec.4: a trigger character NOT at the string's own start is left untouched", _
           VLA_Interpreter.NeutralizeFormulaInjection("total=5"), "total=5"
    CheckV "sec.4: a real number is left completely untouched", _
           VLA_Interpreter.NeutralizeFormulaInjection(CDbl(5)), 5
    CheckV "sec.4: a real Boolean is left completely untouched", _
           VLA_Interpreter.NeutralizeFormulaInjection(True), True

    Dim probe As New Collection
    probe.Add "marker"
    Dim resultObj As Variant
    Set resultObj = VLA_Interpreter.NeutralizeFormulaInjection(probe)
    Report "sec.4: an object input passes through unchanged (Set, not a bare assignment)", _
           resultObj Is probe, "expected the same object reference back"
End Sub

' IN2.6: Excel's named enum constants, purely - one per enum family this
' corpus's own measured set spans (XlSortOrder/XlYesNoGuess/XlDirection/
' XlPasteType/XlLineStyle/XlFixedFormatType/XlCalculation - XlHAlign
' covered by xlascending's sibling xlcenter below), not all 13 - a
' representative sample plus the regression case matters more than
' exhaustively re-listing VLA_Interpreter.bas's own table here. The
' regression case is the point as much as the constants are: a real
' variable whose value happens to equal what a constant WOULD have
' resolved to must still come from the frame, not the table - proving
' ResolveExcelConstant is consulted before, and does not shadow, a normal
' variable lookup for any name outside its own 13.
Private Sub TestInterpreterExcelConstants()
    CheckV "in2.6: xlascending resolves to 1 (XlSortOrder)", _
           VLA_Interpreter.VlaEvalExpression("xlascending"), 1
    CheckV "in2.6: xldescending resolves to 2 (XlSortOrder)", _
           VLA_Interpreter.VlaEvalExpression("xldescending"), 2
    CheckV "in2.6: xlyes resolves to 1 (XlYesNoGuess)", _
           VLA_Interpreter.VlaEvalExpression("xlyes"), 1
    CheckV "in2.6: xlcenter resolves to -4108 (XlHAlign)", _
           VLA_Interpreter.VlaEvalExpression("xlcenter"), -4108
    CheckV "in2.6: xlup resolves to -4162 (XlDirection)", _
           VLA_Interpreter.VlaEvalExpression("xlup"), -4162
    CheckV "in2.6: xlpastevalues resolves to -4163 (XlPasteType)", _
           VLA_Interpreter.VlaEvalExpression("xlpastevalues"), -4163
    CheckV "in2.6: xltypepdf resolves to 0 (XlFixedFormatType)", _
           VLA_Interpreter.VlaEvalExpression("xltypepdf"), 0
    CheckV "in2.6: xlcalculationmanual resolves to -4135 (XlCalculation)", _
           VLA_Interpreter.VlaEvalExpression("xlcalculationmanual"), -4135
    CheckV "in2.6: resolution is case-insensitive, like every other head symbol", _
           VLA_Interpreter.VlaEvalExpression("XlAscending"), 1

    ' Regression: a bound variable takes priority over nothing, because a
    ' name outside the 13-constant table was never touched by this pass -
    ' still just a frame lookup.
    Dim frame As Object
    Set frame = VLA_Interpreter.VlaInterpret("(begin (dim my-order Double) (set! my-order 99))")
    CheckV "in2.6: an ordinary variable is unaffected by the constant table", _
           VLA_Runtime.VlaDictGet(frame, "my-order"), 99
End Sub

' IN.3: VBA's own intrinsic constants (VLA_Interpreter.bas's own
' ResolveVbConstant note has the full reasoning - found live, "Make cell
' C4 yellow." raised "there is nothing stored at key 'vbyellow'" the
' first time instructions.txt reached it). Same shape as
' TestInterpreterExcelConstants above, a different source, deliberately
' not folded into the same pin.
Private Sub TestInterpreterVbConstants()
    CheckV "in.3: vbred resolves to 255", _
           VLA_Interpreter.VlaEvalExpression("vbred"), 255
    CheckV "in.3: vbyellow resolves to 65535", _
           VLA_Interpreter.VlaEvalExpression("vbyellow"), 65535
    CheckV "in.3: vbtextcompare resolves to 1", _
           VLA_Interpreter.VlaEvalExpression("vbtextcompare"), 1
    CheckV "in.3: resolution is case-insensitive, like every other head symbol", _
           VLA_Interpreter.VlaEvalExpression("VbYellow"), 65535

    ' Regression: an ordinary variable is unaffected by the constant table.
    Dim frame As Object
    Set frame = VLA_Interpreter.VlaInterpret("(begin (dim my-flag Double) (set! my-flag 7))")
    CheckV "in.3: an ordinary variable is unaffected by the vb-constant table", _
           VLA_Runtime.VlaDictGet(frame, "my-flag"), 7
End Sub

' IN2.7: EvalOpChain's 18 operators, purely - AS.8's own scan found only
' +/-/> had interpreter cases; this pins the rest. 'is' is deliberately
' absent (needs real object identity - a live-host concern, the same
' "not exhaustive" precedent msgbox/inputbox and IN2.0's object-dispatch
' pins already set for this suite - see VLA_Tests_Host.bas for a real
' Range comparison instead). The 3-operand '<' case is the point as much
' as any single operator: (< 3 2 1) is True, NOT False - EvalOpChain is a
' literal left-to-right FOLD, matching what EmitChain's own generated
' VBA text does ("((3 < 2) < 1)"), not mathematical range-chaining. VBA
' folds (3 < 2) to False, coerces False to 0, then evaluates (0 < 1),
' which is True - pinned exactly so a future reader trusts the fold
' model instead of "fixing" it into something EmitChain doesn't do.
Private Sub TestInterpreterOperators()
    CheckV "in2.7: * chains N-ary, like EmitChain's own text", _
           VLA_Interpreter.VlaEvalExpression("(* 2 3 4)"), 24
    CheckV "in2.7: & concatenates, not forced through CDbl", _
           VLA_Interpreter.VlaEvalExpression("(& ""a"" ""b"" ""c"")"), "abc"
    CheckV "in2.7: / divides", _
           VLA_Interpreter.VlaEvalExpression("(/ 10 4)"), 2.5
    CheckV "in2.7: backslash is VBA's integer division", _
           VLA_Interpreter.VlaEvalExpression("(\ 10 3)"), 3
    CheckV "in2.7: = compares equal", _
           VLA_Interpreter.VlaEvalExpression("(= 1 1)"), True
    CheckV "in2.7: <> compares unequal", _
           VLA_Interpreter.VlaEvalExpression("(<> 1 2)"), True
    CheckV "in2.7: < now supports strings, not just CDbl-forced numbers", _
           VLA_Interpreter.VlaEvalExpression("(< ""abc"" ""abd"")"), True
    CheckV "in2.7: > folds N-ary, closing its own pre-existing 2-operand-only gap", _
           VLA_Interpreter.VlaEvalExpression("(> 5 3 1)"), False
    CheckV "in2.7: < folds left-to-right, not mathematical range-chaining (see this Sub's own header note)", _
           VLA_Interpreter.VlaEvalExpression("(< 3 2 1)"), True
    CheckV "in2.7: <= compares", _
           VLA_Interpreter.VlaEvalExpression("(<= 5 5)"), True
    CheckV "in2.7: >= compares", _
           VLA_Interpreter.VlaEvalExpression("(>= 3 5)"), False
    CheckV "in2.7: and", _
           VLA_Interpreter.VlaEvalExpression("(and true false)"), False
    CheckV "in2.7: or", _
           VLA_Interpreter.VlaEvalExpression("(or false true)"), True
    CheckV "in2.7: xor", _
           VLA_Interpreter.VlaEvalExpression("(xor true true)"), False
    CheckV "in2.7: imp", _
           VLA_Interpreter.VlaEvalExpression("(imp false true)"), True
    CheckV "in2.7: eqv", _
           VLA_Interpreter.VlaEvalExpression("(eqv true false)"), False
    CheckV "in2.7: mod", _
           VLA_Interpreter.VlaEvalExpression("(mod 10 3)"), 1
    CheckV "in2.7: like, VBA pattern matching", _
           VLA_Interpreter.VlaEvalExpression("(like ""hello"" ""h*"")"), True
    CheckV "in2.7: not", _
           VLA_Interpreter.VlaEvalExpression("(not true)"), False

    ' Regression: existing +/- chains (EvalChain, untouched by this pass)
    ' still work exactly as before.
    CheckV "in2.7: + still chains via the original EvalChain, unchanged", _
           VLA_Interpreter.VlaEvalExpression("(+ 1 2 3)"), 6
    CheckV "in2.7: unary - still negates, unchanged", _
           VLA_Interpreter.VlaEvalExpression("(- 5)"), -5
End Sub

' ---------------------------------------------------------------------
'  REPLEVAL.0: EvalExpr's own new "quote" case (VLA_Interpreter.bas) -
'  found live, via eval's own first real use, when eval "(quote (1 2
'  3))" raised "'1' is not a form ..." instead of yielding an array.
'  TestArrayPrimitive's own precedent followed exactly: an array result
'  is checked structurally (IsArray/LBound/UBound/elements), never
'  CStr'd - CStr on a VBA array raises Type mismatch, the same hazard
'  that test's own header note already names.
' ---------------------------------------------------------------------
Private Sub TestInterpreterQuote()
    Dim arr As Variant
    arr = VLA_Interpreter.VlaEvalExpression("(quote (1 2 3))")
    Dim ok As Boolean, detail As String
    If IsArray(arr) Then
        ok = (LBound(arr) = 0) And (UBound(arr) = 2) _
             And (arr(0) = 1) And (arr(1) = 2) And (arr(2) = 3)
        detail = "array, length " & (UBound(arr) - LBound(arr) + 1)
    Else
        detail = "not an array: " & CStr(arr)
    End If
    Report "repl-eval: (quote (1 2 3)) evaluates to a real array of numbers, not a call", ok, detail

    ' A quoted bare symbol becomes a STRING at runtime - QuoteDatum's
    ' own convention on the compile side (VLA.bas), mirrored exactly:
    ' quoting a symbol never invents a distinct "symbol" runtime value.
    CheckV "repl-eval: (quote foo) yields the string ""foo"", matching QuoteDatum's own convention", _
           VLA_Interpreter.VlaEvalExpression("(quote foo)"), "foo"

    ' A string literal inside quote stays exactly that string - never
    ' re-wrapped, never treated as a symbol.
    CheckV "repl-eval: a string literal inside quote is unwrapped to its own content", _
           VLA_Interpreter.VlaEvalExpression("(quote ""hi"")"), "hi"

    ' Mixed data - number, string literal, bare symbol - each converted
    ' by its OWN rule, recursively, never by evaluating any element as
    ' code (a symbol here, "three", would raise if EvalExpr's dynamic
    ' dispatch ever touched it as a call head).
    Dim mixed As Variant
    mixed = VLA_Interpreter.VlaEvalExpression("(quote (1 ""two"" three))")
    Dim mixedOk As Boolean, mixedDetail As String
    If IsArray(mixed) Then
        mixedOk = (mixed(0) = 1) And (mixed(1) = "two") And (mixed(2) = "three")
        mixedDetail = "elements: " & CStr(mixed(0)) & " / " & CStr(mixed(1)) & " / " & CStr(mixed(2))
    Else
        mixedDetail = "not an array: " & CStr(mixed)
    End If
    Report "repl-eval: quote converts each element by its own rule - number, string, symbol-as-string", mixedOk, mixedDetail

    ' The empty list quotes to a real zero-length array, not an error -
    ' TestArrayPrimitive's own (array) pin, same shape, same reason.
    ' NOT named "empty" - that collides with VBA's own Empty keyword
    ' (the uninitialized-Variant literal), a real compile-time syntax
    ' error caught live, not a style choice.
    Dim emptyArr As Variant
    emptyArr = VLA_Interpreter.VlaEvalExpression("(quote ())")
    Report "repl-eval: (quote ()) is a real zero-length array, not an error", _
           IsArray(emptyArr) And (UBound(emptyArr) < LBound(emptyArr)), "got: " & TypeName(emptyArr)

    ' Arity: exactly one datum, same message EmitQuote's own compile-side
    ' check raises (VLA.bas) - checked directly against the interpreter,
    ' not assumed to match just because the wording was copied.
    Dim d As String
    On Error Resume Next
    Err.Clear
    VLA_Interpreter.VlaEvalExpression("(quote a b)")
    d = Err.Description
    On Error GoTo 0
    Report "repl-eval: quote refuses more than one datum", _
           InStr(1, d, "quote takes exactly one datum", vbTextCompare) > 0, "got: " & d
End Sub

' ---------------------------------------------------------------------
'  G6: `array` - list-valued slots' own new runtime primitive
'  (VLA.bas's EmitExpr, VLA_Interpreter.bas's EvalExpr). Pinned
'  directly against both backends here, not only through a translated
'  English sentence (TestG6, VLA_Tests_Grammar.bas, covers that
'  composition) - so a break in either backend's own case arm is
'  caught even if TestG6 didn't happen to exercise this exact shape,
'  matching AS.8's own backend-parity discipline. The interpreter side
'  can't go through CheckV: CStr on a VBA array raises "Type mismatch",
'  the same class of gotcha CStr on a possible Object already needed a
'  guard for elsewhere in this codebase (quote-if's own error message,
'  VLA.bas) - checked structurally instead, never CStr'd.
' ---------------------------------------------------------------------
Private Sub TestArrayPrimitive()
    ' Compile backend: a real VBA Array(...) literal.
    Dim vbaOut As String
    vbaOut = TryTranspile("g6-array: emits a VBA Array(...) literal", _
                          "(sub t () (debug-print (array ""a"" ""b"" ""c"")))")
    If Len(vbaOut) > 0 Then
        CheckFrags "g6-array: emits a VBA Array(...) literal", vbaOut, Array("Array(""a"", ""b"", ""c"")")
    End If

    ' Interpreter backend: a real Variant array, right length and
    ' elements, in order.
    Dim arrResult As Variant
    arrResult = VLA_Interpreter.VlaEvalExpression("(array ""a"" ""b"" ""c"")")
    Dim arrOk As Boolean, arrDetail As String
    If IsArray(arrResult) Then
        arrOk = (LBound(arrResult) = 0) And (UBound(arrResult) = 2) _
                And (arrResult(0) = "a") And (arrResult(1) = "b") And (arrResult(2) = "c")
        arrDetail = "array, length " & (UBound(arrResult) - LBound(arrResult) + 1)
    Else
        arrDetail = "not an array: " & CStr(arrResult)
    End If
    Report "g6-array: the interpreter evaluates (array ...) to a real Variant array, right length and order", _
           arrOk, arrDetail

    ' Zero items - (array) with nothing to list - is a real zero-length
    ' array, never an error. A field-list slot always matches at least
    ' one item (MatchRefToken, VLA_English.bas), but the primitive
    ' itself makes no such assumption, so this is checked directly
    ' rather than left implicit.
    Dim arrEmpty As Variant
    arrEmpty = VLA_Interpreter.VlaEvalExpression("(array)")
    Dim emptyOk As Boolean
    emptyOk = IsArray(arrEmpty) And (UBound(arrEmpty) < LBound(arrEmpty))
    Report "g6-array: (array) with no items is a real zero-length array, not an error", _
           emptyOk, "got: " & TypeName(arrEmpty)
End Sub

' L-INTERPOLATE: (interpolate tpl :key val ...) - both backends' own
' header notes (VLA.bas's EmitInterpolateCall, VLA_Interpreter.bas's
' EvalInterpolateCall) have the full design. Named "interpolate", not
' CL's "format" - this codebase's own format-as-currency/-percent/
' -date (english.vla) already own "format" for visual cell styling.
' Pure throughout - VlaEvalExpression needs no live workbook for any
' expression here, and TryTranspile only ever inspects generated text,
' never runs it.
Private Sub TestLInterpolate()
    ' Compile backend: an ordinary "&" chain, same shape the two
    ' DATALOG macros this primitive replaces already hand-spliced.
    Dim vbaOut As String
    vbaOut = TryTranspile("l-interpolate: emits a literal/hole ""&"" chain", _
                          "(sub t () (debug-print (interpolate ""a{x}b"" :x 5)))")
    If Len(vbaOut) > 0 Then
        CheckFrags "l-interpolate: emits a literal/hole ""&"" chain", vbaOut, _
                   Array("""a"" & 5 & ""b""")
    End If

    ' A hole reused twice in one template - the whole reason this
    ' primitive uses named holes rather than positional {0}/{1}
    ' (datalog-filter-place/datalog-chain-place each reuse a hole this
    ' way) - costs the caller nothing extra; the value is supplied once.
    vbaOut = TryTranspile("l-interpolate: a repeated hole reuses its one supplied value", _
                          "(sub t () (debug-print (interpolate ""{a}-{a}"" :a ""hi"")))")
    If Len(vbaOut) > 0 Then
        CheckFrags "l-interpolate: a repeated hole reuses its one supplied value", vbaOut, _
                   Array("""hi"" & ""-"" & ""hi""")
    End If

    ' "{{"/"}}" escape to literal braces - no hole, no keyword args
    ' needed at all.
    vbaOut = TryTranspile("l-interpolate: doubled braces escape to a literal brace", _
                          "(sub t () (debug-print (interpolate ""{{literal}}"")))")
    If Len(vbaOut) > 0 Then
        CheckFrags "l-interpolate: doubled braces escape to a literal brace", vbaOut, _
                   Array("""{literal}""")
    End If

    ' Interpreter backend: real substitution, real VBA "&" coercion on
    ' a non-string hole value (42, a Double) - no separate coercion
    ' rule, matching EvalOpChain's own "&" case exactly.
    CheckV "l-interpolate: the interpreter substitutes named holes", _
           VLA_Interpreter.VlaEvalExpression("(interpolate ""Hello {name}!"" :name ""World"")"), _
           "Hello World!"
    CheckV "l-interpolate: a non-string hole value coerces the same way ""&"" does", _
           VLA_Interpreter.VlaEvalExpression("(interpolate ""n={n}"" :n 42)"), _
           "n=42"
    CheckV "l-interpolate: a repeated hole reuses its one supplied value (interpreter)", _
           VLA_Interpreter.VlaEvalExpression("(interpolate ""{a}-{a}"" :a ""hi"")"), _
           "hi-hi"

    ' Refusals - every malformed-template/argument case gets its own
    ' named id (VLA_Messages.bas), never a bare type-mismatch.
    CheckInterpolateRefusal "l-interpolate: an unknown hole refuses by name", _
        "(interpolate ""{oops}"")", "'oops'"
    CheckInterpolateRefusal "l-interpolate: an unused keyword argument refuses by name", _
        "(interpolate ""no holes here"" :x 1)", "'x'"
    CheckInterpolateRefusal "l-interpolate: a computed (non-literal) template refuses", _
        "(interpolate somevar :y 1)", "computed value"
    CheckInterpolateRefusal "l-interpolate: an unclosed hole refuses", _
        "(interpolate ""abc{def"")", "unclosed hole"
    CheckInterpolateRefusal "l-interpolate: an empty hole refuses", _
        "(interpolate ""abc{}def"")", "empty hole"
    CheckInterpolateRefusal "l-interpolate: a positional argument where a keyword is expected refuses", _
        "(interpolate ""{x}"" 5)", "keyword pairs"
    CheckInterpolateRefusal "l-interpolate: a dangling keyword with no value refuses", _
        "(interpolate ""{x}"" :x)", "missing its value"
End Sub

' L-INTERPOLATE: exercises a refusal through BOTH backends from one
' bare interpolate expression - the interpreter via VlaEvalExpression,
' the compiler via VlaTranspile with the SAME expression wrapped in
' "(sub t () (debug-print ...))" (TestArrayPrimitive's own compile-
' backend shape). Both sides must refuse, and the SAME wantFrag must
' appear in both descriptions - AS.8 parity applied to a refusal, not
' just a success. Every case here is a STRUCTURAL check (a literal-
' template test, an arg-shape test) that never needs a bound variable -
' "somevar"/"{x}" et al never have to actually resolve, since both
' refusals fire by inspecting the raw form, before either backend would
' ever evaluate/emit the pieces inside it.
Private Sub CheckInterpolateRefusal(ByVal name As String, ByVal interpolateExpr As String, ByVal wantFrag As String)
    Dim d As String
    On Error Resume Next
    VLA_Interpreter.VlaEvalExpression interpolateExpr
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    Report name & " (interpreter)", InStr(1, d, wantFrag, vbTextCompare) > 0, "got: " & d

    d = ""
    On Error Resume Next
    VlaTranspile "(sub t () (debug-print " & interpolateExpr & "))"
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    Report name & " (compiler)", InStr(1, d, wantFrag, vbTextCompare) > 0, "got: " & d
End Sub

' IN2.3: do-until and select, purely - no worksheet needed (for-each's
' own pin lives in VLA_Tests_Host.bas instead, over a real Range,
' matching how the corpus actually uses it - "Remember range B2:B4 as
' results. For each r in results..." - rather than the interpreter's
' unsupported (new Collection), a separate, pre-existing gap this item
' does not touch). exit-do is proven through a nested if, not a bare
' loop body, specifically to prove mLoopBreak really propagates through
' ExecBody/ExecIf rather than only working when exit-do is the loop's
' own direct child.
Private Sub TestInterpreterControlFlow()
    Dim frame As Object

    Set frame = VLA_Interpreter.VlaInterpret( _
        "(begin (dim n Double) (set! n 1)" & _
        " (do-until (> n 5) (set! n (+ n 1))))")
    CheckV "in2.3: do-until loops while its condition is false, testing before each pass", _
           VLA_Runtime.VlaDictGet(frame, "n"), 6

    Set frame = VLA_Interpreter.VlaInterpret( _
        "(begin (dim n Double) (set! n 1)" & _
        " (do-until (> n 100)" & _
        "   (set! n (+ n 1))" & _
        "   (if (> n 3) (then (exit-do)))))")
    CheckV "in2.3: exit-do stops the loop immediately, from inside a nested if", _
           VLA_Runtime.VlaDictGet(frame, "n"), 4

    ' exit-do's own sibling (ExecFor shares the same mLoopBreak check) -
    ' functionally supported since IN2.3 but never pinned by name until
    ' now (AS.8's own scanner found the gap).
    Set frame = VLA_Interpreter.VlaInterpret( _
        "(begin (dim n Double) (set! n 0)" & _
        " (for (i 1 10)" & _
        "   (set! n (+ n 1))" & _
        "   (if (> n 3) (then (exit-for)))))")
    CheckV "in2.3: exit-for stops the loop immediately, from inside a nested if", _
           VLA_Runtime.VlaDictGet(frame, "n"), 4

    Set frame = VLA_Interpreter.VlaInterpret( _
        "(begin (dim label)" & _
        " (select 2" & _
        "   (case (1) (set! label ""one""))" & _
        "   (case (2 3) (set! label ""two-or-three""))" & _
        "   (case-else (set! label ""other""))))")
    CheckV "in2.3: select matches a multi-value case, first match wins", _
           VLA_Runtime.VlaDictGet(frame, "label"), "two-or-three"

    Set frame = VLA_Interpreter.VlaInterpret( _
        "(begin (dim label)" & _
        " (select 99" & _
        "   (case (1) (set! label ""one""))" & _
        "   (case (2 3) (set! label ""two-or-three""))" & _
        "   (case-else (set! label ""other""))))")
    CheckV "in2.3: select falls to case-else when nothing matches", _
           VLA_Runtime.VlaDictGet(frame, "label"), "other"
End Sub

' IN.10: user-procedure call/return - pins mirroring the real shapes
' instructions_golden.vla's own 8 call sites use (positional, named-with-
' defaults, nullary-as-value, nullary-statement), plus recursion,
' exit-sub's unwind, and the latent bug this item's own roadmap text
' named directly: before this pass, EVERY top-level sub's body ran
' once unconditionally, param-blind, in file order - "never-called"
' below proves a declared-but-uncalled procedure now stays dormant.
Private Sub TestInterpreterCallReturn()
    Dim frame As Object

    ' Positional binding + (return <expr>) - instructions.txt's own (tax 100).
    Set frame = VLA_Interpreter.VlaInterpret( _
        "(function double-it ((amount Variant)) (return (* amount 2)))" & vbCrLf & _
        "(sub main () (dim total) (set! total (double-it 21)))")
    CheckV "in.10: positional argument binding + return", _
           VLA_Runtime.VlaDictGet(frame, "total"), 42

    ' Named binding, every default omitted - instructions.txt's own (stamp).
    Set frame = VLA_Interpreter.VlaInterpret( _
        "(function stamp ((optional row-number Variant 1) (optional value Variant ""ok""))" & _
        " (return (& row-number ""-"" value)))" & vbCrLf & _
        "(sub main () (dim result) (set! result (stamp)))")
    CheckV "in.10: named-argument call with every default applied", _
           VLA_Runtime.VlaDictGet(frame, "result"), "1-ok"

    ' Named binding, both supplied - instructions.txt's own
    ' (stamp :row-number 2 :value "beta").
    Set frame = VLA_Interpreter.VlaInterpret( _
        "(function stamp ((optional row-number Variant 1) (optional value Variant ""ok""))" & _
        " (return (& row-number ""-"" value)))" & vbCrLf & _
        "(sub main () (dim result) (set! result (stamp :row-number 2 :value ""beta"")))")
    CheckV "in.10: named-argument call with both arguments supplied", _
           VLA_Runtime.VlaDictGet(frame, "result"), "2-beta"

    ' Nullary function called bare, as a value - instructions.txt's own (vat-rate).
    Set frame = VLA_Interpreter.VlaInterpret( _
        "(function half () (return (/ 10 2)))" & vbCrLf & _
        "(sub main () (dim r) (set! r (half)))")
    CheckV "in.10: nullary function call in expression position", _
           VLA_Runtime.VlaDictGet(frame, "r"), 5

    ' Nullary sub called as a statement - instructions.txt's own (tidy-up) -
    ' observed through the effect log (the callee's own locals are not
    ' visible to the caller, by design, so debug-print is the only
    ' externally observable proof the body actually ran).
    Set frame = VLA_Interpreter.VlaInterpret( _
        "(sub announce () (debug-print ""tidied""))" & vbCrLf & _
        "(sub main () (announce))")
    CheckV "in.10: nullary statement call actually runs the callee's body", _
           CBool(InStr(VLA_Interpreter.VlaInterpreterEffectLog(), "tidied") > 0), True

    ' Recursion through a real call stack (not the emitter's L18 TCO -
    ' an explicit non-goal for this item, real recursion is correct
    ' without it).
    Set frame = VLA_Interpreter.VlaInterpret( _
        "(function fact ((n Variant))" & _
        " (if (> n 1) (then (return (* n (fact (- n 1))))) (else (return 1))))" & vbCrLf & _
        "(sub main () (dim r) (set! r (fact 5)))")
    CheckV "in.10: recursive user-procedure calls via a real call stack", _
           VLA_Runtime.VlaDictGet(frame, "r"), 120

    ' exit-sub unwinds immediately, through an enclosing if - mLoopBreak's
    ' own shape (IN2.3), scoped to the call frame instead of a loop.
    Set frame = VLA_Interpreter.VlaInterpret( _
        "(sub early ((flag Variant)) (debug-print ""before"")" & _
        " (if flag (then (exit-sub))) (debug-print ""after""))" & vbCrLf & _
        "(sub main () (early true))")
    Dim logTrue As String
    logTrue = VLA_Interpreter.VlaInterpreterEffectLog()
    CheckV "in.10: exit-sub stops the callee immediately, skipping later statements", _
           CBool(InStr(logTrue, "before") > 0 And InStr(logTrue, "after") = 0), True

    Set frame = VLA_Interpreter.VlaInterpret( _
        "(sub early ((flag Variant)) (debug-print ""before"")" & _
        " (if flag (then (exit-sub))) (debug-print ""after""))" & vbCrLf & _
        "(sub main () (early false))")
    Dim logFalse As String
    logFalse = VLA_Interpreter.VlaInterpreterEffectLog()
    CheckV "in.10: without exit-sub the callee's remaining statements still run", _
           CBool(InStr(logFalse, "before") > 0 And InStr(logFalse, "after") > 0), True

    ' exit-sub's function-shaped sibling (same "Case exit-sub, exit-function"
    ' line) - a bare exit-function unwinds immediately AND leaves the
    ' return value at Empty, the same default an unassigned real VBA
    ' Function returns (mProcReturnValue is never touched, since
    ' (return 99) never runs).
    Set frame = VLA_Interpreter.VlaInterpret( _
        "(function early-return ((flag Variant))" & _
        " (if flag (then (exit-function))) (return 99))" & vbCrLf & _
        "(sub main () (dim r) (set! r (early-return true)))")
    CheckV "in.10: exit-function stops the callee immediately, before its own return", _
           VLA_Runtime.VlaDictGet(frame, "r"), ""

    ' The bug this item's own roadmap text named: before this pass,
    ' EVERY top-level sub's body ran once, param-blind, in file order.
    ' "never-called" must now stay dormant, and main must still run.
    Set frame = VLA_Interpreter.VlaInterpret( _
        "(sub never-called () (debug-print ""should-not-run""))" & vbCrLf & _
        "(sub main () (dim x) (set! x 1))")
    CheckV "in.10: a declared-but-uncalled procedure no longer executes", _
           CBool(InStr(VLA_Interpreter.VlaInterpreterEffectLog(), "should-not-run") > 0), False
    CheckV "in.10: main still runs normally alongside other declared procedures", _
           VLA_Runtime.VlaDictGet(frame, "x"), 1
End Sub

' IN.3: module-level scope visibility, IN.10's own named non-goal
' ("a call frame is fresh and ISOLATED - it does not see the caller's/
' module's own top-level dim/const bindings") measured against the real
' corpus rather than left an open question: instructions.txt's "Define
' hot-pink as ..." is a true top-level (const hot-pink ...), and
' "To tidy-up:" - a real called procedure - reads it. Also regresses the
' "const" ExecStmt case this same corpus read found missing entirely
' (grepped: no "const" case existed anywhere in this module before this
' pass) - both fixes are exercised by every check below, since none of
' them would get past the top-level (const ...) declaration otherwise.
Private Sub TestInterpreterModuleScope()
    Dim frame As Object

    ' The real gap: a module-level const, read from inside a CALLED
    ' procedure's own isolated frame (instructions.txt's hot-pink/tidy-up
    ' shape exactly - a nullary sub reading a top-level const).
    Set frame = VLA_Interpreter.VlaInterpret( _
        "(const greeting ""hi"")" & vbCrLf & _
        "(function echo () (return greeting))" & vbCrLf & _
        "(sub main () (dim r) (set! r (echo)))")
    CheckV "in.3: a called procedure reads a module-level const (hot-pink/tidy-up shape)", _
           VLA_Runtime.VlaDictGet(frame, "r"), "hi"

    ' Shadowing: a same-named PARAMETER in the callee must win over the
    ' module-level fallback, not be silently overridden by it - the
    ' fallback only fires when the call frame itself has no binding.
    Set frame = VLA_Interpreter.VlaInterpret( _
        "(const x 1)" & vbCrLf & _
        "(function shadow ((x Variant)) (return x))" & vbCrLf & _
        "(sub main () (dim r) (set! r (shadow 99)))")
    CheckV "in.3: a callee's own parameter shadows a same-named module-level const", _
           VLA_Runtime.VlaDictGet(frame, "r"), 99

    ' A genuinely undefined name - present in NEITHER the call frame nor
    ' module scope - still raises loudly, proving the fallback narrows
    ' the miss rather than swallowing every unbound-name error.
    Dim raised As Boolean
    On Error Resume Next
    VLA_Interpreter.VlaInterpret _
        "(function bad () (return nowhere))" & vbCrLf & "(sub main () (bad))"
    raised = (Err.Number <> 0)
    On Error GoTo 0
    Report "in.3: a name unbound in both the call frame and module scope still raises", _
           raised, "VlaInterpret did not raise for a truly undefined name"
End Sub

' IN.3: (new Collection) - VLA_Interpreter.bas's own EvalExpr "new" Case
' note has the full reasoning. Found live, not read: english.vla's
' V3.1 "Create a list called ..."/"Append ... to ..." macros
' ((begin (dim found-items Collection) (obj-set! found-items (new
' Collection)))) raised "there is nothing stored at key 'Collection'" -
' the type-name argument was being evaluated as an EXPRESSION (a plain
' variable lookup) before this fix, since EvalExpr had no dedicated case
' for "new" and fell through to the generic dynamic-head dispatch, which
' evaluates every remaining list element as an argument.
Private Sub TestInterpreterNewCollection()
    Dim frame As Object
    Set frame = VLA_Interpreter.VlaInterpret("(begin (dim c) (obj-set! c (new Collection)))")
    Report "in.3: (new Collection) returns a real Collection object", _
           TypeName(VLA_Runtime.VlaDictGet(frame, "c")) = "Collection", _
           "got " & TypeName(VLA_Runtime.VlaDictGet(frame, "c"))

    ' A genuinely unsupported type still refuses in words - not the
    ' misleading "undefined variable" miss this fix replaces.
    Dim raised As Boolean
    On Error Resume Next
    VLA_Interpreter.VlaInterpret "(begin (dim c) (obj-set! c (new Dictionary)))"
    raised = (Err.Number <> 0)
    On Error GoTo 0
    Report "in.3: an unsupported (new ...) type still refuses in words", raised, _
           "VlaInterpret did not raise for an unsupported new-type"

    ' IN.11: the actual real-corpus shape this pin's own siblings never
    ' exercised - creating a Collection AND appending a real value AND
    ' reading it back. "Set pick-check to item 2 of found-items."
    ' (instructions.txt line 243, "found-items" built exactly this way -
    ' Create a list called .../Append ... to ...) returned a raw-
    ' pointer-shaped garbage number instead of the real appended value
    ' once VerifyReportInterpreter finally reached it live - the real
    ' cause was a genuine VBA quirk in CallByNameArgs/CallByNameArgsVoid
    ' (VLA_Interpreter.bas's own IN.11 note on CallByNameArgs has the
    ' full mechanism, found only after a standalone, zero-shared-state
    ' repro outside this project's own test infrastructure - a throwaway
    ' scratch module, not kept here), now fixed at its source. This pin
    ' is the real-corpus-shaped regression proof.
    Dim frame2 As Object
    Set frame2 = VLA_Interpreter.VlaInterpret( _
        "(begin (dim c) (obj-set! c (new Collection)) (. c add 300) (. c add 400) (dim x) (set! x (vlaitem c 2)))")
    CheckV "in.11: appending a number to a (new Collection) and reading it back (via vlaitem)", _
           VLA_Runtime.VlaDictGet(frame2, "x"), 400

    ' The same append, read back via real VBA .Item directly on the
    ' Collection pulled out of the frame - not through this
    ' interpreter's own read dispatch at all (ruling out a SEPARATE
    ' read-side bug were this ever to regress: if this pin passes but
    ' the one above does not, the problem moved to TryRuntimeHelper's
    ' own Application.Run path, not .add's own storage). A first
    ' attempt read back via "(. c item 2)" instead - DynamicGet's own
    ' CallByNameArgs path - but Collection's "Item" turned out to be
    ' its own separate, unrelated gap (unresolvable via CallByName's
    ' VbGet/VbMethod fallback at all, the same "Borders"-shaped
    ' parameterized-default-property issue found once already) -
    ' raising before ever reaching this comparison, a distraction from
    ' the actual question, not chased further here.
    Dim frame3 As Object
    Set frame3 = VLA_Interpreter.VlaInterpret( _
        "(begin (dim c) (obj-set! c (new Collection)) (. c add 300) (. c add 400))")
    Dim rawColl As Collection
    Set rawColl = VLA_Runtime.VlaDictGet(frame3, "c")
    CheckV "in.11: appending a number to a (new Collection), read back via real VBA .Item (bypassing this interpreter's own read dispatch entirely)", _
           rawColl.Item(2), 400
End Sub

' SEC.1 Tier 2: the CallByName fallback that used to sit at the bottom
' of DynamicGet/DynamicCall/DynamicSet is gone - anything not in their
' own native Select Case now refuses in words instead of reaching
' arbitrary late-bound dispatch. Collection.Remove is real, host-free
' (no Excel needed, unlike almost every other member this pass
' touched), and was perfectly resolvable through the old CallByName
' fallback - a plain method, none of the parameterized-default-property
' sharp edges IN.3/IN.11 hit elsewhere - so proving it is refused now
' proves the fallback itself is gone, not merely that this one member
' was never reached before. Pre-seeded real state (a genuine
' (new Collection) with one real element), not a bypass toggle -
' SEC.2's own "no test-bypass toggle" discipline, applied here too.
Private Sub TestInterpreterSec1DynamicMemberRefused()
    Dim raised As Boolean
    Dim desc As String
    On Error Resume Next
    Err.Clear
    VLA_Interpreter.VlaInterpret "(begin (dim c) (obj-set! c (new Collection)) (. c add 5) (. c remove 1))"
    raised = (Err.Number <> 0)
    desc = Err.Description
    On Error GoTo 0
    Report "sec.1 tier 2: '.remove' on a real Collection (real, CallByName-resolvable before this fix) is refused in words, not silently dispatched", _
           raised And InStr(1, desc, "native dynamic-dispatch allowlist", vbTextCompare) > 0, _
           "wanted a SEC.1 refusal naming the allowlist, got: raised=" & raised & " desc=" & desc
End Sub

' IN3.6: on-error/goto/label/resume - pinned directly against raw VLA
' in Case "try"'s own shape (VLA_English.bas's ParseStmt), not just the
' four bare forms parity already checks: (on-error goto <label>) armed,
' a real error jumps to recovery; the SAME shape with no error skips
' recovery via its own (goto <label-past-it>); (on-error resume-next)
' silently absorbs a real error and continues to the next statement;
' and - the item's own named correctness hazard, checked directly, not
' assumed - this interpreter's OWN diagnostic signaling (an unresolved
' head) still escapes an armed handler rather than being swallowed as
' if it were the program's own recoverable error. All five use "(/ 1
' 0)" for the real, genuinely-trappable error: real VBA division,
' unwrapped by any of this project's own Err.Raise calls, so it always
' carries VBA's own default Source - never one of the three this
' project's own diagnostics use - the same real-vs-own distinction a
' live Excel run (instructions.txt's "Go to sheet Nowhere-Land") exercises
' with a COM error instead, without needing Excel to pin it here.
Private Sub TestInterpreterOnError()
    Dim frame As Object

    ' A real error under (on-error goto ...): the failing statement's
    ' own remainder ((set! result 999), never reached) is skipped, the
    ' recovery paragraph runs, and vla-problem's own shape (a plain
    ' variable set from err.description) reads the real failure text.
    Set frame = VLA_Interpreter.VlaInterpret( _
        "(begin" & vbCrLf & _
        "  (dim problem)" & vbCrLf & _
        "  (dim result)" & vbCrLf & _
        "  (on-error goto errf)" & vbCrLf & _
        "  (set! result (/ 1 0))" & vbCrLf & _
        "  (set! result 999)" & vbCrLf & _
        "  (goto trydone)" & vbCrLf & _
        "  (label errf)" & vbCrLf & _
        "  (set! problem err.description)" & vbCrLf & _
        "  (resume tryr)" & vbCrLf & _
        "  (label tryr)" & vbCrLf & _
        "  (on-error goto 0)" & vbCrLf & _
        "  (set! result -1)" & vbCrLf & _
        "  (label trydone)" & vbCrLf & _
        "  (on-error goto 0))")
    CheckV "in3.6: a real error under (on-error goto ...) runs the recovery, not the rest of the body", _
           VLA_Runtime.VlaDictGet(frame, "result"), -1
    Report "in3.6: err.description is readable in recovery, before (resume ...) clears it", _
           Len(CStr(VLA_Runtime.VlaDictGet(frame, "problem"))) > 0, _
           "problem was empty - err.description did not read the real failure"

    ' No error: the body's own (goto <past-recovery>) fires instead,
    ' and the recovery paragraph never runs at all.
    Set frame = VLA_Interpreter.VlaInterpret( _
        "(begin" & vbCrLf & _
        "  (dim result)" & vbCrLf & _
        "  (on-error goto errf)" & vbCrLf & _
        "  (set! result 42)" & vbCrLf & _
        "  (goto trydone)" & vbCrLf & _
        "  (label errf)" & vbCrLf & _
        "  (set! result -1)" & vbCrLf & _
        "  (label trydone)" & vbCrLf & _
        "  (on-error goto 0))")
    CheckV "in3.6: (on-error goto ...) with no error skips the recovery paragraph entirely", _
           VLA_Runtime.VlaDictGet(frame, "result"), 42

    ' (on-error resume-next) - show-all-rows's own shape: a real error
    ' is silently absorbed (the failing statement's own assignment
    ' never happens), and execution just continues at the next one.
    Set frame = VLA_Interpreter.VlaInterpret( _
        "(begin" & vbCrLf & _
        "  (dim x)" & vbCrLf & _
        "  (on-error resume-next)" & vbCrLf & _
        "  (set! x (/ 1 0))" & vbCrLf & _
        "  (set! x 5)" & vbCrLf & _
        "  (on-error goto 0))")
    CheckV "in3.6: (on-error resume-next) absorbs a real error and continues at the next statement", _
           VLA_Runtime.VlaDictGet(frame, "x"), 5

    ' The named correctness hazard, checked directly: this
    ' interpreter's OWN diagnostic (an unresolved head - no such form,
    ' place helper, dotted global, built-in, or user procedure) must
    ' still escape an armed handler, not be swallowed as a "recoverable
    ' program error" the way the real (/ 1 0) failure above correctly
    ' was.
    Dim raisedOwn As Boolean
    On Error Resume Next
    VLA_Interpreter.VlaInterpret _
        "(begin (on-error goto errf) (totally-unresolved-head-in3-6) (label errf))"
    raisedOwn = (Err.Number <> 0)
    On Error GoTo 0
    Report "in3.6: this interpreter's own diagnostic raise still escapes an armed on-error handler", _
           raisedOwn, "VlaInterpret did not raise - an interpreter-internal error was silently trapped as if it were a program error"

    ' A goto/resume target that does not exist anywhere is a real bug
    ' (a mistyped label), not a silently stuck flag - raised loudly by
    ' name, the same AS.6 discipline this whole module already follows.
    Dim raisedMissing As Boolean
    On Error Resume Next
    VLA_Interpreter.VlaInterpret "(begin (goto nowhere-in3-6))"
    raisedMissing = (Err.Number <> 0)
    On Error GoTo 0
    Report "in3.6: a goto/resume target label that is never found still raises, rather than silently doing nothing", _
           raisedMissing, "VlaInterpret did not raise for an unresolved goto target"
End Sub

' F5.0: proves VlaPushContext/VlaPopContext round-trip VLA.bas's
' compile-time state, not just that the depth counter moves. Two
' transpiles that each define a differently-named macro, observed
' through the existing Public VlaMacroDoc (reads mMacros/mSubDocs, the
' two object-typed fields most likely to reveal a shallow-copy or
' missed-field bug - mMacros a Dictionary since PDICT.0, mSubDocs still
' a Collection, both captured by reference the same way): the nested
' compile's "beta" must not leak out, and
' popping must bring "alpha" back exactly as it was, not merely make
' VlaContextDepth read 0 again. Guarded the same way TestDotCount is
' (AS.6): a raise anywhere here reports FAIL under this pin's own name
' instead of escaping and killing the suite.
Private Sub TestContextPushPop()
    Dim d As String
    Dim startDepth As Long
    startDepth = VlaContextDepth()

    On Error Resume Next
    VlaTranspile "(defmacro (alpha x) ""alpha's docstring"" (+ x 1))"
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    If Len(d) > 0 Then
        Report "f5: context push/pop round-trips compile state", False, "outer transpile error: " & d
        Exit Sub
    End If
    Dim outerDoc As String
    outerDoc = VlaMacroDoc("alpha")

    On Error Resume Next
    VlaPushContext
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    If Len(d) > 0 Then
        Report "f5: context push/pop round-trips compile state", False, "VlaPushContext error: " & d
        Exit Sub
    End If
    Dim depthAfterPush As Long
    depthAfterPush = VlaContextDepth()

    On Error Resume Next
    VlaTranspile "(defmacro (beta y) ""beta's docstring"" (- y 1))"
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    If Len(d) > 0 Then
        Report "f5: context push/pop round-trips compile state", False, "nested transpile error: " & d
        VlaPopContext   ' balance the stack even after a failed nested compile
        Exit Sub
    End If
    Dim innerAlphaDoc As String, innerBetaDoc As String
    innerAlphaDoc = VlaMacroDoc("alpha")
    innerBetaDoc = VlaMacroDoc("beta")

    On Error Resume Next
    VlaPopContext
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    If Len(d) > 0 Then
        Report "f5: context push/pop round-trips compile state", False, "VlaPopContext error: " & d
        Exit Sub
    End If
    Dim depthAfterPop As Long
    depthAfterPop = VlaContextDepth()
    Dim restoredAlphaDoc As String, restoredBetaDoc As String
    restoredAlphaDoc = VlaMacroDoc("alpha")
    restoredBetaDoc = VlaMacroDoc("beta")

    Dim ok As Boolean
    ok = (depthAfterPush = startDepth + 1)
    ok = ok And (depthAfterPop = startDepth)
    ok = ok And (Len(outerDoc) > 0)
    ok = ok And (innerAlphaDoc = "")
    ok = ok And (innerBetaDoc = "beta's docstring")
    ok = ok And (restoredAlphaDoc = outerDoc)
    ok = ok And (restoredBetaDoc = "")
    Report "f5: context push/pop round-trips compile state", ok, _
           "depths " & startDepth & "/" & depthAfterPush & "/" & depthAfterPop & _
           ", outerDoc=[" & outerDoc & "] innerAlpha=[" & innerAlphaDoc & "] innerBeta=[" & innerBetaDoc & _
           "] restoredAlpha=[" & restoredAlphaDoc & "] restoredBeta=[" & restoredBetaDoc & "]"
End Sub

' F5.0: the underflow guard - VlaPopContext must raise, not silently
' no-op, when called with nothing left to restore. Same On-Error-and-
' check shape AssertClaimRefusal uses elsewhere in this file for
' "this must fail." Assumes TestContextPushPop (dispatched immediately
' before this one) left the stack balanced at 0; the precondition check
' below reports that assumption's own failure honestly rather than
' misreporting the underflow guard as broken if it does not hold.
Private Sub TestContextUnderflow()
    Dim startDepth As Long
    startDepth = VlaContextDepth()
    If startDepth <> 0 Then
        Report "f5: an unbalanced VlaPopContext raises", False, _
               "test precondition failed: context depth is " & startDepth & ", expected 0 (a prior test left it unbalanced)"
        Exit Sub
    End If
    Dim d As String
    On Error Resume Next
    VlaPopContext
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    Report "f5: an unbalanced VlaPopContext raises", Len(d) > 0, "expected a raise, got: " & IIf(Len(d) > 0, d, "(no error)")
End Sub









' CO.4 - grammar semantic versioning. VLA.bas's own CO.4 banner carries
' the decision (the grammar's compatibility version IS
' VLA_RELEASE_VERSION under SD-14's triggers; no second axis); these
' pins hold the ORDERING and the REFUSAL that decision leans on.
'
' The pin that matters most is "0.9.0 is older than 0.10.0": a string
' compare gets that backwards, and it is the one wrong answer that
' would silently mis-gate a phrasebook's `requires: version:` for real
' once minor numbers reach double digits. It is pinned first, and on
' purpose, rather than left to be noticed at 0.10.0.
Private Sub TestCo4VersionSemver()
    Dim mj As Long, mn As Long, pt As Long
    Dim ok As Boolean
    Dim d As String

    ' --- parse: the shapes that must read ---
    ok = VlaVersionParse("0.5.1", mj, mn, pt)
    Report "co4: 0.5.1 parses to 0/5/1", _
           ok And mj = 0 And mn = 5 And pt = 1, _
           "got ok=" & ok & " " & mj & "/" & mn & "/" & pt
    ok = VlaVersionParse("  0.5.1  ", mj, mn, pt)
    Report "co4: surrounding whitespace is trimmed", _
           ok And mj = 0 And mn = 5 And pt = 1, _
           "got ok=" & ok & " " & mj & "/" & mn & "/" & pt
    ok = VlaVersionParse("999999999.0.0", mj, mn, pt)
    Report "co4: a 9-digit part still parses (Long headroom)", _
           ok And mj = 999999999, "got ok=" & ok & " major=" & mj

    ' --- parse: the shapes that must NOT read, each for its own reason ---
    Report "co4: empty text is not a version", _
           Not VlaVersionParse("", mj, mn, pt), "empty text parsed"
    Report "co4: two numbers is not a version", _
           Not VlaVersionParse("1.2", mj, mn, pt), "1.2 parsed"
    Report "co4: four numbers is not a version", _
           Not VlaVersionParse("1.2.3.4", mj, mn, pt), "1.2.3.4 parsed"
    Report "co4: an empty part is not a number", _
           Not VlaVersionParse("1..3", mj, mn, pt), "1..3 parsed"
    Report "co4: a trailing dot is not a version", _
           Not VlaVersionParse("1.2.", mj, mn, pt), "1.2. parsed"
    Report "co4: words are not a version", _
           Not VlaVersionParse("banana", mj, mn, pt), "banana parsed"
    ' The deliberate scope line, pinned so a future author who adds
    ' pre-release ordering has to change a test that states the reason.
    Report "co4: a -beta suffix does not parse (no ordering defined)", _
           Not VlaVersionParse("0.5.1-beta", mj, mn, pt), "0.5.1-beta parsed"
    Report "co4: a signed part is not a version", _
           Not VlaVersionParse("-1.2.3", mj, mn, pt), "-1.2.3 parsed"
    Report "co4: a 10-digit part is refused, not overflowed", _
           Not VlaVersionParse("1234567890.0.0", mj, mn, pt), "1234567890.0.0 parsed"

    ' A failed parse leaves the out-params at 0, never half-filled.
    mj = 7: mn = 7: pt = 7
    VlaVersionParse "1.2", mj, mn, pt
    Report "co4: a failed parse zeroes its out-params", _
           mj = 0 And mn = 0 And pt = 0, "got " & mj & "/" & mn & "/" & pt

    ' --- ordering ---
    Report "co4: 0.9.0 is OLDER than 0.10.0 (numeric, not string, order)", _
           VlaVersionCompare("0.9.0", "0.10.0") = -1, _
           "got " & VlaVersionCompare("0.9.0", "0.10.0") & " - a string compare would say +1 here"
    Report "co4: 0.10.0 is NEWER than 0.9.0", _
           VlaVersionCompare("0.10.0", "0.9.0") = 1, _
           "got " & VlaVersionCompare("0.10.0", "0.9.0")
    Report "co4: equal versions compare 0", _
           VlaVersionCompare("0.5.1", "0.5.1") = 0, _
           "got " & VlaVersionCompare("0.5.1", "0.5.1")
    Report "co4: MAJOR outranks MINOR and PATCH", _
           VlaVersionCompare("1.0.0", "0.99.99") = 1, _
           "got " & VlaVersionCompare("1.0.0", "0.99.99")
    Report "co4: MINOR outranks PATCH", _
           VlaVersionCompare("0.6.0", "0.5.99") = 1, _
           "got " & VlaVersionCompare("0.6.0", "0.5.99")
    Report "co4: PATCH breaks the tie", _
           VlaVersionCompare("0.5.2", "0.5.1") = 1, _
           "got " & VlaVersionCompare("0.5.2", "0.5.1")

    ' --- the refusal: malformed compares as nothing, on either side ---
    d = ""
    On Error Resume Next
    VlaVersionCompare "banana", "0.5.1"
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    Report "co4: a malformed LEFT side refuses in words", _
           InStr(d, "not a version this build can compare") > 0, _
           "expected the vla-version-malformed refusal, got: " & IIf(Len(d) > 0, d, "(no error)")

    d = ""
    On Error Resume Next
    VlaVersionCompare "0.5.1", "banana"
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    Report "co4: a malformed RIGHT side refuses in words", _
           InStr(d, "not a version this build can compare") > 0, _
           "expected the vla-version-malformed refusal, got: " & IIf(Len(d) > 0, d, "(no error)")

    ' --- VlaVersionAtLeast, against the real running VLA_RELEASE_VERSION ---
    Report "co4: this build is at least 0.0.0", _
           VlaVersionAtLeast("0.0.0"), _
           "VLA_RELEASE_VERSION=" & VLA_RELEASE_VERSION & " did not clear 0.0.0"
    Report "co4: this build is at least its own version", _
           VlaVersionAtLeast(VLA_RELEASE_VERSION), _
           "VLA_RELEASE_VERSION=" & VLA_RELEASE_VERSION & " did not clear itself"
    Report "co4: this build is NOT at least 999.0.0", _
           Not VlaVersionAtLeast("999.0.0"), _
           "VLA_RELEASE_VERSION=" & VLA_RELEASE_VERSION & " claimed to clear 999.0.0"

    ' A typo'd requires: line refuses; it does not quietly read as unmet.
    d = ""
    On Error Resume Next
    VlaVersionAtLeast "0.5"
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    Report "co4: a typo'd minimum refuses rather than reading as unmet", _
           InStr(d, "not a version this build can compare") > 0, _
           "expected the vla-version-malformed refusal, got: " & IIf(Len(d) > 0, d, "(no error)")

    ' The shipped constant itself must be readable by the very scheme
    ' that now depends on it - the pin that would fire the day someone
    ' hand-edits VLA_RELEASE_VERSION into a shape release.ps1's regex
    ' still accepts but this ordering cannot read.
    Report "co4: VLA_RELEASE_VERSION itself parses as MAJOR.MINOR.PATCH", _
           VlaVersionParse(VLA_RELEASE_VERSION, mj, mn, pt), _
           "VLA_RELEASE_VERSION=" & VLA_RELEASE_VERSION & " is not readable by CO.4's own ordering"
End Sub

' F.10 - (requires-<namespace> "<value>") in phrasebooks. The parse is
' shared; the ENFORCEMENT deliberately is not, and several of these
' pins exist to keep that split from being quietly collapsed later.
'
' The two that matter most:
'   - "version is checked before any other tag" holds the owner's own
'     unknown-namespace decision together. Refusing an unknown
'     namespace is only humane because a phrasebook that declares its
'     version gets the precise message instead of the vague one, and
'     that is purely a matter of evaluation order, which nothing else
'     would catch if it regressed.
'   - "capability is NOT checked on the host-free text path" is
'     SEC.2's own trap, pinned so it cannot be re-learned live: that
'     path is VLA_Browser.bas's, documented as never showing a dialog
'     and tracked by name in check_translate_purity.ps1.
Private Sub TestF10Requires()
    Dim v As String
    Dim d As String
    Dim got As String

    ' --- met, unmet, and the two malformed shapes ---
    EnglishResetGrammar
    v = "(requires-version ""0.0.0"")" & vbLf & _
        "(english-vla ""zzmet cell {r:text}"" (set! (range {r}) 1))"
    On Error Resume Next
    Err.Clear
    EnglishLoadVocabularyText v, "selftest-vocab"
    d = Err.Description
    On Error GoTo 0
    Report "f10: a met version requirement loads normally", Err.Number = 0, d

    EnglishResetGrammar
    v = "(requires-version ""999.0.0"")"
    d = ""
    On Error Resume Next
    Err.Clear
    EnglishLoadVocabularyText v, "selftest-vocab"
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    Report "f10: an unmet version requirement refuses, naming both versions", _
           InStr(d, "needs Frazaro 999.0.0") > 0 And InStr(d, VLA_RELEASE_VERSION) > 0, _
           "expected the unmet-version refusal naming 999.0.0 and " & VLA_RELEASE_VERSION & ", got: " & IIf(Len(d) > 0, d, "(no error)")

    ' A typo'd version must refuse AS a typo - never read as merely
    ' unmet, which would tell an author to upgrade Frazaro to fix a
    ' spelling mistake. CO.4's VlaVersionParse is what makes this
    ' checkable with no error handler around it.
    EnglishResetGrammar
    v = "(requires-version ""banana"")"
    d = ""
    On Error Resume Next
    Err.Clear
    EnglishLoadVocabularyText v, "selftest-vocab"
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    Report "f10: a malformed version refuses as malformed, not as unmet", _
           InStr(d, "is not a version") > 0, _
           "expected the malformed-version refusal, got: " & IIf(Len(d) > 0, d, "(no error)")

    EnglishResetGrammar
    v = "(requires-version)"
    d = ""
    On Error Resume Next
    Err.Clear
    EnglishLoadVocabularyText v, "selftest-vocab"
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    Report "f10: a requirement with no value refuses", _
           InStr(d, "names no value") > 0, _
           "expected the missing-value refusal, got: " & IIf(Len(d) > 0, d, "(no error)")

    ' --- the unknown-namespace policy, and what makes it humane ---
    EnglishResetGrammar
    v = "(requires-signature ""acme"")"
    d = ""
    On Error Resume Next
    Err.Clear
    EnglishLoadVocabularyText v, "selftest-vocab"
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    Report "f10: an unknown namespace refuses rather than being ignored", _
           InStr(d, "is not a kind of requirement") > 0, _
           "expected the unknown-namespace refusal, got: " & IIf(Len(d) > 0, d, "(no error)")

    ' VERSION-FIRST ORDERING. The unknown tag is written FIRST here on
    ' purpose: file order must not decide which refusal an author
    ' sees. Without the ordering rule this reports "unknown
    ' requirement 'signature'" and the author never learns the one
    ' fact that would have told them what to do.
    EnglishResetGrammar
    v = "(requires-signature ""acme"")" & vbLf & _
        "(requires-version ""999.0.0"")"
    d = ""
    On Error Resume Next
    Err.Clear
    EnglishLoadVocabularyText v, "selftest-vocab"
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    Report "f10: version is checked before any other tag, whatever the file order", _
           InStr(d, "needs Frazaro 999.0.0") > 0, _
           "expected the version refusal to win over the unknown-namespace one, got: " & IIf(Len(d) > 0, d, "(no error)")

    ' --- the two-site split (SEC.2's trap, pinned) ---
    EnglishResetGrammar
    v = "(requires-capability ""sendmail"")" & vbLf & _
        "(english-vla ""zzcap cell {r:text}"" (set! (range {r}) 1))"
    d = ""
    On Error Resume Next
    Err.Clear
    EnglishLoadVocabularyText v, "selftest-vocab"
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    Report "f10: capability is NOT checked on the host-free text path", _
           Len(d) = 0, _
           "the host-free path enforced a consent-shaped namespace, which is SEC.2's own correction undone: " & d

    ' --- form: recognized, refused, not silently passed ---
    EnglishResetGrammar
    v = "(requires-form ""paint cell"")"
    d = ""
    On Error Resume Next
    Err.Clear
    EnglishLoadVocabularyText v, "selftest-vocab"
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    Report "f10: form is understood but refuses as not-yet-enforceable", _
           InStr(d, "not yet enforceable") > 0, _
           "expected the form-unsupported refusal, got: " & IIf(Len(d) > 0, d, "(no error)")

    ' --- the scan is lexical: comments and string literals are not
    '     declarations (VocabTextHasRawForm's own reason, same shape) ---
    EnglishResetGrammar
    v = "; (requires-version ""999.0.0"")" & vbLf & _
        "(english-vla ""zzcomment cell {r:text}"" (set! (range {r}) 1))"
    d = ""
    On Error Resume Next
    Err.Clear
    EnglishLoadVocabularyText v, "selftest-vocab"
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    Report "f10: a requirement inside a comment is not a requirement", _
           Len(d) = 0, "a commented-out declaration gated the load: " & d

    EnglishResetGrammar
    v = "(english-vla ""zzsay cell {r:text}""" & vbLf & _
        "    (set! (range {r}) ""(requires-version \""999.0.0\"")""))"
    d = ""
    On Error Resume Next
    Err.Clear
    EnglishLoadVocabularyText v, "selftest-vocab"
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    Report "f10: a requirement inside a string literal is not a requirement", _
           Len(d) = 0, "a declaration quoted inside string data gated the load: " & d

    ' --- refuses BEFORE anything registers, wherever the tag sits ---
    ' The declaration is written BELOW the rule deliberately: the
    ' pre-pass reads the whole file before dispatch begins, so
    ' position cannot change what loads. A dispatch-time arm would
    ' have registered this rule first and then refused.
    EnglishResetGrammar
    v = "(english-vla ""zzlate cell {r:text}"" (set! (range {r}) 1))" & vbLf & _
        "(requires-version ""999.0.0"")"
    On Error Resume Next
    Err.Clear
    EnglishLoadVocabularyText v, "selftest-vocab"
    On Error GoTo 0
    got = ""
    On Error Resume Next
    Err.Clear
    got = EnglishToVla("Zzlate cell B2.")
    On Error GoTo 0
    Report "f10: an unmet requirement refuses before a single rule registers", _
           Len(got) = 0, _
           "a rule from a refused phrasebook registered anyway: " & got

    ' --- a generator cannot smuggle one in ---
    ' The pre-pass reads raw source text, so a declaration that only
    ' exists after macro expansion was never checked. Honouring it
    ' would be honouring an unchecked requirement.
    EnglishResetGrammar
    v = "(defmacro (zzgen) ""g"" (requires-version ""0.0.0""))" & vbLf & _
        "(zzgen)"
    d = ""
    On Error Resume Next
    Err.Clear
    EnglishLoadVocabularyText v, "selftest-vocab"
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    Report "f10: a requirement produced by expansion refuses rather than counting", _
           InStr(d, "produced by a generator's expansion") > 0, _
           "expected the from-expansion refusal, got: " & IIf(Len(d) > 0, d, "(no error)")

    ' --- the shipped corpus declares nothing, and must keep loading ---
    EnglishResetGrammar
End Sub

' L2.1: transpile that cannot crash the suite - a raise reports FAIL
' under the pin's own name and returns "", so the caller skips its
' check (exactly one Report per pin either way) and every section
' after this one still runs.
Public Function TryTranspile(ByVal name As String, ByVal src As String) As String
    Dim d As String
    On Error Resume Next
    TryTranspile = VlaTranspile(src)
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    If Len(d) > 0 Then
        Report name, False, "transpile error: " & d
        TryTranspile = ""
    End If
End Function

' L2.2: TryTranspile's sibling for the English half of an
' engine-generated pipeline - same contract, same reason.
Public Function TryEnglish(ByVal name As String, ByVal sentences As String) As String
    Dim d As String
    On Error Resume Next
    TryEnglish = EnglishToVla(sentences)
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    If Len(d) > 0 Then
        Report name, False, "translate error: " & d
        TryEnglish = ""
    End If
End Function

Public Sub AssertVla(ByVal name As String, ByVal src As String, ParamArray frags() As Variant)
    Dim t As String, d As String
    On Error Resume Next
    t = VlaTranspile(src)
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    If Len(d) > 0 Then
        Report "emit: " & name, False, "transpile error: " & d
        Exit Sub
    End If
    ' VBA forbids forwarding a ParamArray to another procedure;
    ' copying it into a plain Variant first is the sanctioned move.
    Dim fr As Variant
    fr = frags
    CheckFrags "emit: " & name, t, fr
End Sub

Public Sub AssertEnglish(ByVal name As String, ByVal sentence As String, ByVal frag As String)
    Dim t As String, d As String
    On Error Resume Next
    t = EnglishToVla(sentence)
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    If Len(d) > 0 Then
        Report "english: " & name, False, "translate error: " & d
        Exit Sub
    End If
    Report "english: " & name, InStr(1, Norm(t), Norm(frag), vbTextCompare) > 0, _
           "missing '" & frag & "' in: " & Left$(Norm(t), 160)
End Sub

Public Sub CheckFrags(ByVal name As String, ByVal text As String, frags As Variant)
    Dim i As Long
    Dim nt As String
    nt = Norm(text)
    For i = LBound(frags) To UBound(frags)
        If InStr(1, nt, Norm(CStr(frags(i))), vbTextCompare) = 0 Then
            Report name, False, "missing '" & frags(i) & "' in: " & Left$(nt, 160)
            Exit Sub
        End If
    Next
    Report name, True, ""
End Sub

Public Sub CheckV(ByVal name As String, ByVal got As Variant, ByVal want As Variant)
    Report name, CStr(got) = CStr(want), "got '" & got & "', wanted '" & want & "'"
End Sub

' QUASIQUOTE: factors out TestVlaExpand's own inline "expect
' VlaExpandText to raise, check the message" pattern (its "reader
' error raises" case) - reused here often enough (every reserved-name
' and guard-rail assertion) to earn a shared helper rather than
' repeating the On Error dance at each call site.
Public Sub CheckExpandErr(ByVal name As String, ByVal source As String, ByVal wantFrag As String)
    Dim d As String
    Dim fired As Long
    On Error Resume Next
    VlaExpandText source, True, fired
    d = ""
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    Report name, InStr(1, d, wantFrag, vbTextCompare) > 0, "got: " & d
End Sub

Public Sub Report(ByVal name As String, ByVal ok As Boolean, ByVal detail As String)
    If ok Then
        mPass = mPass + 1
        Debug.Print "  PASS  " & name
    Else
        mFail = mFail + 1
        Debug.Print "  FAIL  " & name & " - " & detail
        ' Immediate Window scrollback is finite and this suite is long
        ' enough to outrun it - the recap at the end of VlaSelfTest
        ' reprints name + detail together so a failure near the start
        ' is not lost by the time the summary line prints.
        If Not mFailedNames Is Nothing Then mFailedNames.Add name & " - " & detail
    End If
End Sub

' AS7.1: see this module's own top-of-file note.
Public Sub VlaSelfTestLastResult(ByRef outPass As Long, ByRef outFail As Long, ByRef outFailed As Collection)
    outPass = mPass
    outFail = mFail
    Set outFailed = mFailedNames
End Sub

Public Function Norm(ByVal s As String) As String
    s = Replace(s, vbCrLf, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, vbTab, " ")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    Norm = Trim$(s)
End Function
