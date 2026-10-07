Attribute VB_Name = "VLA_Runtime"
Option Explicit
' SPDX-License-Identifier: 0BSD
' Copyright 2026 Spreadsheet Company. This module's text above the
' EN_RUNTIME INJECT BOUNDARY is copied into your workbook as
' Frazaro_EN_Runtime under the BSD Zero Clause License: use it, change
' it, ship it, no attribution required. Generated code beside it is
' yours outright (OUTPUT-EXCEPTION.md in the Frazaro repository).
' SIG.0: this notice sits INSIDE the injectable region on purpose, so
' it travels with the code it licenses; tools/check_spdx.ps1 fails if
' it ever drifts below the boundary.
Public Const VLA_RUNTIME_VERSION As String = "SEC.15"
' SEC.15: VlaSetFormula, the one formula sink on both backends (VLA.bas's
' emitter writes Call VlaSetFormula(obj, text) where it wrote .Formula2 =,
' and VLA_Interpreter's Case "formula" calls it), and VlaFormulaEgress,
' the pure scan it asks first: a formula that calls WEBSERVICE, FILTERXML,
' HYPERLINK, RTD, IMAGE, STOCKHISTORY, an XLM function that reaches a
' program or a DLL, Sheets' IMPORT* and GOOGLE* fetches or Calc's DDE, or
' that holds a DDE link (a | outside a string), is refused by name before
' anything is written (rt-formula-egress). Held to scripts/egress.txt and
' scripts/egress_golden.txt (VlaWriteEgressGolden, VLA_Tests.bas), which
' the core's egress.rs reproduces; the list is the one place in the VBA,
' and tools/check_egress_golden.ps1 holds it equal to the core's.
' DATALOG13.0: the VlaDict family (and VLA_Relation's join index) test
' for the fallback with VlaDictIsFallback (Is Nothing, then TypeOf) instead
' of asking TypeName(d) = "Dictionary" - TypeName on the late-bound
' Dictionary cost 0.148 ms a call and most of a DATALOG row. The same
' behaviour on both representations and on Nothing; the note at
' VlaDictSet says why the fallback is now the branch tested for.
' PF4B.0: VlaSlabRead/VlaSlabWrite - PRODUCT · PERFORMANCE's PF.4b, the
' array-slab bulk read/write-back helpers PF.4c's own for-each-row will
' build on. One Range.Value read/write regardless of row count, instead
' of one Cells(i, j) COM call per cell - the same shape PF.7's own
' VlaEmbeddedText fix (VLA.bas) already proved for a sheet-column read.
' Self-contained, not a delegation to VLA_Relation.SourceToArray despite
' that function already doing almost exactly this for SQL/DATALOG - a
' real design correction caught before landing, not after: this
' module's own code above the EN_RUNTIME INJECT BOUNDARY (below) is
' copied verbatim into a standalone user workbook with NO other add-in
' module available, so a call out to VLA_Relation would raise "Sub or
' Function not defined" the moment the add-in wasn't present - exactly
' the V5.3 problem this whole module was split out to solve, and
' exactly the "export" scenario IN.9 already ties PF.4's own
' performance claim to (this is the heavy-workload case, not just
' interactive add-in use). Duplicating the small amount of read logic
' needed follows this module's own existing R7 precedent (its private
' Fold, duplicated from VLA_Identity for the identical reason). Neither
' function is ListObject-aware or strips a header row - that's a SQL/
' DATALOG-specific convention (a header row isn't a "fact"), not a
' general one; for-each-row reads exactly the physical range it's
' given, matching what a hand-written per-cell loop over that same
' range would have iterated. .Value throughout, not .Value2 - this
' needs to behave exactly like the Cells(i, j).Value loop it replaces
' (real Date/Currency typing preserved), not merely fast. Tests:
' TestArraySlabHelpers (VLA_Tests_Host.bas).
' LX2.0: attempted first, then mostly reverted the same session - 22 of
' this module's 23 raw Err.Raise sites were migrated to
' VLA_Messages.RaiseMsg, then live-caught crashing "Compile and Trace"
' in a fresh workbook ("Variable not defined", VLA_Messages unresolved)
' and reverted back to raw Err.Raise. Root cause: this module's own
' EN_RUNTIME INJECT BOUNDARY discipline (below), already documented and
' already enforced once for VLA_Identity (see Fold's own comment,
' ~line 406) - code above the boundary is copied verbatim into a user's
' workbook as Frazaro_EN_Runtime, which carries NO other add-in module,
' VLA_Messages included. All 23 of this module's raw sites sit above
' the boundary except one (runtime-helpers-unreadable, inside the
' below-boundary injector Sub itself, add-in-side only) - that one
' alone stays migrated. Named functionally here rather than by
' identifier, on purpose: this note sits ABOVE the boundary, and that
' Sub's own name is exactly what the fence check (VLA_Tests_Host.bas's
' TestRuntimeModule) scans injectable text for - spelling it out here
' would fail that check the moment this comment ships, caught live,
' the same session, the same way the file's own older comment already
' warned this exact mistake would (~line 126 by original numbering,
' "a same-session lesson, not a hypothetical" - now twice true).
' R7/rule-12's own "self-contained, no cross-module Private
' calls" discipline already covered this; this pass just didn't check
' it against a module-level Public call before migrating. F.14's own
' ratchet ceiling for this module reflects the reverted count (22), not
' zero - ~line 1204 has the one legitimate exception's own note.
' LX2.1, later session: the 22 sites reverted above now route through a
' second, self-contained mini-catalogue (RaiseRuntimeMsg and its own
' RuntimeCatalogue/RuntimeAddEntries/RuntimeSubstituteSlots, right after
' Fold, ~line 445) duplicated above this boundary rather than calling
' VLA_Messages - the option this entry's prior session named and left
' unbuilt. F.14's ratchet ceiling for this module drops from 22 to 4
' (the mini-catalogue's own internal checks and chokepoint call, same
' shape as VLA_Messages's own 4) - see tools/check_raise_ratchet.ps1.
' GPIVOT.4: the CRUD-completeness survey's own findings, scoped and
' adjudicated with the owner (docs/BETA_REARVIEW.md's G-PIVOT entry has
' the full conversation) - pivot-rename (VlaPivotRename), pivot-clear
' (VlaPivotClear), pivot-source (VlaPivotChangeSource - ChangePivotCache
' over a fresh PivotCache, not a direct PivotCache.SourceData
' reassignment, which real-world reports say is unreliable; auto-
' refreshes, the owner's own call), pivot-sort/pivot-sort-by-value
' (VlaPivotSort, one shared Sub - AutoSort's own Field-parameter
' contract reconciled from Microsoft's self-contradictory docs: a
' field's own SourceName for its own-label sort, a value field's
' rendered CAPTION for a value-field sort, with an explicit raised
' error when a SourceName backs more than one value field at once -
' SalesPivot's own deliberately-unguarded Revenue/Units, owner's own
' call to error rather than silently guess), and pivot-remove-field
' (a 4th "hidden" kind added to the existing VlaPivotSetOrientation,
' the missing inverse of pivot-rows/-columns/-filters, scoped to row/
' column/filter fields only - extending it to value fields raises the
' same ambiguity as sort AND an unverified question of whether
' Orientation=xlHidden can even target one of several same-source data-
' field instances, so it stays filed rather than guessed at).
' pivot-sort's own english.vla dispatch crashed live on its first
' version (run-time error 5, a bare unquoted "descending" symbol where
' VlaPivotSort expected the string "descending") - see the fix's own
' comment in scripts/english.vla, right at the two now-doubled defmacro
' pairs it produced. This Sub's own contract (direction as a plain
' string parameter) never changed; only how english.vla reaches it did.
' GPIVOT.3: pivot-layout/pivot-subtotals-hide/-show/pivot-blank-line-
' add/-remove (scripts/english.vla) - the report-layout knobs closing
' out G-PIVOT, scoped with the owner before writing any code (three
' real forks: which of compact/tabular/outline to expose - all three,
' RowAxisLayout's own full range; whether to widen scope past row-axis
' form - yes, PivotField.Subtotals and PivotField.LayoutBlankLine are
' real, distinct, per-field knobs; and verb-authoring shape). New
' VlaPivotSetRowLayout/VlaPivotSetSubtotals/VlaPivotSetBlankLine, same
' shared-Tier-2-helper-plus-thin-defmacro shape as every other pivot
' verb. RowAxisLayout is the one genuine outlier in this whole family:
' table-WIDE, not per-field, and with NO object-model readback at all
' for its own XlLayoutRowType state (Microsoft's own docs confirm this;
' see that Sub's own header note) - FIRST VERSION wrongly assumed
' PivotField.LayoutForm could substitute, caught live (both a compact
' and an outline check read back the same value): LayoutForm is a
' different, older per-field enum (XlLayoutFormType - xlOutline/
' xlTabular only) entirely unrelated to RowAxisLayout's own parameter
' type. VerifyReportChecks does not attempt a state check for this verb
' at all now - the same honestly-thin "call succeeds" limit
' VlaPivotRefresh's own comment already accepts - and instead checks
' the one real, answerable question: whether RowAxisLayout disturbs an
' already-collapsed field's per-PivotItem ShowDetail state (the same
' Compact-Form interaction that broke VlaPivotSetShowDetail's first
' version) - not yet confirmed either way; awaiting a clean re-run.
' Same pass also
' retrofitted pivot-rows/-columns/-filters and pivot-values-sum/-count/
' -average (renamed from -avg) onto the owner's own {d:...} functor
' authoring convention (scripts/english.vla) - no runtime behavior
' change, source-shape only.
' GPIVOT.2: VlaPivotRefresh/VlaPivotRefreshAll/VlaPivotDelete/
' VlaPivotSetShowDetail (pivot-refresh/-refresh-all/-delete/-collapse/
' -expand, scripts/english.vla). New RequirePivotTableByName factors
' out the lookup-or-raise pair that was about to be duplicated a 5th
' time; VlaPivotSetOrientation/VlaPivotAddValues refactored to use it
' too (no behavior change). VlaPivotRefreshAll is a real, narrower
' alternative to the already-shipped "Refresh everything." - pivots
' only, never an external data connection - built on that technical
' distinction, not on pareto.txt's own say-so. FIRST VERSION of
' VlaPivotSetShowDetail set PivotField.ShowDetail directly - wrong,
' caught live (run-time error 1004) before this version was ever
' committed, not by tracing: CreatePivotTable's own default row layout
' is Compact Form, where PivotField.ShowDetail is unreliable and
' expand/collapse is actually driven per PivotItem - see that Sub's
' own header note for the full reasoning.
' GPIVOT.1: VlaPivotAddValues gains a func parameter ("sum"/"count"/
' "average", same plain-string convention as VlaPivotSetOrientation's
' own "kind") so pivot-values-sum/-count/-avg (scripts/english.vla)
' can share it - pivot-full's own call site updated to pass "sum"
' explicitly, preserving its existing v1 behavior unchanged.
' GPIVOT.0: VlaPivotSetOrientation/VlaPivotAddValues - the field-
' orientation and values half of G-PIVOT's own pivot-rows/-columns/
' -filters/-full (scripts/english.vla, BETA_REARVIEW.md's own G-PIVOT
' entry has the full scoping). One shared helper for row/column/filter
' (kind is a plain string, never the raw XL constant - that mapping
' stays entirely inside this real host VBA), plus a new
' FindPivotTableByName since none of these sentences name a sheet and
' Excel has no Workbook.PivotTables, only a per-worksheet one.
' N1.0: VlaCheckSheetName / VlaCheckRangeName - name-range and
' add-sheet-called both hit Excel's own naming validation (Range.Name /
' Worksheet.Name), which is stricter than VLA's own text slots and
' previously surfaced raw as whatever cryptic message Excel gives
' (1004, "The syntax of this name isn't correct" - the report that
' prompted this). Both macros now check first and refuse in words
' (SD-5/SD-10) with a computed legal alternative, before the real
' assignment ever runs. Uses the module's own private Fold (line ~314),
' not VLA_Identity.Fold - this code is above the inject boundary and
' the injected copy has no VLA_Identity to call.
' LX3.0: the VlaDict fallback's key-folding now uses a private, invariant
' ASCII Fold (duplicated from VLA_Identity per R7 - this module's text
' crosses the injection boundary and must compile alone); the
' post-boundary manifest scan calls VLA_Identity.Fold directly, since
' that code never leaves the add-in. R6/SD-8.
' IO6.0: RT_MODULE was "EN_Runtime", now "Frazaro_EN_Runtime" - the
' name of the module the injector below the boundary writes into a
' user's workbook. The EN_RUNTIME INJECT BOUNDARY marker (the string
' InjectBoundaryMark self-scans for) is unchanged - it is a build-
' internal sentinel, not the module's actual name, and renaming it
' buys nothing. (This note stays ABOVE the boundary, so it must never
' name the injector Sub by identifier - VLA_Tests.bas's
' TestRuntimeModule fails its fence check the moment injectable text
' contains it - a same-session lesson, not a hypothetical.)

' =====================================================================
'  VLA_Runtime - the runtime helper zoo: the procedures a GENERATED
'  program calls (VlaColor, VlaCount, VlaItem, the VlaDict family,
'  VlaFindRow, VlaSendMail, VlaEnsureSheet). Split out of VLA_English
'  at V5.3 for one reason, proved by the first standalone run: a
'  program compiled into the USER's workbook cannot call a Public
'  helper that lives only in the (locked) add-in - VBA has no cross-
'  project call, so "Range(...).Font.Color = vlacolor(hot_pink)"
'  raised "Sub or Function not defined" the moment the dev workbook
'  (where these were in scope) was closed.
'
'  The fix has two halves, both in VLA_Build/VLA_IDE:
'    * the BUILD ships this module inside the add-in (so the add-in's
'      own tooling can still call the helpers), AND
'    * the COMPILE step injects a copy of this module into the user's
'      workbook as Frazaro_EN_Runtime beside the generated
'      Frazaro_EN_<program> module, so the generated code's calls
'      resolve IN-PROJECT.
'  One source of truth (this file); the injected copy is exported
'  from it at build time, never hand-maintained.
'
'  These are self-contained (rule 12): they lean on no Private in
'  any other module, so the exported copy compiles alone.
'
'  S3 adds the MESSAGE SEAM here for the same topology reason the
'  module exists at all: every user-facing message the product speaks
'  goes through VlaShowError, and this is the one module present in
'  EVERY world - beside the generated module in a user workbook
'  (Frazaro_EN_Runtime), inside the add-in, and in the dev workbook - so one
'  function can carry the product's voice everywhere, and the test
'  harness can flip VlaMessageCapture to READ dialogs instead of
'  clicking them (the V5.2 lesson: an untestable dialog talked
'  nonsense for the project's entire history). A USER's own dialogs
'  (the Show/Say sentences) deliberately do NOT pass through the
'  seam - capture must never eat a dialog the program itself asked
'  for.
' =====================================================================

' S3: message-seam state. Session-scoped; capture defaults OFF, so a
' user workbook's Frazaro_EN_Runtime copy behaves exactly like MsgBox unless
' a harness flips it. (Module-level declarations stay above the first
' procedure - the D1.1 rule.)
Private mMsgCaptureOn As Boolean
Private mMsgLog As Collection
Private mTraceOn As Boolean        ' S5: runtime trace - off by default,
Private mTraceLog As Collection    ' so a user Run pays one flag check
                                   ' per step and nothing more
Private mTraceCacheLoaded As Boolean ' S5.1: False after VBA project
                                   ' state loss - the reload signal
Private Const TRACE_NAME As String = "VLAt_TraceOn"
                                   ' S5.2: declared HERE, not beside
                                   ' its machinery - the D1.1 rule's
                                   ' second victim in this project
Private Const PROVENANCE_NAME As String = "VLAt_RunProvenance"
                                   ' the VerifyReports stale-read bug's
                                   ' own fix - see VlaStampRunProvenance
Private mStopArmed As Boolean      ' U.25: the Run button is waiting for
Private mStopped As Boolean        ' a stop - see VlaReportStop
Private mStopStep As String
Private mStopWhat As String

' =====================================================================
'  S3: the message seam. One function carries every product-voiced
'  dialog; the styling (vbExclamation, default title) lives here, in
'  one place, instead of in every call site and every generated
'  program. With capture on, messages append to a log instead of
'  popping - dialog CONTENT becomes a pinnable artifact.
' =====================================================================
Public Sub VlaShowError(ByVal text As String, _
                        Optional ByVal title As String = "Frazaro")
    ' S3.2 (owner catch): the default title is the BRAND, one place -
    ' the first standalone dialog said "VLA English" to a user who
    ' has only ever met Frazaro, because the engine's old title
    ' survived in the seam default while the IDE had long since been
    ' rebranded. The seam is exactly where such drift goes to die.
    If mMsgCaptureOn Then
        If mMsgLog Is Nothing Then Set mMsgLog = New Collection
        mMsgLog.Add title & ": " & text
    Else
        MsgBox text, vbExclamation, title
    End If
End Sub

' The informational sibling: same seam, same capture, gentler icon.
Public Sub VlaShowInfo(ByVal text As String, _
                       Optional ByVal title As String = "Frazaro")
    If mMsgCaptureOn Then
        If mMsgLog Is Nothing Then Set mMsgLog = New Collection
        mMsgLog.Add title & ": " & text
    Else
        MsgBox text, vbInformation, title
    End If
End Sub

' Enabling clears the log (each capture session starts clean);
' disabling restores real dialogs and keeps the log readable.
Public Sub VlaMessageCapture(ByVal onOff As Boolean)
    mMsgCaptureOn = onOff
    If onOff Then Set mMsgLog = New Collection
End Sub

' Everything captured since capture was enabled, one message per
' line, "title: text". Empty string when nothing was captured.
Public Function VlaCapturedMessages() As String
    If mMsgLog Is Nothing Then Exit Function
    Dim r As String
    Dim e As Variant
    For Each e In mMsgLog
        If Len(r) > 0 Then r = r & vbCrLf
        r = r & e
    Next
    VlaCapturedMessages = r
End Function

' U.25: a compiled program's stop. Its vla-report-error calls this with
' the step it stopped at, that step's sentence and line, and what went
' wrong. When the Run button started it, the button armed this first
' (VlaArmStopReport), so the stop is only recorded - the button puts the
' sheets back and then says what happened, once. Run any other way (its
' main from Excel's Macros list, or a standalone copy), nothing is
' waiting, so it is shown here through the seam, as it always was. Only
' the first stop is kept: a step failing inside an action of the
' program's own is reported by that action, and its caller goes on.
Public Sub VlaReportStop(ByVal stepNo As Long, ByVal stepText As String, ByVal what As String)
    If mStopArmed Then
        If Not mStopped Then
            mStopped = True
            mStopStep = stepText
            mStopWhat = what
        End If
    Else
        VlaShowError "Something went wrong at step " & stepNo & ":" & vbCrLf & vbCrLf & _
                     stepText & vbCrLf & vbCrLf & what
    End If
End Sub

' U.25: the Run button arms this before it runs a compiled program and
' disarms it after, reading VlaStopReport in between. Either way the
' record starts empty.
Public Sub VlaArmStopReport(ByVal onOff As Boolean)
    mStopArmed = onOff
    mStopped = False
    mStopStep = ""
    mStopWhat = ""
End Sub

' U.25: "" when the armed run did not stop; otherwise the step's text
' (its sentence and " [line N]"), a line feed, and what went wrong.
Public Function VlaStopReport() As String
    If mStopped Then VlaStopReport = mStopStep & vbLf & mStopWhat
End Function

' =====================================================================
'  S5: the runtime trace. The run was observable at exactly two
'  points - the failing step and the end state - and nothing observed
'  the middle of a successful-but-wrong run. Every tracked step now
'  carries a guarded call: when the trace is OFF (always, unless a
'  harness or a developer flips it) the cost is one Boolean check;
'  when ON, each step logs its number and its ORIGINAL sentence in
'  execution order - loop passes appear once per pass, which is the
'  point. F1's parity harness is the named consumer: two modes, two
'  traces, one diff, and the first divergent step has a name.
'  Capped so a runaway loop cannot eat the session.
' =====================================================================
' S5.1 (owner smoke-test catch): the flag CANNOT live only in module
' state, because the Run itself wipes it - injecting EN_<program>
' into the project that hosts its own run (the dev topology) resets
' every module-level variable in that project before main's first
' step, so a flag flipped in the Immediate window was dead by the
' time the first guard read it. The durable half is a hidden
' workbook NAME - Names are workbook DATA, not VBA state, and
' survive recompiles - and the module cache keeps the per-step cost
' at one Boolean check: state loss clears mTraceCacheLoaded, which
' is precisely the signal to reload from the Name once and re-cache.
' (In the injected topology, Frazaro_EN_Runtime's ThisWorkbook is the
' USER workbook, so tracing a standalone run means flipping ITS copy:
'   Application.Run "'Book1'!Frazaro_EN_Runtime.VlaTrace", True )
' (TRACE_NAME lives in the declarations block at the top of the
'  module - the D1.1 rule, relearned on this pass's maiden compile:
'  "Only comments may appear after End Sub..." at this very spot.)
Public Function VlaTraceOn() As Boolean
    If Not mTraceCacheLoaded Then
        mTraceOn = ReadTraceName()
        mTraceCacheLoaded = True
    End If
    VlaTraceOn = mTraceOn
End Function

' Called by generated code at each step, behind a VlaTraceOn guard.
Public Sub VlaTraceStep(ByVal n As Long, ByVal stepText As String)
    If Not VlaTraceOn() Then Exit Sub
    If mTraceLog Is Nothing Then Set mTraceLog = New Collection
    If mTraceLog.Count >= 20000 Then
        If CStr(mTraceLog.Item(mTraceLog.Count)) <> "(trace capped at 20000 entries)" Then
            mTraceLog.Add "(trace capped at 20000 entries)"
        End If
        Exit Sub
    End If
    mTraceLog.Add Right$("      " & n, 6) & "  " & stepText
End Sub

' Enabling clears (each measurement starts clean); disabling stops
' logging and keeps the log readable.
Public Sub VlaTrace(ByVal onOff As Boolean)
    mTraceOn = onOff
    mTraceCacheLoaded = True
    WriteTraceName onOff
    If onOff Then Set mTraceLog = New Collection
End Sub

Private Function ReadTraceName() As Boolean
    On Error Resume Next
    Dim v As String
    v = CStr(ThisWorkbook.Names(TRACE_NAME).RefersTo)
    On Error GoTo 0
    ReadTraceName = (InStr(v, "1") > 0)
End Function

Private Sub WriteTraceName(ByVal onOff As Boolean)
    On Error Resume Next
    ThisWorkbook.Names(TRACE_NAME).Delete
    On Error GoTo 0
    On Error Resume Next
    ThisWorkbook.Names.Add Name:=TRACE_NAME, _
                           RefersTo:="=" & IIf(onOff, "1", "0"), _
                           Visible:=False
    On Error GoTo 0
End Sub

' Owner-caught live (a multi-round debugging session that turned out to
' be chasing a harness bug, not a rule bug): VerifyReport() only ever
' reads whatever is CURRENTLY on the Output/GStruct sheets - it has no
' way to tell whether that state came from a real emitter Run or was
' left behind by an interpreter Run that happened to run more recently
' (VerifyReportInterpreter deletes and rebuilds those same sheets as
' part of its own live run). The old defense was a comment ("VerifyReport
' must run FIRST") - true, but silently violated the moment anyone runs
' Compile-and-Trace and Interpret-and-Trace in the "wrong" order while
' debugging both backends, which produced a mislabeled, misleading
' report instead of a loud one. Stamped by RunProgram (VLA_IDE.bas,
' "emitter") after a real compiled Run and by VerifyReportInterpreter
' (VLA_Tests_Host.bas, "interpreter") after its own live interpret -
' VerifyReport reads it before trusting the sheet.
Public Sub VlaStampRunProvenance(ByVal tag As String)
    On Error Resume Next
    ThisWorkbook.Names(PROVENANCE_NAME).Delete
    On Error GoTo 0
    On Error Resume Next
    ThisWorkbook.Names.Add Name:=PROVENANCE_NAME, _
                           RefersTo:="=" & Chr$(34) & tag & Chr$(34), _
                           Visible:=False
    On Error GoTo 0
End Sub

Public Function VlaReadRunProvenance() As String
    On Error Resume Next
    Dim v As String
    v = CStr(ThisWorkbook.Names(PROVENANCE_NAME).RefersTo)
    On Error GoTo 0
    ' RefersTo comes back as ="tag" (with the leading = and the quotes) -
    ' strip both rather than assume a fixed length, so an empty/missing
    ' Name (v = "") falls through to an empty result instead of erroring.
    Dim q1 As Long, q2 As Long
    q1 = InStr(v, Chr$(34))
    q2 = InStrRev(v, Chr$(34))
    If q1 > 0 And q2 > q1 Then
        VlaReadRunProvenance = Mid$(v, q1 + 1, q2 - q1 - 1)
    End If
End Function

Public Function VlaTraceReport() As String
    If mTraceLog Is Nothing Then
        VlaTraceReport = "(no trace recorded - VlaTrace True, then Run)"
        Exit Function
    End If
    Dim r As String
    r = "===== RUN TRACE (" & mTraceLog.Count & " step executions) =====" & vbCrLf
    Dim e As Variant
    For Each e In mTraceLog
        r = r & e & vbCrLf
    Next
    VlaTraceReport = r
End Function

' Convert "#RRGGBB" (or "RRGGBB", or a numeric VBA color) to the Long
' VBA expects. Hex color codes read red-first; VBA's color Longs are
' byte-reversed (BGR), so going through RGB() here is what makes
' #FF69B4 come out hot pink instead of powder blue.
' G-FORMAT slice 2: also the eight named colors the phrasebook's
' {c:color} slot accepts (VLA_English.IsColorWord's own list), so
' "Add a border colored red around ..." works - that slot hands its
' word over as text. Checked BEFORE the hex path, and it has to be:
' "yellow" is six letters and would otherwise be read as a malformed
' hex code. Folded invariantly (SD-8), not with LCase$. A widening
' only: every string this used to accept means what it meant, and a
' name used to refuse.
Public Function VlaColor(ByVal v As Variant) As Long
    If VarType(v) = vbString Then
        Dim s As String
        s = Trim$(CStr(v))
        Select Case Fold(s)
            Case "black": VlaColor = vbBlack: Exit Function
            Case "white": VlaColor = vbWhite: Exit Function
            Case "red": VlaColor = vbRed: Exit Function
            Case "green": VlaColor = vbGreen: Exit Function
            Case "blue": VlaColor = vbBlue: Exit Function
            Case "yellow": VlaColor = vbYellow: Exit Function
            Case "magenta": VlaColor = vbMagenta: Exit Function
            Case "cyan": VlaColor = vbCyan: Exit Function
        End Select
        If Left$(s, 1) = "#" Then s = Mid$(s, 2)
        If Len(s) <> 6 Then GoTo bad
        On Error GoTo bad
        VlaColor = RGB(CLng("&H" & Mid$(s, 1, 2)), _
                       CLng("&H" & Mid$(s, 3, 2)), _
                       CLng("&H" & Mid$(s, 5, 2)))
        Exit Function
bad:
        RaiseRuntimeMsg "rt-color-invalid", "value", v
    Else
        VlaColor = CLng(v)
    End If
End Function

' =====================================================================
'  Runtime helpers for the expansion vocabulary (the VlaColor pattern:
'  a small VBA function absorbs an API sharp edge so one clean
'  sentence can exist).
' =====================================================================

' Row of the first cell in a column whose value equals `what`, or 0
' if absent - Range.Find returns Nothing on a miss, which would be an
' instant runtime error in a template; 0 lets programs test for it:
'   Put row of "Widget" in column A into found-row.
'   If found-row is 0, show "not found".
' How many things a value holds. Lists (Collections) answer their own
' Count property; anything else - a range, most commonly - goes through
' Excel's COUNT exactly as "count of" always has (it counts numbers).
' The property-vs-worksheet-function split is absorbed here, once, so
' "count of" means one thing to the person typing it.
Public Function VlaCount(ByVal v As Variant) As Double
    ' V3: Dictionary joins Collection - both spell it .Count, and the
    ' lookup fallback IS a Collection, so one branch serves them all.
    If TypeName(v) = "Collection" Or TypeName(v) = "Dictionary" Then
        VlaCount = v.Count
    Else
        VlaCount = Application.WorksheetFunction.count(v)
    End If
End Function

' B7: element access, Set-vs-Let absorbed (Collection items are
' plain values; Range.Item returns a Range object, which then acts
' as a value through its default property).
Public Function VlaItem(ByVal coll As Variant, ByVal idx As Variant) As Variant
    If IsObject(coll.Item(idx)) Then
        Set VlaItem = coll.Item(idx)
    Else
        VlaItem = coll.Item(idx)
    End If
End Function

Public Function VlaFirst(ByVal coll As Variant) As Variant
    If IsObject(coll.Item(1)) Then
        Set VlaFirst = coll.Item(1)
    Else
        VlaFirst = coll.Item(1)
    End If
End Function

Public Function VlaLast(ByVal coll As Variant) As Variant
    If IsObject(coll.Item(coll.Count)) Then
        Set VlaLast = coll.Item(coll.Count)
    Else
        VlaLast = coll.Item(coll.Count)
    End If
End Function

' =====================================================================
'  V3: keyed memory - the VlaDict helper family. One sharp-edge
'  bundle absorbed behind four small helpers, the VlaColor exchange
'  rate: Scripting.Dictionary where it exists (late-bound, so no
'  reference; case-insensitive keys via TextCompare, matching every
'  text comparison in the language), and a pair-Collection fallback
'  where it doesn't (Mac has no Scripting runtime - the V5 audit's
'  honesty about platforms, paid here in ~30 lines). The fallback
'  stores (display-key, value) pairs keyed by lowercase key, so both
'  representations agree on: add-or-replace semantics, the FIRST
'  spelling of a key surviving replacement, case-insensitive lookup,
'  insertion-ordered keys, .Count, and a loud miss. Set-vs-Let is
'  branched per item (CollGet's lesson, hazard 11), once, here.
' =====================================================================

' Invariant ASCII case fold, DUPLICATED from VLA_Identity rather than
' called cross-module (R7): this function sits above the EN_RUNTIME
' INJECT BOUNDARY below, so its text is copied verbatim into user
' workbooks as Frazaro_EN_Runtime, which has no VLA_Identity to call - exactly
' the "Sub or Function not defined" class VLA_Runtime was split out to
' avoid in the first place (V5.3). Keep in sync with VLA_Identity.Fold
' by hand; it is six lines and does not drift.
Private Function Fold(ByVal s As String) As String
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

' =====================================================================
'  LX2.1: a second, self-contained message catalogue, duplicated above
'  the EN_RUNTIME INJECT BOUNDARY below for the identical reason Fold
'  (above) is duplicated rather than calling VLA_Identity.Fold: code up
'  here ships into a user's workbook as Frazaro_EN_Runtime, which
'  carries no other add-in module, VLA_Messages included (this file's
'  own header note has the live-crash story that first proved it).
'  Same id-plus-named-parameters shape as VLA_Messages.RaiseMsg, same
'  {slotName} substitution syntax, same "rebuilt every call, never
'  cached at module level" reasoning (a VBA recompile drops module
'  state - LESSONS.md XXXII) - but its own table and its own
'  chokepoint, since nothing above this boundary may call anything in
'  VLA_Messages.bas. Ids use an "rt-" prefix so they read as visibly
'  this module's own catalogue at a glance; SD-9's never-reused
'  discipline still applies globally regardless of prefix; no rt- id
'  collides with any existing VLA_Messages id, checked, not assumed.
'  Covers the 22 refusal sites that sit above the boundary (see this
'  file's own header note for that count's history); the one site
'  below it (runtime-helpers-unreadable) already routes through the
'  real VLA_Messages.RaiseMsg, add-in-side only.
' =====================================================================

Private Function RuntimeCatalogue() As Collection
    Dim m As New Collection
    RuntimeAddEntries m
    Set RuntimeCatalogue = m
End Function

Private Sub RuntimeAddEntries(ByVal m As Collection)
    RuntimeAddMsg m, "rt-color-invalid", 5, "VLA-English", "'{value}' is not a color - use one of red, yellow, black, blue, cyan, green, magenta or white, or ""#RRGGBB"", like ""#FF69B4"""
    RuntimeAddMsg m, "rt-dict-key-missing", 5, "VLA-English", "there is nothing stored at key '{key}'"
    RuntimeAddMsg m, "rt-mail-attachment-not-found", 53, "VLA-English", "Attachment not found: {path}"
    RuntimeAddMsg m, "rt-mail-outlook-unavailable", 5, "VLA-English", "Could not start Outlook to create the email"
    RuntimeAddMsg m, "rt-sheet-name-empty", 5, "VLA-Runtime", "a sheet name can't be empty."
    RuntimeAddMsg m, "rt-sheet-name-too-long", 5, "VLA-Runtime", "the sheet name ""{name}"" is {len} characters - sheet names top out at 31."
    RuntimeAddMsg m, "rt-sheet-name-bad-char", 5, "VLA-Runtime", "the sheet name ""{name}"" can't contain '{char}' - sheet names may not use : \ / ? * [ ] characters."
    RuntimeAddMsg m, "rt-sheet-name-apostrophe", 5, "VLA-Runtime", "the sheet name ""{name}"" can't start or end with an apostrophe."
    RuntimeAddMsg m, "rt-sheet-name-reserved-history", 5, "VLA-Runtime", """History"" is reserved by Excel for its own change-tracking sheet - pick another name."
    RuntimeAddMsg m, "rt-sheet-already-exists", 5, "VLA-Runtime", "a sheet called ""{name}"" already exists - pick another name, or use ""Work on sheet {name}"" to reuse it."
    RuntimeAddMsg m, "rt-workbook-name-empty", 5, "VLA-Runtime", "a workbook name can't be empty."
    RuntimeAddMsg m, "rt-workbook-name-too-long", 5, "VLA-Runtime", "the name ""{name}"" is {len} characters - workbook names top out at 255."
    RuntimeAddMsg m, "rt-workbook-name-bad-start", 5, "VLA-Runtime", "the name ""{name}"" has to start with a letter, underscore, or backslash."
    RuntimeAddMsg m, "rt-workbook-name-bad-char", 5, "VLA-Runtime", "the name ""{name}"" can't contain '{char}' - workbook names may only use letters, numbers, periods, and underscores. Try ""{suggestion}"" instead."
    RuntimeAddMsg m, "rt-workbook-name-cell-like", 5, "VLA-Runtime", "the name ""{name}"" looks like a cell reference - Excel won't accept that as a workbook name."
    RuntimeAddMsg m, "rt-pivot-not-found", 5, "VLA-Runtime", "no pivot table named '{name}' in this workbook."
    RuntimeAddMsg m, "rt-pivot-orientation-unknown-kind", 5, "VLA-Runtime", "VlaPivotSetOrientation: unknown kind '{kind}' - expected row, column, filter, or hidden."
    RuntimeAddMsg m, "rt-pivot-values-unknown-function", 5, "VLA-Runtime", "VlaPivotAddValues: unknown function '{func}' - expected sum, count, or average."
    RuntimeAddMsg m, "rt-pivot-row-layout-unknown-kind", 5, "VLA-Runtime", "VlaPivotSetRowLayout: unknown kind '{kind}' - expected compact, tabular, or outline."
    RuntimeAddMsg m, "rt-pivot-sort-unknown-direction", 5, "VLA-Runtime", "VlaPivotSort: unknown direction '{direction}' - expected ascending or descending."
    RuntimeAddMsg m, "rt-pivot-sort-value-field-not-found", 5, "VLA-Runtime", "VlaPivotSort: '{field}' is not a value field in pivot table '{table}'."
    RuntimeAddMsg m, "rt-pivot-sort-value-field-ambiguous", 5, "VLA-Runtime", "VlaPivotSort: '{field}' has more than one aggregation in pivot table '{table}'s values - pivot-sort cannot tell which one to sort by."
    RuntimeAddMsg m, "rt-freeze-panes-nonpositive", 5, "VLA-Runtime", "Freeze needs at least 1 row - got {n}. Use ""Unfreeze the panes."" to remove frozen panes instead."
    RuntimeAddMsg m, "rt-fill-series-unknown-kind", 5, "VLA-Runtime", "VlaFillSeries: unknown kind '{kind}' - expected linear or growth."
    RuntimeAddMsg m, "rt-number-format-decimals", 5, "VLA-Runtime", "'{value}' is not a number of decimal places - use a whole number from 0 to 30, like 2."
    RuntimeAddMsg m, "rt-number-format-unknown-kind", 5, "VLA-Runtime", "VlaNumberFormatCode: unknown kind '{kind}' - expected number, number-separated, percent, dollars, euros, pounds, or accounting- followed by one of those three currencies."
    RuntimeAddMsg m, "rt-column-outside-range", 5, "VLA-Runtime", "column {col} is not one of range {range}'s own columns - name a column inside the range."
    RuntimeAddMsg m, "rt-filters-elsewhere", 5, "VLA-Runtime", "this sheet already has filter buttons on {existing}, and a sheet holds one set - say ""Remove the filters."" first, then add them to {range}."
    RuntimeAddMsg m, "rt-filter-needs-number", 5, "VLA-Runtime", "'{value}' is not a number - ""greater than"" and ""less than"" compare numbers, like 100 or 2.5."
    RuntimeAddMsg m, "rt-filter-unknown-kind", 5, "VLA-Runtime", "VlaFilterCriterion: unknown kind '{kind}' - expected at-least, at-most, contains, greater, or less."
    RuntimeAddMsg m, "rt-filter-range-empty", 5, "VLA-Runtime", "range {range} is empty - filters need a header row with data below it."
    RuntimeAddMsg m, "rt-table-arguments-empty", 5, "VLA-Runtime", "a question needs at least one data table to read - name the Tables the rules mention, like Staff, Shifts, and Leave."
    RuntimeAddMsg m, "rt-table-arguments-bad-name", 5, "VLA-Runtime", "'{name}' cannot name a data table in a formula - a Table name starts with a letter or underscore and continues in letters, digits, underscores and periods, with no spaces."
    RuntimeAddMsg m, "rt-text-unknown-op", 5, "VLA-Runtime", "VlaTextOp: unknown change '{op}' - expected upper, lower, capitalize-after-space, capitalize-after-non-letter, remove-extra-spaces, or remove-non-printing."
    RuntimeAddMsg m, "rt-text-unknown-kind", 5, "VLA-Runtime", "{proc}: unknown kind '{kind}' - expected {expected}."
    RuntimeAddMsg m, "rt-text-not-one-value", 5, "VLA-Runtime", "this needs one value - a piece of text, a number, or one cell - not a list or a range of several cells."
    RuntimeAddMsg m, "rt-text-error-value", 5, "VLA-Runtime", "the value is an error ({value}), not text - fix the cell it came from first."
    RuntimeAddMsg m, "rt-text-marker-empty", 5, "VLA-Runtime", "the text to look for is empty - name a character or a word, like ""-""."
    RuntimeAddMsg m, "rt-text-marker-missing", 5, "VLA-Runtime", "'{marker}' is not in '{text}' - where some values may not have it, check first with ""If ... contains ..."", which looks for it the same way."
    RuntimeAddMsg m, "rt-text-pad-character", 5, "VLA-Runtime", "'{value}' is not one character - pad with a single character, like ""0"" or "" ""."
    RuntimeAddMsg m, "rt-text-pad-width", 5, "VLA-Runtime", "'{value}' is not a number of characters - use a whole number, like 5."
    RuntimeAddMsg m, "rt-join-cell-error", 5, "VLA-Runtime", "cell {cell} holds an error ({value}) - a list is made of values, so fix or clear that cell first."
    RuntimeAddMsg m, "rt-split-separator-empty", 5, "VLA-Runtime", "the text to split at is empty - name a character or a word, like "","" or "" - ""."
    RuntimeAddMsg m, "rt-split-columns-not-empty", 5, "VLA-Runtime", "splitting column {column} needs {need} to its right, {columns}, down to row {row}, and cell {cell} holds something - clear those cells or insert columns first, since a split never writes over anything."
    RuntimeAddMsg m, "rt-split-too-wide", 5, "VLA-Runtime", "splitting column {column} needs {need} to its right, past the sheet's last column."
    ' SEC.15: the formula sink's refusal, raised by VlaSetFormula on both
    ' backends. The formula is shown whole to 80 characters, then cut; the
    ' cell is the target's sheet and address.
    RuntimeAddMsg m, "rt-formula-egress", 5, "VLA-Runtime", "the formula {formula} uses {name}, which can reach outside the workbook on its own (the network, another program, or a link to one), so Frazaro refuses to write it. Nothing was written to {cell}."
End Sub

Private Sub RuntimeAddMsg(ByVal m As Collection, ByVal id As String, ByVal errNum As Long, _
                           ByVal source As String, ByVal template As String)
    On Error GoTo dup
    m.Add Array(errNum, source, template), id
    Exit Sub
dup:
    ' Raw by necessity, not oversight, same reasoning as VLA_Messages's
    ' own AddMsg: the chokepoint that would report this via
    ' RaiseRuntimeMsg is the thing that just failed to build its own
    ' lookup table, so it cannot depend on itself here.
    Err.Raise 5, "VLA-Runtime", "message id '" & id & "' is already registered - ids are never reused (SD-9's discipline); pick a different one"
End Sub

' The one chokepoint this module's above-boundary refusals call.
' ParamArray copied into a plain Variant before forwarding, same VBA
' workaround VLA_Messages.RaiseMsg already documents ("Invalid
' ParamArray use" otherwise, a real language rule).
Private Sub RaiseRuntimeMsg(ByVal id As String, ParamArray kv() As Variant)
    Dim cat As Collection
    Set cat = RuntimeCatalogue()
    Dim rec As Variant
    On Error GoTo unknown
    rec = cat.Item(id)
    On Error GoTo 0
    Dim kvArr As Variant
    kvArr = kv
    Err.Raise CLng(rec(0)), CStr(rec(1)), RuntimeSubstituteSlots(CStr(rec(2)), kvArr)
    Exit Sub
unknown:
    Err.Raise 5, "VLA-Runtime", "RaiseRuntimeMsg: unknown message id '" & id & "'"
End Sub

Private Function RuntimeSubstituteSlots(ByVal tmpl As String, ByRef kv As Variant) As String
    Dim r As String, i As Long, openPos As Long, closePos As Long
    i = 1
    Do While i <= Len(tmpl)
        openPos = InStr(i, tmpl, "{")
        If openPos = 0 Then
            r = r & Mid$(tmpl, i)
            Exit Do
        End If
        closePos = InStr(openPos, tmpl, "}")
        If closePos = 0 Then
            r = r & Mid$(tmpl, i)
            Exit Do
        End If
        r = r & Mid$(tmpl, i, openPos - i)
        r = r & CStr(RuntimeSlotValue(Mid$(tmpl, openPos + 1, closePos - openPos - 1), kv))
        i = closePos + 1
    Loop
    RuntimeSubstituteSlots = r
End Function

Private Function RuntimeSlotValue(ByVal slotName As String, ByRef kv As Variant) As Variant
    Dim i As Long
    For i = LBound(kv) To UBound(kv) Step 2
        If CStr(kv(i)) = slotName Then
            RuntimeSlotValue = kv(i + 1)
            Exit Function
        End If
    Next i
    Err.Raise 5, "VLA-Runtime", "RaiseRuntimeMsg: no value supplied for slot '{" & slotName & "}'"
End Function

Public Function VlaDictNew() As Object
    Dim d As Object
    On Error Resume Next
    Set d = CreateObject("Scripting.Dictionary")
    On Error GoTo 0
    If Not d Is Nothing Then
        d.CompareMode = vbTextCompare
        Set VlaDictNew = d
        Exit Function
    End If
    Set VlaDictNew = New Collection
End Function

' DATALOG.13: every wrapper in this family asks VlaDictIsFallback(d)
' and never TypeName(d). TypeName on a late-bound COM object asks it
' for its type information, measured at 0.148 ms a call (as much as
' creating the dictionary), and a DATALOG row made about a dozen of
' these calls: 85% of its cost. TypeOf on a VBA Collection and Is
' Nothing each measure 0.000 ms. The fallback is the tested branch
' because a Collection here means the host had no Scripting runtime.
' Nothing keeps the fallback it always had: VlaDictHas answers False
' and VlaDictGet misses loudly, which VLA_Interpreter's module-scope
' read relies on (its And evaluates VlaDictHas even when mModuleFrame
' is Nothing).
' tools/check_vladict_guard.ps1 holds TypeName out of these wrappers.
'
' Two tests, never one line: TypeOf RAISES error 91 on Nothing rather
' than answering False, and VBA's Or evaluates both operands, so
' "d Is Nothing Or TypeOf d Is Collection" raised on exactly the
' Nothing it was written to admit - caught by the owner's live pass,
' 2026-09-18, eval of an unbound name in a fresh session.
Private Function VlaDictIsFallback(ByVal d As Object) As Boolean
    If d Is Nothing Then
        VlaDictIsFallback = True
    Else
        VlaDictIsFallback = TypeOf d Is Collection
    End If
End Function

Public Sub VlaDictSet(ByVal d As Object, ByVal k As Variant, ByVal v As Variant)
    Dim ks As String
    ks = CStr(k)
    If VlaDictIsFallback(d) Then
        ' Fallback: replace = remove + re-add, keeping the FIRST
        ' spelling from the old pair so the two representations
        ' cannot be told apart from the sentence side.
        Dim old As Collection
        Dim disp As String
        disp = ks
        On Error Resume Next
        Set old = d.Item(Fold(ks))
        On Error GoTo 0
        If Not old Is Nothing Then
            disp = CStr(old.Item(1))
            d.Remove Fold(ks)
        End If
        Dim pair As New Collection
        pair.Add disp
        pair.Add v
        d.Add pair, Fold(ks)
    Else
        If IsObject(v) Then
            Set d.Item(ks) = v
        Else
            d.Item(ks) = v           ' add-or-replace in one move;
                                     ' TextCompare keeps the first
                                     ' spelling of a matched key
        End If
    End If
End Sub

Public Function VlaDictGet(ByVal d As Object, ByVal k As Variant) As Variant
    Dim ks As String
    ks = CStr(k)
    Dim hit As Boolean
    If VlaDictIsFallback(d) Then
        Dim pair As Collection
        On Error Resume Next
        Set pair = d.Item(Fold(ks))
        On Error GoTo 0
        hit = Not pair Is Nothing
        If hit Then
            If IsObject(pair.Item(2)) Then
                Set VlaDictGet = pair.Item(2)
            Else
                VlaDictGet = pair.Item(2)
            End If
        End If
    Else
        hit = d.Exists(ks)
        If hit Then
            If IsObject(d.Item(ks)) Then
                Set VlaDictGet = d.Item(ks)
            Else
                VlaDictGet = d.Item(ks)
            End If
        End If
    End If
    If Not hit Then
        ' Loud miss at the step, naming the key - the form-79 lookup
        ' precedent; the step dialog names the sentence.
        RaiseRuntimeMsg "rt-dict-key-missing", "key", ks
    End If
End Function

' IN.3: a non-raising existence check, the same hit-detection VlaDictGet
' already does internally, duplicated rather than factored out (VlaDictGet
' would need a ByRef "suppress the raise" flag threaded through every
' existing call site otherwise). Needed so a caller can test "is this key
' bound HERE" before deciding whether to fall back to a different dict -
' the interpreter's module-scope read fallback (VLA_Interpreter.bas) is
' the first caller.
Public Function VlaDictHas(ByVal d As Object, ByVal k As Variant) As Boolean
    Dim ks As String
    ks = CStr(k)
    If VlaDictIsFallback(d) Then
        Dim pair As Collection
        On Error Resume Next
        Set pair = d.Item(Fold(ks))
        On Error GoTo 0
        VlaDictHas = Not pair Is Nothing
    Else
        VlaDictHas = d.Exists(ks)
    End If
End Function

Public Function VlaDictKeys(ByVal d As Object) As Collection
    ' Keys as a plain Collection - identical walk order (insertion)
    ' under both representations, so "For each k in keys of prices"
    ' is deterministic everywhere.
    Dim r As New Collection
    Dim e As Variant
    If VlaDictIsFallback(d) Then
        For Each e In d
            r.Add e.Item(1)
        Next
    Else
        For Each e In d.Keys
            r.Add e
        Next
    End If
    Set VlaDictKeys = r
End Function

' V3.1: the entries as (key, value) pairs - 2-element Collections -
' for "For each pair in <lookup>". Identical under both
' representations: insertion order, and values as of the walk's
' start (a Store during the walk replaces the ENTRY in the
' lookup, but pairs already dealt keep what they held - the two
' branches agree because replacement builds NEW pairs in the
' fallback while the Dictionary branch copies at walk time).
Public Function VlaDictPairs(ByVal d As Object) As Collection
    Dim r As New Collection
    Dim e As Variant
    Dim pair As Collection
    If VlaDictIsFallback(d) Then
        For Each e In d
            r.Add e                  ' the fallback's entries ARE pairs
        Next
    Else
        For Each e In d.Keys
            Set pair = New Collection
            pair.Add e
            pair.Add d.Item(e)
            r.Add pair
        Next
    End If
    Set VlaDictPairs = r
End Function

Public Function VlaPairKey(ByVal p As Variant) As Variant
    VlaPairKey = p.Item(1)           ' keys are always text - Let is right
End Function

Public Function VlaPairValue(ByVal p As Variant) As Variant
    If IsObject(p.Item(2)) Then      ' Set-vs-Let, branched per item
        Set VlaPairValue = p.Item(2)
    Else
        VlaPairValue = p.Item(2)
    End If
End Function

' PF.4b: the bulk-read half of PF.4's array-slab runtime helpers -
' PF.4c's own for-each-row will read its whole range once through this
' instead of one Cells(i, j) COM call per cell, the same shape PF.7's
' VlaEmbeddedText fix (VLA.bas) already proved out for a sheet-column
' read. Deliberately NOT a call to VLA_Relation.SourceToArray, despite
' that function already doing almost exactly this for SQL/DATALOG: this
' module's code above the EN_RUNTIME INJECT BOUNDARY (below) is copied
' verbatim into a user's own workbook with NO other add-in module
' available - VLA_Relation would not exist there, exactly the V5.3
' problem this whole module was split out to solve in the first place,
' and exactly the scenario IN.9 already ties PF.4's own performance
' claim to (the export, not just interactive Frazaro-add-in use). Self-
' contained instead, the same R7 "duplicate across the injection
' boundary" precedent this module's own Fold already uses. Also
' deliberately NOT ListObject-aware and does not strip a header row -
' SourceToArray's own table-header-stripping is a SQL/DATALOG-specific
' convention (a header row isn't a "fact"), not a general one; for-
' each-row reads exactly the physical range it's given, matching what
' a hand-written per-cell loop over that same range would have
' iterated, byte-for-byte - no surprise substitution of what the loop
' actually walks. .Value, not .Value2 (VLA_Relation's own choice) -
' this needs to behave exactly like the Cells(i, j).Value loop it
' replaces (real Date/Currency typing preserved), not merely fast; the
' 1-cell scalar quirk (Range.Value collapses to a bare scalar on a
' single-cell range) is wrapped before use, same fix as VlaEmbeddedText's.
Public Function VlaSlabRead(ByVal src As Range) As Variant
    If src.Cells.Count = 1 Then
        Dim one(1 To 1, 1 To 1) As Variant
        one(1, 1) = src.Value
        VlaSlabRead = one
    Else
        VlaSlabRead = src.Value
    End If
End Function

' PF.4b: the write-back half - one bulk write regardless of row count,
' the other side of VlaSlabRead's own saving. Resizes FROM destRange's
' own top-left corner to match arr's own dimensions exactly, rather
' than trusting the caller to have already sized destRange correctly -
' for PF.4c's in-place-update case (write back to the same range just
' read) this is a same-size no-op resize, but it costs nothing and
' removes any ambiguity for every caller, not just that one. No 1-cell
' special case needed on this side - unlike reading, assigning a 2D
' array (even a 1x1 one) to a range whose cell count matches its
' element count is unambiguous; only READING a single cell collapses
' the dimension, writing never does.
Public Sub VlaSlabWrite(ByRef arr As Variant, ByVal destRange As Range)
    Dim nRows As Long, nCols As Long
    nRows = UBound(arr, 1) - LBound(arr, 1) + 1
    nCols = UBound(arr, 2) - LBound(arr, 2) + 1
    destRange.Resize(nRows, nCols).Value = arr
End Sub

Public Function VlaFindRow(ByVal what As Variant, ByVal columnRef As String) As Long
    Dim hit As Range
    On Error GoTo missing
    Set hit = ActiveSheet.Columns(columnRef).Find( _
        What:=what, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If hit Is Nothing Then Exit Function
    VlaFindRow = hit.Row
    Exit Function
missing:
    VlaFindRow = 0
End Function

' Create an email in Outlook and DISPLAY it as a draft - deliberately
' not sent: the person reviews and clicks Send themselves, which is
' both safer and avoids Outlook's programmatic-send security prompts.
Public Sub VlaSendMail(ByVal toAddr As Variant, ByVal subject As Variant, _
                       ByVal body As Variant, Optional ByVal attachPath As Variant = "")
    ' SEC.8: gated on the provenance of the workbook carrying the
    ' program. The draft is displayed, never sent (see the comment
    ' above), so a human click still stands between this and any mail
    ' leaving - but the draft arrives pre-addressed, pre-written, and
    ' able to attach ANY local file the program can name, which is
    ' exactly the shape of an exfiltration primitive that only needs
    ' one careless click. Guarded at the top of this Sub rather than at
    ' the dispatch site because this is a VLA_Runtime helper reachable
    ' by name, so the Sub itself is the only true chokepoint.
    VLA_Provenance.VlaProvenanceGuardCaptured "email something out of Excel"

    Dim ol As Object, m As Object
    On Error Resume Next
    Set ol = GetObject(, "Outlook.Application")
    On Error GoTo 0
    If ol Is Nothing Then
        On Error GoTo noOutlook
        Set ol = CreateObject("Outlook.Application")
        On Error GoTo 0
    End If
    Set m = ol.CreateItem(0)                  ' 0 = olMailItem
    m.To = CStr(toAddr)
    m.Subject = CStr(subject)
    m.Body = CStr(body)
    If Len(CStr(attachPath)) > 0 Then
        If Len(Dir$(CStr(attachPath))) > 0 Then
            m.Attachments.Add CStr(attachPath)
        Else
            RaiseRuntimeMsg "rt-mail-attachment-not-found", "path", attachPath
        End If
    End If
    m.Display
    Exit Sub
noOutlook:
    RaiseRuntimeMsg "rt-mail-outlook-unavailable"
End Sub

' Ensure a sheet exists (create at the end if missing) - the runtime
' half of "Work on sheet X.", which means "this is my workspace":
' create if needed, then go. (Contrast "Go to sheet X.", which fails
' if the sheet is missing - navigation, not declaration.)
Public Sub VlaEnsureSheet(ByVal name As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets(name)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ActiveWorkbook.Worksheets.Add( _
            After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
        ws.Name = name
    End If
End Sub

' SD-5/SD-10: refuse in words, naming the reason and the alternative -
' never let Excel's own naming errors (1004, "The syntax of this name
' isn't correct") reach a user raw. name-range and add-sheet-called
' both call the matching check below BEFORE their real .Name
' assignment, so a bad name fails as an ordinary step - the same
' "Excel says: ..." wording every other refusal gets, just with
' Frazaro's own reason in place of Excel's opaque one - instead of
' whatever VBA's default error dialog happens to say. Pure string
' checks, no live Excel object needed, so they're deterministic and
' unit-testable without a workbook (TestHelpers does both).
'
' Sheet names: <=31 chars, none of : \ / ? * [ ], no leading/trailing
' apostrophe, not the reserved word "History". Duplicate-name checking
' is deliberately NOT here - that is workbook state, a different
' failure class from pure syntax, and Excel's own message for it is
' already reasonably clear.
Public Sub VlaCheckSheetName(ByVal nm As String)
    If Len(nm) = 0 Then
        RaiseRuntimeMsg "rt-sheet-name-empty"
    End If
    If Len(nm) > 31 Then
        RaiseRuntimeMsg "rt-sheet-name-too-long", "name", nm, "len", Len(nm)
    End If
    Dim bad As String, i As Long
    bad = ":\/?*[]"
    For i = 1 To Len(bad)
        If InStr(nm, Mid$(bad, i, 1)) > 0 Then
            RaiseRuntimeMsg "rt-sheet-name-bad-char", "name", nm, "char", Mid$(bad, i, 1)
        End If
    Next
    If Left$(nm, 1) = "'" Or Right$(nm, 1) = "'" Then
        RaiseRuntimeMsg "rt-sheet-name-apostrophe", "name", nm
    End If
    If Fold(nm) = "history" Then
        RaiseRuntimeMsg "rt-sheet-name-reserved-history"
    End If
End Sub

' IN.12: add-sheet-called's pre-flight duplicate check. Create-only means
' a duplicate name must fail BEFORE Worksheets.Add runs, not after - the
' macro used to add the sheet first and only discover the duplicate at
' the rename step (Set ActiveSheet.Name = s), by which point Excel had
' already created it under its own default name (Sheet2, Sheet3, ...).
' On a workbook rerun - exactly what instructions.txt's own Try: block around
' "Add sheet called ..." exists to tolerate - that left one unnamed,
' unclaimed sheet behind every time, forever, which is where the stray
' Sheet2xx/Sheet3xx sheets (IN.12) came from. Checking existence first
' means the duplicate is refused before anything is created, so a rerun
' is safe with nothing left over. Separate from VlaCheckSheetName on
' purpose, same reasoning as that function's own comment: this is
' workbook state, not name syntax, and needs a live workbook to answer.
Public Sub VlaCheckSheetAbsent(ByVal nm As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets(nm)
    On Error GoTo 0
    If Not ws Is Nothing Then
        RaiseRuntimeMsg "rt-sheet-already-exists", "name", nm
    End If
End Sub

' Workbook (defined) names: start with a letter, underscore, or
' backslash; body letters/digits/periods/underscores only (no hyphens,
' no spaces); <=255 chars; can't read as a cell reference (Excel
' rejects "A1"-shaped names outright). The message computes and offers
' a legal alternative - illegal characters replaced with underscores -
' so the fix is one copy-paste away, without silently substituting it
' for the user (SD-10: never diverge from what they wrote without
' saying so).
Public Sub VlaCheckRangeName(ByVal nm As String)
    If Len(nm) = 0 Then
        RaiseRuntimeMsg "rt-workbook-name-empty"
    End If
    If Len(nm) > 255 Then
        RaiseRuntimeMsg "rt-workbook-name-too-long", "name", nm, "len", Len(nm)
    End If
    Dim first As String
    first = Left$(nm, 1)
    If Not (first Like "[A-Za-z]" Or first = "_" Or first = "\") Then
        RaiseRuntimeMsg "rt-workbook-name-bad-start", "name", nm
    End If
    Dim i As Long, c As String, badChar As String, suggestion As String
    suggestion = nm
    For i = 1 To Len(nm)
        c = Mid$(nm, i, 1)
        If Not (c Like "[A-Za-z0-9._]") Then
            If Len(badChar) = 0 Then badChar = c
            Mid$(suggestion, i, 1) = "_"
        End If
    Next
    If Len(badChar) > 0 Then
        RaiseRuntimeMsg "rt-workbook-name-bad-char", "name", nm, "char", badChar, "suggestion", suggestion
    End If
    If IsCellLikeName(nm) Then
        RaiseRuntimeMsg "rt-workbook-name-cell-like", "name", nm
    End If
End Sub

' A1-shaped text (1-3 letters then 1+ digits, nothing else) - the
' shape Excel itself refuses as a defined name, regardless of whether
' the letters form a real column.
Private Function IsCellLikeName(ByVal nm As String) As Boolean
    Dim i As Long, n As Long
    n = Len(nm)
    i = 1
    Do While i <= n And Mid$(nm, i, 1) Like "[A-Za-z]"
        i = i + 1
    Loop
    If i = 1 Or i > 4 Then Exit Function
    Dim digitsStart As Long
    digitsStart = i
    Do While i <= n And Mid$(nm, i, 1) Like "[0-9]"
        i = i + 1
    Loop
    IsCellLikeName = (i = n + 1 And i > digitsStart)
End Function

' G-STRUCT (pareto.txt sections 2-5): the marked (!) verbs whose COM
' sequence is too stateful or too index-sensitive to reduce to one
' defmacro expression - move-column and banded-rows genuinely need a
' real helper; copy-style/copy-formats and header-style turned out not
' to (see scripts/english.vla's own notes at each rule, the same
' "the marker described something else" correction G-PIVOT's own
' pivot-create entry already made once).

' move-column: Cut + Insert(Shift:=xlToRight) is a single atomic Excel
' operation - the pending-cut Insert removes the source column AND
' closes the gap it leaves behind in one call, so no manual index-shift
' arithmetic is needed regardless of whether destCol sits before or
' after srcCol (pareto.txt's own "Cut + Insert (helper: keeps formats)"
' target, and Insert-with-a-pending-cut is what keeps the formats -
' unlike Copy+PasteSpecial, nothing has to be reassembled). Self-move
' (srcCol = destCol) is a no-op, not a refusal - "move column C before
' column C" is a degenerate instruction, not an error.
Public Sub VlaMoveColumn(ByVal srcCol As String, ByVal destCol As String)
    If UCase$(srcCol) = UCase$(destCol) Then Exit Sub
    Columns(srcCol).Cut
    Columns(destCol).Insert Shift:=xlToRight
End Sub

' freeze-panes: "Freeze the first {n:expr} rows." follows the active
' window exactly like the already-shipped freeze-top-row does (select
' Rows(n+1), set FreezePanes) - the only guard is n <= 0, a degenerate
' "freeze zero rows" instruction that would otherwise silently select
' row 1 and do nothing useful; "Unfreeze the panes." already exists as
' the honest way to remove frozen panes, so the message points there
' instead of guessing what the writer meant.
Public Sub VlaFreezePanes(ByVal n As Long)
    If n <= 0 Then
        RaiseRuntimeMsg "rt-freeze-panes-nonpositive", "n", CStr(n)
    End If
    Rows(n + 1).Select
    ActiveWindow.FreezePanes = True
End Sub

' delete-blank-rows: SpecialCells(xlCellTypeBlanks) raises a real Excel
' error (1004, "No cells were found") when the range has no blank cells
' at all - pareto.txt's own marker note ("helper: guard the no-blanks
' error"). Guarded to a silent no-op instead: "there was nothing to
' delete" is not a failure worth a teaching refusal. EntireRow.Delete on
' the (possibly multi-area) blanks range deletes every matched row in
' one call - Excel resolves a multi-area delete as a single operation,
' so this does not hit the reverse-iteration hazard a hand-written
' row-by-row loop would (pareto.txt's own delete-matching-rows entry
' names that exact bug).
Public Sub VlaDeleteBlankRows(ByVal rng As Range)
    Dim blanks As Range
    On Error Resume Next
    Set blanks = rng.SpecialCells(xlCellTypeBlanks)
    On Error GoTo 0
    If blanks Is Nothing Then Exit Sub
    blanks.EntireRow.Delete
End Sub

' trim-sheet: shrink the active sheet to its real used extent. Find's
' own SearchDirection:=xlPrevious walking backward from A1 is the
' standard way to find the true last non-blank cell (UsedRange itself
' is unreliable - it remembers cells that were ever formatted, even
' after their contents are cleared). A completely empty sheet (Find
' returns Nothing) is a no-op, not a refusal.
'
' Owner-caught live, 2026-08-28 (bisect_j/k, TER-4): a plain
' Find(What:="*", LookIn:=xlFormulas) only matches cell CONTENT - a
' row/column with nothing but formatting on it (bold, filled, no
' value) was invisible to it, so on a sparsely-valued sheet the last
' row/column with real CONTENT could sit well before the last row/
' column with real FORMATTING, and this Sub would then delete the
' formatted-but-content-empty row/column whole, mistaking it for
' trailing-empty. A single Find by row order was also unreliable for
' the COLUMN bound specifically (it returns the cell with the highest
' ROW number, whose own column isn't necessarily the sheet's true
' rightmost - a separate, pre-existing latent bug in the same family).
' Fixed below: lastRow/lastCol are each the max across a row-order AND
' a column-order content Find, plus format-only Finds
' (SearchFormat:=True) covering VLA's own Tier-2 formatting vocabulary
' - bold and interior fill. This is NOT a fully generic "any
' formatting" detector - Excel's SearchFormat matches a specific
' format, it has no "anything non-default" wildcard - so a future
' formatting-only rule outside bold/fill (e.g. a font-color-only or
' border-only cell with no bold, no fill, no content) would need a
' matching FindFormat addition here, the same narrow/hand-maintained
' pattern as ResolveExcelConstant and DynamicNamedCall elsewhere in
' this codebase.
'
' UNRESOLVED, SEPARATELY - see docs/BETA_REARVIEW.md's TERRARIUM TER-4
' for the full incident, including a correction logged the same
' night. This Sub's own EntireColumn.Delete loop, specifically,
' reproducibly reverted PasteSpecial/Cut-mediated formatting elsewhere
' on the sheet, in a program that also calls move-column ("Move column
' B before column D."). Ruled out, each confirmed by direct test: the
' sheet's absolute physical edge (a bounded delete margin still
' reproduced it), Application.CutCopyMode state (explicit clear made no
' difference), bulk-vs-one-at-a-time deletion (row-by-row still
' reproduced it), Find itself (disabling both delete loops, Find still
' running, fixed it - so Find is innocent of THIS bug specifically),
' the row-delete loop (never reproduces it alone), and execution ORDER
' relative to move-column (reordering move-column to run after this
' Sub instead of before did NOT fix it). A follow-up bisection round
' (removing one candidate at a time from hide/unhide, group/ungroup,
' row insert/delete, move-column, freeze-panes, each alone) also never
' fixed it - but that round, and the "remove all five" control it was
' checked against, both ran against the OLD content-only Find above
' and may themselves have been confounded by the very bug this note
' just fixed (a sparse control sheet silently skipping the delete loop
' entirely looks identical to "the bug didn't fire") - a corrected
' re-run (`scripts/bisect_*2_*.txt`) is in progress as of this note.
' EntireRow/EntireColumn.Delete is not inherently unsafe -
' delete-blank-rows's own blanks.EntireRow.Delete (above) never shows
' this. No known fix or workaround for THIS bug as of this note - the
' Sub below should still be treated as carrying a live, unresolved
' correctness risk on its EntireColumn.Delete loop, even though the
' Find bound-detection bug is now fixed.
Public Sub VlaTrimSheet()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim lastRow As Long, lastCol As Long
    lastRow = 0
    lastCol = 0

    VlaTrimExtendBounds lastRow, lastCol, ws.Cells.Find(What:="*", _
        After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, _
        SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    VlaTrimExtendBounds lastRow, lastCol, ws.Cells.Find(What:="*", _
        After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, _
        SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)

    Application.FindFormat.Clear
    Application.FindFormat.Font.Bold = True
    VlaTrimExtendBounds lastRow, lastCol, ws.Cells.Find(What:="", _
        After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, _
        SearchOrder:=xlByRows, SearchDirection:=xlPrevious, SearchFormat:=True)
    VlaTrimExtendBounds lastRow, lastCol, ws.Cells.Find(What:="", _
        After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, _
        SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, SearchFormat:=True)

    Application.FindFormat.Clear
    Application.FindFormat.Interior.Pattern = xlSolid
    VlaTrimExtendBounds lastRow, lastCol, ws.Cells.Find(What:="", _
        After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, _
        SearchOrder:=xlByRows, SearchDirection:=xlPrevious, SearchFormat:=True)
    VlaTrimExtendBounds lastRow, lastCol, ws.Cells.Find(What:="", _
        After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, _
        SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, SearchFormat:=True)
    Application.FindFormat.Clear

    If lastRow = 0 And lastCol = 0 Then Exit Sub
    If lastRow = 0 Then lastRow = 1
    If lastCol = 0 Then lastCol = 1

    Dim rowCap As Long, colCap As Long
    rowCap = lastRow + 2000
    If rowCap > ws.Rows.Count Then rowCap = ws.Rows.Count
    colCap = lastCol + 200
    If colCap > ws.Columns.Count Then colCap = ws.Columns.Count
    Dim r As Long, c As Long
    For r = rowCap To lastRow + 1 Step -1
        ws.Rows(r).Delete
    Next r
    For c = colCap To lastCol + 1 Step -1
        ws.Columns(c).Delete
    Next c
End Sub

' Shared by VlaTrimSheet's several Find calls (content, row-order and
' column-order; format-only, row-order and column-order): widens
' lastRow/lastCol to cover whatever cell was found, or does nothing if
' this particular Find came back empty.
Private Sub VlaTrimExtendBounds(ByRef lastRow As Long, ByRef lastCol As Long, ByVal found As Range)
    If found Is Nothing Then Exit Sub
    If found.Row > lastRow Then lastRow = found.Row
    If found.Column > lastCol Then lastCol = found.Column
End Sub

' banded-rows: a static, one-time loop (owner's call over a live
' Conditional Formatting rule - no cf-* machinery exists yet, pareto.txt
' section 7 is a later, unstarted pass, and a static fill is what
' clear-formats/clear-all already know how to undo). Shades every
' OTHER row starting at the range's own 2nd row, not its 1st - the
' common zebra-stripe convention that leaves row 1 (often a header)
' unbanded; documented here since pareto.txt's own surface doesn't say
' which parity to pick.
Public Sub VlaBandRows(ByVal rng As Range, ByVal colorVal As Long)
    Dim i As Long
    For i = 2 To rng.Rows.Count Step 2
        rng.Rows(i).Interior.Color = colorVal
    Next i
End Sub

' fill-series: DataSeries has no "starting value" parameter of its own
' - the value has to already be sitting in the range's first cell
' before DataSeries runs, so "starting at {n:expr}" means writing n
' there first. kind is a plain runtime string, not a template-glued
' identifier (the same VlaPivotSetOrientation kind-dispatch shape
' above) because scripts/english.vla reaches this from four separate
' rules (linear/growth x with-step/without), never from a captured
' {d:...} slot.
Public Sub VlaFillSeries(ByVal rng As Range, ByVal startVal As Double, _
                          ByVal stepVal As Double, ByVal kind As String)
    Dim t As XlDataSeriesType
    Select Case kind
        Case "linear": t = xlLinear
        Case "growth": t = xlGrowth
        Case Else
            RaiseRuntimeMsg "rt-fill-series-unknown-kind", "kind", kind
    End Select
    rng.Item(1).Value = startVal
    rng.DataSeries Type:=t, Step:=stepVal
End Sub

' G-FORMAT slice 2 (pareto.txt section 6) and EN.3's policy, in one
' place: a named number format from english.vla's "Format ... as ..."
' sentences, turned into an Excel format code. The policy is the
' owner's (2026-09-10): the SYMBOL follows the data, the ORDER follows
' the reader. A currency is named in the sentence - dollars, euros,
' pounds - because it is a fact about the numbers: dollar figures
' opened on a German laptop are still dollars. Separators need nothing
' here: Range.NumberFormat always reads US-syntax codes, and Excel
' draws "," and "." from the reading machine's own settings when it
' displays them. Dates and times are not built here at all -
' english.vla's date and time macros set Excel's own system date and
' time codes, which follow the reading machine by design.
'
' Built here, not written literally in each macro, for two reasons:
' the decimal count is a runtime value ("with {n} decimals"), and the
' euro and pound signs are made with ChrW rather than typed - a
' compiled program's text passes through the machine's ANSI code page
' on its way into the VBA project, where a non-Western code page would
' mangle a literal euro or pound sign before Excel ever saw it.
'
' decimals: a whole number from 0 to 30 (Excel's own ceiling for a
' format code). Anything else refuses by name here, rather than
' reaching Excel, whose own error names neither the sentence's number
' nor the fix. Empty and True are refused too, although VBA calls both
' numeric - an unset value is not a decimal count.
Public Function VlaNumberFormatCode(ByVal kind As String, ByVal decimals As Variant) As String
    If IsObject(decimals) Then GoTo badDecimals
    If IsEmpty(decimals) Or IsNull(decimals) Then GoTo badDecimals
    If VarType(decimals) = vbBoolean Then GoTo badDecimals
    If Not IsNumeric(decimals) Then GoTo badDecimals
    Dim n As Double
    n = CDbl(decimals)
    If n <> Fix(n) Or n < 0 Or n > 30 Then GoTo badDecimals

    Dim places As String
    If n > 0 Then places = "." & String$(CLng(n), "0")

    Select Case kind
        Case "number": VlaNumberFormatCode = "0" & places
        Case "number-separated": VlaNumberFormatCode = "#,##0" & places
        Case "percent": VlaNumberFormatCode = "0" & places & "%"
        Case "dollars", "euros", "pounds"
            VlaNumberFormatCode = CurrencySymbolCode(kind) & "#,##0" & places
        ' Excel's own accounting layout: the sign pinned left ("* "
        ' fills the gap), negatives in brackets, zero as a dash with
        ' room for the decimals ("?" per place), text left alone.
        Case "accounting-dollars", "accounting-euros", "accounting-pounds"
            Dim sym As String
            sym = CurrencySymbolCode(Mid$(kind, Len("accounting-") + 1))
            VlaNumberFormatCode = "_(" & sym & "* #,##0" & places & "_);" & _
                                  "_(" & sym & "* (#,##0" & places & ");" & _
                                  "_(" & sym & "* ""-""" & String$(CLng(n), "?") & "_);" & _
                                  "_(@_)"
        Case Else
            RaiseRuntimeMsg "rt-number-format-unknown-kind", "kind", kind
    End Select
    Exit Function
badDecimals:
    Dim shown As String
    If IsObject(decimals) Then
        shown = "that"
    ElseIf IsEmpty(decimals) Or IsNull(decimals) Then
        shown = "(nothing)"
    Else
        shown = CStr(decimals)
    End If
    RaiseRuntimeMsg "rt-number-format-decimals", "value", shown
End Function

' The piece of a format code that shows a currency's sign. Every sign
' is escaped with a backslash (Excel's "show the next character as it
' is"), the dollar included: a bare "$" in a format code is NOT a
' literal - Excel reads it as the READING machine's own currency
' symbol, so "$#,##0.00" showed "£1,234.56" on a UK-region run (found
' live, owner's run, 2026-09-10; a US machine cannot tell the two
' apart). The euro and pound signs are built with ChrW - see
' VlaNumberFormatCode's own note on code pages.
' G-PROLOG slice 1: a G6 list of Table names -> the literal argument
' text an =PROLOG(...) / =DATALOG(...) formula needs after its program
' string. VLA has no string join, and an Excel formula's Table arguments
' must be LITERAL - a cell holding "Staff, Shifts, Leave" is text, not
' three ranges - so the join has to happen while the formula is being
' written, which is here. VlaNumberFormatCode is the precedent: a small
' pure helper that builds a piece of Excel syntax out of a runtime
' value, and refuses by name rather than handing Excel something it
' will reject in its own words.
'
' WHY THIS REFUSES, and why it is not a general string join. Every name
' here is spliced straight into a formula string, so a name carrying a
' comma, a quote or a bracket does not produce a bad Table reference -
' it produces a DIFFERENT FORMULA, with arguments the writer never
' asked for. A bare {t:text-list} item is one word and cannot do that,
' but G2's own "quotes are the door" means a writer may quote a name,
' and a quoted token carries anything. So each name is held to Excel's
' own rule for what may name a Table: it starts with a letter, an
' underscore or a backslash, and continues in letters, digits,
' underscores and periods. A general VlaJoinNames(items, separator)
' could not make this check - it would not know its items were Table
' names - and would let the corrupted formula through silently.
'
' The names arrive folded to lowercase (every bare word in a sentence
' does), which costs nothing: an Excel formula resolves a Table name
' case-insensitively, and PROLOG/DATALOG fold a predicate name through
' VLA_Identity.Fold on both sides of the match.
Public Function VlaTableArguments(ByVal names As Variant) As String
    If Not IsArray(names) Then RaiseRuntimeMsg "rt-table-arguments-empty"
    Dim lo As Long, hi As Long
    lo = LBound(names)
    hi = UBound(names)
    If hi < lo Then RaiseRuntimeMsg "rt-table-arguments-empty"
    Dim outText As String
    Dim i As Long
    For i = lo To hi
        Dim nm As String
        nm = CStr(names(i))
        If Not IsTableNameShape(nm) Then
            RaiseRuntimeMsg "rt-table-arguments-bad-name", "name", nm
        End If
        If Len(outText) > 0 Then outText = outText & ", "
        outText = outText & nm
    Next i
    VlaTableArguments = outText
End Function

' Excel's own rule for a Table (or defined-name) identifier, narrowed to
' the part a formula argument needs. Checked character by character
' rather than with Like, so the accepted set is readable and no locale's
' character classes can widen it (SD-8).
Private Function IsTableNameShape(ByVal nm As String) As Boolean
    If Len(nm) = 0 Then Exit Function
    Dim i As Long
    Dim c As String
    For i = 1 To Len(nm)
        c = Mid$(nm, i, 1)
        If i = 1 Then
            If Not ((c >= "a" And c <= "z") Or (c >= "A" And c <= "Z") _
                    Or c = "_" Or c = "\") Then Exit Function
        Else
            If Not ((c >= "a" And c <= "z") Or (c >= "A" And c <= "Z") _
                    Or (c >= "0" And c <= "9") Or c = "_" Or c = ".") Then Exit Function
        End If
    Next i
    IsTableNameShape = True
End Function

Private Function CurrencySymbolCode(ByVal currencyName As String) As String
    Select Case currencyName
        Case "dollars": CurrencySymbolCode = "\$"
        Case "euros": CurrencySymbolCode = "\" & ChrW$(8364)
        Case "pounds": CurrencySymbolCode = "\" & ChrW$(163)
        Case Else
            RaiseRuntimeMsg "rt-number-format-unknown-kind", "kind", currencyName
    End Select
End Function

' G-SORTFILTER (pareto.txt section 8): "by column B", "where column C" -
' a sheet column letter, the way every other column slot in english.vla
' reads - turned into that column's position inside the range, which is
' what a sort key (rng.Columns(n)) and Range.AutoFilter's Field both
' want. A column outside the range is refused by name: Excel's own
' answers (a sort key outside the data, a Field of 0 or past the last
' column) name neither the column nor the fix. The slot's shape check
' passes any one to three letters, so a letter past Excel's last column
' (ZZZ) is caught here too, as the same refusal.
Public Function VlaColumnInRange(ByVal rng As Range, ByVal colLetters As String) As Long
    Dim col As Long
    On Error Resume Next
    col = rng.Worksheet.Range(colLetters & "1").Column
    On Error GoTo 0
    If col < rng.Column Or col > rng.Column + rng.Columns.Count - 1 Then
        RaiseRuntimeMsg "rt-column-outside-range", "col", UCase$(colLetters), "range", rng.Address(False, False)
    End If
    VlaColumnInRange = col - rng.Column + 1
End Function

' G-SORTFILTER: "Add filters to range A1:D50." Range.AutoFilter called
' with no arguments is a TOGGLE - on a range that already has filter
' buttons it takes them away - so a sentence that says "add" is made to
' only ever add: already there is a no-op. A sheet holds one set of
' filter buttons, so filters already on a DIFFERENT range are refused by
' name rather than silently moved (and their conditions dropped). A lone
' cell is compared as the block around it, which is the range Excel
' itself filters when handed one cell. A range inside a table shows the
' table's own buttons - calling AutoFilter there would toggle them off;
' "Remove the filters." does not reach a table's buttons (G-TABLES,
' pareto.txt section 9, owns tables).
Public Sub VlaAddFilters(ByVal rng As Range)
    If Not rng.ListObject Is Nothing Then
        rng.ListObject.ShowAutoFilter = True
        Exit Sub
    End If
    If FiltersAlreadyOn(rng) Then Exit Sub
    rng.AutoFilter
End Sub

' G-SORTFILTER: the Field for a "Filter ... to show rows where column C"
' sentence - VlaColumnInRange's position - after the same checks
' VlaAddFilters makes. It exists so a refused filter sentence changes
' nothing (found reading the owner's first live run, 2026-09-10): the
' filter macros used to call VlaAddFilters first and build the condition
' after, so a refused condition ("greater than "abc"") would have left
' the buttons behind. Range.
' AutoFilter WITH a Field is not a toggle - it turns the buttons on
' itself - so the macros now make one call, and every refusal (buttons
' on another range, an empty range, a column outside the range, a value
' that is not a number) comes while its arguments are worked out, before
' Excel is touched. A range inside a table filters through the table's
' own buttons, so the one-set-per-sheet check does not apply to it.
Public Function VlaFilterField(ByVal rng As Range, ByVal colLetters As String) As Long
    ' Called for its refusals only: buttons already on this range, or
    ' none yet, both let the filter call go ahead.
    If rng.ListObject Is Nothing Then Call FiltersAlreadyOn(rng)
    VlaFilterField = VlaColumnInRange(rng, colLetters)
End Function

' Shared by VlaAddFilters and VlaFilterField: True when the sheet's
' filter buttons are already on this range, False when the sheet has
' none; refuses by name when they are on a different range, and when the
' range is empty - Excel's own refusal for an empty range asks the
' writer to "select a single cell", which no sentence can do. A lone
' cell stands for the block around it, as Excel reads it.
Private Function FiltersAlreadyOn(ByVal rng As Range) As Boolean
    Dim target As Range
    If rng.Cells.Count = 1 Then
        Set target = rng.CurrentRegion
    Else
        Set target = rng
    End If
    Dim want As String
    want = target.Address(False, False)
    Dim ws As Worksheet
    Set ws = rng.Worksheet
    If ws.AutoFilterMode Then
        Dim have As String
        have = ws.AutoFilter.Range.Address(False, False)
        If StrComp(have, want, vbTextCompare) <> 0 Then
            RaiseRuntimeMsg "rt-filters-elsewhere", "existing", have, "range", want
        End If
        FiltersAlreadyOn = True
        Exit Function
    End If
    If Application.WorksheetFunction.CountA(target) = 0 Then
        RaiseRuntimeMsg "rt-filter-range-empty", "range", want
    End If
End Function

' G-SORTFILTER: the text of one AutoFilter condition, from the value a
' "Filter ... to show rows where column C ..." sentence names. Pure - a
' function of its arguments, no workbook - so every condition it builds
' is pinned in the pure suite (and it is an F.16 candidate).
'   at-least, at-most - the two halves of "is": english.vla passes
'     Criteria1:=at-least, Operator:=xlAnd, Criteria2:=at-most. For a
'     NUMBER they are ">=100" and "<=100", not one "=100": an "equals"
'     condition is matched against the cell's DISPLAYED text (as the
'     filter dropdown lists displayed values - taken on reading, not yet
'     seen here), so "=100" could miss a cell showing $100.00; two
'     comparisons match the value whatever its format, and are right
'     either way. For TEXT both halves are the same "=text",
'     with the text's own wildcards escaped (below) - text is what the
'     cell shows, so equals is right for it, including text that looks
'     like a number ("007"). An empty value is "=", Excel's own
'     condition for a blank cell.
'   contains - "*text*", the text's own wildcards escaped, so "contains
'     "5*3"" finds 5*3, not every row.
'   greater, less - ">100", "<100". A number only: text refuses by
'     name rather than quietly comparing alphabetically.
' Every number is written with Str$, which always uses "." - AutoFilter
' reads a condition from VBA the US way, so a comma-decimal machine's
' own CStr (2,5) would be read as text. ("given", not "val": a local
' named like VBA's own Val function is a known trap.)
Public Function VlaFilterCriterion(ByVal kind As String, ByVal v As Variant) As String
    Dim given As Variant
    If IsObject(v) Then
        given = v.Value
    Else
        given = v
    End If
    If IsNull(given) Then given = Empty
    Dim isNum As Boolean
    Select Case VarType(given)
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, vbByte, vbDate
            isNum = True
    End Select

    Select Case kind
        Case "at-least", "at-most"
            If isNum Then
                If kind = "at-least" Then
                    VlaFilterCriterion = ">=" & InvariantNumber(given)
                Else
                    VlaFilterCriterion = "<=" & InvariantNumber(given)
                End If
            ElseIf IsEmpty(given) Then
                VlaFilterCriterion = "="
            Else
                VlaFilterCriterion = "=" & EscapeFilterWildcards(CStr(given))
            End If
        Case "contains"
            If isNum Then
                VlaFilterCriterion = "*" & InvariantNumber(given) & "*"
            Else
                VlaFilterCriterion = "*" & EscapeFilterWildcards(CStr(given)) & "*"
            End If
        Case "greater", "less"
            If Not isNum Then
                Dim shown As String
                If IsEmpty(given) Then shown = "(nothing)" Else shown = CStr(given)
                RaiseRuntimeMsg "rt-filter-needs-number", "value", shown
            End If
            If kind = "greater" Then
                VlaFilterCriterion = ">" & InvariantNumber(given)
            Else
                VlaFilterCriterion = "<" & InvariantNumber(given)
            End If
        Case Else
            RaiseRuntimeMsg "rt-filter-unknown-kind", "kind", kind
    End Select
End Function

' A number as AutoFilter reads it from VBA: "." for the decimal point on
' every machine (Str$ never localises), no leading space, and a leading
' zero before a bare point (Str$ writes 0.5 as " .5").
Private Function InvariantNumber(ByVal n As Variant) As String
    Dim s As String
    s = Trim$(Str$(CDbl(n)))
    If Left$(s, 1) = "." Then
        s = "0" & s
    ElseIf Left$(s, 2) = "-." Then
        s = "-0" & Mid$(s, 2)
    End If
    InvariantNumber = s
End Function

' AutoFilter's wildcards are * and ?, escaped with ~ (which escapes
' itself), so a value is matched as written. ~ first, or the other two
' escapes would be escaped again.
Private Function EscapeFilterWildcards(ByVal s As String) As String
    s = Replace(s, "~", "~~")
    s = Replace(s, "*", "~*")
    EscapeFilterWildcards = Replace(s, "?", "~?")
End Function

' G-TEXT slice 1 (pareto.txt section 12): one change to one piece of
' text, purely, so every rule below is pinned in the pure suite with no
' workbook. Where Excel has a worksheet function for the change, this is
' that function written out rather than called - WorksheetFunction would
' tie the pure suite to a host - and the host suite holds each one
' against Excel's own function as its oracle.
'   upper, lower                  UCase/LCase: Excel's UPPER and LOWER.
'   capitalize-after-non-letter   Excel's PROPER: a letter is a capital
'                                 when it starts the text or comes right
'                                 after any character that is not a
'                                 letter, and small otherwise. "o'neil"
'                                 is "O'Neil"; PROPER's own quirk,
'                                 "don't" is "Don'T", comes with it.
'   capitalize-after-space        the same, but only a space starts a
'                                 word: "don't" is "Don't", while "3rd"
'                                 and "(note)" stay as they are, since
'                                 their first letter does not come right
'                                 after a space. A space is an ordinary
'                                 or non-breaking one, a tab or a line
'                                 break.
'   remove-extra-spaces           Excel's TRIM, with a non-breaking space
'                                 (160) counted as a space (the owner's
'                                 call, 2026-09-25): every run of them
'                                 becomes one ordinary space, and none is
'                                 left at either end.
'   remove-non-printing           Excel's CLEAN: characters 0 to 31 go,
'                                 tabs and line breaks among them, and
'                                 nothing else does.
' The rule that says which one a sentence meant is in the sentence
' ("after any space|non-letter"); the owner's call over a default that
' would hide it. An unknown change refuses by name, and VlaTextInRange
' asks here before it touches a cell, so a refused sentence changes
' nothing. Slice 2 made the text one value of any kind (TextOfValue's
' header says how each reads), so the same changes serve a variable:
' "Set tidy to name with extra spaces removed."
Public Function VlaTextOp(ByVal v As Variant, ByVal op As String) As String
    Dim s As String
    s = TextOfValue(v)
    Select Case op
        Case "upper": VlaTextOp = UCase$(s)
        Case "lower": VlaTextOp = LCase$(s)
        Case "capitalize-after-non-letter": VlaTextOp = CapitalizeWords(s, False)
        Case "capitalize-after-space": VlaTextOp = CapitalizeWords(s, True)
        Case "remove-extra-spaces": VlaTextOp = RemoveExtraSpaces(s)
        Case "remove-non-printing": VlaTextOp = RemoveNonPrinting(s)
        Case Else
            RaiseRuntimeMsg "rt-text-unknown-op", "op", op
    End Select
End Function

' The two capitalize rules differ in one question only: what may come
' right before a letter that starts a word. Letters that start no word
' are made small, as PROPER does - "MCDONALD" is "Mcdonald" either way.
' The text is changed in place with the Mid$ statement, one character at
' a time, so a long cell is not rebuilt by concatenation.
Private Function CapitalizeWords(ByVal s As String, ByVal afterSpaceOnly As Boolean) As String
    Dim r As String
    r = s
    Dim i As Long
    Dim ch As String
    Dim startsWord As Boolean
    startsWord = True
    For i = 1 To Len(r)
        ch = Mid$(r, i, 1)
        If IsCasedLetter(ch) Then
            If startsWord Then
                Mid$(r, i, 1) = UCase$(ch)
            Else
                Mid$(r, i, 1) = LCase$(ch)
            End If
            startsWord = False
        ElseIf afterSpaceOnly Then
            startsWord = IsTextSpace(ch)
        Else
            startsWord = True
        End If
    Next
    CapitalizeWords = r
End Function

' A letter, for capitalizing: A to Z and a to z; the letters of Latin-1
' (192 to 255 but for the two signs at 215 and 247, so the sharp s at 223
' counts although VBA gives it no capital); and past Latin-1, any character
' with an upper and a lower case. AscW is signed above &H7FFF, so a
' negative code simply means past Latin-1.
Private Function IsCasedLetter(ByVal ch As String) As Boolean
    Select Case AscW(ch)
        Case 65 To 90, 97 To 122, 192 To 214, 216 To 246, 248 To 255
            IsCasedLetter = True
        Case 0 To 191, 215, 247
            IsCasedLetter = False
        Case Else
            IsCasedLetter = (UCase$(ch) <> LCase$(ch))
    End Select
End Function

' What "after any space" counts as a space: an ordinary or non-breaking
' space, a tab, or a line break (a new line in a cell starts a word as
' visibly as a space does).
Private Function IsTextSpace(ByVal ch As String) As Boolean
    Select Case AscW(ch)
        Case 32, 160, 9, 10, 13
            IsTextSpace = True
    End Select
End Function

' Excel's TRIM, non-breaking spaces included: a run of spaces is held
' back until the next character that is not one, so it is written as
' one ordinary space between two words and not at all at either end.
Private Function RemoveExtraSpaces(ByVal s As String) As String
    Dim buf As String
    buf = Space$(Len(s))
    Dim n As Long
    Dim i As Long
    Dim code As Long
    Dim pending As Boolean
    For i = 1 To Len(s)
        code = AscW(Mid$(s, i, 1))
        If code = 32 Or code = 160 Then
            If n > 0 Then pending = True
        Else
            If pending Then
                n = n + 1
                Mid$(buf, n, 1) = " "
                pending = False
            End If
            n = n + 1
            Mid$(buf, n, 1) = Mid$(s, i, 1)
        End If
    Next
    RemoveExtraSpaces = Left$(buf, n)
End Function

' Excel's CLEAN: the 32 control characters of 7-bit ASCII go, and
' nothing past them (CLEAN leaves 127 and the other high controls too).
Private Function RemoveNonPrinting(ByVal s As String) As String
    Dim buf As String
    buf = Space$(Len(s))
    Dim n As Long
    Dim i As Long
    Dim code As Long
    For i = 1 To Len(s)
        code = AscW(Mid$(s, i, 1))
        If code < 0 Or code > 31 Then
            n = n + 1
            Mid$(buf, n, 1) = Mid$(s, i, 1)
        End If
    Next
    RemoveNonPrinting = Left$(buf, n)
End Function

' G-TEXT slice 1: the cells a text sentence changes are the ones holding
' TEXT, and no others. A formula stays even when it shows text (writing
' would replace it with its own result), and so do numbers, dates,
' TRUE/FALSE, errors and blank cells, so "Make column B upper case."
' never turns a date into text. A whole column is read only as far as
' the sheet's used range; each area is read once, in bulk, and a cell is
' written only if its text changed.
'
' A cell that held text still holds text afterwards (the owner's call,
' 2026-09-25). Excel reads a value written from VBA as if it were typed,
' so "  00123 " trimmed to "00123", or "true" upper-cased to "TRUE",
' would come back as the number 123 or the logical TRUE. Such a cell is
' written again behind an apostrophe, Excel's own mark for text, and
' which ones need it is found by asking Excel after the write, not by
' predicting its reader, which differs by region. Text that begins with
' = + - or @ is marked before it is written at all, so upper-casing
' "=hyperlink(...)" can never make a formula: SEC.4's rule, the one
' VLA_Interpreter's NeutralizeFormulaInjection applies, repeated here
' rather than called because this module must compile alone in a user's
' workbook.
Public Sub VlaTextInRange(ByVal rng As Range, ByVal op As String)
    Dim probe As String
    probe = VlaTextOp(vbNullString, op)
    Dim target As Range
    Set target = Application.Intersect(rng, rng.Worksheet.UsedRange)
    If target Is Nothing Then Exit Sub
    Dim area As Range
    For Each area In target.Areas
        TextInArea area, op
    Next
End Sub

' One rectangle of the range. HasFormula answers for the whole area:
' True when every cell holds a formula (nothing to do), False when none
' does, and Null when some do - only then is each text cell asked.
Private Sub TextInArea(ByVal area As Range, ByVal op As String)
    Dim hf As Variant
    hf = area.HasFormula
    If VarType(hf) = vbBoolean Then
        If hf Then Exit Sub
    End If
    Dim checkEach As Boolean
    checkEach = IsNull(hf)
    Dim vals As Variant
    vals = area.Value
    If Not IsArray(vals) Then
        If VarType(vals) = vbString Then RewriteTextCell area.Cells(1, 1), CStr(vals), op, checkEach
        Exit Sub
    End If
    Dim r As Long
    Dim c As Long
    For r = 1 To UBound(vals, 1)
        For c = 1 To UBound(vals, 2)
            If VarType(vals(r, c)) = vbString Then
                RewriteTextCell area.Cells(r, c), CStr(vals(r, c)), op, checkEach
            End If
        Next
    Next
End Sub

' One text cell: changed, then written so it reads back as the same
' text. Text changed to nothing (a cell of spaces, trimmed) leaves the
' cell blank; Value, not ClearContents, so the first cell of a merged
' block can be emptied too.
Private Sub RewriteTextCell(ByVal cell As Range, ByVal s As String, ByVal op As String, _
                            ByVal checkFormula As Boolean)
    If checkFormula Then
        If cell.HasFormula Then Exit Sub
    End If
    Dim t As String
    t = VlaTextOp(s, op)
    If t = s Then Exit Sub
    If Len(t) = 0 Then
        cell.Value = Empty
        Exit Sub
    End If
    Select Case Left$(t, 1)
        Case "=", "+", "-", "@"
            cell.Value = "'" & t
            Exit Sub
    End Select
    cell.Value = t
    If VarType(cell.Value) <> vbString Then cell.Value = "'" & t
End Sub

' G-TEXT slice 2 (pareto.txt section 12): text in a variable. One value
' read as text, for every change made to a value rather than to cells:
' text as it is; a number as VBA's CStr writes it (this machine's decimal
' mark, the same text "joined with" gives); a date as this machine's
' short date; TRUE and FALSE as Excel shows them; nothing as nothing. One
' cell gives its value. A list, a range of several cells, or an error
' value refuses by name, since none of them is one piece of text.
' (v Is Nothing is asked on its own line: TypeOf on Nothing raises.)
Private Function TextOfValue(ByVal v As Variant) As String
    Dim given As Variant
    If IsObject(v) Then
        If v Is Nothing Then RaiseRuntimeMsg "rt-text-not-one-value"
        If Not TypeOf v Is Range Then RaiseRuntimeMsg "rt-text-not-one-value"
        If v.Cells.CountLarge <> 1 Then RaiseRuntimeMsg "rt-text-not-one-value"
        given = v.Value
    Else
        given = v
    End If
    If IsArray(given) Then RaiseRuntimeMsg "rt-text-not-one-value"
    If IsError(given) Then RaiseRuntimeMsg "rt-text-error-value", "value", CStr(given)
    Select Case VarType(given)
        Case vbEmpty, vbNull
            TextOfValue = vbNullString
        Case vbBoolean
            If given Then TextOfValue = "TRUE" Else TextOfValue = "FALSE"
        Case Else
            TextOfValue = CStr(given)
    End Select
End Function

' A value inside a refusal: long text cut short so the message stays
' readable, and nothing shown as "(nothing)".
Private Function ShownText(ByVal s As String) As String
    If Len(s) = 0 Then
        ShownText = "(nothing)"
    ElseIf Len(s) > 60 Then
        ShownText = Left$(s, 57) & "..."
    Else
        ShownText = s
    End If
End Function

' "Set part to the text before|after [the last] "-" in code." Excel's
' TEXTBEFORE and TEXTAFTER, with two of the owner's calls (2026-09-25):
' the marker is found in any case, exactly as the core's "If code
' contains "-"" finds it (InStr with vbTextCompare, the same call), so
' the two sentences can never disagree; and a marker that is not there
' refuses by name, as Excel's own #N/A does, rather than giving the whole
' text or nothing and letting a wrong value run on. A program expecting
' gaps asks "If code contains "-"" first. "the last" searches from the
' end. The side and the occurrence are checked before anything is found,
' so an unknown one refuses the same way whatever the text is.
Public Function VlaTextBeside(ByVal v As Variant, ByVal marker As Variant, _
                              ByVal side As String, ByVal which As String) As String
    If side <> "before" And side <> "after" Then
        RaiseRuntimeMsg "rt-text-unknown-kind", "proc", "VlaTextBeside", "kind", side, "expected", "before or after"
    End If
    If which <> "first" And which <> "last" Then
        RaiseRuntimeMsg "rt-text-unknown-kind", "proc", "VlaTextBeside", "kind", which, "expected", "first or last"
    End If
    Dim s As String
    Dim m As String
    s = TextOfValue(v)
    m = TextOfValue(marker)
    If Len(m) = 0 Then RaiseRuntimeMsg "rt-text-marker-empty"
    Dim pos As Long
    If which = "first" Then
        pos = InStr(1, s, m, vbTextCompare)
    Else
        pos = InStrRev(s, m, -1, vbTextCompare)
    End If
    If pos = 0 Then RaiseRuntimeMsg "rt-text-marker-missing", "marker", m, "text", ShownText(s)
    If side = "before" Then
        VlaTextBeside = Left$(s, pos - 1)
    Else
        VlaTextBeside = Mid$(s, pos + Len(m))
    End If
End Function

' "Set code to id padded on the left|right with "0" to 5 characters."
' The side and the character are both in the sentence (the owner's
' call), so 42 becomes "00042" and a fixed-width field "ab  " only when
' the sentence says so. Text already that long or longer is never cut -
' Excel's RIGHT("00000"&A1,5) would cut it, and a silently shortened
' code is worse than a long one. The pad must be exactly one character
' and the width a whole number of characters; each refuses by name.
Public Function VlaTextPad(ByVal v As Variant, ByVal padWith As Variant, _
                           ByVal width As Variant, ByVal side As String) As String
    If side <> "left" And side <> "right" Then
        RaiseRuntimeMsg "rt-text-unknown-kind", "proc", "VlaTextPad", "kind", side, "expected", "left or right"
    End If
    Dim s As String
    Dim c As String
    s = TextOfValue(v)
    c = TextOfValue(padWith)
    If Len(c) <> 1 Then RaiseRuntimeMsg "rt-text-pad-character", "value", c
    Dim wholeWidth As Boolean
    Select Case VarType(width)
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, vbByte
            If width = Int(width) Then
                If width >= 0 And width <= 32767 Then wholeWidth = True
            End If
    End Select
    If Not wholeWidth Then RaiseRuntimeMsg "rt-text-pad-width", "value", ShownText(TextOfValue(width))
    Dim n As Long
    n = CLng(width)
    If Len(s) >= n Then
        VlaTextPad = s
    ElseIf side = "left" Then
        VlaTextPad = String$(n - Len(s), c) & s
    Else
        VlaTextPad = s & String$(n - Len(s), c)
    End If
End Function

' "Set names to range A2:A9 as one list [separated by "; "]." G-PROLOG
' slice 5's own words for a list in one cell, so one phrase means one
' thing: ", " between items unless the sentence names another, row by
' row (TEXTJOIN's order), blank cells skipped. Each cell gives its value
' read as TextOfValue reads one - a number without its currency format,
' a date as this machine's short date - not its displayed text, which
' shows "####" in a column too narrow and would make the answer depend on
' column width (the owner's call). A whole column is read only as far as
' the used range. A cell holding an error refuses, naming the cell.
' Items are gathered first and joined once, so a long column is not
' rebuilt by concatenation.
Public Function VlaJoinRange(ByVal rng As Range, ByVal separator As Variant) As String
    Dim sep As String
    sep = TextOfValue(separator)
    Dim target As Range
    Set target = Application.Intersect(rng, rng.Worksheet.UsedRange)
    If target Is Nothing Then Exit Function
    Dim items As Collection
    Set items = New Collection
    Dim area As Range
    Dim vals As Variant
    Dim r As Long
    Dim c As Long
    For Each area In target.Areas
        vals = area.Value
        If IsArray(vals) Then
            For r = 1 To UBound(vals, 1)
                For c = 1 To UBound(vals, 2)
                    AddJoinItem items, vals(r, c), area, r, c
                Next
            Next
        Else
            AddJoinItem items, vals, area, 1, 1
        End If
    Next
    If items.Count = 0 Then Exit Function
    Dim parts() As String
    ReDim parts(1 To items.Count)
    Dim i As Long
    For i = 1 To items.Count
        parts(i) = items.Item(i)
    Next
    VlaJoinRange = Join(parts, sep)
End Function

' One cell's contribution to a list: nothing for a blank cell, a refusal
' naming the cell for an error, and otherwise its value as text. The
' cell's address is worked out only for the refusal.
Private Sub AddJoinItem(ByVal items As Collection, ByVal given As Variant, _
                        ByVal area As Range, ByVal r As Long, ByVal c As Long)
    If IsError(given) Then
        RaiseRuntimeMsg "rt-join-cell-error", "cell", area.Cells(r, c).Address(False, False), "value", CStr(given)
    End If
    If IsEmpty(given) Then Exit Sub
    If VarType(given) = vbString Then
        If Len(given) = 0 Then Exit Sub
    End If
    items.Add TextOfValue(given)
End Sub

' G-TEXT slice 3 (pareto.txt sections 12 and 13): find and replace, and
' split. Four forks the owner decided (2026-09-28), each named where its
' helper sits.
'
' "Replace "N/A" with 0 in range A1:D50." (and its column and "on this
' sheet" siblings) changes VALUES: a cell holding a formula is never
' touched, and "... in the formulas of range ..." is the sentence that
' changes formulas and nothing else (the owner's call). Until this slice
' the sentence was Excel's own Range.Replace, which edits formula text
' too - Replace "A" with "B" in a column turned =A1*2 into =B1*2 - and
' which enters each cell it changes again as if typed, so a result
' starting with = became a live formula. That was a way around SEC.4's
' rule, which every other value Frazaro writes obeys; narrowing a
' shipped sentence's meaning was the owner's call on security grounds.
'
' What a value cell matches against is its own entry - what Excel's
' Replace itself reads for a constant (.Formula: text as it is, a number
' or a date as Excel keeps it) - found in any case, anywhere in the cell,
' Excel's starting settings, the ones slice 1 pinned. A changed entry is
' written back the way Excel's Replace writes it, as if typed, so "N/A"
' replaced by 0 is still the number 0, with two exceptions: text that
' would begin with = + - or @ and is not a plain number is kept as text
' behind an apostrophe (SEC.4's rule, repeated here because this module
' must compile alone in a user's workbook), and a cell marked as text
' with an apostrophe stays text. A whole column or sheet is read only as
' far as the used range, each area in bulk, and a cell is written only if
' its entry changed.
Public Sub VlaReplaceInRange(ByVal rng As Range, ByVal what As Variant, _
                             ByVal replacement As Variant, ByVal where As String)
    If where <> "values" And where <> "formulas" Then
        RaiseRuntimeMsg "rt-text-unknown-kind", "proc", "VlaReplaceInRange", "kind", where, "expected", "values or formulas"
    End If
    Dim w As String
    Dim b As String
    w = TextOfValue(what)
    b = TextOfValue(replacement)
    If Len(w) = 0 Then RaiseRuntimeMsg "rt-text-marker-empty"
    Dim target As Range
    Set target = Application.Intersect(rng, rng.Worksheet.UsedRange)
    If target Is Nothing Then Exit Sub
    Dim area As Range
    For Each area In target.Areas
        If where = "values" Then
            ReplaceValuesInArea area, w, b
        Else
            ReplaceFormulasInArea area, w, b
        End If
    Next
End Sub

' One rectangle, values only. HasFormula answers for the whole area as
' in slice 1: True leaves it alone, Null means each matching cell is
' asked before it is changed.
Private Sub ReplaceValuesInArea(ByVal area As Range, ByVal w As String, ByVal b As String)
    Dim hf As Variant
    hf = area.HasFormula
    If VarType(hf) = vbBoolean Then
        If hf Then Exit Sub
    End If
    Dim checkEach As Boolean
    checkEach = IsNull(hf)
    Dim f As Variant
    f = area.Formula
    If Not IsArray(f) Then
        ReplaceInCell area.Cells(1, 1), CStr(f), w, b, checkEach
        Exit Sub
    End If
    Dim r As Long
    Dim c As Long
    For r = 1 To UBound(f, 1)
        For c = 1 To UBound(f, 2)
            If InStr(1, f(r, c), w, vbTextCompare) > 0 Then
                ReplaceInCell area.Cells(r, c), CStr(f(r, c)), w, b, checkEach
            End If
        Next
    Next
End Sub

' One cell whose entry holds what is being replaced. Emptied when nothing
' is left, as Excel's Replace does.
Private Sub ReplaceInCell(ByVal cell As Range, ByVal s As String, ByVal w As String, _
                          ByVal b As String, ByVal checkFormula As Boolean)
    If InStr(1, s, w, vbTextCompare) = 0 Then Exit Sub
    If checkFormula Then
        If cell.HasFormula Then Exit Sub
    End If
    Dim t As String
    t = Replace(s, w, b, 1, -1, vbTextCompare)
    If t = s Then Exit Sub
    If Len(t) = 0 Then
        cell.Value = Empty
        Exit Sub
    End If
    Select Case Left$(t, 1)
        Case "=", "+", "-", "@"
            If Not IsPlainNumberText(t) Then
                cell.Value = "'" & t
                Exit Sub
            End If
    End Select
    ' Written as Value, never Formula: by here nothing that begins with
    ' = + - or @ is left but a plain number, so Excel's reading of it as
    ' typed can only make a number, a date or a logical - the same as its
    ' own Replace - and this helper is no place a formula can be written
    ' from (check_no_network's formula-write sites stay two).
    If cell.PrefixCharacter = "'" Then
        cell.Value = "'" & t
    Else
        cell.Value = t
    End If
End Sub

' One rectangle, formulas only: Excel's own Replace, on the formula cells
' and no others, with the settings pinned so no earlier Find leaks in.
' (A mixed area has at least two cells, so SpecialCells reads the area
' itself - on one cell it would read the whole sheet.)
Private Sub ReplaceFormulasInArea(ByVal area As Range, ByVal w As String, ByVal b As String)
    Dim hf As Variant
    hf = area.HasFormula
    Dim formulaCells As Range
    If VarType(hf) = vbBoolean Then
        If Not hf Then Exit Sub
        Set formulaCells = area
    Else
        Set formulaCells = area.SpecialCells(xlCellTypeFormulas)
    End If
    formulaCells.Replace What:=w, Replacement:=b, LookAt:=xlPart, MatchCase:=False, _
                         SearchFormat:=False, ReplaceFormat:=False
End Sub

' A text that is plainly a number and nothing more: an optional minus,
' digits, and at most one point with digits after it - no leading zero
' before another digit, no exponent, no thousands mark, and at most 15
' digits, so reading it as a number loses nothing a person would miss.
' "42", "-3.5" and "0.25" are; "0042", "1E5", "1,000", "3/4", "+5" and a
' 16-digit account number are not.
Private Function IsPlainNumberText(ByVal s As String) As Boolean
    Dim body As String
    body = s
    If Left$(body, 1) = "-" Then body = Mid$(body, 2)
    If Len(body) = 0 Then Exit Function
    Dim intPart As String
    Dim fracPart As String
    Dim dot As Long
    dot = InStr(body, ".")
    If dot = 0 Then
        intPart = body
    Else
        intPart = Left$(body, dot - 1)
        fracPart = Mid$(body, dot + 1)
        If Len(fracPart) = 0 Then Exit Function
        If fracPart Like "*[!0-9]*" Then Exit Function
    End If
    If Len(intPart) = 0 Then Exit Function
    If intPart Like "*[!0-9]*" Then Exit Function
    If Len(intPart) > 1 And Left$(intPart, 1) = "0" Then Exit Function
    If Len(intPart) + Len(fracPart) > 15 Then Exit Function
    IsPlainNumberText = True
End Function

' "Split column C by "," as text|reading numbers." The column's own cell
' keeps the first piece and the rest go into the columns to its right,
' as Excel's Text to Columns does, with the owner's two calls
' (2026-09-28). A split never writes over anything: when a cell those
' pieces need holds something, down to the column's last used row,
' nothing changes and the refusal names that cell. And the sentence says
' what the pieces are: "as text" keeps every piece text ("0042" keeps its
' zeros), "reading numbers" makes a plain number a number (IsPlainNumberText
' - "42" and "-3.5", never "0042", "3/4" or "1E5"), and neither ever makes
' a date. Only text is split: numbers, dates, formulas and blank cells stay
' as they are. The separator is found in any case, as "contains" finds
' text, and the pieces are kept exactly, spaces included. A piece that
' would begin with = + - or @ is written as text (SEC.4). The pieces to the
' right go in with one write, then any Excel read as something other than
' text are written again behind an apostrophe - slice 1's way, asking
' Excel after the write rather than predicting its reader.
Public Sub VlaSplitColumn(ByVal col As Range, ByVal separator As Variant, ByVal how As String)
    If how <> "text" And how <> "numbers" Then
        RaiseRuntimeMsg "rt-text-unknown-kind", "proc", "VlaSplitColumn", "kind", how, "expected", "text or numbers"
    End If
    Dim sep As String
    sep = TextOfValue(separator)
    If Len(sep) = 0 Then RaiseRuntimeMsg "rt-split-separator-empty"
    Dim target As Range
    Set target = Application.Intersect(col.Columns(1), col.Worksheet.UsedRange)
    If target Is Nothing Then Exit Sub
    Dim hf As Variant
    hf = target.HasFormula
    If VarType(hf) = vbBoolean Then
        If hf Then Exit Sub
    End If
    Dim n As Long
    n = target.Rows.Count
    Dim vals As Variant
    vals = target.Value
    Dim pieces() As Variant
    ReDim pieces(1 To n)
    Dim widest As Long
    Dim r As Long
    Dim s As Variant
    Dim parts() As String
    For r = 1 To n
        If n = 1 Then s = vals Else s = vals(r, 1)
        If VarType(s) = vbString Then
            If InStr(1, s, sep, vbTextCompare) > 0 Then
                If Not SplitSkipsFormula(target.Cells(r, 1), IsNull(hf)) Then
                    parts = Split(s, sep, -1, vbTextCompare)
                    pieces(r) = parts
                    If UBound(parts) > widest Then widest = UBound(parts)
                End If
            End If
        End If
    Next
    If widest = 0 Then Exit Sub

    Dim ws As Worksheet
    Set ws = target.Worksheet
    Dim need As String
    need = widest & " empty column"
    If widest > 1 Then need = need & "s"
    If target.Column + widest > ws.Columns.Count Then
        RaiseRuntimeMsg "rt-split-too-wide", "column", ColumnLetters(ws, target.Column), "need", need
    End If
    Dim room As Range
    Set room = target.Offset(0, 1).Resize(n, widest)
    Dim held As String
    held = FirstCellHolding(room)
    If Len(held) > 0 Then
        Dim span As String
        span = ColumnLetters(ws, target.Column + 1)
        If widest > 1 Then span = span & " to " & ColumnLetters(ws, target.Column + widest)
        RaiseRuntimeMsg "rt-split-columns-not-empty", "column", ColumnLetters(ws, target.Column), "need", need, _
                        "columns", span, "row", CStr(target.Row + n - 1), "cell", held
    End If

    Dim out() As Variant
    ReDim out(1 To n, 1 To widest)
    Dim k As Long
    For r = 1 To n
        If IsArray(pieces(r)) Then
            parts = pieces(r)
            For k = 1 To UBound(parts)
                out(r, k) = PieceValue(parts(k), how)
            Next
        End If
    Next
    room.Value = out
    Dim got As Variant
    got = room.Value
    For r = 1 To n
        If IsArray(pieces(r)) Then
            parts = pieces(r)
            For k = 1 To UBound(parts)
                If n = 1 And widest = 1 Then
                    KeepPieceText room.Cells(r, k), out(r, k), got, parts(k)
                Else
                    KeepPieceText room.Cells(r, k), out(r, k), got(r, k), parts(k)
                End If
            Next
            WritePiece target.Cells(r, 1), parts(0), how
        End If
    Next
End Sub

' A text cell in a column that holds formulas is asked whether it is one.
Private Function SplitSkipsFormula(ByVal cell As Range, ByVal checkEach As Boolean) As Boolean
    If checkEach Then SplitSkipsFormula = cell.HasFormula
End Function

' What one piece is written as: nothing for an empty piece, a number for
' a plain number when the sentence reads numbers, and otherwise text -
' marked with an apostrophe already when it begins with = + - or @, or
' with an apostrophe of its own, which Excel would otherwise take as
' its mark and drop.
Private Function PieceValue(ByVal p As String, ByVal how As String) As Variant
    If Len(p) = 0 Then
        PieceValue = Empty
    ElseIf how = "numbers" And IsPlainNumberText(Trim$(p)) Then
        PieceValue = Val(Trim$(p))
    Else
        Select Case Left$(p, 1)
            Case "=", "+", "-", "@", "'"
                PieceValue = "'" & p
            Case Else
                PieceValue = p
        End Select
    End If
End Function

' A piece meant as text that Excel read as a number, a date or a logical
' is written again behind an apostrophe, so it reads back as it was. One
' already marked is text by then.
Private Sub KeepPieceText(ByVal cell As Range, ByVal written As Variant, ByVal readBack As Variant, _
                          ByVal p As String)
    If VarType(written) <> vbString Then Exit Sub
    If Left$(written, 1) = "'" Then Exit Sub
    If VarType(readBack) <> vbString Then cell.Value = "'" & p
End Sub

' The column's own cell, keeping the first piece.
Private Sub WritePiece(ByVal cell As Range, ByVal p As String, ByVal how As String)
    Dim v As Variant
    v = PieceValue(p, how)
    cell.Value = v
    If VarType(v) = vbString Then
        If VarType(cell.Value) <> vbString Then cell.Value = "'" & p
    End If
End Sub

' The first cell of a rectangle holding anything - a value or a formula,
' even one showing nothing - read row by row; "" when every cell is empty.
Private Function FirstCellHolding(ByVal room As Range) As String
    Dim f As Variant
    f = room.Formula
    If Not IsArray(f) Then
        If Len(f) > 0 Then FirstCellHolding = room.Cells(1, 1).Address(False, False)
        Exit Function
    End If
    Dim r As Long
    Dim c As Long
    For r = 1 To UBound(f, 1)
        For c = 1 To UBound(f, 2)
            If Len(f(r, c)) > 0 Then
                FirstCellHolding = room.Cells(r, c).Address(False, False)
                Exit Function
            End If
        Next
    Next
End Function

' A column's letters, "D" or "AB", from its number.
Private Function ColumnLetters(ByVal ws As Worksheet, ByVal colNum As Long) As String
    ColumnLetters = Split(ws.Cells(1, colNum).Address(True, False), "$")(0)
End Function

' "Set r to the row of the first cell in range A1:D50 containing "x"."
' (and "the column of", and "in column C") - the owner's call: a search
' hands back numbers the "column number ... row ..." sentences can use,
' and 0 when nothing matches, as the shipped "row of ... in column" does;
' pareto's "the first cell" itself could not be kept in a variable, since
' a variable set to a cell holds its value (IN.11). "If range A1:D50
' contains "x"" asks the same question first. A cell matches when its
' value, read as text the way TextOfValue reads one, contains the text in
' any case - exactly as "If code contains "x"" matches one value, so the
' two never disagree. The search goes row by row (each area of the range
' in turn), blank cells and error values skipped: an error holds no text.
Public Function VlaFindText(ByVal rng As Range, ByVal lookFor As Variant, ByVal what As String) As Long
    If what <> "row" And what <> "column" Then
        RaiseRuntimeMsg "rt-text-unknown-kind", "proc", "VlaFindText", "kind", what, "expected", "row or column"
    End If
    VlaFindText = SearchText(rng, lookFor, what)
End Function

' "Set n to how many cells in range A2:A99 contain "x"." The same match,
' counted: text, not COUNTIF's wildcards, so "*" and "?" are themselves.
Public Function VlaCountText(ByVal rng As Range, ByVal lookFor As Variant) As Long
    VlaCountText = SearchText(rng, lookFor, "count")
End Function

' The one walk both answers share: the first match's row or column, or
' how many cells match.
Private Function SearchText(ByVal rng As Range, ByVal lookFor As Variant, ByVal what As String) As Long
    Dim t As String
    t = TextOfValue(lookFor)
    If Len(t) = 0 Then RaiseRuntimeMsg "rt-text-marker-empty"
    Dim target As Range
    Set target = Application.Intersect(rng, rng.Worksheet.UsedRange)
    If target Is Nothing Then Exit Function
    Dim found As Long
    Dim area As Range
    Dim vals As Variant
    Dim r As Long
    Dim c As Long
    Dim v As Variant
    For Each area In target.Areas
        vals = area.Value
        For r = 1 To area.Rows.Count
            For c = 1 To area.Columns.Count
                If IsArray(vals) Then v = vals(r, c) Else v = vals
                If ValueHoldsText(v, t) Then
                    If what = "row" Then
                        SearchText = area.Row + r - 1
                        Exit Function
                    ElseIf what = "column" Then
                        SearchText = area.Column + c - 1
                        Exit Function
                    End If
                    found = found + 1
                End If
            Next
        Next
    Next
    If what = "count" Then SearchText = found
End Function

Private Function ValueHoldsText(ByVal v As Variant, ByVal t As String) As Boolean
    If IsError(v) Or IsEmpty(v) Then Exit Function
    ValueHoldsText = (InStr(1, TextOfValue(v), t, vbTextCompare) > 0)
End Function

' G-PIVOT rule #1 (pareto.txt section 10, "Pivot tables"): the two-step
' COM sequence pareto.txt's own target names - PivotCaches.Create then
' CreatePivotTable - lives here rather than in a defmacro because it is
' two statements, not a nested expression, and (unlike make-table's
' single-expression .Add) doesn't reduce to one "." form at all.
' Dispatched exactly like every other VLA_Runtime helper: a direct call
' from the compiled backend, Application.Run by name from the
' interpreter (VLA_Interpreter.bas's TryRuntimeHelper) - no new
' mechanism, the same one VlaEnsureSheet above already proves works for
' object-model side effects, not just pure-value helpers.
Public Sub VlaPivotCreate(ByVal sourceRange As Range, ByVal destCell As Range, ByVal tableName As String)
    Dim pc As PivotCache
    Set pc = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=sourceRange)
    pc.CreatePivotTable TableDestination:=destCell, TableName:=tableName
End Sub

' G-PIVOT: Excel has no Workbook.PivotTables - only a per-Worksheet
' one - and none of pivot-rows/-columns/-filters/-full's own sentences
' name a sheet (unlike pivot-create, which gets one implicitly via
' destCell), so finding a pivot by name alone means searching every
' worksheet. PivotTables(name) raises rather than returning Nothing on
' a per-sheet miss (VLA_Tests_Host.bas's own G-PIVOT host check already
' established this, guarding the same way), so each probe is guarded
' individually rather than the whole loop.
Private Function FindPivotTableByName(ByVal tableName As String) As PivotTable
    Dim ws As Worksheet
    Dim pvt As PivotTable
    For Each ws In ActiveWorkbook.Worksheets
        On Error Resume Next
        Set pvt = Nothing
        Set pvt = ws.PivotTables(tableName)
        On Error GoTo 0
        If Not pvt Is Nothing Then
            Set FindPivotTableByName = pvt
            Exit Function
        End If
    Next
End Function

' GPIVOT.2: lookup-or-raise, factored out once the "If pvt Is Nothing
' Then Err.Raise..." pair was about to be duplicated a 5th time
' (VlaPivotSetOrientation/VlaPivotAddValues already had their own
' copy; VlaPivotRefresh/VlaPivotDelete/VlaPivotSetShowDetail below all
' need the identical check) - pure internal cleanup, extensibility over
' a coin-toss per the owner's own standing default, no behavior change
' for any existing caller.
Private Function RequirePivotTableByName(ByVal tableName As String) As PivotTable
    Set RequirePivotTableByName = FindPivotTableByName(tableName)
    If RequirePivotTableByName Is Nothing Then
        RaiseRuntimeMsg "rt-pivot-not-found", "name", tableName
    End If
End Function

' G6 + G-PIVOT: the field-orientation half of pivot-rows/-columns/
' -filters (pareto.txt section 10) - one shared helper, not three,
' since ".PivotFields(name).Orientation = xl...Field" is the same
' per-field-name loop regardless of which orientation is being set.
' kind is a plain string ("row"/"column"/"filter"), not the raw XL
' constant - the row/column/filter DEFMACROs each bake their own kind
' in as a template literal (scripts/english.vla), so the xl*Field
' constants stay entirely inside this real host VBA and never cross
' into the VLA layer or the interpreter's own constant table at all.
' fields is a real runtime array - G6's own `array` primitive
' (VLA.bas's EmitExpr, VLA_Interpreter.bas's EvalExpr), the value a
' {name:text-list} slot resolves to - looped with a plain For Each,
' the only loop shape this needs (there is no expand-time answer here;
' this is genuine Excel-automation-time work, TABLESPEC-SCALE's own
' finding that a runtime list needs a runtime loop, not a macro one).
' GPIVOT.4: "hidden" added as a 4th kind - pivot-remove-field's own
' body, the missing inverse of pivot-rows/-columns/-filters (there was
' previously no way to take a field back OUT of the layout at all).
' Reuses this exact helper rather than a new one: xlHidden is just
' another XlPivotFieldOrientation member, the same per-field-name loop
' applies unchanged. Deliberately scoped to row/column/filter fields
' only, not value fields - a field that's a DATA field more than once
' (SalesPivot's own Revenue, summed AND counted, by deliberate design -
' see VlaPivotAddValues's own header note) raises the same "which one"
' question VlaPivotSort's own ambiguity check below has to answer, and
' unlike that check, it isn't yet known whether Orientation = xlHidden
' can even target ONE of several same-source data-field instances at
' all, or only removes them as a block - unverified against the real
' API, so left out rather than guessed at.
Public Sub VlaPivotSetOrientation(ByVal tableName As String, ByVal fields As Variant, ByVal kind As String)
    Dim pvt As PivotTable
    Set pvt = RequirePivotTableByName(tableName)
    Dim orient As XlPivotFieldOrientation
    Select Case kind
        Case "row": orient = xlRowField
        Case "column": orient = xlColumnField
        Case "filter": orient = xlPageField
        Case "hidden": orient = xlHidden
        Case Else
            RaiseRuntimeMsg "rt-pivot-orientation-unknown-kind", "kind", kind
    End Select
    Dim f As Variant
    For Each f In fields
        pvt.PivotFields(CStr(f)).Orientation = orient
    Next
End Sub

' G-PIVOT: shared by pivot-full's own "values of" clause AND the three
' standalone pivot-values-sum/-count/-avg verbs (scripts/english.vla) -
' func is a plain string ("sum"/"count"/"average"), not the raw XL
' constant, the exact same reasoning VlaPivotSetOrientation's own
' "kind" parameter already gives (the xlSum/xlCount/xlAverage mapping
' stays entirely inside this real host VBA, never crossing into the
' VLA layer or the interpreter's own constant table). fields is a real
' runtime array, same as that helper - {name:text-list} slots, more
' than one value field addable in one sentence ("Add Revenue, Cost to
' pivot X as a sum."), a deliberate generalization past pareto.txt's
' own pre-G6 single-field draft (owner's own call: "collapsing the
' semantics of two sentences into one is always a win for brevity").
' No guard against re-adding an already-present field/function pair -
' AddDataField's own native behavior (a second, distinctly-captioned
' data field, never a silent replace) is left exactly as Excel gives
' it, owner's own explicit call: users are allowed to abuse Excel to
' their own detriment, same as everywhere else in this project.
Public Sub VlaPivotAddValues(ByVal tableName As String, ByVal fields As Variant, ByVal func As String)
    Dim pvt As PivotTable
    Set pvt = RequirePivotTableByName(tableName)
    Dim aggFunc As XlConsolidationFunction
    Select Case func
        Case "sum": aggFunc = xlSum
        Case "count": aggFunc = xlCount
        Case "average": aggFunc = xlAverage
        Case Else
            RaiseRuntimeMsg "rt-pivot-values-unknown-function", "func", func
    End Select
    Dim f As Variant
    For Each f In fields
        pvt.AddDataField pvt.PivotFields(CStr(f)), , aggFunc
    Next
End Sub

' G-PIVOT: refresh one named pivot from its source data - narrower
' than "Refresh everything." (refresh-everything, scripts/english.vla,
' ActiveWorkbook.RefreshAll), which also touches external data
' connections/queries this verb never does.
Public Sub VlaPivotRefresh(ByVal tableName As String)
    RequirePivotTableByName(tableName).RefreshTable
End Sub

' G-PIVOT: refresh EVERY pivot table in the workbook - genuinely
' narrower than "Refresh everything.", not merely a rephrasing of it:
' this loop never touches an external data connection/query the way
' ActiveWorkbook.RefreshAll does, so it is the correct tool for "just
' the pivots," independent of which one pareto.txt happened to name.
Public Sub VlaPivotRefreshAll()
    Dim ws As Worksheet
    Dim pvt As PivotTable
    For Each ws In ActiveWorkbook.Worksheets
        For Each pvt In ws.PivotTables
            pvt.RefreshTable
        Next
    Next
End Sub

' G-PIVOT: delete a pivot table by name. TableRange2 (unlike
' TableRange1) includes the filter/page-field area, so clearing it
' removes the whole pivot, not just its data body.
Public Sub VlaPivotDelete(ByVal tableName As String)
    RequirePivotTableByName(tableName).TableRange2.Clear
End Sub

' G-PIVOT: pivot-collapse/pivot-expand's own shared helper - one
' Boolean, not a "kind" string like orientation/values, since
' ShowDetail is a genuine two-state property, not a three-way choice.
' fields is a real runtime array, same {name:text-list} convention as
' every other pivot verb - collapsing or expanding several fields in
' one sentence.
' FIRST VERSION set PivotField.ShowDetail directly - wrong, caught
' live (run-time error 1004, "Application-defined or object-defined
' error"), not by tracing. A pivot table Excel just created via
' CreatePivotTable defaults to Compact Form row layout, and in Compact
' Form expand/collapse is driven per PIVOT ITEM - the actual "+"/"-"
' control in the real UI - not per field; PivotField.ShowDetail only
' works reliably in Tabular/Outline layout, which nothing here sets.
' Looping PivotItems instead works in any layout, matching what the
' UI's own expand/collapse control actually does.
Public Sub VlaPivotSetShowDetail(ByVal tableName As String, ByVal fields As Variant, ByVal showDetail As Boolean)
    Dim pvt As PivotTable
    Set pvt = RequirePivotTableByName(tableName)
    Dim f As Variant, pi As PivotItem
    For Each f In fields
        For Each pi In pvt.PivotFields(CStr(f)).PivotItems
            pi.ShowDetail = showDetail
        Next
    Next
End Sub

' GPIVOT.3: pivot-layout's own helper. PivotTable.RowAxisLayout is a
' table-WIDE method - unlike every other pivot verb, it takes no field
' name at all, it re-forms EVERY row field to the given layout in one
' call, matching pareto.txt's own "Show pivot {n:text} in tabular
' form." (only the table is named). kind is a plain string ("compact"/
' "tabular"/"outline"), same convention as VlaPivotSetOrientation's own
' "kind" - the xl*Row constant mapping stays entirely inside this real
' host VBA.
' FIRST VERSION's own comment here claimed PivotField.LayoutForm could
' confirm this took effect - wrong, caught live: LayoutForm is type
' XlLayoutFormType (Microsoft's own docs - xlOutline/xlTabular ONLY, no
' compact option, defaults to xlTabular), a different, older per-field
' property entirely unrelated to XlLayoutRowType (this Sub's own
' parameter type). Per Microsoft's own RowAxisLayout reference, there
' is genuinely NO object-model readback for XlLayoutRowType state -
' VerifyReportChecks (VLA_Tests_Host.bas) does not attempt one; this
' verb's own live proof is necessarily thinner than every other pivot
' verb's, the same honestly-named limit VlaPivotRefresh's own comment
' already accepts (the call succeeds, not a visible state change).
' What VerifyReportChecks DOES check, because it's a real and
' answerable question: whether RowAxisLayout re-forming every row
' field at once disturbs an already-collapsed field's own per-
' PivotItem ShowDetail state (VlaPivotSetShowDetail's own header note
' has the full history of that exact class of interaction) - checked
' fresh after a layout switch, not assumed independent.
Public Sub VlaPivotSetRowLayout(ByVal tableName As String, ByVal kind As String)
    Dim pvt As PivotTable
    Set pvt = RequirePivotTableByName(tableName)
    Select Case kind
        Case "compact": pvt.RowAxisLayout xlCompactRow
        Case "tabular": pvt.RowAxisLayout xlTabularRow
        Case "outline": pvt.RowAxisLayout xlOutlineRow
        Case Else
            RaiseRuntimeMsg "rt-pivot-row-layout-unknown-kind", "kind", kind
    End Select
End Sub

' GPIVOT.3: pivot-subtotals-hide/-show's own helper. PivotField.Subtotals
' is a 12-element Boolean array (one per aggregation function; index 1
' is "Automatic", the only one Excel turns on by default and the only
' one any verb here ever sets) - set per-index, not by assigning a
' whole array, the documented-safe usage. {name:text-list} throughout,
' same convention as every other pivot verb. Not yet live-verified.
Public Sub VlaPivotSetSubtotals(ByVal tableName As String, ByVal fields As Variant, ByVal showSubtotals As Boolean)
    Dim pvt As PivotTable
    Set pvt = RequirePivotTableByName(tableName)
    Dim f As Variant, pf As PivotField, i As Integer
    For Each f In fields
        Set pf = pvt.PivotFields(CStr(f))
        For i = 1 To 12
            pf.Subtotals(i) = (showSubtotals And i = 1)
        Next
    Next
End Sub

' GPIVOT.3: pivot-blank-line-add/-remove's own helper.
' PivotField.LayoutBlankLine is a genuine two-state per-field Boolean -
' inserts a blank row after each item of that field, independent of
' both ShowDetail and Subtotals. {name:text-list} throughout, same
' convention as every other pivot verb. Not yet live-verified.
Public Sub VlaPivotSetBlankLine(ByVal tableName As String, ByVal fields As Variant, ByVal addBlankLine As Boolean)
    Dim pvt As PivotTable
    Set pvt = RequirePivotTableByName(tableName)
    Dim f As Variant
    For Each f In fields
        pvt.PivotFields(CStr(f)).LayoutBlankLine = addBlankLine
    Next
End Sub

' GPIVOT.4: pivot-rename's own helper - a straightforward rename,
' PivotTable names being unique WORKBOOK-wide (stricter than sheet
' names). No friendly duplicate-name guard, deliberately: N1.0's own
' VlaCheckSheetName/VlaCheckRangeName guards exist because Excel's
' NAMING SYNTAX validation is stricter than VLA's own text slots and
' surfaced a cryptic 1004 - a different problem from a duplicate-name
' COLLISION, which is what renaming actually risks. pivot-create itself
' (same TableName parameter, same collision surface) has never gotten
' one either; adding a guard only here would be new asymmetry, not
' consistency.
Public Sub VlaPivotRename(ByVal tableName As String, ByVal newName As String)
    RequirePivotTableByName(tableName).Name = newName
End Sub

' GPIVOT.4: pivot-clear's own helper. PivotTable.ClearTable resets
' every field to unassigned (plus subtotals/grouping/calculated items,
' none of which this grammar builds) but keeps the PivotTable object
' and its cache alive - a real, distinct verb from pivot-delete
' (VlaPivotDelete), which removes the object entirely.
Public Sub VlaPivotClear(ByVal tableName As String)
    RequirePivotTableByName(tableName).ClearTable
End Sub

' GPIVOT.4: pivot-source's own helper. FIRST DRAFT (never committed)
' assumed PivotCache.SourceData could just be reassigned directly -
' checked against real-world reports before writing any code, not
' assumed safe: direct SourceData reassignment is widely reported as
' unreliable in practice. The robust, documented path is a fresh
' PivotCache over the new range, swapped in via ChangePivotCache -
' ChangePivotCache errors only for an OLE DB/external-connection
' source, never a plain worksheet range, which is all this grammar
' ever builds. Auto-refreshes (owner's own call): changing a pivot's
' source and leaving it showing stale data until a separate "Refresh
' pivot X." sentence would surprise anyone who just repointed it.
Public Sub VlaPivotChangeSource(ByVal tableName As String, ByVal newRange As Range)
    Dim pvt As PivotTable
    Set pvt = RequirePivotTableByName(tableName)
    Dim newCache As PivotCache
    Set newCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=newRange)
    pvt.ChangePivotCache newCache
    pvt.RefreshTable
End Sub

' GPIVOT.4: pivot-sort's and pivot-sort-by-value's own shared helper -
' one Sub, not two, since PivotField.AutoSort is the same call either
' way, just a different Field-string argument. valueFieldName = ""
' means "sort by the field's own labels" (pivot-sort); a real value-
' field name means "sort by that value field's numbers"
' (pivot-sort-by-value).
' Two separate english-vla rules dispatch here (scripts/english.vla),
' not one rule with an omissible "by {v:text}" clause - checked against
' VLA_English.bas's own optional-literal mechanism before assuming it
' would work: G1's [optional] syntax matches ONE bare word, never a
' multi-word literal phrase carrying its own slot, so "by value field
' {v:text}" can't be made optional as a unit the way a single word can.
' Two ordinary rules is the proven-safe alternative - the exact same
' shape pivot-create/pivot-full already use for "same opening clause,
' different length" dispatch (TryPhrase's own end-of-sentence gate
' correctly falls through the shorter rule when the sentence continues
' past where it would end).
' Microsoft's own AutoSort reference is internally inconsistent - its
' prose says the Field parameter takes a field's unique SourceName, but
' its own example passes a data field's rendered CAPTION ("Sum of
' Sales"), not a bare source name. Reconciled, not just picked one:
' sorting a field by its OWN labels really does use its own SourceName
' (the prose is correct for that case); sorting by a VALUE field needs
' that value field's caption instead (the example is correct for that
' case) - two different, real scenarios, not a documentation error to
' route around. This collides directly with VlaPivotAddValues's own
' deliberately-unguarded design: a SourceName can legitimately back
' MORE than one value field at once (SalesPivot's own Revenue, summed
' AND counted). Owner's own explicit call for that case: raise a clear
' error naming the ambiguity rather than silently guessing an
' aggregation - this verb has no way to let a sentence specify WHICH
' one (that would need a caption-aware slot this grammar doesn't have),
' so the error is a genuine dead end today, named honestly as one
' rather than oversold as resolvable in-sentence.
Public Sub VlaPivotSort(ByVal tableName As String, ByVal fieldName As String, ByVal direction As String, ByVal valueFieldName As String)
    Dim pvt As PivotTable
    Set pvt = RequirePivotTableByName(tableName)
    Dim ord As XlSortOrder
    Select Case direction
        Case "ascending": ord = xlAscending
        Case "descending": ord = xlDescending
        Case Else
            RaiseRuntimeMsg "rt-pivot-sort-unknown-direction", "direction", direction
    End Select
    Dim pf As PivotField
    Set pf = pvt.PivotFields(fieldName)
    If Len(valueFieldName) = 0 Then
        pf.AutoSort ord, pf.SourceName
    Else
        Dim df As PivotField, matchCaption As String, matches As Long
        matches = 0
        For Each df In pvt.DataFields
            If StrComp(df.SourceName, valueFieldName, vbTextCompare) = 0 Then
                matchCaption = df.Name
                matches = matches + 1
            End If
        Next
        If matches = 0 Then
            RaiseRuntimeMsg "rt-pivot-sort-value-field-not-found", "field", valueFieldName, "table", tableName
        ElseIf matches > 1 Then
            RaiseRuntimeMsg "rt-pivot-sort-value-field-ambiguous", "field", valueFieldName, "table", tableName
        End If
        pf.AutoSort ord, matchCaption
    End If
End Sub

' =====================================================================
'  SEC.15: THE FORMULA SINK. Every formula a program writes goes
'  through VlaSetFormula, on both backends: VLA.bas's emitter writes
'  Call VlaSetFormula(obj, text) for exactly (set! (. obj formula) v),
'  where it wrote obj.Formula2 = text, and VLA_Interpreter's DynamicSet
'  calls it from its Case "formula". It writes through Formula2, as both
'  did (Range.Formula auto-inserts @ on anything that could spill), and
'  asks VlaFormulaEgress first: a text that calls a function which
'  reaches outside the workbook on its own, or holds a DDE link, is
'  refused by name (rt-formula-egress) and nothing is written. SEC.4
'  neutralizes a VALUE beginning with = because a value is data; a
'  formula that calls WEBSERVICE is never data, and an apostrophe would
'  turn the planting into text under a green row, so this refuses.
'
'  The scan lives here and not in VLA_Refers.bas, the reader, for the
'  reason SEC.4's rule is repeated in VlaTextInRange rather than called:
'  the compiled program runs in a workbook holding this module alone.
'  It is pure (text in, a name out, no Excel object), so the pure suite
'  and VlaWriteEgressGolden (VLA_Tests.bas) call it with no workbook,
'  and the core's egress.rs is held to its golden, scripts/egress.txt to
'  scripts/egress_golden.txt. The list below is the one place in the
'  VBA; tools/check_egress_golden.ps1 holds it equal to the core's and
'  every name reached by a case.
'
'  WHAT IS SCANNED: any text written through the formula member,
'  whatever it begins with (strict first: narrowing later to texts
'  beginning with = would loosen, where the reverse would break
'  programs under SD-4). A number, a date or a Boolean is written as
'  before. A cell handed as the text is read first, since Formula2
'  would have taken its value, so a planted text behind an apostrophe
'  cannot be laundered into a formula through a second cell.
'
'  THE RULE: a string literal ("...", a doubled quote inside), a quoted
'  name ('...') and a bracket group ([...], nested) are stepped over, so
'  ="WEBSERVICE(", ='WEBSERVICE'!A1 and =Sales[[#This Row],[a|b]] are
'  written; a run of identifier characters (letters, digits, _ and .,
'  anything past ASCII) followed by spaces, tabs or line breaks and then
'  ( is a call; the file prefixes _xlfn. and _xlws. are stripped from
'  its front, as many as stand there; the rest is compared without case
'  to the list, whole, so WEBSERVICE2(, MYWEBSERVICE( and WEBSERVICE.X(
'  are other functions and REGISTER.ID( is its own name; a name with no
'  ( after it is a defined name or a sheet, not a call. A | outside the
'  three is a DDE link (=cmd|'/c calc'!A0), reported as DDE. The first
'  hit in text order is named, in the list's upper-case spelling.
'
'  THE LIST, each with its reason (the full entry is SEC.15's in
'  docs/BETA_REARVIEW.md): WEBSERVICE, an HTTP request from a cell;
'  FILTERXML, the host's XML engine over text, whose handling of an
'  external entity is not verified here, and no use without WEBSERVICE;
'  HYPERLINK, a click opens a URL or a path, and a UNC path offers the
'  user's credentials; RTD, a registered COM server started by ProgID;
'  IMAGE, a picture fetched from a URL on calculation (Excel 365 and
'  Sheets); STOCKHISTORY, Excel's data service, the ticker sent out;
'  CALL, REGISTER and REGISTER.ID, XLM's DLL function by name; EXEC,
'  XLM starts a program; INITIATE, EXECUTE, POKE and REQUEST, XLM's DDE
'  conversation; SEND.MAIL, XLM mail; DDE, Calc's function of that
'  name; IMPORTXML, IMPORTDATA, IMPORTHTML, IMPORTRANGE and IMPORTFEED,
'  Sheets fetches a URL or another spreadsheet when the file opens
'  there; GOOGLEFINANCE and GOOGLETRANSLATE, Sheets sends the arguments
'  to Google. The XLM names run only on a macro sheet, which a program
'  never makes but a workbook may hold; the Sheets names are #NAME? in
'  Excel, and the .xlsx the core writes is the file Sheets imports.
' =====================================================================

' The function a formula text calls that reaches outside the workbook,
' in upper case, or DDE for a DDE link; "" when it holds none.
Public Function VlaFormulaEgress(ByVal formulaText As String) As String
    Dim listed As String
    listed = " webservice filterxml hyperlink rtd image stockhistory call register register.id exec initiate execute poke request send.mail dde importxml importdata importhtml importrange importfeed googlefinance googletranslate "
    Dim n As Long, i As Long, k As Long
    Dim ch As String, tok As String
    n = Len(formulaText)
    i = 1
    Do While i <= n
        ch = Mid$(formulaText, i, 1)
        If ch = """" Then
            i = EgressSkipDelimited(formulaText, i, """")
        ElseIf ch = "'" Then
            i = EgressSkipDelimited(formulaText, i, "'")
        ElseIf ch = "[" Then
            i = EgressSkipBrackets(formulaText, i)
        ElseIf ch = "|" Then
            VlaFormulaEgress = "DDE"
            Exit Function
        ElseIf EgressIsIdentChar(ch) Then
            k = i
            Do While k <= n
                If Not EgressIsIdentChar(Mid$(formulaText, k, 1)) Then Exit Do
                k = k + 1
            Loop
            tok = Fold(Mid$(formulaText, i, k - i))
            i = k
            Do While i <= n
                ch = Mid$(formulaText, i, 1)
                If ch <> " " And ch <> vbTab And ch <> vbCr And ch <> vbLf Then Exit Do
                i = i + 1
            Loop
            If Mid$(formulaText, i, 1) = "(" Then
                Do While Left$(tok, 6) = "_xlfn." Or Left$(tok, 6) = "_xlws."
                    tok = Mid$(tok, 7)
                Loop
                If InStr(1, listed, " " & tok & " ", vbBinaryCompare) > 0 Then
                    VlaFormulaEgress = UCase$(tok)
                    Exit Function
                End If
            End If
        Else
            i = i + 1
        End If
    Loop
End Function

' The position after the closing delimiter of the delimited text at j (a
' string literal's ", a quoted name's '), a doubled delimiter staying
' inside; the end of the text when it never closes.
Private Function EgressSkipDelimited(ByRef s As String, ByVal j As Long, ByVal d As String) As Long
    Dim k As Long, n As Long
    n = Len(s)
    k = j + 1
    Do While k <= n
        If Mid$(s, k, 1) <> d Then
            k = k + 1
        ElseIf Mid$(s, k + 1, 1) = d Then
            k = k + 2
        Else
            EgressSkipDelimited = k + 1
            Exit Function
        End If
    Loop
    EgressSkipDelimited = n + 1
End Function

' The position after the ] that closes the [ at i, nested brackets
' counted and a ' inside taken as the escape it is there; the end of the
' text when it never closes.
Private Function EgressSkipBrackets(ByRef s As String, ByVal i As Long) As Long
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
                EgressSkipBrackets = k
                Exit Function
            End If
        End If
    Loop
    EgressSkipBrackets = n + 1
End Function

' A letter, a digit, _ or ., or any character past ASCII (a letter in
' another alphabet, as a name or a sheet may hold).
Private Function EgressIsIdentChar(ByVal ch As String) As Boolean
    Dim a As Long
    If Len(ch) <> 1 Then Exit Function
    a = AscW(ch)
    If (a >= 65 And a <= 90) Or (a >= 97 And a <= 122) Or (a >= 48 And a <= 57) Then
        EgressIsIdentChar = True
    ElseIf ch = "_" Or ch = "." Then
        EgressIsIdentChar = True
    Else
        EgressIsIdentChar = (a < 0 Or a > 127)
    End If
End Function

' Range.Formula2 = text, the formula sink of both backends, after the
' scan. A refusal raised here lands on the ordinary VBA call stack of the
' program's own handler (the scaffold's vla-fail on both backends, and
' TryRuntimeHelper's native Case for a phrasebook that calls it, IN.15).
Public Sub VlaSetFormula(ByVal target As Object, ByVal v As Variant)
    Dim toWrite As Variant
    If IsObject(v) Then
        ' A cell handed as the text: Formula2 would take its value, so take
        ' it here and scan that. Anything else is written as it came.
        On Error Resume Next
        toWrite = v.Value
        If Err.Number <> 0 Then
            Err.Clear
            Set toWrite = v
        End If
        On Error GoTo 0
    Else
        toWrite = v
    End If
    If VarType(toWrite) = vbString Then
        Dim found As String
        found = VlaFormulaEgress(CStr(toWrite))
        If Len(found) > 0 Then
            RaiseRuntimeMsg "rt-formula-egress", "formula", EgressShown(CStr(toWrite)), "name", found, "cell", EgressCellLabel(target)
        End If
    End If
    target.Formula2 = toWrite
End Sub

' The formula as the refusal shows it: whole to 80 characters, then cut.
Private Function EgressShown(ByVal s As String) As String
    If Len(s) > 80 Then
        EgressShown = Left$(s, 80) & "..."
    Else
        EgressShown = s
    End If
End Function

' The target as the refusal names it, Sheet1!B3, or "the cell" when the
' target is not a range.
Private Function EgressCellLabel(ByVal target As Object) As String
    On Error Resume Next
    EgressCellLabel = target.Worksheet.Name & "!" & target.Address(False, False)
    On Error GoTo 0
    If Len(EgressCellLabel) = 0 Then EgressCellLabel = "the cell"
End Function

' === EN_RUNTIME INJECT BOUNDARY ======================================
'  Everything ABOVE this line is the injectable runtime: the exact
'  text placed into the user's workbook as module Frazaro_EN_Runtime. The
'  machinery BELOW is add-in-side only and never ships into user
'  workbooks (it references ThisWorkbook, which means the add-in).
' =====================================================================

' V5.3: put the runtime helpers into the user's workbook as module
' Frazaro_EN_Runtime, so the generated Frazaro_EN_<program> module's helper calls
' (vlacolor, vlacount, the VlaDict family...) resolve IN-PROJECT.
' The first standalone run proved the need: with the dev workbook
' closed, "Range(...).Font.Color = vlacolor(hot_pink)" raised "Sub
' or Function not defined" - VBA has no cross-project call, and the
' helpers lived only inside the (locked) add-in.
'
' Source acquisition is two-tiered because DEPLOY locks the shipped
' project, and component access to a protected project raises:
'   1. the live VLA_Runtime component of THIS project - the dev
'      workbook, or an unlocked add-in;
'   2. the very-hidden VLAr_Source sheet the BUILD writes into the
'      add-in while the dev project is still readable.
' The injected text stops at the INJECT BOUNDARY above, so user
' workbooks receive helpers only, never this machinery. Idempotent:
' Frazaro_EN_Runtime is rewritten on every Run, so helper fixes propagate
' with no stale copies. Requires the same Trust Center setting the
' program compile already needs - no new deployment ask.
Public Sub VlaInjectRuntime(ByVal targetWb As Workbook)
    ' IO.6: was "EN_Runtime" - renamed alongside OUT_MODULE (VLA_IDE.bas)
    ' to prevent, not just catch, a name collision with a foreign module.
    Const RT_MODULE As String = "Frazaro_EN_Runtime"
    ' Maiden-run incident #2 (owner report): when the TARGET is the
    ' project this module lives in - the dev workbook hosting its own
    ' run - the helpers are already in-project as VLA_Runtime, and
    ' injecting Frazaro_EN_Runtime beside them creates the very "Ambiguous
    ' name" this module exists to prevent. Skip - and self-heal: a
    ' Frazaro_EN_Runtime left there by the pre-guard bug is removed (best-
    ' effort; removal is deferred, hazard 11, gone by next compile).
    ' The dual-load state (add-in AND dev workbook open at once) can
    ' still recreate the collision from the add-in's side - that is
    ' the state DEPLOY already forbids for development.
    '
    ' IN.4: the same collision, one object-identity check away from
    ' this guard's reach - caught live, exporting a standalone copy of
    ' the DEV workbook (SaveCopyAs, VLA_Tests_Host.bas's own new
    ' TestExportSkeleton), not the add-in reinjecting into itself.
    ' targetWb Is ThisWorkbook is an OBJECT check: it catches "the exact
    ' same live workbook" but not "a freshly reopened copy that carries
    ' the same MODULES under a different object identity" - which is
    ' exactly what SaveCopyAs-then-reopen of a workbook that already
    ' has VLA_Runtime.bas as one of its own components produces. A real
    ' end-user's host workbook never has this problem (it never carries
    ' the engine's own source), but the owner's own dev/test workbook
    ' does, every time - and testing Export from the workbook every
    ' other feature in this project has been tested from is not a rare
    ' path to leave broken. Checked by NAME, not identity: does
    ' targetWb already have a component called "VLA_Runtime" - the
    ' same self-heal-and-skip this guard already does for the identity
    ' case, now covering the name case too.
    Dim hasEngineRuntime As Boolean
    On Error Resume Next
    Dim probeComp As Object
    Set probeComp = targetWb.VBProject.VBComponents("VLA_Runtime")
    On Error GoTo 0
    hasEngineRuntime = Not (probeComp Is Nothing)
    If targetWb Is ThisWorkbook Or hasEngineRuntime Then
        On Error Resume Next
        targetWb.VBProject.VBComponents.Remove targetWb.VBProject.VBComponents(RT_MODULE)
        On Error GoTo 0
        Exit Sub
    End If
    Dim code As String
    code = RuntimeSourceText()
    If Len(Trim$(code)) = 0 Then
        VLA_Messages.RaiseMsg "runtime-helpers-unreadable"
    End If
    code = "' Frazaro_EN_Runtime - Frazaro's runtime helpers, rewritten at every Run." & vbCrLf & _
           "' Generated: do not edit here; this workbook stays runnable" & vbCrLf & _
           "' even where the Frazaro add-in is not installed." & vbCrLf & code

    Dim proj As Object, comp As Object
    Set proj = targetWb.VBProject
    On Error Resume Next
    Set comp = proj.VBComponents(RT_MODULE)
    On Error GoTo 0
    If comp Is Nothing Then
        Set comp = proj.VBComponents.Add(1)      ' 1 = vbext_ct_StdModule
        comp.Name = RT_MODULE
    End If
    If comp.CodeModule.CountOfLines > 0 Then
        comp.CodeModule.DeleteLines 1, comp.CodeModule.CountOfLines
    End If
    comp.CodeModule.AddFromString code
End Sub

' S4.2: the runtime's own manifest - every Public vla-name this module
' provides, lowercase, underscore-stripped, space-delimited with a
' space on each end (built for InStr " name " membership tests). The
' resolve check compares generated code's helper calls against this
' list, so a dialect template's typo'd helper refuses at CHECK with
' its name - instead of dying at VBA's untrappable compile modal (the
' S4.0/S4.1 empirical verdict). Self-scanning RuntimeSourceText keeps
' it maintenance-free: a helper added to this module is in the
' manifest by existing, in both source tiers (live module or the
' build's VLAr_Source sheet). Machinery, below the inject boundary -
' user workbooks never need it.
Public Function VlaHelperManifest() As String
    Dim src As String
    On Error Resume Next
    src = RuntimeSourceText()
    On Error GoTo 0
    If Len(Trim$(src)) = 0 Then Exit Function   ' cannot read -> empty
                                                ' (the check treats an
                                                ' empty manifest as
                                                ' "cannot verify" and
                                                ' passes - soft-failure)
    Dim r As String
    Dim ln As Variant
    For Each ln In Split(Replace(src, vbCrLf, vbLf), vbLf)
        Dim t As String
        t = VLA_Identity.Fold(Trim$(CStr(ln)))
        Dim nm As String
        nm = ""
        If Left$(t, 11) = "public sub " Then
            nm = Mid$(t, 12)
        ElseIf Left$(t, 16) = "public function " Then
            nm = Mid$(t, 17)
        End If
        If Len(nm) > 0 Then
            Dim p As Long
            p = InStr(nm, "(")
            If p > 0 Then nm = Left$(nm, p - 1)
            nm = Trim$(Replace(nm, "_", ""))
            If Left$(nm, 3) = "vla" Then r = r & " " & nm
        End If
    Next
    ' OPTIMIZE.3 slice 5: the one helper this Frazaro provides from outside
    ' this module - the command behind "Optimize cell C1.", which runs the
    ' OPTIMIZE engine, so it lives beside that engine (VLA_Optimize) and
    ' only the interpreter reaches it (its own Case; the emitter refuses it
    ' by name). Named here, or the resolve check refuses every program for
    ' the phrasebook macro that calls it (live-caught, the owner's first
    ' add-in build with the slice).
    If Len(r) > 0 Then r = r & " vlaoptimizecell"
    If Len(r) > 0 Then VlaHelperManifest = r & " "
End Function

' The injectable text, public: the BUILD pours it into the add-in's
' VLAr_Source sheet, and the self-test pins its boundary.
Public Function VlaRuntimeInjectText() As String
    VlaRuntimeInjectText = RuntimeSourceText()
End Function

' Maiden-compile incident (D1.1's lesson re-learned, caught by the
' compile gate exactly as designed): these two lived here as module-
' level Consts, AFTER the procedures - "Only comments may appear
' after End Sub..." - and moving them to the module top was not an
' option either, because the top is the INJECTABLE region and a
' Const whose literal IS the boundary marker, sitting above the
' fence, would be found first by the trim search and truncate the
' inject text at itself. Functions carry both instead: legal at any
' position, and the marker is built by CONCATENATION so no source
' line except the fence itself ever contains the contiguous marker.
Private Function InjectBoundaryMark() As String
    InjectBoundaryMark = "=== EN_RUNTIME INJECT " & "BOUNDARY ==="
End Function

Public Function VlaRuntimeSheetName() As String
    VlaRuntimeSheetName = "VLAr_Source"
End Function

' The injectable text: tier 1 reads this very module (CodeModule
' excludes Attribute lines, includes Option Explicit); tier 2 reads
' the build-written sheet. Both trim at the boundary marker.
Private Function RuntimeSourceText() As String
    Dim s As String
    On Error Resume Next
    Dim cm As Object
    Set cm = ThisWorkbook.VBProject.VBComponents("VLA_Runtime").CodeModule
    If Not cm Is Nothing Then
        If cm.CountOfLines > 0 Then s = cm.Lines(1, cm.CountOfLines)
    End If
    On Error GoTo 0
    If Len(s) = 0 Then
        On Error Resume Next
        Dim sh As Worksheet
        Set sh = ThisWorkbook.Worksheets(VlaRuntimeSheetName())
        On Error GoTo 0
        If Not sh Is Nothing Then
            Dim n As Long, i As Long
            n = sh.Cells(sh.Rows.Count, 1).End(xlUp).Row
            For i = 1 To n
                s = s & CStr(sh.Cells(i, 1).Value) & vbCrLf
            Next
        End If
    End If
    RuntimeSourceText = TrimAtBoundary(s)
End Function

Private Function TrimAtBoundary(ByVal s As String) As String
    Dim p As Long
    p = InStr(1, s, InjectBoundaryMark())
    If p = 0 Then
        TrimAtBoundary = s                       ' soft-degrade: whole text
        Exit Function
    End If
    ' back up to the start of the marker's comment line
    Dim ls As Long
    ls = InStrRev(s, vbCrLf & "'", p)
    If ls = 0 Then ls = InStrRev(s, Chr$(10) & "'", p)
    If ls > 0 Then TrimAtBoundary = Left$(s, ls - 1) Else TrimAtBoundary = Left$(s, p - 1)
End Function
