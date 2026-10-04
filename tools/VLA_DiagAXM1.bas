Attribute VB_Name = "VLA_DiagAXM1"
Option Explicit

' Standalone diagnostic for AXM.1 - NOT part of the VLA project, and called by
' nothing in Frazaro. It measures how big a real model is when read as
' relations, so that numbers decide REFLECT's shape for AXM.8 (a cone per
' question, or a used-range scan per sheet) before any REFLECT code exists. It
' is a microscope, not a first draft of REFLECT: it follows no reference and
' parses no formula beyond the sheet name in front of a "!".
'
' WHERE IT RUNS. Import it into a NEW BLANK WORKBOOK that is never saved. Run
' every entry point from Excel's Macros dialog (Alt+F8, Macros in: All Open
' Workbooks), which finds a macro whichever workbook is active; its lines print
' to the Immediate window (Alt+F11, then Ctrl+G). Typed in the Immediate window,
' a command is looked up in the VBE's active project only, and a command that
' adds a workbook (AXM1Fixture's Workbooks.Add) or the opening of one (the
' model) can make that the new workbook's project: the next command typed then
' reads "Sub or Function not defined" (the owner's first run, 2026-10-01). It
' measures the workbook active in Excel, which you open yourself, READ-ONLY: it
' refuses a workbook that is not read-only, and refuses its own. The steps are in
' archive/axm1_live_steps.md; the predictions and the decision rule, written
' before the first run, are in docs/BETA_REARVIEW.md under AXM.1.
'
' WHAT IT NEVER DOES. It never writes into the workbook it measures: it only
' reads UsedRange, Value2, Formula, FormulaR1C1, SpecialCells, Precedents,
' Names, ListObjects and LinkSources. It never opens a file (an opened
' workbook's Workbook_Open would run), never changes an Application setting,
' makes no network call and has no Declare. The one workbook it writes is the
' fixture AXM1Fixture makes with Workbooks.Add. Every line it prints carries
' counts and times only: sheets by position (1 to N in the Sheets collection's
' order, hidden and chart sheets included), cones by the order you measure
' them. No file name, sheet name, address, formula text or value is printed.
' To see which sheet is number 7, type  ? ActiveWorkbook.Sheets(7).Name  in
' the Immediate window; the answer stays on your machine.
'
' ENTRY POINTS, each a macro with no arguments, for the Macros dialog:
'   AXM1Fixture   control 1. Builds a new workbook whose every count was
'                 derived by hand beforehand, measures it and four cones in
'                 it, and prints an AXM1K| verdict per line: ok, or WRONG with
'                 each field that differs.
'   AXM1Sample    control 2. The active workbook must be Frazaro Sample
'                 Data.xlsx (built by tools/build_examples.ps1), open
'                 read-only; its counts were derived from its own XML.
'   AXM1Measure   the model: every sheet, then the workbook.
'   AXM1Cone      one cone: the active cell of the measured workbook. Click a
'                 cell, then Alt+F8, AXM1Cone; ten times. AXM1ConeAgain 3,
'                 typed in the Immediate window, measures cone 3 again (a
'                 macro with an argument is not in the Macros dialog).
'   AXM1Report    the decision rule over the cones, then LAST the one line to
'                 paste back (AXM1|1|...).
'
' PER SHEET (AXM1S| lines, after a legend line):
'   - the used range's rows, columns and cells, and how far it runs past the
'     last non-blank cell (rows, columns): UsedRange counts formatting too;
'   - non-blank cells (Value2 not Empty) and formula cells, counted from the
'     arrays, never cell by cell. A cell is a formula when its Formula text
'     starts with "=" and is not also its value: a text cell typed as '=x
'     reads back "=x" from both. Excel's own count (SpecialCells) prints
'     beside it, taken after the timed reads so that it warms nothing;
'   - the time to read the used range as one array: Value2, Formula and
'     FormulaR1C1, three reads, in each of three passes, the mean printed;
'   - formula characters, the longest formula, the distinct R1C1 formulas
'     this sheet adds (what a parse-once cache would parse), the formulas
'     containing INDIRECT( or OFFSET( (what AXM.7 will refuse), its Tables;
'   - its sheet-qualified references, by a deliberately crude count: the
'     sheet name in front of each "!" outside a string or a quoted name -
'     Sales!, 'Q1 Data'!, 'Bob''s'!, a 3D span Jan:Dec! (every sheet
'     between), and [Book.xlsx]Sheet1! counted as another workbook. It reads
'     no address, no defined name, no structured reference and nothing inside
'     INDIRECT; those are AXM.7's, and everything built on this count is
'     blind to them. The count makes the sheet graph: sheet A links to sheet
'     B when a formula on A names B.
' PER WORKBOOK (AXM1W|): the totals, the largest used range, each pass's
' time, defined names (hidden, #REF!, naming another workbook) with Excel's own
' _xlfn. placeholders counted apart, Tables, and links to other workbooks
' (counts only, never their paths).
'
' A CONE (AXM1C|). Its floor is Excel's own Precedents of the cell, all
' levels, on the cell's own sheet only - Excel stops at the sheet boundary,
' and Microsoft documents DirectPrecedents as working only on the active
' sheet, which is why a cone's cell is one you click - clipped to the used
' range, each cell counted once, split into non-blank and formula cells.
' Excel raises 1004 when a cell has no precedents on its own sheet: a floor of
' 0. Its ceiling: the sheets the cone's formulas name, followed through the
' sheet graph, and every non-blank cell on them (the cell's own sheet whole
' when the graph leads back to it). Timed: Precedents itself; the floor read
' one area at a time (the cone shape's read) beside its sheet's whole read
' (the scan's).
'
' THE DECISION RULE (AXM1R|), written before the first run and settled with
' the owner 2026-09-30. A question's price is its reads plus its facts at
' 0.0328 ms a fact. Facts = non-blank cells + 2 x formula cells (a cell fact,
' a formula fact, and at least one refers fact); 0.0328 ms is DATALOG.14's
' measured scan (10,000 rows in 0.328 s), the cheapest thing any engine does
' with a fact. The budget is 2 s a question (DATALOG.12's settled "fine"). P
' prices the whole workbook; S, a cone as a scan (its own sheet and every
' sheet it reaches, read whole); C, a cone as a cone (its floor, read area by
' area). The first clause that holds is the verdict - the simplest shape that
' fits the budget:
'   1. P <= 2 s          SCAN THE WORKBOOK  certain
'   2. median S <= 2 s   SCAN PER SHEET     optimistic: the reach is blind to
'                                           names, structured references and
'                                           INDIRECT, so AXM.7 confirms it
'   3. median C <= 2 s   CONE               on the floor: AXM.7 sizes the
'                                           same ten cones before AXM.8
'   4. otherwise         RANGES             cones too big as cells: engines
'                                           take areas; AXM.8 is re-scoped
' The median is the typical question; the worst prints beside it, and does
' not decide.
'
' TRAPS HANDLED. A one-cell range's Value2 is a scalar, not an array
' (MakeGrid). A used range is read in chunks of at most 2^19 cells, so no
' array exhausts memory, and one past 2^25 cells is not read at all (its
' counts come from SpecialCells, and its line says so). Timer steps by about
' 16 ms, and is a Single, so after 18:12 by 1/128 s: every read is timed and
' summed, over three passes - a sum of many short readings is unbiased, but a
' sheet under ~50 ms is a few ticks, so trust the totals. A cone's timings
' repeat until a quarter of a second has passed. A first pass over 30 s is not
' repeated. VBA's And and Or never short-circuit, so no test leans on one.

Private Const CHUNK_CELLS As Long = 524288          ' 2^19: the most cells one array read holds
Private Const CEILING_CELLS As Double = 33554432#   ' 2^25: a used range past this is not read
Private Const GRID_CELLS As Double = 16777216#      ' 2^24: the largest box a cone is counted once per cell in
Private Const DISTINCT_CAP As Long = 1048576        ' the most distinct R1C1 formulas kept
Private Const PASSES As Long = 3
Private Const PASS_GUARD_S As Double = 30#          ' a first pass slower than this is not repeated
Private Const REP_SECONDS As Double = 0.25          ' a cone's timings repeat until this much has passed
Private Const REP_MAX As Long = 200
Private Const MAX_CONES As Long = 10
Private Const FACT_MS As Double = 0.0328            ' DATALOG.14: a 10,000-row scan in 0.328 s
Private Const BUDGET_S As Double = 2#               ' DATALOG.12's settled "fine"
Private Const SAMPLE_BOOK As String = "Frazaro Sample Data.xlsx"
Private Const FIXTURE_SHEETS As String = "Inputs,Calc,Q1 Data,Bob's,Hidden,VeryHidden,Empty,One,Inflated,FxChart"
Private Const SHEET_FIELDS As String = "kind|visibility|rows|cols|cells|non-blank|formulas|excel formulas|past last cell|formula chars|longest|distinct r1c1 new|qualified refs|sheets named|external refs|unresolved|indirect or offset|tables"
Private Const BOOK_FIELDS As String = "sheets|worksheets|charts|hidden|very hidden|cells|non-blank|formulas|excel formulas|largest sheet|largest cells|formula chars|longest|distinct r1c1|qualified refs|sheet links|external refs|unresolved|indirect or offset|names|hidden names|broken names|external names|tables|excel links|ole links|not read"
Private Const CONE_FIELDS As String = "precedents error|floor cells|floor non-blank|floor formulas|off-sheet areas|qualified refs|sheets named|external refs|reach sheets|back to own sheet|ceiling non-blank|floor facts|scan facts"

Private Type SheetStats
    kind As String              ' ws, macro, chart, or another TypeName
    vis As String               ' V visible, H hidden, VH very hidden
    prot As String              ' y when its contents are protected
    urAddr As String            ' re-read as the same cells on every pass; never printed
    readable As Boolean
    nRows As Double
    nCols As Double
    nCells As Double
    nonblank As Double
    formulas As Double
    xlFormulas As Double        ' Excel's own count; -1 when SpecialCells failed
    lastRow As Long             ' the last non-blank row and column, within the used range; 0 when none
    lastCol As Long
    tV As Double                ' seconds, summed over the passes
    tF As Double
    tR As Double
    chunks As Long
    fChars As Double
    longest As Long
    distinctNew As Double
    qual As Double
    named As Long
    ext As Double
    unres As Double
    nIndirect As Double
    nTables As Long
    remark As String
End Type

Private Type ConeStats
    floorCells As Double
    floorNon As Double
    floorF As Double
    qual As Double
    ext As Double
    unres As Double
End Type

' The measured workbook, kept between AXM1Measure, AXM1Cone and AXM1Report.
Private gMeasured As Boolean
Private gBookFullName As String     ' compared, never printed
Private gN As Long
Private gSheetIx As Object          ' sheet name -> position, compared without case, as Excel compares them
Private gNon() As Double
Private gFor() As Double
Private gReadS() As Double          ' the mean Value2 + Formula read of the whole used range, seconds
Private gUrAddr() As String         ' never printed
Private gExtRefs() As Double
Private gNotRead() As Boolean
Private gEdge() As Boolean          ' gEdge(a, b): a formula on sheet a names sheet b
Private gDistinctFull As Boolean
Private gVersion As String
Private gSheets As Long
Private gWorksheets As Long
Private gCells As Double
Private gW As Double                ' non-blank cells
Private gF As Double                ' formula cells
Private gLargestIx As Long
Private gLargestCells As Double
Private gValS As Double
Private gForS As Double
Private gRcS As Double
Private gDistinct As Double
Private gNotReadCount As Long
Private gConeLegend As Boolean
Private gConeSet(1 To MAX_CONES) As Boolean
Private gConeFloor(1 To MAX_CONES) As Double
Private gConeCeil(1 To MAX_CONES) As Double
Private gConeScanP(1 To MAX_CONES) As Double
Private gConeFloorP(1 To MAX_CONES) As Double
' the controls' tally
Private gChecks As Long
Private gOks As Long

' ===========================================================================
' Entry points
' ===========================================================================

Public Sub AXM1Fixture()
    Dim fx As Workbook
    gChecks = 0
    gOks = 0
    Set fx = BuildFixture()
    If fx Is Nothing Then
        Debug.Print "AXM1K|fixture|" & gOks & " of " & gChecks & " ok"
        Exit Sub
    End If
    MeasureWorkbook fx, "fixture"
    ' four cones, each measured with its own sheet active, as Precedents needs
    FixtureCone fx, "Calc", "D10", 1
    FixtureCone fx, "Calc", "G2", 2
    FixtureCone fx, "Calc", "H2", 3
    FixtureCone fx, "Q1 Data", "C2", 4
    ' the probe, checked against nothing: Precedents of a cell whose sheet is NOT active
    fx.Worksheets("Inputs").Activate
    MeasureCone fx.Worksheets("Calc").Range("D10"), 0, "", "probe"
    fx.Worksheets("Calc").Activate
    Debug.Print "AXM1K|fixture|" & gOks & " of " & gChecks & " ok"
    Debug.Print "=== AXM1 fixture done. AXM1Report (Alt+F8) next checks the rule's arithmetic; then close the fixture without saving. ==="
End Sub

Public Sub AXM1Sample()
    Dim wb As Workbook
    If Not TargetOk(wb, True) Then Exit Sub
    gChecks = 0
    gOks = 0
    MeasureWorkbook wb, "sample"
    Debug.Print "AXM1K|sample|" & gOks & " of " & gChecks & " ok"
End Sub

Public Sub AXM1Measure()
    Dim wb As Workbook
    If Not TargetOk(wb, False) Then Exit Sub
    MeasureWorkbook wb, ""
    Debug.Print "=== AXM1 measured. Next: click a cell in the model, Alt+F8, AXM1Cone - ten cells - then AXM1Report. ==="
End Sub

Public Sub AXM1Cone()
    ConeAt 0
End Sub

Public Sub AXM1ConeAgain(ByVal slot As Long)
    ConeAt slot
End Sub

Private Sub ConeAt(ByVal slot As Long)
    Dim c As Range, k As Long
    If Not gMeasured Then
        Debug.Print "AXM1 refused: nothing is measured yet. Run AXM1Measure first; the cones use its sheet counts."
        Exit Sub
    End If
    If ActiveWorkbook Is Nothing Then
        Debug.Print "AXM1 refused: no workbook is active. Click a cell in the measured workbook, then try again."
        Exit Sub
    End If
    If ActiveWorkbook.FullName <> gBookFullName Then
        Debug.Print "AXM1 refused: the active workbook is not the one last measured. Click a cell in it, then try again."
        Exit Sub
    End If
    On Error Resume Next
    Set c = ActiveCell
    On Error GoTo 0
    If c Is Nothing Then
        Debug.Print "AXM1 refused: no cell is active (a chart sheet?). Click a cell on a worksheet, then try again."
        Exit Sub
    End If
    If slot = 0 Then
        For k = 1 To MAX_CONES
            If Not gConeSet(k) Then Exit For
        Next k
        If k > MAX_CONES Then
            Debug.Print "AXM1 refused: " & MAX_CONES & " cones are measured. AXM1ConeAgain <n> measures cone n again; AXM1Report reports."
            Exit Sub
        End If
    ElseIf slot < 1 Or slot > MAX_CONES Then
        Debug.Print "AXM1 refused: a cone is numbered 1 to " & MAX_CONES & "."
        Exit Sub
    Else
        k = slot
    End If
    MeasureCone c, k, "", CStr(k)
End Sub

Public Sub AXM1Report()
    Dim k As Long, n As Long, sp() As Double, fp() As Double
    Dim pAll As Double, medS As Double, worstS As Double, medF As Double, worstF As Double
    Dim clause As Long, verdict As String, caveat As String, floorList As String, ceilList As String
    If Not gMeasured Then
        Debug.Print "AXM1 refused: nothing is measured yet. Run AXM1Measure (or a control) first."
        Exit Sub
    End If
    ReDim sp(1 To MAX_CONES)
    ReDim fp(1 To MAX_CONES)
    For k = 1 To MAX_CONES
        If gConeSet(k) Then
            n = n + 1
            sp(n) = gConeScanP(k)
            fp(n) = gConeFloorP(k)
            floorList = floorList & ";" & Cnt(gConeFloor(k))
            ceilList = ceilList & ";" & Cnt(gConeCeil(k))
        Else
            floorList = floorList & ";-"
            ceilList = ceilList & ";-"
        End If
    Next k
    If n = 0 Then
        Debug.Print "AXM1 refused: no cone is measured yet. Click a cell in the measured workbook, then AXM1Cone."
        Exit Sub
    End If
    pAll = gValS + gForS + (gW + 2 * gF) * FACT_MS / 1000#
    medS = MedianOf(sp, n)
    worstS = MaxOf(sp, n)
    medF = MedianOf(fp, n)
    worstF = MaxOf(fp, n)
    ' The rule, as written before the first run: the simplest shape that fits the budget.
    If pAll <= BUDGET_S Then
        clause = 1
        verdict = "SCAN THE WORKBOOK"
        caveat = "certain"
    ElseIf medS <= BUDGET_S Then
        clause = 2
        verdict = "SCAN PER SHEET"
        caveat = "optimistic: the reach is blind to names, structured references and INDIRECT; AXM.7 confirms"
    ElseIf medF <= BUDGET_S Then
        clause = 3
        verdict = "CONE"
        caveat = "on the floor: AXM.7 sizes the same cones before AXM.8 commits"
    Else
        clause = 4
        verdict = "RANGES"
        caveat = "cones too big as cells: engines take areas; AXM.8 is re-scoped"
    End If
    If n < MAX_CONES Then caveat = caveat & "; only " & n & " of " & MAX_CONES & " cones"
    If gNotReadCount > 0 Then caveat = caveat & "; " & gNotReadCount & " sheet(s) not read, so P and S are lower bounds"
    Debug.Print "AXM1R|cones|whole workbook s (P)|median scan s (S)|worst scan s|median floor s (C)|worst floor s|budget s|clause|verdict|caveat"
    Debug.Print "AXM1R|" & n & "|" & S3(pAll) & "|" & S3(medS) & "|" & S3(worstS) & "|" & S3(medF) & "|" & S3(worstF) & "|" & _
                Format$(BUDGET_S, "0") & "|" & clause & "|" & verdict & "|" & caveat
    ' LAST: the one line to paste back.
    Debug.Print "AXM1|1|" & gVersion & "|sheets " & gSheets & " (" & gWorksheets & " worksheets)|cells " & Cnt(gCells) & _
                "|non-blank " & Cnt(gW) & "|formulas " & Cnt(gF) & "|largest " & Cnt(gLargestCells) & " (sheet " & _
                gLargestIx & ")|read s: value2 " & S3(gValS) & ", formula " & S3(gForS) & ", r1c1 " & S3(gRcS) & _
                "|distinct r1c1 " & Cnt(gDistinct) & "|cones " & n & "|floor " & Mid$(floorList, 2) & "|ceiling " & _
                Mid$(ceilList, 2) & "|P " & S3(pAll) & " s|S median " & S3(medS) & " s, worst " & S3(worstS) & _
                " s|C median " & S3(medF) & " s, worst " & S3(worstF) & " s|clause " & clause & ": " & verdict & _
                "|not read " & gNotReadCount
End Sub

' ===========================================================================
' The workbook: every sheet, three passes, then the totals
' ===========================================================================

Private Function TargetOk(ByRef wb As Workbook, ByVal wantSample As Boolean) As Boolean
    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        Debug.Print "AXM1 refused: no workbook is active. Click into the workbook to measure, then run the macro again (Alt+F8)."
        Exit Function
    End If
    If wb Is ThisWorkbook Then
        Debug.Print "AXM1 refused: the active workbook is this diagnostic's own. Click into the workbook to measure, then run the macro again (Alt+F8)."
        Exit Function
    End If
    If wantSample Then
        If wb.Name <> SAMPLE_BOOK Then
            Debug.Print "AXM1 refused: AXM1Sample measures " & SAMPLE_BOOK & " only. Open it read-only and click into it."
            Exit Function
        End If
    End If
    If Not wb.ReadOnly Then
        Debug.Print "AXM1 refused: the active workbook is not open read-only. Close it, then File > Open, select it, and choose Open Read-Only from the arrow beside Open."
        Exit Function
    End If
    TargetOk = True
End Function

Private Sub MeasureWorkbook(ByVal wb As Workbook, ByVal expectKey As String)
    Dim n As Long, i As Long, p As Long, passesRun As Long
    Dim st() As SheetStats
    Dim passS(1 To PASSES) As Double
    Dim dict As Object, ws As Worksheet
    Dim hits() As Boolean

    n = wb.Sheets.Count
    ResetState n
    ReDim st(1 To n)
    Set dict = CreateObject("Scripting.Dictionary")
    gVersion = "Excel " & Application.Version & " " & Bitness()
    Debug.Print "=== AXM1 measure, " & TargetLabel(expectKey) & ": " & n & " sheets, " & gVersion & ", calculation " & _
                CalcWord() & ", " & Format$(Now, "yyyy-mm-dd hh:nn") & " ==="
    Debug.Print "    Timer steps by ~16 ms: a sheet under ~50 ms is a few ticks; trust the totals. Leave Excel alone while it runs."

    SurveySheets wb, st, n

    ' The timed passes. The first also counts; the others only read again.
    For p = 1 To PASSES
        If p > 1 Then
            If passS(1) > PASS_GUARD_S Then Exit For
        End If
        For i = 1 To n
            If st(i).readable Then
                Set ws = wb.Sheets(i)
                If p = 1 Then ReDim hits(1 To n)
                passS(p) = passS(p) + SafeReadSheet(ws, st(i), (p = 1), dict, hits)
                If p = 1 Then RecordEdges i, hits, st(i)
            End If
            DoEvents
        Next i
        passesRun = p
    Next p

    CrossCheck wb, st, n
    For i = 1 To n
        gNon(i) = st(i).nonblank
        gFor(i) = st(i).formulas
        gUrAddr(i) = st(i).urAddr
        gExtRefs(i) = st(i).ext
        If IsGrid(st(i)) Then
            If st(i).readable Then
                gReadS(i) = (st(i).tV + st(i).tF) / passesRun
            Else
                gNotRead(i) = True
            End If
        End If
    Next i

    Debug.Print "AXM1S|sheet|kind|vis|prot|rows|cols|cells|non-blank|formulas|excel formulas|past last cell (rows,cols)|" & _
                "value2 ms|formula ms|r1c1 ms|chunks|formula chars|longest|distinct r1c1 new|qualified refs|sheets named|" & _
                "external refs|unresolved|indirect or offset|tables|remark"
    For i = 1 To n
        PrintSheet i, st(i), passesRun
        If Len(expectKey) > 0 Then Check "sheet " & i, SheetKey(st(i)), SheetWant(expectKey, i), SHEET_FIELDS
    Next i
    BookLine wb, st, n, passS, passesRun, expectKey
    gBookFullName = wb.FullName
    gMeasured = True
    Set dict = Nothing
End Sub

Private Sub ResetState(ByVal n As Long)
    Dim k As Long
    gMeasured = False
    gBookFullName = ""
    gN = n
    ReDim gNon(1 To n)
    ReDim gFor(1 To n)
    ReDim gReadS(1 To n)
    ReDim gUrAddr(1 To n)
    ReDim gExtRefs(1 To n)
    ReDim gNotRead(1 To n)
    ReDim gEdge(1 To n, 1 To n)
    Set gSheetIx = CreateObject("Scripting.Dictionary")
    gSheetIx.CompareMode = vbTextCompare
    gDistinctFull = False
    gConeLegend = False
    For k = 1 To MAX_CONES
        gConeSet(k) = False
        gConeFloor(k) = 0
        gConeCeil(k) = 0
        gConeScanP(k) = 0
        gConeFloorP(k) = 0
    Next k
End Sub

' Kind, visibility, protection, Tables and the used range's size, before any
' timed read.
Private Sub SurveySheets(ByVal wb As Workbook, ByRef st() As SheetStats, ByVal n As Long)
    Dim i As Long, sh As Object, ws As Worksheet, ur As Range
    For i = 1 To n
        Set sh = wb.Sheets(i)
        st(i).kind = SheetKind(sh)
        st(i).vis = VisCode(sh.Visible)
        If Not gSheetIx.Exists(sh.Name) Then gSheetIx.Add sh.Name, i
        If IsGrid(st(i)) Then
            Set ws = sh
            If ws.ProtectContents Then st(i).prot = "y" Else st(i).prot = "n"
            If st(i).kind = "ws" Then st(i).nTables = ws.ListObjects.Count
            Set ur = ws.UsedRange
            st(i).urAddr = ur.Address
            st(i).nRows = ur.Rows.Count
            st(i).nCols = ur.Columns.Count
            st(i).nCells = st(i).nRows * st(i).nCols
            st(i).readable = (st(i).nCells <= CEILING_CELLS)
            If Not st(i).readable Then st(i).remark = "not read: past 2^25 cells, counts from SpecialCells"
        Else
            st(i).prot = "-"
            st(i).remark = "no used range"
        End If
    Next i
End Sub

Private Function SafeReadSheet(ByVal ws As Worksheet, ByRef st As SheetStats, ByVal counting As Boolean, _
                               ByVal dict As Object, ByRef hits() As Boolean) As Double
    On Error GoTo Fail
    SafeReadSheet = ReadSheet(ws, st, counting, dict, hits)
    Exit Function
Fail:
    st.readable = False
    st.remark = "read failed: error " & Err.Number & " (" & Err.Description & "), counts from SpecialCells"
End Function

' One pass over one sheet's used range, in chunks of whole rows: three timed
' reads per chunk, then (first pass only) the counting, untimed.
Private Function ReadSheet(ByVal ws As Worksheet, ByRef st As SheetStats, ByVal counting As Boolean, _
                           ByVal dict As Object, ByRef hits() As Boolean) As Double
    Dim ur As Range, blk As Range, nR As Long, nC As Long, per As Long, r1 As Long, r2 As Long
    Dim v As Variant, f As Variant, rc As Variant, t0 As Double, t As Double, total As Double
    Set ur = ws.Range(st.urAddr)
    nR = ur.Rows.Count
    nC = ur.Columns.Count
    per = RowsPerChunk(nC)
    r1 = 1
    Do While r1 <= nR
        r2 = r1 + per - 1
        If r2 > nR Then r2 = nR
        t0 = Timer
        Set blk = ws.Range(ur.Cells(r1, 1), ur.Cells(r2, nC))
        v = blk.Value2
        t = SecondsSince(t0)
        st.tV = st.tV + t
        total = total + t
        t0 = Timer
        f = blk.Formula
        t = SecondsSince(t0)
        st.tF = st.tF + t
        total = total + t
        t0 = Timer
        rc = blk.FormulaR1C1
        t = SecondsSince(t0)
        st.tR = st.tR + t
        total = total + t
        If counting Then
            st.chunks = st.chunks + 1
            MakeGrid v
            MakeGrid f
            MakeGrid rc
            CountChunk v, f, rc, r1 - 1, st, dict, hits
        End If
        v = Empty
        f = Empty
        rc = Empty
        r1 = r2 + 1
    Loop
    ReadSheet = total
End Function

Private Sub CountChunk(ByRef v As Variant, ByRef f As Variant, ByRef rc As Variant, ByVal rowOff As Long, _
                       ByRef st As SheetStats, ByVal dict As Object, ByRef hits() As Boolean)
    Dim i As Long, j As Long, nR As Long, nC As Long, fs As String
    nR = UBound(v, 1)
    nC = UBound(v, 2)
    For i = 1 To nR
        For j = 1 To nC
            If Not IsEmpty(v(i, j)) Then
                st.nonblank = st.nonblank + 1
                st.lastRow = rowOff + i
                If j > st.lastCol Then st.lastCol = j
            End If
            If Len(f(i, j)) > 1 Then
                fs = f(i, j)
                If IsFormulaText(fs, v(i, j)) Then
                    st.formulas = st.formulas + 1
                    st.fChars = st.fChars + Len(fs)
                    If Len(fs) > st.longest Then st.longest = Len(fs)
                    If InStr(1, fs, "INDIRECT(", vbBinaryCompare) > 0 Then
                        st.nIndirect = st.nIndirect + 1
                    ElseIf InStr(1, fs, "OFFSET(", vbBinaryCompare) > 0 Then
                        st.nIndirect = st.nIndirect + 1
                    End If
                    ScanRefs fs, hits, st.qual, st.ext, st.unres
                    AddDistinct dict, CStr(rc(i, j)), st
                End If
            End If
        Next j
    Next i
End Sub

Private Sub AddDistinct(ByVal dict As Object, ByVal rcText As String, ByRef st As SheetStats)
    If gDistinctFull Then Exit Sub
    If dict.Exists(rcText) Then Exit Sub
    If dict.Count >= DISTINCT_CAP Then
        gDistinctFull = True
        Exit Sub
    End If
    dict.Add rcText, Empty
    st.distinctNew = st.distinctNew + 1
End Sub

Private Sub RecordEdges(ByVal host As Long, ByRef hits() As Boolean, ByRef st As SheetStats)
    Dim t As Long
    For t = 1 To gN
        If hits(t) Then
            If t <> host Then
                gEdge(host, t) = True
                st.named = st.named + 1
            End If
        End If
    Next t
End Sub

' Excel's own formula count, taken after the timed reads; and, for a sheet
' that was not read, Excel's counts stand in for the arrays'.
Private Sub CrossCheck(ByVal wb As Workbook, ByRef st() As SheetStats, ByVal n As Long)
    Dim i As Long, ws As Worksheet, ur As Range, consts As Double
    For i = 1 To n
        If IsGrid(st(i)) Then
            Set ws = wb.Sheets(i)
            Set ur = ws.Range(st(i).urAddr)
            st(i).xlFormulas = SpecialCount(ur, xlCellTypeFormulas)
            If Not st(i).readable Then
                consts = SpecialCount(ur, xlCellTypeConstants)
                st(i).formulas = 0
                st(i).nonblank = 0
                If st(i).xlFormulas > 0 Then st(i).formulas = st(i).xlFormulas
                st(i).nonblank = st(i).formulas
                If consts > 0 Then st(i).nonblank = st(i).nonblank + consts
            End If
        End If
    Next i
End Sub

Private Function SpecialCount(ByVal ur As Range, ByVal which As Long) As Double
    Dim sc As Range, en As Long
    On Error Resume Next
    Set sc = ur.SpecialCells(which)
    en = Err.Number
    On Error GoTo 0
    If en = 1004 Then
        SpecialCount = 0                    ' "No cells were found."
    ElseIf en <> 0 Then
        SpecialCount = -1
    Else
        SpecialCount = CDbl(sc.CountLarge)
    End If
End Function

Private Sub PrintSheet(ByVal i As Long, ByRef st As SheetStats, ByVal passesRun As Long)
    Dim msV As String, msF As String, msR As String
    If Not IsGrid(st) Then
        Debug.Print "AXM1S|" & i & "|" & st.kind & "|" & st.vis & "|-|-|-|-|-|-|-|-|-|-|-|-|-|-|-|-|-|-|-|-|-|" & st.remark
        Exit Sub
    End If
    If st.readable Then
        msV = Ms(st.tV / passesRun)
        msF = Ms(st.tF / passesRun)
        msR = Ms(st.tR / passesRun)
    Else
        msV = "-"
        msF = "-"
        msR = "-"
    End If
    Debug.Print "AXM1S|" & i & "|" & st.kind & "|" & st.vis & "|" & st.prot & "|" & Cnt(st.nRows) & "|" & Cnt(st.nCols) & "|" & _
                Cnt(st.nCells) & "|" & Cnt(st.nonblank) & "|" & Cnt(st.formulas) & "|" & Cnt(st.xlFormulas) & "|" & _
                PastText(st) & "|" & msV & "|" & msF & "|" & msR & "|" & st.chunks & "|" & Cnt(st.fChars) & "|" & _
                Cnt(st.longest) & "|" & Cnt(st.distinctNew) & "|" & Cnt(st.qual) & "|" & st.named & "|" & Cnt(st.ext) & "|" & _
                Cnt(st.unres) & "|" & Cnt(st.nIndirect) & "|" & st.nTables & "|" & st.remark
End Sub

Private Function SheetKey(ByRef st As SheetStats) As String
    If Not IsGrid(st) Then
        SheetKey = st.kind & "|" & st.vis
        Exit Function
    End If
    SheetKey = st.kind & "|" & st.vis & "|" & st.nRows & "|" & st.nCols & "|" & st.nCells & "|" & st.nonblank & "|" & _
               st.formulas & "|" & st.xlFormulas & "|" & PastText(st) & "|" & st.fChars & "|" & st.longest & "|" & _
               st.distinctNew & "|" & st.qual & "|" & st.named & "|" & st.ext & "|" & st.unres & "|" & st.nIndirect & "|" & _
               st.nTables
End Function

Private Function PastText(ByRef st As SheetStats) As String
    If Not st.readable Then
        PastText = "?"
    ElseIf st.nonblank = 0 Then
        PastText = "-"
    Else
        PastText = (st.nRows - st.lastRow) & "," & (st.nCols - st.lastCol)
    End If
End Function

Private Sub BookLine(ByVal wb As Workbook, ByRef st() As SheetStats, ByVal n As Long, ByRef passS() As Double, _
                     ByVal passesRun As Long, ByVal expectKey As String)
    Dim i As Long, nWs As Long, nCharts As Long, nHid As Long, nVHid As Long, nNotRead As Long
    Dim tCells As Double, tNon As Double, tFor As Double, tXl As Double, xlBad As Boolean
    Dim tChars As Double, tLongest As Long, tDistinct As Double, tQual As Double, tLinks As Double
    Dim tExtRefs As Double, tUnres As Double, tInd As Double, tTables As Long
    Dim bigIx As Long, bigCells As Double, sV As Double, sF As Double, sR As Double
    Dim nmAll As Long, nmHid As Long, nmBroken As Long, nmExt As Long, nmXlfn As Long, lkX As Long, lkO As Long
    Dim passText As String, got As String

    For i = 1 To n
        If st(i).kind = "chart" Then nCharts = nCharts + 1
        If st(i).vis = "H" Then nHid = nHid + 1
        If st(i).vis = "VH" Then nVHid = nVHid + 1
        If IsGrid(st(i)) Then
            nWs = nWs + 1
            tCells = tCells + st(i).nCells
            tNon = tNon + st(i).nonblank
            tFor = tFor + st(i).formulas
            If st(i).xlFormulas < 0 Then xlBad = True Else tXl = tXl + st(i).xlFormulas
            tChars = tChars + st(i).fChars
            If st(i).longest > tLongest Then tLongest = st(i).longest
            tDistinct = tDistinct + st(i).distinctNew
            tQual = tQual + st(i).qual
            tLinks = tLinks + st(i).named
            tExtRefs = tExtRefs + st(i).ext
            tUnres = tUnres + st(i).unres
            tInd = tInd + st(i).nIndirect
            tTables = tTables + st(i).nTables
            If st(i).nCells > bigCells Then
                bigCells = st(i).nCells
                bigIx = i
            End If
            If st(i).readable Then
                sV = sV + st(i).tV / passesRun
                sF = sF + st(i).tF / passesRun
                sR = sR + st(i).tR / passesRun
            Else
                nNotRead = nNotRead + 1
            End If
        End If
    Next i
    CountNames wb, nmAll, nmHid, nmBroken, nmExt, nmXlfn
    lkX = LinkCount(wb, xlExcelLinks)
    lkO = LinkCount(wb, xlOLELinks)
    For i = 1 To passesRun
        If i > 1 Then passText = passText & " / "
        passText = passText & Format$(passS(i), "0.000")
    Next i

    Debug.Print "AXM1W|sheets|worksheets|charts|hidden|very hidden|cells|non-blank|formulas|excel formulas|" & _
                "largest used range (sheet: cells)|value2 s|formula s|r1c1 s|scan s (value2+formula)|passes s|" & _
                "formula chars|longest|distinct r1c1|qualified refs|sheet links|external refs|unresolved|" & _
                "indirect or offset|names|hidden names|broken names|external names|_xlfn names|tables|excel links|" & _
                "ole links|not read"
    Debug.Print "AXM1W|" & n & "|" & nWs & "|" & nCharts & "|" & nHid & "|" & nVHid & "|" & Cnt(tCells) & "|" & _
                Cnt(tNon) & "|" & Cnt(tFor) & "|" & IIf(xlBad, "err", Cnt(tXl)) & "|" & bigIx & ": " & Cnt(bigCells) & "|" & _
                S3(sV) & "|" & S3(sF) & "|" & S3(sR) & "|" & S3(sV + sF) & "|" & passText & "|" & Cnt(tChars) & "|" & _
                Cnt(tLongest) & "|" & Cnt(tDistinct) & IIf(gDistinctFull, " (capped)", "") & "|" & Cnt(tQual) & "|" & _
                Cnt(tLinks) & "|" & Cnt(tExtRefs) & "|" & Cnt(tUnres) & "|" & Cnt(tInd) & "|" & nmAll & "|" & nmHid & "|" & _
                nmBroken & "|" & nmExt & "|" & nmXlfn & "|" & tTables & "|" & LinkText(lkX) & "|" & LinkText(lkO) & "|" & _
                nNotRead
    ' The check leaves the _xlfn names out: whether Excel makes one depends on
    ' its version (control 1's INDIRECT gets an _xlfn.SINGLE in Excel 365 only).
    If Len(expectKey) > 0 Then
        got = n & "|" & nWs & "|" & nCharts & "|" & nHid & "|" & nVHid & "|" & tCells & "|" & tNon & "|" & tFor & "|" & _
              IIf(xlBad, -1, tXl) & "|" & bigIx & "|" & bigCells & "|" & tChars & "|" & tLongest & "|" & tDistinct & "|" & _
              tQual & "|" & tLinks & "|" & tExtRefs & "|" & tUnres & "|" & tInd & "|" & nmAll & "|" & nmHid & "|" & _
              nmBroken & "|" & nmExt & "|" & tTables & "|" & lkX & "|" & lkO & "|" & nNotRead
        Check "workbook", got, BookWant(expectKey), BOOK_FIELDS
    End If

    gSheets = n
    gWorksheets = nWs
    gCells = tCells
    gW = tNon
    gF = tFor
    gLargestIx = bigIx
    gLargestCells = bigCells
    gValS = sV
    gForS = sF
    gRcS = sR
    gDistinct = tDistinct
    gNotReadCount = nNotRead
End Sub

' Every defined name: hidden ones, those whose Refers To holds #REF! (or
' cannot be read), and those naming another workbook. An _xlfn. name is
' counted apart: Excel's own placeholder, Refers To =#NAME?, for a function or
' operator older Excel lacks (_xlfn.SINGLE is the @ of implicit intersection,
' which Excel 365 adds to a formula written through Formula), not a name anyone
' defined (control 1's run, 2026-10-01).
Private Sub CountNames(ByVal wb As Workbook, ByRef nAll As Long, ByRef nHid As Long, ByRef nBroken As Long, _
                       ByRef nOutside As Long, ByRef nXlfn As Long)
    Dim nm As Object, rt As String, en As Long, scratch() As Boolean
    Dim q As Double, e As Double, u As Double
    ReDim scratch(1 To gN)
    For Each nm In wb.Names
        If StrComp(Left$(nm.Name, 6), "_xlfn.", vbTextCompare) = 0 Then
            nXlfn = nXlfn + 1
        Else
            nAll = nAll + 1
            If Not nm.Visible Then nHid = nHid + 1
            rt = ""
            On Error Resume Next
            rt = nm.RefersTo
            en = Err.Number
            On Error GoTo 0
            If en <> 0 Then
                nBroken = nBroken + 1
            ElseIf InStr(1, rt, "#REF!", vbBinaryCompare) > 0 Then
                nBroken = nBroken + 1
            Else
                e = 0
                ScanRefs rt, scratch, q, e, u
                If e > 0 Then nOutside = nOutside + 1
            End If
        End If
    Next nm
End Sub

' How many workbooks (or OLE and DDE sources) the workbook links to. Counted
' only: the paths are never read into anything that prints.
Private Function LinkCount(ByVal wb As Workbook, ByVal which As Long) As Long
    Dim v As Variant, en As Long
    On Error Resume Next
    v = wb.LinkSources(which)
    en = Err.Number
    On Error GoTo 0
    If en <> 0 Then
        LinkCount = -1
    ElseIf IsArray(v) Then
        LinkCount = UBound(v) - LBound(v) + 1
    End If
End Function

' ===========================================================================
' A cone: Excel's Precedents for the floor, the sheet graph for the ceiling
' ===========================================================================

Private Sub MeasureCone(ByVal c As Range, ByVal k As Long, ByVal want As String, ByVal label As String)
    Dim ws As Worksheet, host As Long, isF As Boolean, ownF As String
    Dim prec As Range, en As Long, tP As Double, repsP As Long, t0 As Double
    Dim urA As Range, ar As Range, cl As Range, clips As Collection
    Dim raw As Double, nAreas As Long, offAreas As Long
    Dim minR As Long, maxR As Long, minC As Long, maxC As Long, rEnd As Long, cEnd As Long
    Dim useGrid As Boolean, mark() As Byte
    Dim tR As Double, repsR As Long, coneRead As Double, readText As String
    Dim hits() As Boolean, reach() As Boolean, cs As ConeStats
    Dim s As Long, nNamed As Long, nReach As Long, back As Boolean, reachExt As Boolean
    Dim ceil As Double, floorFacts As Double, scanFacts As Double, scanRead As Double
    Dim floorP As Double, scanP As Double, remark As String

    On Error GoTo Fail
    Set ws = c.Worksheet
    ' The sheet's position as the measure numbered it, found by its name: what
    ' Index counts is not relied on, and a sheet moved since reads as moved.
    host = SheetIndexOf(ws.Name)
    If host = 0 Then
        Debug.Print "AXM1 refused: this sheet was not there when the workbook was measured. Measure it again."
        Exit Sub
    End If
    If ws.Parent.Sheets(host).Name <> ws.Name Then
        Debug.Print "AXM1 refused: the sheets have moved since the workbook was measured. Measure it again."
        Exit Sub
    End If
    isF = c.HasFormula

    ' --- Excel's own walk: all levels, this sheet only ----------------------
    Do
        Set prec = Nothing
        t0 = Timer
        On Error Resume Next
        Set prec = c.Precedents
        en = Err.Number
        On Error GoTo Fail
        tP = tP + SecondsSince(t0)
        repsP = repsP + 1
    Loop While tP < REP_SECONDS And repsP < REP_MAX

    ' --- the floor, clipped to the used range --------------------------------
    Set urA = ws.Range(gUrAddr(host))
    Set clips = New Collection
    If en = 0 Then
        raw = CDbl(prec.CountLarge)
        nAreas = prec.Areas.Count
        For Each ar In prec.Areas
            If ar.Worksheet.Name <> ws.Name Then
                offAreas = offAreas + 1
            Else
                Set cl = Application.Intersect(ar, urA)
                If Not cl Is Nothing Then clips.Add cl
            End If
        Next ar
    End If
    If clips.Count > 0 Then
        minR = ws.Rows.Count + 1
        minC = ws.Columns.Count + 1
        For Each cl In clips
            If cl.Row < minR Then minR = cl.Row
            rEnd = cl.Row + cl.Rows.Count - 1
            If rEnd > maxR Then maxR = rEnd
            If cl.Column < minC Then minC = cl.Column
            cEnd = cl.Column + cl.Columns.Count - 1
            If cEnd > maxC Then maxC = cEnd
        Next cl
        useGrid = (CDbl(maxR - minR + 1) * CDbl(maxC - minC + 1) <= GRID_CELLS)
        If useGrid Then ReDim mark(1 To maxR - minR + 1, 1 To maxC - minC + 1)
        ' the cone shape's read: every clipped area, one read each, timed
        Do
            For Each cl In clips
                tR = tR + TimeBlockRead(ws, cl)
            Next cl
            repsR = repsR + 1
        Loop While tR < REP_SECONDS And repsR < REP_MAX
        coneRead = tR / repsR
    End If

    ' --- counted once per cell, untimed; and the sheets its formulas name ----
    ReDim hits(1 To gN)
    For Each cl In clips
        CountConeBlock ws, cl, useGrid, mark, minR, minC, cs, hits
    Next cl
    If isF Then
        ownF = c.Formula
        ScanRefs ownF, hits, cs.qual, cs.ext, cs.unres
    End If

    ' --- the ceiling: those sheets, followed through the sheet graph ---------
    ReDim reach(1 To gN)
    For s = 1 To gN
        If hits(s) Then
            If s <> host Then
                reach(s) = True
                nNamed = nNamed + 1
            End If
        End If
    Next s
    CloseReach reach
    back = reach(host)
    If back Then ceil = gNon(host) Else ceil = cs.floorNon
    scanFacts = gNon(host) + 2 * gFor(host)
    scanRead = gReadS(host)
    If gNotRead(host) Then remark = remark & "its own sheet was not read; "
    For s = 1 To gN
        If reach(s) Then
            If s <> host Then
                nReach = nReach + 1
                ceil = ceil + gNon(s)
                scanFacts = scanFacts + gNon(s) + 2 * gFor(s)
                scanRead = scanRead + gReadS(s)
                If gExtRefs(s) > 0 Then reachExt = True
                If gNotRead(s) Then remark = remark & "reaches a sheet not read; "
            End If
        End If
    Next s
    floorFacts = cs.floorNon + 2 * cs.floorF
    floorP = coneRead + floorFacts * FACT_MS / 1000#
    scanP = scanRead + scanFacts * FACT_MS / 1000#

    If en = 1004 Then remark = remark & "no precedents on its own sheet (1004); "
    If en <> 0 And en <> 1004 Then remark = remark & "Precedents raised " & en & "; "
    If offAreas > 0 Then remark = remark & "Precedents left the sheet; "
    If clips.Count > 0 And Not useGrid Then remark = remark & "not counted once per cell: its box is past 2^24 cells; "
    If cs.ext > 0 Or reachExt Then remark = remark & "reaches another workbook; "
    If cs.unres > 0 Then remark = remark & Cnt(cs.unres) & " qualifier(s) name no sheet here; "
    If repsR > 0 Then readText = Ms(coneRead) & " (" & repsR & ")" Else readText = "-"

    If Not gConeLegend Then
        Debug.Print "AXM1C|cone|sheet|formula|precedents|raw cells|areas|off-sheet areas|floor cells|floor non-blank|" & _
                    "floor formulas|precedents ms (reps)|cone read ms (reps)|sheet read ms|qualified refs|sheets named|" & _
                    "external refs|reach sheets|back to own sheet|ceiling non-blank|floor facts|scan facts|" & _
                    "floor price s|scan price s|remark"
        gConeLegend = True
    End If
    Debug.Print "AXM1C|" & label & "|" & host & "|" & YN(isF) & "|" & PrecText(en) & "|" & Cnt(raw) & "|" & nAreas & "|" & _
                offAreas & "|" & Cnt(cs.floorCells) & "|" & Cnt(cs.floorNon) & "|" & Cnt(cs.floorF) & "|" & _
                Ms(tP / repsP) & " (" & repsP & ")|" & readText & "|" & Ms(gReadS(host)) & "|" & Cnt(cs.qual) & "|" & _
                nNamed & "|" & Cnt(cs.ext) & "|" & nReach & "|" & YN(back) & "|" & Cnt(ceil) & "|" & Cnt(floorFacts) & "|" & _
                Cnt(scanFacts) & "|" & S3(floorP) & "|" & S3(scanP) & "|" & remark
    If k >= 1 Then
        gConeSet(k) = True
        gConeFloor(k) = cs.floorNon
        gConeCeil(k) = ceil
        gConeScanP(k) = scanP
        gConeFloorP(k) = floorP
    End If
    If Len(want) > 0 Then
        Check "cone " & label, en & "|" & cs.floorCells & "|" & cs.floorNon & "|" & cs.floorF & "|" & offAreas & "|" & _
              cs.qual & "|" & nNamed & "|" & cs.ext & "|" & nReach & "|" & YN(back) & "|" & ceil & "|" & floorFacts & "|" & _
              scanFacts, want, CONE_FIELDS
    End If
    Exit Sub
Fail:
    Debug.Print "AXM1C|" & label & "|failed: error " & Err.Number & " (" & Err.Description & ")"
    If Len(want) > 0 Then Check "cone " & label, "failed", want, CONE_FIELDS
End Sub

' The cone shape's read of one area: Value2 and Formula, in chunks, timed
' together with making the Range, which a reader following references pays.
Private Function TimeBlockRead(ByVal ws As Worksheet, ByVal rng As Range) As Double
    Dim nR As Long, nC As Long, per As Long, r1 As Long, r2 As Long
    Dim blk As Range, v As Variant, f As Variant, t0 As Double, t As Double
    nR = rng.Rows.Count
    nC = rng.Columns.Count
    per = RowsPerChunk(nC)
    r1 = 1
    Do While r1 <= nR
        r2 = r1 + per - 1
        If r2 > nR Then r2 = nR
        t0 = Timer
        Set blk = ws.Range(rng.Cells(r1, 1), rng.Cells(r2, nC))
        v = blk.Value2
        f = blk.Formula
        t = t + SecondsSince(t0)
        v = Empty
        f = Empty
        r1 = r2 + 1
    Loop
    TimeBlockRead = t
End Function

' One clipped area of the floor, each cell counted once across overlapping
' areas when the box fits the grid.
Private Sub CountConeBlock(ByVal ws As Worksheet, ByVal rng As Range, ByVal useGrid As Boolean, ByRef mark() As Byte, _
                           ByVal minR As Long, ByVal minC As Long, ByRef cs As ConeStats, ByRef hits() As Boolean)
    Dim nR As Long, nC As Long, per As Long, r1 As Long, r2 As Long, i As Long, j As Long
    Dim blk As Range, v As Variant, f As Variant, fs As String
    Dim baseR As Long, baseC As Long, gi As Long, gj As Long, fresh As Boolean
    nR = rng.Rows.Count
    nC = rng.Columns.Count
    per = RowsPerChunk(nC)
    r1 = 1
    Do While r1 <= nR
        r2 = r1 + per - 1
        If r2 > nR Then r2 = nR
        Set blk = ws.Range(rng.Cells(r1, 1), rng.Cells(r2, nC))
        v = blk.Value2
        f = blk.Formula
        MakeGrid v
        MakeGrid f
        baseR = blk.Row
        baseC = blk.Column
        For i = 1 To UBound(v, 1)
            For j = 1 To UBound(v, 2)
                fresh = True
                If useGrid Then
                    gi = baseR + i - minR
                    gj = baseC + j - minC
                    If mark(gi, gj) = 0 Then
                        mark(gi, gj) = 1
                    Else
                        fresh = False
                    End If
                End If
                If fresh Then
                    cs.floorCells = cs.floorCells + 1
                    If Not IsEmpty(v(i, j)) Then cs.floorNon = cs.floorNon + 1
                    fs = f(i, j)
                    If IsFormulaText(fs, v(i, j)) Then
                        cs.floorF = cs.floorF + 1
                        ScanRefs fs, hits, cs.qual, cs.ext, cs.unres
                    End If
                End If
            Next j
        Next i
        v = Empty
        f = Empty
        r1 = r2 + 1
    Loop
End Sub

Private Sub CloseReach(ByRef reach() As Boolean)
    Dim s As Long, t As Long, grew As Boolean
    Do
        grew = False
        For s = 1 To gN
            If reach(s) Then
                For t = 1 To gN
                    If gEdge(s, t) Then
                        If Not reach(t) Then
                            reach(t) = True
                            grew = True
                        End If
                    End If
                Next t
            End If
        Next s
    Loop While grew
End Sub

' ===========================================================================
' The crude count: the sheet name in front of each "!"
' ===========================================================================

' Walks the formula once, left to right, skipping strings ("...") and quoted
' names ('...'), and reads back from each "!" that is outside both.
Private Sub ScanRefs(ByRef fs As String, ByRef hits() As Boolean, ByRef qual As Double, ByRef ext As Double, _
                     ByRef unres As Double)
    Dim i As Long, n As Long, ch As String, inText As Boolean, inName As Boolean
    If InStr(1, fs, "!", vbBinaryCompare) = 0 Then Exit Sub
    n = Len(fs)
    For i = 1 To n
        ch = Mid$(fs, i, 1)
        If inText Then
            If ch = """" Then inText = False        ' a doubled "" closes and reopens: the same state
        ElseIf inName Then
            If ch = "'" Then inName = False         ' a doubled '' likewise
        ElseIf ch = """" Then
            inText = True
        ElseIf ch = "'" Then
            inName = True
        ElseIf ch = "!" Then
            ReadQualifier fs, i, hits, qual, ext, unres
        End If
    Next i
End Sub

Private Sub ReadQualifier(ByRef fs As String, ByVal p As Long, ByRef hits() As Boolean, ByRef qual As Double, _
                          ByRef ext As Double, ByRef unres As Double)
    Dim e As Long, s As Long, s2 As Long, q As String, ch As String
    e = p - 1
    qual = qual + 1
    If e < 1 Then
        unres = unres + 1
        Exit Sub
    End If
    If Mid$(fs, e, 1) = "'" Then
        ' a quoted name: back to its opening quote; '' inside it is one quote
        s = e - 1
        Do While s >= 1
            If Mid$(fs, s, 1) <> "'" Then
                s = s - 1
            ElseIf s = 1 Then
                Exit Do
            ElseIf Mid$(fs, s - 1, 1) = "'" Then
                s = s - 2
            Else
                Exit Do
            End If
        Loop
        If s < 1 Then
            unres = unres + 1
            Exit Sub
        End If
        q = Replace(Mid$(fs, s + 1, e - s - 1), "''", "'")
        If InStr(1, q, "[", vbBinaryCompare) > 0 Then
            ext = ext + 1                           ' '[Book.xlsx]Sheet 1'! or 'C:\dir\[Book.xlsx]Sheet 1'!
            Exit Sub
        End If
    Else
        ' an unquoted name: back over the characters a name is made of
        s = e
        Do While s >= 1
            If IsNameChar(Mid$(fs, s, 1)) Then
                s = s - 1
            Else
                Exit Do
            End If
        Loop
        If s = e Then                               ' nothing a name is made of in front of the "!"
            unres = unres + 1
            Exit Sub
        End If
        q = Mid$(fs, s + 1, e - s)
        If s >= 1 Then
            ch = Mid$(fs, s, 1)
            If ch = "]" Then                        ' [Book.xlsx]Sheet1! or [1]Sheet1!
                ext = ext + 1
                Exit Sub
            ElseIf ch = ":" Then                    ' Jan:Dec! - a 3D span; read its first sheet too
                s2 = s - 1
                Do While s2 >= 1
                    If IsNameChar(Mid$(fs, s2, 1)) Then
                        s2 = s2 - 1
                    Else
                        Exit Do
                    End If
                Loop
                q = Mid$(fs, s2 + 1, s - s2 - 1) & ":" & q
            End If
        End If
    End If
    MarkSheets q, hits, unres
End Sub

' One sheet, or a 3D span's every sheet from first to last by position.
Private Sub MarkSheets(ByVal q As String, ByRef hits() As Boolean, ByRef unres As Double)
    Dim c As Long, a As Long, b As Long, i As Long
    c = InStr(1, q, ":", vbBinaryCompare)
    If c = 0 Then
        a = SheetIndexOf(q)
        If a = 0 Then
            unres = unres + 1
        Else
            hits(a) = True
        End If
        Exit Sub
    End If
    a = SheetIndexOf(Left$(q, c - 1))
    b = SheetIndexOf(Mid$(q, c + 1))
    If a = 0 Or b = 0 Then
        unres = unres + 1
        Exit Sub
    End If
    If a > b Then
        i = a
        a = b
        b = i
    End If
    For i = a To b
        hits(i) = True
    Next i
End Sub

Private Function SheetIndexOf(ByVal nm As String) As Long
    If gSheetIx.Exists(nm) Then SheetIndexOf = gSheetIx.Item(nm)
End Function

Private Function IsNameChar(ByVal ch As String) As Boolean
    Dim a As Long
    a = AscW(ch)
    If a < 0 Then a = a + 65536
    Select Case a
    Case 48 To 57, 65 To 90, 97 To 122, 95, 46      ' 0-9 A-Z a-z _ .
        IsNameChar = True
    Case Is > 127                                   ' letters past ASCII
        IsNameChar = True
    End Select
End Function

' ===========================================================================
' The controls: the fixture, built here, and its expectations
' ===========================================================================

' Ten sheets, every count derived by hand before any run (the expectations
' below say how). Every sheet exists before any formula names it: a formula
' naming a sheet that does not exist yet makes Excel ask for a file.
Private Function BuildFixture() As Workbook
    Dim wb As Workbook, ch As Chart, r As Long, i As Long, gotNames As String, stage As String
    On Error GoTo Fail
    stage = "a new workbook"
    Set wb = Workbooks.Add(xlWBATWorksheet)
    stage = "its sheets"
    wb.Worksheets(1).Name = "Inputs"
    AddSheet wb, "Calc"
    AddSheet wb, "Q1 Data"
    AddSheet wb, "Bob's"
    AddSheet wb, "Hidden"
    AddSheet wb, "VeryHidden"
    AddSheet wb, "Empty"
    AddSheet wb, "One"
    AddSheet wb, "Inflated"
    stage = "the chart sheet"
    ' Made while every sheet is still empty, so it plots nothing. Charts.Add
    ' does not reliably honour After: (the owner's run, 2026-10-01: ten sheets,
    ' not in the order built), so the chart is moved to the end if need be.
    Set ch = wb.Charts.Add(After:=wb.Sheets(wb.Sheets.Count))
    ch.Name = "FxChart"
    If wb.Sheets(wb.Sheets.Count).Name <> "FxChart" Then ch.Move After:=wb.Sheets(wb.Sheets.Count)
    stage = "the name FxRate"
    wb.Names.Add Name:="FxRate", RefersTo:="=Inputs!$B$1"
    stage = "the name FxHidden"
    wb.Names.Add Name:="FxHidden", RefersTo:="=Calc!$D$10", Visible:=False
    ' A name turned #REF! as a model's names are, by deleting what it refers
    ' to: Names.Add refuses "=#REF!" itself, with error 1004 (the owner's first
    ' run, 2026-10-01). Column D of the still-empty Inputs goes; FxRate's B
    ' stays where it is.
    stage = "the name FxBroken"
    wb.Names.Add Name:="FxBroken", RefersTo:="=Inputs!$D$1"
    wb.Worksheets("Inputs").Columns(4).Delete
    stage = "the cells"
    With wb.Worksheets("Inputs")
        .Range("A1").Value = "Rate"
        .Range("B1").Value = 0.1
        .Range("A2").Value = "Base"
        .Range("B2").Value = 1000
        .Range("A3").Value = "Units"
        .Range("B3").Value = 12
    End With
    With wb.Worksheets("Calc")
        .Range("A1").Value = "Item"
        .Range("B1").Value = "Qty"
        .Range("C1").Value = "Tax"
        .Range("D1").Value = "Total"
        For r = 2 To 9
            .Cells(r, 2).Value = r - 1
            .Cells(r, 3).Formula = "=B" & r & "*Inputs!$B$1"
            .Cells(r, 4).Formula = "=B" & r & "+C" & r
        Next r
        .Range("D10").Formula = "=SUM(D2:D9)"
        .Range("E10").Formula = "=""Total! ""&D10"
        .Range("G2").Formula = "=SUM(B2:B5)+SUM(B4:B9)"
        .Range("H2").Formula = "=1+2"
        .Range("I2").Value = "'=looks like a formula"
    End With
    With wb.Worksheets("Q1 Data")
        .Range("A1").Value = "Link"
        .Range("B2").Formula = "=Calc!D10+'Bob''s'!B2"
        .Range("B3").Formula = "=SUM(Inputs:Calc!B2)"
        .Range("C2").Formula = "=B2*2"
    End With
    With wb.Worksheets("Bob's")
        .Range("A1").Value = "Name"
        .Range("B1").Value = "Amount"
        .Range("A2").Value = "Ann"
        .Range("B2").Value = 10
        .Range("A3").Value = "Bo"
        .Range("B3").Value = 20
        .ListObjects.Add(xlSrcRange, .Range("A1:B3"), , xlYes).Name = "FxPeople"
    End With
    With wb.Worksheets("Hidden")
        .Range("A1").Value = "Secret"
        .Range("B1").Formula = "=FxRate*2"
        .Range("B2").Formula = "=INDIRECT(""Inputs!B2"")"
    End With
    wb.Worksheets("VeryHidden").Range("A1").Formula = "=Hidden!B1+1"
    wb.Worksheets("One").Range("C3").Value = 42
    With wb.Worksheets("Inflated")
        .Range("B2").Value = "x"
        .Range("C2").Value = 1
        .Range("B3").Value = "y"
        .Range("C3").Value = 2
        .Range("B4").Value = "z"
        .Range("C4").Value = 3
        .Range("K40").Interior.Color = RGB(255, 255, 0)     ' formatting alone: the used range now runs to K40
    End With
    stage = "hiding two sheets"
    wb.Worksheets("Hidden").Visible = xlSheetHidden
    wb.Worksheets("VeryHidden").Visible = xlSheetVeryHidden
    stage = "the check of the build"
    ' The sheets in their actual order, printed whole when they are not the
    ' order built: the fixture's own names, nothing of the model's.
    For i = 1 To wb.Sheets.Count
        If i > 1 Then gotNames = gotNames & ","
        gotNames = gotNames & wb.Sheets(i).Name
    Next i
    Check "fixture build", gotNames, FIXTURE_SHEETS, "sheets in order"
    If gotNames <> FIXTURE_SHEETS Then Exit Function
    wb.Activate
    wb.Worksheets("Calc").Activate
    Set BuildFixture = wb
    Exit Function
Fail:
    Debug.Print "AXM1K|fixture build|WRONG: error " & Err.Number & " (" & Err.Description & ") while making " & stage & _
                ". Close the new workbook without saving."
    gChecks = gChecks + 1
End Function

Private Sub AddSheet(ByVal wb As Workbook, ByVal nm As String)
    Dim ws As Worksheet
    Set ws = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    ws.Name = nm
End Sub

Private Sub FixtureCone(ByVal fx As Workbook, ByVal sheetName As String, ByVal addr As String, ByVal k As Long)
    fx.Activate
    fx.Worksheets(sheetName).Activate
    MeasureCone fx.Worksheets(sheetName).Range(addr), k, ConeWant(k), CStr(k)
End Sub

' Sheet fields, in SHEET_FIELDS' order: kind|vis|rows|cols|cells|non-blank|
' formulas|excel formulas|past last cell|formula chars|longest|distinct r1c1
' new|qualified refs|sheets named|external refs|unresolved|indirect or
' offset|tables.
'
' THE FIXTURE, derived by hand:
'  1 Inputs     A1:B3, three labels and three numbers.
'  2 Calc       A1:I10 (10 x 9 = 90). Non-blank 33: A1:D1 4, B2:B9 8, C2:C9 8,
'               D2:D9 8, then D10 E10 G2 H2 I2. Formulas 20: C2:C9 =B2*Inputs!$B$1
'               (15 chars), D2:D9 =B2+C2 (6), D10 =SUM(D2:D9) (11), E10
'               ="Total! "&D10 (14), G2 =SUM(B2:B5)+SUM(B4:B9) (22), H2 =1+2 (4);
'               I2 is text typed as '=..., not a formula. Chars 8x15 + 8x6 + 11
'               + 14 + 22 + 4 = 219. Distinct R1C1 6 (C2:C9 are one, D2:D9 are
'               one). Qualified 8, the C column's Inputs!; the "!" inside E10's
'               string is not one. Sheets named 1.
'  3 Q1 Data    A1:C3 (9). Non-blank 4: A1, B2, B3, C2. B2 =Calc!D10+'Bob''s'!B2
'               (21 chars), B3 =SUM(Inputs:Calc!B2) (20), C2 =B2*2 (5): 46.
'               Qualified 3: Calc!, 'Bob''s'!, and the 3D span Inputs:Calc!.
'               Sheets named 3: Calc, Bob's, Inputs.
'  4 Bob's      A1:B3, a Table of two rows under two headers.
'  5 Hidden     hidden. A1:B2 (4), non-blank 3. B1 =FxRate*2 (9 chars) reaches
'               Inputs through a defined name, which the crude count cannot
'               see; B2 =INDIRECT("Inputs!B2") (22) keeps its "!" inside a
'               string. Chars 31, qualified 0, INDIRECT 1.
'  6 VeryHidden very hidden. A1 alone, =Hidden!B1+1 (12): a one-cell used
'               range, whose Value2 is a scalar.
'  7 Empty      never written: its used range is A1, blank.
'  8 One        C3 = 42 alone: a one-cell used range away from A1.
'  9 Inflated   B2:C4 hold six values; K40 is only formatted, so the used range
'               is B2:K40 (39 x 10 = 390) and runs 36 rows and 8 columns past
'               the last non-blank cell, C4.
' 10 FxChart    a chart sheet: no used range.
' Workbook: 508 cells, 60 non-blank, 26 formulas, the largest used range
' sheet 9's 390, 308 formula chars, 12 distinct R1C1 formulas, 12 qualified
' references, 5 sheet links (Calc to Inputs; Q1 Data to Calc, Bob's and
' Inputs; VeryHidden to Hidden), 1 INDIRECT, 3 names (1 hidden, 1 #REF!), 1
' Table.
'
' THE SAMPLE, Frazaro Sample Data.xlsx as tools/build_examples.ps1 builds it
' (2026-09-30), derived from its XML: no sheet has a <dimension>, so Excel
' computes each used range from the cells it holds, and every cell the file
' holds has a value. "Start Here" holds B2:F28; Inventory's header reaches H1,
' its rows only G. Prediction: a column width alone (Start Here's column A;
' Expenses' and Invoices' F; Ledger's G) does not stretch the used range.
Private Function SheetWant(ByVal expectKey As String, ByVal i As Long) As String
    If expectKey = "fixture" Then
        Select Case i
        Case 1: SheetWant = "ws|V|3|2|6|6|0|0|0,0|0|0|0|0|0|0|0|0|0"
        Case 2: SheetWant = "ws|V|10|9|90|33|20|20|0,0|219|22|6|8|1|0|0|0|0"
        Case 3: SheetWant = "ws|V|3|3|9|4|3|3|0,0|46|21|3|3|3|0|0|0|0"
        Case 4: SheetWant = "ws|V|3|2|6|6|0|0|0,0|0|0|0|0|0|0|0|0|1"
        Case 5: SheetWant = "ws|H|2|2|4|3|2|2|0,0|31|22|2|0|0|0|0|1|0"
        Case 6: SheetWant = "ws|VH|1|1|1|1|1|1|0,0|12|12|1|1|1|0|0|0|0"
        Case 7: SheetWant = "ws|V|1|1|1|0|0|0|-|0|0|0|0|0|0|0|0|0"
        Case 8: SheetWant = "ws|V|1|1|1|1|0|0|0,0|0|0|0|0|0|0|0|0|0"
        Case 9: SheetWant = "ws|V|39|10|390|6|0|0|36,8|0|0|0|0|0|0|0|0|0"
        Case 10: SheetWant = "chart|V"
        End Select
    ElseIf expectKey = "sample" Then
        Select Case i
        Case 1: SheetWant = "ws|V|27|5|135|76|0|0|0,0|0|0|0|0|0|0|0|0|0"
        Case 2: SheetWant = "ws|V|41|7|287|287|0|0|0,0|0|0|0|0|0|0|0|0|0"
        Case 3: SheetWant = "ws|V|26|5|130|130|0|0|0,0|0|0|0|0|0|0|0|0|0"
        Case 4: SheetWant = "ws|V|31|5|155|155|0|0|0,0|0|0|0|0|0|0|0|0|0"
        Case 5: SheetWant = "ws|V|21|6|126|126|0|0|0,0|0|0|0|0|0|0|0|0|0"
        Case 6: SheetWant = "ws|V|21|8|168|148|0|0|0,0|0|0|0|0|0|0|0|0|0"
        Case 7: SheetWant = "ws|V|11|3|33|33|0|0|0,0|0|0|0|0|0|0|0|0|1"
        Case 8: SheetWant = "ws|V|6|3|18|18|0|0|0,0|0|0|0|0|0|0|0|0|1"
        Case 9: SheetWant = "ws|V|4|2|8|8|0|0|0,0|0|0|0|0|0|0|0|0|1"
        End Select
    End If
End Function

' Workbook fields, in BOOK_FIELDS' order.
Private Function BookWant(ByVal expectKey As String) As String
    If expectKey = "fixture" Then
        BookWant = "10|9|1|1|1|508|60|26|26|9|390|308|22|12|12|5|0|0|1|3|1|1|0|1|0|0|0"
    ElseIf expectKey = "sample" Then
        BookWant = "9|9|0|0|0|1060|981|0|0|2|287|0|0|0|0|0|0|0|0|0|0|0|0|3|0|0|0"
    End If
End Function

' Cone fields, in CONE_FIELDS' order. Calc holds 33 non-blank cells and 20
' formulas, so a scan of Calc alone is 33 + 2 x 20 = 73 facts.
'  1 Calc!D10 =SUM(D2:D9): Precedents B2:D9, 24 cells, 16 of them formulas
'    (C and D). The C column names Inputs 8 times; Inputs names nothing, so
'    the reach is Inputs alone. Ceiling 24 + Inputs' 6 = 30. Floor facts
'    24 + 2 x 16 = 56; scan facts 73 + 6 = 79.
'  2 Calc!G2 =SUM(B2:B5)+SUM(B4:B9): two overlapping ranges, B2:B9 counted
'    once: 8. No formula among them, no sheet named: ceiling 8, floor facts 8.
'  3 Calc!H2 =1+2: no precedents, Excel raises 1004. Floor 0, ceiling 0.
'  4 'Q1 Data'!C2 =B2*2: Precedents B2 alone (its own precedents are on other
'    sheets), a formula naming Calc and Bob's; Calc leads on to Inputs, so the
'    reach is 3 sheets and never back to Q1 Data. Ceiling 1 + 33 + 6 + 6 = 46.
'    Floor facts 1 + 2 = 3; scan facts (4 + 2 x 3) + 73 + 6 + 6 = 95.
Private Function ConeWant(ByVal k As Long) As String
    Select Case k
    Case 1: ConeWant = "0|24|24|16|0|8|1|0|1|n|30|56|79"
    Case 2: ConeWant = "0|8|8|0|0|0|0|0|0|n|8|8|73"
    Case 3: ConeWant = "1004|0|0|0|0|0|0|0|0|n|0|0|73"
    Case 4: ConeWant = "0|1|1|1|0|2|2|0|3|n|46|3|95"
    End Select
End Function

Private Sub Check(ByVal label As String, ByVal got As String, ByVal want As String, ByVal fieldNames As String)
    Dim g As Variant, w As Variant, fn As Variant, i As Long, msg As String
    gChecks = gChecks + 1
    If got = want Then
        gOks = gOks + 1
        Debug.Print "AXM1K|" & label & "|ok"
        Exit Sub
    End If
    g = Split(got, "|")
    w = Split(want, "|")
    fn = Split(fieldNames, "|")
    For i = 0 To UBound(w)
        If i > UBound(g) Then
            msg = msg & "; " & FieldName(fn, i) & " missing, expected " & w(i)
        ElseIf g(i) <> w(i) Then
            msg = msg & "; " & FieldName(fn, i) & " got " & g(i) & ", expected " & w(i)
        End If
    Next i
    If Len(msg) = 0 Then msg = "; got " & got & ", expected " & want
    Debug.Print "AXM1K|" & label & "|WRONG:" & Mid$(msg, 2)
End Sub

Private Function FieldName(ByRef fn As Variant, ByVal i As Long) As String
    If i <= UBound(fn) Then FieldName = fn(i) Else FieldName = "field " & (i + 1)
End Function

' ===========================================================================
' Small helpers
' ===========================================================================

Private Function SecondsSince(ByVal t0 As Double) As Double
    Dim t As Double
    t = Timer - t0
    If t < 0 Then t = t + 86400#            ' midnight
    SecondsSince = t
End Function

Private Function RowsPerChunk(ByVal nC As Long) As Long
    RowsPerChunk = CHUNK_CELLS \ nC
    If RowsPerChunk < 1 Then RowsPerChunk = 1
End Function

' A one-cell range's Value2 or Formula is a scalar: make it a 1 x 1 array.
Private Sub MakeGrid(ByRef v As Variant)
    Dim one(1 To 1, 1 To 1) As Variant
    If IsArray(v) Then Exit Sub
    one(1, 1) = v
    v = one
End Sub

' A formula's text starts with "=" and is not also its value; a text cell
' typed as '=x reads back "=x" from Formula and from Value2 alike.
Private Function IsFormulaText(ByRef fs As String, ByRef cv As Variant) As Boolean
    If Len(fs) < 2 Then Exit Function
    If AscW(fs) <> 61 Then Exit Function
    If VarType(cv) = vbString Then
        If cv = fs Then Exit Function
    End If
    IsFormulaText = True
End Function

Private Function IsGrid(ByRef st As SheetStats) As Boolean
    IsGrid = (st.kind = "ws" Or st.kind = "macro")
End Function

Private Function SheetKind(ByVal sh As Object) As String
    Select Case TypeName(sh)
    Case "Worksheet"
        If sh.Type = xlWorksheet Then SheetKind = "ws" Else SheetKind = "macro"
    Case "Chart"
        SheetKind = "chart"
    Case Else
        SheetKind = TypeName(sh)
    End Select
End Function

Private Function VisCode(ByVal vis As Long) As String
    Select Case vis
    Case xlSheetVisible: VisCode = "V"
    Case xlSheetHidden: VisCode = "H"
    Case xlSheetVeryHidden: VisCode = "VH"
    Case Else: VisCode = CStr(vis)
    End Select
End Function

Private Function TargetLabel(ByVal expectKey As String) As String
    Select Case expectKey
    Case "fixture": TargetLabel = "control 1, the fixture"
    Case "sample": TargetLabel = "control 2, Frazaro Sample Data"
    Case Else: TargetLabel = "the active workbook"
    End Select
End Function

Private Function CalcWord() As String
    Select Case Application.Calculation
    Case xlCalculationAutomatic: CalcWord = "automatic"
    Case xlCalculationManual: CalcWord = "manual"
    Case xlCalculationSemiautomatic: CalcWord = "automatic except tables"
    Case Else: CalcWord = CStr(Application.Calculation)
    End Select
End Function

Private Function Bitness() As String
#If Win64 Then
    Bitness = "64-bit"
#Else
    Bitness = "32-bit"
#End If
End Function

Private Function MedianOf(ByRef a() As Double, ByVal n As Long) As Double
    Dim b() As Double, i As Long, j As Long, x As Double
    ReDim b(1 To n)
    For i = 1 To n
        b(i) = a(i)
    Next i
    For i = 2 To n                          ' insertion sort: ten values at most
        x = b(i)
        j = i - 1
        Do While j >= 1
            If b(j) <= x Then Exit Do
            b(j + 1) = b(j)
            j = j - 1
        Loop
        b(j + 1) = x
    Next i
    If n Mod 2 = 1 Then
        MedianOf = b((n + 1) \ 2)
    Else
        MedianOf = (b(n \ 2) + b(n \ 2 + 1)) / 2
    End If
End Function

Private Function MaxOf(ByRef a() As Double, ByVal n As Long) As Double
    Dim i As Long
    MaxOf = a(1)
    For i = 2 To n
        If a(i) > MaxOf Then MaxOf = a(i)
    Next i
End Function

Private Function Cnt(ByVal x As Double) As String
    If x < 0 Then
        Cnt = "err"
    Else
        Cnt = Format$(x, "#,##0")
    End If
End Function

Private Function Ms(ByVal secs As Double) As String
    Ms = Format$(secs * 1000#, "0.0")
End Function

Private Function S3(ByVal secs As Double) As String
    S3 = Format$(secs, "0.000")
End Function

Private Function YN(ByVal b As Boolean) As String
    If b Then YN = "y" Else YN = "n"
End Function

Private Function PrecText(ByVal en As Long) As String
    Select Case en
    Case 0: PrecText = "found"
    Case 1004: PrecText = "none"
    Case Else: PrecText = "error " & en
    End Select
End Function

Private Function LinkText(ByVal x As Long) As String
    If x < 0 Then LinkText = "err" Else LinkText = CStr(x)
End Function
