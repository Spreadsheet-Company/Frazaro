Attribute VB_Name = "VLA_HeadTable"
Option Explicit
Public Const VLA_HEADTABLE_VERSION As String = "IN5.0"
' IN5.0: the export-only column adjudicated (see the note below the
' honesty note on Formula) - 9 of the draft's 12 True rows corrected to
' False (sub, function, type, enum, public, private, include, at-line,
' doc), each checked against the actual emitter code rather than the
' draft's own prose. TestHeadTableExportOnly (VLA_Tests.bas) pins the
' resulting 3-symbol set (raw, deflambda, lambda) so a future accidental
' change to this column fails loudly instead of drifting unnoticed.

' =====================================================================
'  VLA_HeadTable - one row per core form: symbol -> aliases, arity,
'  interpreter routine, VBA routine, formula-subset membership,
'  export-only flag. Still NOT the emitter's dispatch itself for six
'  of seven columns (VLA.bas's Select Case ARMS in EmitTop/EmitStmt/
'  EmitExpr/EmitFormula are unchanged and still authoritative; the
'  VBA/interpreter/formula/export-only columns remain a catalog, not
'  codegen) - the CATALOG-first construction order REBUILD.md set out
'  ("all seven columns declared but only the VBA column populated").
'  Rewiring the emitter to dispatch through this table instead of its
'  Select Case arms is a separate, much larger undertaking ("the
'  emitter rewrite") that IN.1 explicitly precedes rather than
'  performs. LX.4 is the ONE EXCEPTION: the Aliases column IS live -
'  VlaHeadTableAliasMap feeds VLA.bas's ResolveHeadAlias, one lookup
'  that runs before every Select Case dispatch, so an aliased head
'  (fijar! for set!) reaches the same Case arm its canonical spelling
'  would. Three rows carry one alias each today (set!/fijar!, if/si,
'  debug-print/depurar) - a working chokepoint, not a translation
'  table; LX.5/a dedicated i18n item is where a real vocabulary lives.
'
'  LAYER:     1
'  MAY CALL:  VLA_Identity (VLA_Messages/VLA_Context/VLA_Forms do not
'             exist yet - add to MAY CALL when they do)
'  SHIPS:     add-in only (VLA_ prefix; not injected)
'  PAYS INTO: IN.1 (this module IS the fork point), LX.4 (the Aliases
'             column and VlaHeadTableAliasMap - see above), IN.5 (the
'             Export-Only column, adjudicated in place below - see the
'             note above the row data), the
'             interpreter (IN.2, when built, needs the same table for
'             a different reason - its own dispatch, not the VBA
'             emitter's - per BETA_ROADMAP.md's own note on this item),
'             R9/SD-5 (no row may end up with a filled VBA cell, an
'             empty interpreter cell, and no refusal - unenforceable
'             today since IN.2 doesn't exist, but the column is here
'             waiting), AS.2 + AS.8 (arms become enumerable, so
'             coverage and parity both get a denominator the moment
'             someone writes the report), F.1 (a declared verb set
'             for phrasebook templates to target), F.7 (the Formula
'             column records what EmitFormula already does today -
'             see the honesty note below)
'  REASON:    ~65 English head words are hardcoded in VLA.bas's four
'             separate Select Case blocks (EmitTop, EmitStmt,
'             EmitExpr, EmitFormula) with exactly one consumer each.
'             One chokepoint now; sixty-five sites twice over the day
'             a second backend (IN.2) or a declared formula subset
'             (F.7) actually needs to ask "does form X support this."
'
'  AN HONESTY NOTE ON THE FORMULA COLUMN, found while building this
'  table rather than assumed going in: EmitFormula's Select Case does
'  not cleanly partition into "supported" and "explicitly refused."
'  Thirteen statement forms ARE explicitly refused with words
'  (set!, obj-set!, dim, begin, while, for, for-each, debug-print,
'  on-error, label, goto, deflambda, include) - but everything else
'  NOT in its explicit list (do-until, select, with, return, xor, the
'  backslash operator, ".", is, like, imp, eqv, and more) silently
'  falls through its Case Else and is treated as a worksheet function
'  call, e.g. "(do-until x)" would emit "DO_UNTIL(x)" rather than
'  refusing. This is exactly the gap F.7 exists to close ("a test
'  that fails when the VBA emitter gains a case the formula dialect
'  neither supports nor refuses"). The Formula column below records
'  three states - "yes" (explicit support), "refused" (explicit,
'  worded refusal), "undeclared" (falls through the permissive
'  Case Else - neither supported nor refused) - rather than a clean
'  yes/no, because collapsing "undeclared" into either would have
'  hidden the exact gap F.7 needs to see.
'
'  IN.5's ADJUDICATION OF THE EXPORT-ONLY COLUMN (the draft above this
'  note used to say "treat this as a draft, not a verdict" - this is
'  that verdict, reasoned through against the actual emitter code, not
'  the draft's own prose. The prose named eight of its twelve True rows
'  by category (raw; type/enum; public/private; deflambda; at-line/doc);
'  sub, function, include, and lambda were True in the row DATA with no
'  justification in the prose at all. Of those four unexplained rows,
'  three turned out wrong (sub, function, include - corrected below) and
'  one turned out right, just under-documented (lambda - kept True).
'
'  True (genuinely export-only - the form's MEANING is "produce a VBA
'  or Excel-formula TEXT artifact," not "have a runtime effect," so no
'  amount of interpreter completeness reaches it): raw (splices opaque,
'  unparsed VBA text - there is nothing to evaluate, only text to emit),
'  deflambda (compiles its body through EmitFormula into Excel LAMBDA
'  formula syntax, a serialization target the interpreter's execution
'  model has no bridge to), lambda (by its own row's VBA-routine column:
'  "falls through EmitExpr's Case Else outside a formula" - it has no
'  meaning outside EmitFormula even in the EMITTER today).
'
'  False, corrected from the draft's True (interpretable in principle -
'  each reasoned through individually, not merely reclassified as a
'  group):
'    sub / function - defining a named, callable procedure has a direct
'      dynamic reading (bind a name to its params+body; dispatch calls
'      into it) that needs no VBA text at all - IN.0.5's own ExecTop
'      already unwraps the outermost (sub main () ...) EnglishToVla
'      wraps every program in, on exactly this reasoning, one level in.
'    type / enum - ExecDim already resolves a type NAME to a default
'      value (Double/String/Boolean/Empty) without needing VBA's own
'      Type to exist; a user-declared type/enum is the same idea one
'      level up (a named field-shape or constant-set the interpreter
'      keeps as data), not an artifact requiring VBA's type system.
'    public / private - visibility is a claim about what OTHER MODULES
'      may call; an interpreted program has no other modules to be
'      visible to, so the wrapper is a no-op under interpretation, not
'      an obstacle to it - unwrap and interpret the inner form.
'    include - checked against the actual pipeline, not assumed: both
'      VlaTranspile AND VlaCompileToForms call SpliceIncludes as their
'      first line, so an include directive is already resolved into
'      plain source text before EITHER backend's dispatch ever runs.
'      It was never really "refused by the interpreter" - it never
'      reaches interpreter dispatch as a form to refuse, on either path.
'    at-line - checked against EmitStmt's actual Case "at-line" arm:
'      its own comment calls it "a zero-runtime annotation form" that
'      "emits the wrapped statements exactly as written," used only to
'      improve the EMITTER's own error messages. Zero runtime effect is
'      the definition of interpretable - unwrap and execute the body.
'    doc - checked against EmitDoc's actual body, which surprised the
'      checking: `(doc name)` is not an annotation at all, it compiles
'      to `Debug.Print "name: " & <that macro's docstring>` - a live
'      runtime statement, computable today through the existing Public
'      VlaMacroDoc(name) accessor VLA.bas already exposes (the same one
'      F.5's own TestContextPushPop pin reads). Ordinary debug-print
'      with one extra lookup, not an export-only form.
'
'  What this pass did NOT do: build the refusal wording into IN.0.5 (the
'  walking skeleton) for raw/deflambda/lambda. The skeleton is explicitly
'  a throwaway measurement exercise IN.2 replaces, not a thing to keep
'  extending - "refuses in words, naming the reason and the alternative"
'  is IN.2's own job once it exists to dispatch through this table; this
'  column is the contract it will be built to honor, per IN.1's own
'  precedent (a data change made before the fork it describes existed).
' =====================================================================

' Column indices into a row (a Collection of seven items, in this
' order) - so callers write row.Item(HT_VBA) instead of a magic 5.
Public Const HT_SYMBOL As Long = 1
Public Const HT_ALIASES As Long = 2
Public Const HT_ARITY As Long = 3
Public Const HT_INTERP As Long = 4
Public Const HT_VBA As Long = 5
Public Const HT_FORMULA As Long = 6
Public Const HT_EXPORTONLY As Long = 7

' The full table, rebuilt fresh on every call - no module-level cache,
' so there is nothing here for a recompile to lose (hazard: "editing a
' running project resets module-level state," LESSONS.md XXXII). The
' table is 65 rows of string literals; rebuilding it costs nothing a
' human would notice.
Public Function VlaHeadTableRows() As Collection
    Set VlaHeadTableRows = BuildRows()
End Function

Public Function VlaHeadTableCount() As Long
    VlaHeadTableCount = BuildRows().Count
End Function

' Look up one row by symbol OR by any declared alias (LX.4: three rows
' carry one alias each today - see BuildRows). A linear scan over 65
' rows when the fast keyed lookup misses - fine for the occasional
' catalog lookup this function is for; VlaHeadTableAliasMap below is
' the O(1) alternative the emitter's own hot path uses instead.
' Returns Nothing if no row matches.
Public Function VlaHeadTableRow(ByVal symbol As String) As Collection
    Dim rows As Collection
    Set rows = BuildRows()
    Dim key As String
    key = VLA_Identity.Fold(symbol)
    On Error Resume Next
    Set VlaHeadTableRow = rows.Item(key)
    On Error GoTo 0
    If Not VlaHeadTableRow Is Nothing Then Exit Function

    ' Not a primary symbol - scan alias lists.
    Dim row As Variant
    Dim aliasList As Variant
    Dim a As Variant
    For Each row In rows
        aliasList = Split(CStr(row.Item(HT_ALIASES)), ",")
        For Each a In aliasList
            If Len(Trim$(CStr(a))) > 0 Then
                If VLA_Identity.Fold(Trim$(CStr(a))) = key Then
                    Set VlaHeadTableRow = row
                    Exit Function
                End If
            End If
        Next
    Next
End Function

' LX.4: the alias -> canonical-symbol map VLA.bas's emitter chokepoint
' (ResolveHeadAlias) consults, keyed for O(1) lookup rather than
' VlaHeadTableRow's linear scan - built once per compile (VlaTranspile/
' VlaCompileToForms set it fresh into VLA.bas's mHeadAliases), not once
' per node. No module-level cache HERE either, same reason the table
' itself has none: a recompile drops module state (LESSONS.md XXXII),
' and rebuilding costs nothing a human would notice.
Public Function VlaHeadTableAliasMap() As Collection
    Dim rows As Collection
    Set rows = BuildRows()
    Dim m As New Collection
    Dim row As Variant
    Dim aliasList As Variant
    Dim a As Variant
    For Each row In rows
        aliasList = Split(CStr(row.Item(HT_ALIASES)), ",")
        For Each a In aliasList
            If Len(Trim$(CStr(a))) > 0 Then
                m.Add CStr(row.Item(HT_SYMBOL)), VLA_Identity.Fold(Trim$(CStr(a)))
            End If
        Next
    Next
    Set VlaHeadTableAliasMap = m
End Function

' One pasteable line for VlaDiagnostics - the counters-line lesson
' (Alpha 5, "baselines must print themselves"): a number nobody prints
' is a number that drifts unnoticed.
Public Function VlaHeadTableSummary() As String
    Dim rows As Collection
    Set rows = BuildRows()
    Dim formulaYes As Long, formulaRefused As Long, formulaUndeclared As Long
    Dim exportOnlyN As Long
    Dim row As Variant
    For Each row In rows
        Select Case CStr(row.Item(HT_FORMULA))
            Case "yes": formulaYes = formulaYes + 1
            Case "refused": formulaRefused = formulaRefused + 1
            Case Else: formulaUndeclared = formulaUndeclared + 1
        End Select
        If CBool(row.Item(HT_EXPORTONLY)) Then exportOnlyN = exportOnlyN + 1
    Next
    VlaHeadTableSummary = rows.Count & " forms (formula: " & formulaYes & " supported, " & _
        formulaRefused & " refused, " & formulaUndeclared & " undeclared; " & _
        exportOnlyN & " export-only; 0 with an interpreter routine - IN.2 not built)"
End Function

Private Sub AddRow(rows As Collection, ByVal symbol As String, ByVal aliases As String, _
                    ByVal arity As String, ByVal interpRoutine As String, ByVal vbaRoutine As String, _
                    ByVal formulaSupport As String, ByVal exportOnly As Boolean)
    Dim row As New Collection
    row.Add symbol
    row.Add aliases
    row.Add arity
    row.Add interpRoutine
    row.Add vbaRoutine
    row.Add formulaSupport
    row.Add exportOnly
    ' Keyed Add: a duplicate symbol raises "key already exists" the
    ' moment BuildRows runs - a free, built-in F.4-style conflict
    ' check for this table's own 65 rows.
    rows.Add row, VLA_Identity.Fold(symbol)
End Sub

' The catalog itself: every head symbol VLA.bas's four Select Case
' blocks (EmitTop, EmitStmt, EmitExpr, EmitFormula) currently
' recognize. Grouped as read from source, not alphabetically, so a
' diff against VLA.bas stays legible.
Private Function BuildRows() As Collection
    Dim rows As New Collection

    ' --- Top-level / declaration forms (EmitTop) ---------------------
    AddRow rows, "sub", "", "variable (name, params, [doc], body)", "", "EmitProc", "undeclared", False
    AddRow rows, "function", "", "variable (name, params, [type], [doc], body)", "", "EmitProc", "undeclared", False
    AddRow rows, "type", "", "variable (name, members)", "", "EmitTypeDef", "undeclared", False
    AddRow rows, "enum", "", "variable (name, members)", "", "EmitEnumDef", "undeclared", False
    AddRow rows, "public", "", "1 (wrapped form)", "", "EmitVisibility", "undeclared", False
    AddRow rows, "private", "", "1 (wrapped form)", "", "EmitVisibility", "undeclared", False
    AddRow rows, "include", "", "1 (path)", "", "refused - whole-line splicer only, never reaches the emitter", "undeclared", False

    ' --- Statement forms (EmitStmt), also legal some places above ----
    AddRow rows, "dim", "", "1-2 (name, [type])", "", "EmitDimCore (via EmitTop or EmitStmt)", "refused", False
    AddRow rows, "const", "", "2 (name, value)", "", "EmitConstCore (via EmitTop or EmitStmt)", "undeclared", False
    AddRow rows, "raw", "", "1 (literal VBA text)", "", "StrLitContent (inline)", "undeclared", True
    AddRow rows, "begin", "", "variable (body)", "", "EmitBody (inline)", "refused", False
    AddRow rows, "deflambda", "", "variable (name, params, [doc], formula-body)", "", "EmitDeflambda", "refused", True
    AddRow rows, "set!", "fijar!", "2 (place, value)", "", "inline in EmitStmt", "refused", False
    AddRow rows, "obj-set!", "", "2 (place, value)", "", "inline in EmitStmt", "refused", False
    AddRow rows, "if", "si", "variable (test, clauses)", "", "EmitIf", "yes", False
    AddRow rows, "for", "", "variable (header, body)", "", "EmitFor", "refused", False
    AddRow rows, "for-each", "", "variable (header, body)", "", "EmitForEach", "refused", False
    AddRow rows, "while", "", "variable (test, body)", "", "inline in EmitStmt", "refused", False
    AddRow rows, "do-until", "", "variable (test, body)", "", "inline in EmitStmt", "undeclared", False
    AddRow rows, "select", "", "variable (expr, cases)", "", "EmitSelect", "undeclared", False
    AddRow rows, "with", "", "variable (expr, body)", "", "inline in EmitStmt", "undeclared", False
    AddRow rows, "return", "", "0-1 (value)", "", "EmitReturn", "undeclared", False
    AddRow rows, "exit-sub", "", "0", "", "inline in EmitStmt", "undeclared", False
    AddRow rows, "exit-function", "", "0", "", "inline in EmitStmt", "undeclared", False
    AddRow rows, "exit-for", "", "0", "", "inline in EmitStmt", "undeclared", False
    AddRow rows, "exit-do", "", "0", "", "inline in EmitStmt", "undeclared", False
    AddRow rows, "redim", "", "variable (name, dims)", "", "EmitRedim", "undeclared", False
    AddRow rows, "on-error", "", "1-2 (resume-next | goto label)", "", "inline in EmitStmt", "undeclared", False
    AddRow rows, "goto", "", "1 (label)", "", "inline in EmitStmt", "undeclared", False
    AddRow rows, "label", "", "1 (name)", "", "inline in EmitStmt", "undeclared", False
    AddRow rows, "quote", "", "1 (datum)", "", "EmitQuote / QuoteDatum", "yes", False
    AddRow rows, "debug-print", "depurar", "variable (values)", "", "inline in EmitStmt", "refused", False
    AddRow rows, "call", "", "variable (name, args)", "", "EmitCallStmt", "undeclared", False
    AddRow rows, ".", "", "variable (obj, member, [args])", "", "EmitDotText", "undeclared", False
    AddRow rows, "at-line", "", "variable (line-number, statements)", "", "inline in EmitStmt", "undeclared", False
    AddRow rows, "then", "", "variable (body) - clause of if only", "", "consumed by EmitIf", "undeclared", False
    AddRow rows, "else", "", "variable (body) - clause of if only", "", "consumed by EmitIf", "undeclared", False
    AddRow rows, "elseif", "", "variable (test, body) - clause of if only", "", "consumed by EmitIf", "undeclared", False
    AddRow rows, "case", "", "variable (values, body) - clause of select only", "", "consumed by EmitSelect", "undeclared", False
    AddRow rows, "case-else", "", "variable (body) - clause of select only", "", "consumed by EmitSelect", "undeclared", False
    AddRow rows, "doc", "", "1 (text)", "", "EmitDoc", "undeclared", False
    AddRow rows, "resume", "", "0-1 (next | label)", "", "inline in EmitStmt", "undeclared", False

    ' --- Expression-only forms (EmitExpr) -----------------------------
    AddRow rows, "new", "", "1 (type name)", "", "inline in EmitExpr", "undeclared", False
    AddRow rows, "lambda", "", "2 (params, formula-body) - formula dialect only", "", "n/a - EmitFormula only; falls through EmitExpr's Case Else outside a formula", "yes", True

    ' --- Operators (EmitExpr's EmitChain / unary cases) ---------------
    AddRow rows, "+", "", "2+ (chain)", "", "EmitChain", "yes", False
    AddRow rows, "-", "", "1 (negate) or 2+ (chain)", "", "EmitChain / inline unary", "yes", False
    AddRow rows, "*", "", "2+ (chain)", "", "EmitChain", "yes", False
    AddRow rows, "&", "", "2+ (chain)", "", "EmitChain", "yes", False
    AddRow rows, "/", "", "2+ (chain)", "", "EmitChain", "yes", False
    AddRow rows, "\", "", "2+ (chain)", "", "EmitChain", "undeclared", False
    AddRow rows, "=", "", "2+ (chain)", "", "EmitChain", "yes", False
    AddRow rows, "<>", "", "2+ (chain)", "", "EmitChain", "yes", False
    AddRow rows, "<", "", "2+ (chain)", "", "EmitChain", "yes", False
    AddRow rows, ">", "", "2+ (chain)", "", "EmitChain", "yes", False
    AddRow rows, "<=", "", "2+ (chain)", "", "EmitChain", "yes", False
    AddRow rows, ">=", "", "2+ (chain)", "", "EmitChain", "yes", False
    AddRow rows, "and", "", "2+ (chain)", "", "EmitChain", "yes", False
    AddRow rows, "or", "", "2+ (chain)", "", "EmitChain", "yes", False
    AddRow rows, "xor", "", "2+ (chain)", "", "EmitChain", "undeclared", False
    AddRow rows, "mod", "", "2 (a, b)", "", "EmitChain (VBA); fixed 2-arg MOD() in formulas", "yes", False
    AddRow rows, "is", "", "2 (chain)", "", "EmitChain", "undeclared", False
    AddRow rows, "like", "", "2 (chain)", "", "EmitChain", "undeclared", False
    AddRow rows, "imp", "", "2+ (chain)", "", "EmitChain", "undeclared", False
    AddRow rows, "eqv", "", "2+ (chain)", "", "EmitChain", "undeclared", False
    AddRow rows, "not", "", "1", "", "inline in EmitExpr", "yes", False

    Set BuildRows = rows
End Function
