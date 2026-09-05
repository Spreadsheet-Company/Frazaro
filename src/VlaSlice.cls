VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VlaSlice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' PNTH.0: a read-only, O(1) VIEW over the tail of an existing Collection -
' cdr/cddr's own structural-sharing representation (MACHINE + OPTIMIZATION's
' P-NTH, BETA_ROADMAP2.md), replacing ListTail's old full-copy walk
' (VLA.bas). "Slice," not "cons cell": it stays a FLAT offset into one
' root Collection rather than a general linked structure, on purpose -
' cdr/cddr (repeated tail-walking) is the evidenced hot path (P-DICT's
' own live before/after confirmed GetMacro was never the majority of
' compile-side "expand" time; this is where the rest of it lives), cons
' (prepending) is not, and a flat view cannot represent "one element
' prepended ahead of a shared tail" without falling back to a real
' cons-cell chain - deliberately not built, since cons was never the
' measured cost. EvalCons (VLA.bas) still copies, unchanged; it was
' widened only to accept a VlaSlice tail without crashing, not to
' avoid copying it.
'
' Never nests: built from another VlaSlice, ListTail flattens to point
' at the SAME root Backing with an adjusted Offset (VLA.bas's own
' ListTail comment has the mechanics), so .Item/.Count stay O(1) no
' matter how many chained cdr's produced a given view - a naive
' wrap-the-wrapper design would have only moved the O(n) cost from
' copying into recursive indirection through N nested views instead.
'
' S3.1 safety, checked deliberately, not assumed: every ListTail call
' still constructs a genuinely FRESH VlaSlice object (a new instance,
' a new address) and VLA.bas's TagLine still explicitly tags it 0,
' exactly the discipline that already protected cdr/cddr's old
' Collection-copy result - the stale-ObjPtr hazard S3.1 named was about
' a FREED address being reused by an unrelated later list; a live,
' still-referenced Backing Collection shared across many slices was
' never that hazard, and nothing here ever tags Backing itself, only
' each view's own distinct wrapper identity.
Public Backing As Collection
Public Offset As Long

Public Property Get Count() As Long
    Count = Backing.Count - Offset
End Property

' 1-based, relative to this view - Item(1) is Backing.Item(Offset + 1).
Public Function Item(ByVal i As Long) As Variant
    If IsObject(Backing.Item(Offset + i)) Then
        Set Item = Backing.Item(Offset + i)
    Else
        Item = Backing.Item(Offset + i)
    End If
End Function
