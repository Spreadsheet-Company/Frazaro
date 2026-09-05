VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VlaFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' F5.0: first pass, one instance is one snapshot of VLA.bas's compile-time
' module state. A plain data bag with no behavior of its own: VBA does not
' let a class module read or write another module's Private variables, so
' the copying lives in VLA.bas (VlaPushContext/VlaPopContext), which owns
' the eighteen fields being copied - this class only carries the values
' between one call and the next. Named VlaFrame, not VlaContext, per
' REBUILD.md's own R11 adjudication: one instance is one frame, and
' VLA.bas's own module (the mContextStack behind VlaPushContext/
' VlaPopContext) is the stack of them.
'
' Scope, stated so it is not assumed wider than it is: this snapshots
' VLA.bas's twenty fields only (the ones "the compiler is not reentrant"
' in BETA_ROADMAP.md's F.5 refers to; APROPOSPLUS added AproposRulesCarry
' beside AproposCarry, and LISTOPS-PROVENANCE added GenRow beside AtLine,
' same shape, same reason each needs the same save/restore).
' VLA_English.bas's own module-level state
' (63 variables, the English grammar/parsing engine) is a separate, later
' item - deliberately out of scope here, the same "first cut sized to the
' actual crossed threshold" discipline F.1/F.8/F.11 already used.

Public Macros As Object            ' PDICT.0: late-bound Scripting.Dictionary,
                                    ' was a Collection - see VLA.bas's own
                                    ' mMacros declaration comment.
Public HeadAliases As Collection
Public SubDocs As Collection
Public AproposCarry As String
Public AproposRulesCarry As String
Public FuncName As String
Public InFunction As Boolean
Public TokLines As Variant         ' P-TOK: was a Collection (now a
                                    ' Long() array in VLA.bas) - a bare
                                    ' array-typed Public member is
                                    ' rejected by VBA in a class/object
                                    ' module ("arrays... not allowed as
                                    ' Public members of object modules"),
                                    ' so this holds the array inside a
                                    ' Variant instead, same trick used
                                    ' wherever a class needs to carry an
                                    ' array field. Assignment (either
                                    ' direction) copies the whole array
                                    ' by value, not by pointer; harmless
                                    ' here since VLA.bas's own mTokLines
                                    ' is never mutated in place after a
                                    ' Tokenize call builds it (same
                                    ' "reset always reassigns fresh"
                                    ' discipline this class's own header
                                    ' comment already relies on for
                                    ' every other field).
Public LineOffset As Long
Public FormLines As Collection
Public TcoName As String
Public TcoParams As Collection
Public ProcEmitted As Boolean
Public IncMap As Collection
Public EmitLine As Long
Public AtLine As Long
Public GenRow As String
Public BudgetOn As Boolean
Public ExpandBudget As Long
Public ExpandFired As Long
