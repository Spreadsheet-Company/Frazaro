Attribute VB_Name = "VLA_Loader"
Option Explicit
Public Const VLA_LOADER_VERSION As String = "LX2.0"
' LX3.0: the identity-deciding LCase$ site now folds through
' VLA_Identity.Fold (invariant ASCII, never locale-dependent) - R6/SD-8.
' LX2.0: this module's 3 raw Err.Raise refusal sites now route through
' VLA_Messages.RaiseMsg with a stable id - SD-2/LX.2's first migrated
' batch. Rendered text and Err.Number/Err.Source are unchanged (verified
' by manual trace at migration time; VLA_Messages.bas's own header
' explains the substitution mechanism).

' =====================================================================
'  VLA_Loader - import standalone .vla files into the running workbook.
'
'  Typical use:
'    VlaRunFile ThisWorkbook.Path & "\scripts\report.vla"
'      -> reads the file, transpiles it, injects the VBA into a module
'         named "report", then runs Public Sub main() in that module.
'
'    VlaImportFile path                -> import/refresh only, no run
'    VlaRunFile path, "build-report"   -> run a different entry point
'
'  Requirements (one-time):
'    * Workbook saved as .xlsm with macros enabled
'    * File > Options > Trust Center > Trust Center Settings >
'      Macro Settings > "Trust access to the VBA project object model"
'
'  Convention: each .vla file defines (sub main () ...) as its entry
'  point. The module name is derived from the file name, so two files
'  with the same base name will overwrite each other's module.
' =====================================================================

' Modules the loader must never overwrite (the transpiler and itself).
Private Function IsProtectedModule(ByVal name As String) As Boolean
    Select Case VLA_Identity.Fold(name)
        Case "vla", "vla_loader", "vla_demo"
            IsProtectedModule = True
    End Select
End Function

' Import (or refresh) a .vla file as a standard module.
' Returns the module name it compiled into.
Public Function VlaImportFile(ByVal filePath As String, _
                              Optional ByVal moduleName As String = "") As String
    If Len(Dir$(filePath)) = 0 Then
        VLA_Messages.RaiseMsg "vla-source-not-found", "path", filePath
    End If
    If Len(moduleName) = 0 Then moduleName = ModuleNameFromPath(filePath)
    If IsProtectedModule(moduleName) Then
        VLA_Messages.RaiseMsg "vla-protected-module-name", "name", moduleName
    End If

    Dim src As String
    src = ReadTextFile(filePath)
    VlaCompileToModule src, moduleName
    VlaImportFile = moduleName
End Function

' Import a .vla file and immediately run its entry point.
Public Sub VlaRunFile(ByVal filePath As String, _
                      Optional ByVal entryPoint As String = "main", _
                      Optional ByVal moduleName As String = "")
    Dim m As String
    m = VlaImportFile(filePath, moduleName)
    ' Same identifier rule as the transpiler: hyphens -> underscores.
    Application.Run m & "." & Replace(entryPoint, "-", "_")
End Sub

' Example button target: assign this to a worksheet button (Insert >
' Shapes or Form Control button > Assign Macro) and adjust the path.
Public Sub VlaReloadAndRun()
    VlaRunFile ThisWorkbook.Path & "\scripts\main.vla"
End Sub

' ---------------------------------------------------------------------
'  Helpers
' ---------------------------------------------------------------------

' Derive a legal VBA module name from a file path:
' "C:\x\gl-report.vla" -> "gl_report"
Private Function ModuleNameFromPath(ByVal filePath As String) As String
    Dim base As String
    Dim p As Long
    base = filePath
    p = InStrRev(base, "\")
    If p = 0 Then p = InStrRev(base, "/")
    If p > 0 Then base = Mid$(base, p + 1)
    p = InStrRev(base, ".")
    If p > 1 Then base = Left$(base, p - 1)

    Dim r As String
    Dim i As Long, c As String
    For i = 1 To Len(base)
        c = Mid$(base, i, 1)
        If c Like "[A-Za-z0-9_]" Then
            r = r & c
        Else
            r = r & "_"                      ' hyphens, spaces, dots...
        End If
    Next
    If Len(r) = 0 Then r = "vla_module"
    If Not Left$(r, 1) Like "[A-Za-z]" Then r = "m" & r
    ModuleNameFromPath = r
End Function

' Public entry point for reading any text file (English sources,
' VLA sources, or anything else): UTF-8 with ANSI fallback, and a
' clear error if the path is wrong.
Public Function VlaReadFile(ByVal filePath As String) As String
    If Len(Dir$(filePath)) = 0 Then
        VLA_Messages.RaiseMsg "vla-file-not-found", "path", filePath
    End If
    VlaReadFile = ReadTextFile(filePath)
End Function

' Read a text file as UTF-8 (via ADODB.Stream), falling back to ANSI.
' Save your .vla files as UTF-8; plain ASCII is safe either way.
Private Function ReadTextFile(ByVal filePath As String) As String
    On Error GoTo ansiFallback
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2                              ' adTypeText
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile filePath
    ReadTextFile = st.ReadText(-1)           ' adReadAll (BOM handled)
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
        ReadTextFile = ""
        Exit Function
    End If
    ReDim b(0 To LOF(f) - 1)
    Get #f, , b
    Close #f
    ReadTextFile = StrConv(b, vbUnicode)
End Function
