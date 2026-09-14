Attribute VB_Name = "VLA_Loader"
Option Explicit
Public Const VLA_LOADER_VERSION As String = "U20.0"
' U20.0: VlaReadFileBytes/VlaWriteFileBytes and a strict, pure-VBA UTF-8
' codec (VlaUtf8Encode/VlaUtf8Decode), for U.20: Lint VLA's writer was
' ANSI Print #, and it turned a phrasebook's pound sign into invalid
' UTF-8. See the block above VlaReadFileBytes.
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

' =====================================================================
'  U.20: byte-honest text files. Every text writer in this codebase was
'  VBA's Open ... For Output plus Print #, which re-encodes through the
'  system ANSI code page: a phrasebook's pound sign (U+00A3) went back to
'  disk as the lone byte A3 (invalid UTF-8), and a character outside that code page
'  would go back as "?". These four replace that path with raw bytes in
'  and out and a strict UTF-8 codec between them - pure VBA, no ADODB,
'  no Declare - so VlaSelfTest covers the codec completely, and it runs
'  the same on Mac.
' =====================================================================

' Every byte of a file into b (0-based); returns the count. b is erased
' first, so an empty file never hands back a previous caller's bytes.
Public Function VlaReadFileBytes(ByVal filePath As String, ByRef b() As Byte) As Long
    Erase b
    Dim f As Integer
    f = FreeFile
    Open filePath For Binary Access Read As #f
    Dim n As Long
    n = LOF(f)
    If n > 0 Then
        ReDim b(0 To n - 1)
        Get #f, , b
    End If
    Close #f
    VlaReadFileBytes = n
End Function

' Replaces filePath with exactly the first n bytes of b. The old file is
' deleted first because Open ... For Binary never truncates - fewer bytes
' written over a longer file would leave its old tail behind - and if it
' could not be deleted (read-only, or open elsewhere) nothing is written.
Public Sub VlaWriteFileBytes(ByVal filePath As String, ByRef b() As Byte, ByVal n As Long)
    On Error Resume Next
    Kill filePath                            ' error 53, no such file, is the ordinary case
    On Error GoTo 0
    Dim f As Integer
    f = FreeFile
    Open filePath For Binary Access Write As #f
    If LOF(f) > 0 Then
        Close #f
        VLA_Messages.RaiseMsg "vla-file-not-replaced", "path", filePath
    End If
    If n > 0 Then
        If UBound(b) = n - 1 Then
            Put #f, , b
        Else
            Dim exact() As Byte
            ReDim exact(0 To n - 1)
            Dim i As Long
            For i = 0 To n - 1
                exact(i) = b(i)
            Next
            Put #f, , exact
        End If
    End If
    Close #f
End Sub

' The UTF-8 bytes of s into b (0-based); returns the count. A surrogate
' pair is one code point, four bytes. A lone surrogate - which no valid
' UTF-8 decode can produce - is written as U+FFFD rather than as a
' sequence no reader would accept.
Public Function VlaUtf8Encode(ByVal s As String, ByRef b() As Byte) As Long
    Erase b
    Dim units As Long
    units = Len(s)
    If units = 0 Then Exit Function
    Dim u() As Byte
    u = s                                    ' UTF-16LE, two bytes per unit
    ReDim b(0 To 3 * units - 1)              ' a pair (2 units) needs 4 bytes, a lone unit at most 3
    Dim j As Long, k As Long, cp As Long, lo As Long
    Do While j < units
        cp = u(2 * j) + u(2 * j + 1) * &H100&
        j = j + 1
        If cp >= &HD800& And cp <= &HDBFF& Then
            lo = -1
            If j < units Then lo = u(2 * j) + u(2 * j + 1) * &H100&
            If lo >= &HDC00& And lo <= &HDFFF& Then
                cp = &H10000 + (cp - &HD800&) * &H400& + (lo - &HDC00&)
                j = j + 1
            Else
                cp = &HFFFD&
            End If
        ElseIf cp >= &HDC00& And cp <= &HDFFF& Then
            cp = &HFFFD&
        End If
        If cp < &H80& Then
            b(k) = cp
            k = k + 1
        ElseIf cp < &H800& Then
            b(k) = &HC0& Or (cp \ &H40&)
            b(k + 1) = &H80& Or (cp And &H3F&)
            k = k + 2
        ElseIf cp < &H10000 Then
            b(k) = &HE0& Or (cp \ &H1000&)
            b(k + 1) = &H80& Or ((cp \ &H40&) And &H3F&)
            b(k + 2) = &H80& Or (cp And &H3F&)
            k = k + 3
        Else
            b(k) = &HF0& Or (cp \ &H40000)
            b(k + 1) = &H80& Or ((cp \ &H1000&) And &H3F&)
            b(k + 2) = &H80& Or ((cp \ &H40&) And &H3F&)
            b(k + 3) = &H80& Or (cp And &H3F&)
            k = k + 4
        End If
    Loop
    ReDim Preserve b(0 To k - 1)
    VlaUtf8Encode = k
End Function

' Strict UTF-8 decode of b(start) .. b(n - 1). True with the text in s;
' or False, s empty, and badAt the 0-based index into b of the first
' byte that does not begin a valid sequence - a stray continuation byte,
' a truncated or overlong sequence, an encoded surrogate, or a code
' point above U+10FFFF. Strict on purpose: a lenient decode turns a bad
' byte into U+FFFD, and writing THAT back would silently change the
' file - the exact failure U.20 exists to stop.
Public Function VlaUtf8Decode(ByRef b() As Byte, ByVal start As Long, ByVal n As Long, _
                              ByRef s As String, ByRef badAt As Long) As Boolean
    s = ""
    badAt = -1
    If n <= start Then
        VlaUtf8Decode = True
        Exit Function
    End If
    Dim u() As Byte
    ReDim u(0 To 2 * (n - start) - 1)        ' never more UTF-16 units than UTF-8 bytes
    Dim i As Long, k As Long, j As Long
    Dim c As Long, cp As Long, need As Long, minCp As Long, hi As Long, lo As Long
    i = start
    Do While i < n
        c = b(i)
        If c < &H80& Then
            cp = c
            need = 0
        ElseIf c >= &HC2& And c <= &HDF& Then
            cp = c And &H1F&
            need = 1
            minCp = &H80&
        ElseIf c >= &HE0& And c <= &HEF& Then
            cp = c And &HF&
            need = 2
            minCp = &H800&
        ElseIf c >= &HF0& And c <= &HF4& Then
            cp = c And &H7&
            need = 3
            minCp = &H10000
        Else
            badAt = i
            Exit Function
        End If
        For j = 1 To need
            If i + j >= n Then
                badAt = i
                Exit Function
            End If
            If (b(i + j) And &HC0&) <> &H80& Then
                badAt = i
                Exit Function
            End If
            cp = (cp * &H40&) Or (b(i + j) And &H3F&)
        Next
        If need > 0 Then
            If cp < minCp Or (cp >= &HD800& And cp <= &HDFFF&) Or cp > &H10FFFF Then
                badAt = i
                Exit Function
            End If
        End If
        If cp < &H10000 Then
            u(2 * k) = cp And &HFF&
            u(2 * k + 1) = cp \ &H100&
            k = k + 1
        Else
            cp = cp - &H10000
            hi = &HD800& + cp \ &H400&
            lo = &HDC00& + (cp And &H3FF&)
            u(2 * k) = hi And &HFF&
            u(2 * k + 1) = hi \ &H100&
            u(2 * k + 2) = lo And &HFF&
            u(2 * k + 3) = lo \ &H100&
            k = k + 2
        End If
        i = i + need + 1
    Loop
    If k > 0 Then
        ReDim Preserve u(0 To 2 * k - 1)
        s = u
    End If
    VlaUtf8Decode = True
End Function
