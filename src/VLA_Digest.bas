Attribute VB_Name = "VLA_Digest"
Option Explicit
Public Const VLA_DIGEST_VERSION As String = "OPTIMIZE.1"
' OPTIMIZE.1 (step 1): a canonical, INJECTIVE encoding of a memo's
' inputs, hashed by the SHA-256 above - VlaWtf8Encode plus the
' VlaKeyBegin/VlaKeyAdd*/VlaKeyHex framing at the foot of this module.
' Two reasons it could not be a call to what already existed:
' VlaSha256HexOfAsciiText goes through StrConv's ANSI code page, so
' "Zoe" with a diaeresis and "Zoe?" could share a key; and
' VLA_Loader.VlaUtf8Encode writes a LONE surrogate as U+FFFD, which is
' right for a file (no reader accepts the raw sequence) and wrong for a
' key (two different strings would collide, and a memo would hand one
' workbook's answer to another's inputs). See the foot of this module.

' =====================================================================
'  VLA_Digest - SEC.11: SHA-256, in VBA, depending on nothing.
'
'  WHY THIS EXISTS. SEC.2 remembers a person's consent to load a raw-
'  bearing phrasebook, keyed by a hash of that phrasebook's content.
'  Until 0.5.3 that key was h = h*31 + b mod 2^32 - a Java-string-hash
'  polynomial. Second preimages against it are trivial to construct by
'  hand (flip two bytes of a ; comment to restore the running value),
'  so a grant given to a phrasebook a person actually read transferred
'  cleanly to a crafted one they never saw. A consent record is worth
'  no more than the identity it is keyed to, so the key had to become
'  a real digest.
'
'  WHY PURE VBA, and not the .NET class the roadmap first proposed.
'  SEC.11 was minted proposing
'  CreateObject("System.Security.Cryptography.SHA256Managed").
'  MEASURED 2026-09-08 on the owner's own development machine, from a
'  real COM host (cscript - no Office object anywhere), in BOTH
'  bitnesses: that CreateObject FAILS, 0x80131700. The registry says
'  why. The ProgID resolves to a CLSID whose InprocServer32 is
'  mscoree.dll with Assembly = "mscorlib, Version=2.0.0.0" - the .NET
'  Framework 2.0 shim. Only v4 is installed; there is no v2.0.50727
'  key at all. That is not an unusual box, it is the Windows 11
'  default: .NET Framework 3.5 ships as an optional feature, switched
'  off. So the proposed mechanism was not merely "not guaranteed
'  everywhere" - it was unavailable on the very machine Frazaro is
'  developed on, and would have made every raw-bearing phrasebook
'  unloadable here.
'
'  The other two candidates were ruled out on house grounds, not
'  taste. CAPICOM is not registered on this machine either (measured,
'  same probe) and Microsoft unshipped it years ago. A bcrypt.dll /
'  advapi32 P/Invoke would work, but SUBSTRATE.md's H.4 census pins
'  this codebase at ONE Declare site (frmCLI.frm's four window calls)
'  and proposes a ratchet to hold it there; paying for a security fix
'  with a new bitness-sensitive P/Invoke surface is a bad trade in
'  exactly the direction that census was taken to prevent.
'
'  So: about 150 lines of arithmetic. No dependency, no fallback path
'  to test, identical on every machine and both bitnesses, and - the
'  part that matters most for an item like this - PURE. It touches no
'  host object, opens no file and makes no network call (SD-13
'  untouched, trivially), which is what lets VlaSelfTest cover it
'  COMPLETELY against the published FIPS 180-4 vectors, rather than
'  partially against whatever the current machine happens to have
'  registered.
'
'  THE ARITHMETIC, and why it looks like this. VBA has no unsigned
'  32-bit integer. Every 32-bit word here lives in a signed Long
'  holding the bit pattern the unsigned value would have, and every
'  operation that could overflow goes through UnsignedOf/WordOf, which
'  widen to Double and fold back mod 2^32. A Double holds integers
'  exactly to 2^53, and every intermediate below stays under 2^32 -
'  ShiftLeft masks off the bits that would leave the word BEFORE
'  multiplying, precisely so the product cannot exceed 2^32 and lose
'  exactness. And, Or, Xor and Not are already bitwise on a Long in
'  VBA, so those need no help.
'
'  THE CONSTANT TABLES are two hex strings rather than 72 scattered
'  literals. Two reasons, both practical: VBA caps a statement at 25
'  line continuations, so a 64-element Array() will not compile; and
'  one string can be diffed character-for-character against FIPS
'  180-4's own printed table, which 64 separate &H literals cannot.
' =====================================================================

Private mK(0 To 63) As Long
Private mH(0 To 7) As Long
Private mReady As Boolean

' OPTIMIZE.1: the eight bytes of a Double, read without a Declare.
' LSet between two user-defined types of the same size copies their
' bytes, which is the whole reason these two exist - SUBSTRATE.md's H.4
' census pins this codebase at ONE Declare site, and a memo key is not
' the place to add a second, bitness-sensitive one (the same trade
' SEC.11's own header records for SHA-256 itself). VBA requires a Type
' in the declarations section, which is why they sit here rather than
' beside VlaDoubleBytes at the foot of the module.
Private Type VlaKeyDouble
    d As Double
End Type

Private Type VlaKeyBytes8
    b(0 To 7) As Byte
End Type

' OPTIMIZE.1: the type tags a framed key field carries. Public so a test
' can name them, and so a reader of a key's bytes can decode it by hand
' (GENSYM's "hand-derivable" instinct, applied to a hash's input).
Public Const VLA_KEY_TAG_BLANK As Long = 0
Public Const VLA_KEY_TAG_TEXT As Long = 1
Public Const VLA_KEY_TAG_NUMBER As Long = 2
Public Const VLA_KEY_TAG_BOOLEAN As Long = 3
Public Const VLA_KEY_TAG_ERROR As Long = 4
Public Const VLA_KEY_TAG_COUNT As Long = 5
Public Const VLA_KEY_TAG_NULL As Long = 6
Public Const VLA_KEY_TAG_OTHER As Long = 7

' --- 32-bit word arithmetic on a signed Long -------------------------

' The bit pattern of x, read as an unsigned 32-bit value.
Private Function UnsignedOf(ByVal x As Long) As Double
    If x < 0 Then
        UnsignedOf = CDbl(x) + 4294967296#
    Else
        UnsignedOf = CDbl(x)
    End If
End Function

' d folded mod 2^32 and read back as the signed Long carrying that bit
' pattern - the inverse of UnsignedOf.
Private Function WordOf(ByVal d As Double) As Long
    Dim v As Double
    v = d - Int(d / 4294967296#) * 4294967296#
    If v >= 2147483648# Then
        WordOf = CLng(v - 4294967296#)
    Else
        WordOf = CLng(v)
    End If
End Function

Private Function AddW(ByVal x As Long, ByVal y As Long) As Long
    AddW = WordOf(UnsignedOf(x) + UnsignedOf(y))
End Function

' Logical (not arithmetic) right shift: the sign bit must travel like
' any other bit, which VBA's own \ on a negative Long would not do.
Private Function ShiftRight(ByVal x As Long, ByVal n As Long) As Long
    ShiftRight = WordOf(Int(UnsignedOf(x) / (2 ^ n)))
End Function

' The masking step is load-bearing, not defensive: without it the
' product could reach 2^63 and leave the range in which a Double holds
' integers exactly.
Private Function ShiftLeft(ByVal x As Long, ByVal n As Long) As Long
    Dim v As Double, keep As Double
    v = UnsignedOf(x)
    keep = 2 ^ (32 - n)
    v = v - Int(v / keep) * keep
    ShiftLeft = WordOf(v * (2 ^ n))
End Function

Private Function RotR(ByVal x As Long, ByVal n As Long) As Long
    RotR = ShiftRight(x, n) Or ShiftLeft(x, 32 - n)
End Function

' Eight hex digits to the signed Long carrying that bit pattern.
' Parsed a digit at a time rather than through Val("&H...") because
' VBA's rules for when an 8-digit hex literal is a Long and when it
' overflows are exactly the kind of thing that differs between hosts.
Private Function HexWord(ByVal h As String) As Long
    Dim d As Double, i As Long, p As Long
    d = 0
    For i = 1 To Len(h)
        p = InStr(1, "0123456789abcdef", LCase$(Mid$(h, i, 1))) - 1
        d = d * 16 + p
    Next i
    HexWord = WordOf(d)
End Function

Private Function WordHex(ByVal x As Long) As String
    Dim d As Double, s As String, i As Long, dig As Long
    d = UnsignedOf(x)
    s = ""
    For i = 1 To 8
        dig = CLng(d - Int(d / 16) * 16)
        s = Mid$("0123456789ABCDEF", dig + 1, 1) & s
        d = Int(d / 16)
    Next i
    WordHex = s
End Function

Private Sub EnsureTables()
    If mReady Then Exit Sub
    Dim kText As String, hText As String, i As Long
    ' FIPS 180-4 s4.2.2 - the first 32 bits of the fractional parts of
    ' the cube roots of the first 64 primes.
    kText = "428a2f9871374491b5c0fbcfe9b5dba53956c25b59f111f1923f82a4ab1c5ed5" & _
            "d807aa9812835b01243185be550c7dc372be5d7480deb1fe9bdc06a7c19bf174" & _
            "e49b69c1efbe47860fc19dc6240ca1cc2de92c6f4a7484aa5cb0a9dc76f988da" & _
            "983e5152a831c66db00327c8bf597fc7c6e00bf3d5a7914706ca635114292967" & _
            "27b70a852e1b21384d2c6dfc53380d13650a7354766a0abb81c2c92e92722c85" & _
            "a2bfe8a1a81a664bc24b8b70c76c51a3d192e819d6990624f40e3585106aa070" & _
            "19a4c1161e376c082748774c34b0bcb5391c0cb34ed8aa4a5b9cca4f682e6ff3" & _
            "748f82ee78a5636f84c878148cc7020890befffaa4506cebbef9a3f7c67178f2"
    For i = 0 To 63
        mK(i) = HexWord(Mid$(kText, i * 8 + 1, 8))
    Next i
    ' FIPS 180-4 s5.3.3 - the first 32 bits of the fractional parts of
    ' the square roots of the first 8 primes.
    hText = "6a09e667bb67ae853c6ef372a54ff53a510e527f9b05688c1f83d9ab5be0cd19"
    For i = 0 To 7
        mH(i) = HexWord(Mid$(hText, i * 8 + 1, 8))
    Next i
    mReady = True
End Sub

' --- the digest ------------------------------------------------------

' SHA-256 of the first n bytes of b, as 64 uppercase hex digits.
' n = 0 is legal and hashes the empty message; b may be unallocated in
' that case, which is why nothing indexes it before the loop bound.
Public Function VlaSha256Hex(ByRef b() As Byte, ByVal n As Long) As String
    EnsureTables

    Dim total As Long
    total = ((n + 9 + 63) \ 64) * 64
    Dim m() As Byte
    ReDim m(0 To total - 1)

    Dim i As Long
    For i = 0 To n - 1
        m(i) = b(i)
    Next i
    m(n) = 128                          ' the mandatory 1 bit, then zeros

    ' Message length in BITS, big-endian, in the last eight bytes. A
    ' Double carries this exactly for any file VBA could have read.
    Dim bits As Double
    bits = CDbl(n) * 8
    For i = 0 To 7
        m(total - 1 - i) = CByte(bits - Int(bits / 256) * 256)
        bits = Int(bits / 256)
    Next i

    Dim h0 As Long, h1 As Long, h2 As Long, h3 As Long
    Dim h4 As Long, h5 As Long, h6 As Long, h7 As Long
    h0 = mH(0): h1 = mH(1): h2 = mH(2): h3 = mH(3)
    h4 = mH(4): h5 = mH(5): h6 = mH(6): h7 = mH(7)

    Dim w(0 To 63) As Long
    Dim wa As Long, wb As Long, wc As Long, wd As Long
    Dim we As Long, wf As Long, wg As Long, wh As Long
    Dim s0 As Long, s1 As Long, ch As Long, maj As Long
    Dim ep0 As Long, ep1 As Long, t1 As Long, t2 As Long
    Dim blk As Long, blockBase As Long, t As Long

    For blk = 0 To total \ 64 - 1
        blockBase = blk * 64
        For t = 0 To 15
            w(t) = WordOf(CDbl(m(blockBase + t * 4)) * 16777216# _
                        + CDbl(m(blockBase + t * 4 + 1)) * 65536# _
                        + CDbl(m(blockBase + t * 4 + 2)) * 256# _
                        + CDbl(m(blockBase + t * 4 + 3)))
        Next t
        For t = 16 To 63
            s0 = RotR(w(t - 15), 7) Xor RotR(w(t - 15), 18) Xor ShiftRight(w(t - 15), 3)
            s1 = RotR(w(t - 2), 17) Xor RotR(w(t - 2), 19) Xor ShiftRight(w(t - 2), 10)
            w(t) = AddW(AddW(AddW(w(t - 16), s0), w(t - 7)), s1)
        Next t

        wa = h0: wb = h1: wc = h2: wd = h3
        we = h4: wf = h5: wg = h6: wh = h7

        For t = 0 To 63
            ep1 = RotR(we, 6) Xor RotR(we, 11) Xor RotR(we, 25)
            ch = (we And wf) Xor ((Not we) And wg)
            t1 = AddW(AddW(AddW(AddW(wh, ep1), ch), mK(t)), w(t))
            ep0 = RotR(wa, 2) Xor RotR(wa, 13) Xor RotR(wa, 22)
            maj = (wa And wb) Xor (wa And wc) Xor (wb And wc)
            t2 = AddW(ep0, maj)
            wh = wg
            wg = wf
            wf = we
            we = AddW(wd, t1)
            wd = wc
            wc = wb
            wb = wa
            wa = AddW(t1, t2)
        Next t

        h0 = AddW(h0, wa): h1 = AddW(h1, wb): h2 = AddW(h2, wc): h3 = AddW(h3, wd)
        h4 = AddW(h4, we): h5 = AddW(h5, wf): h6 = AddW(h6, wg): h7 = AddW(h7, wh)
    Next blk

    VlaSha256Hex = WordHex(h0) & WordHex(h1) & WordHex(h2) & WordHex(h3) & _
                   WordHex(h4) & WordHex(h5) & WordHex(h6) & WordHex(h7)
End Function

' SEC.11's actual key: SHA-256 over only those bytes that are not tab,
' LF, CR or space, plus (through `counted`) how many took part.
'
' The filter is inherited deliberately from the 0.5.1 polynomial this
' replaces, and it is load-bearing rather than incidental - see
' EnglishSourceHash's own header in VLA_SentenceEngine.bas for the
' full argument. In one line: this repo's .gitattributes checks .bas
' and .vla out as CRLF, so a raw-byte digest would give one phrasebook
' two different identities on two machines and silently invalidate
' consent on both. Skipping whitespace costs nothing cryptographically
' - an attacker hunting a second preimage was always free to vary the
' non-whitespace bytes, and that is the whole space SHA-256 is hard
' over. What the filter removes is not attacker freedom, it is a
' portability defect.
Public Function VlaSha256HexSkippingWhitespace(ByRef b() As Byte, ByVal n As Long, ByRef counted As Long) As String
    Dim packed() As Byte
    Dim i As Long, k As Long
    If n > 0 Then ReDim packed(0 To n - 1)
    k = 0
    For i = 0 To n - 1
        If b(i) <> 9 And b(i) <> 10 And b(i) <> 13 And b(i) <> 32 Then
            packed(k) = b(i)
            k = k + 1
        End If
    Next i
    counted = k
    VlaSha256HexSkippingWhitespace = VlaSha256Hex(packed, k)
End Function

' --- text conveniences, for VlaSelfTest's published vectors ----------
'
' StrConv vbFromUnicode uses the system ANSI code page, so these are
' exact for ASCII and nothing else. Every vector pinned against them
' is ASCII on purpose; anything that must be byte-honest (a real
' phrasebook, which may carry a BOM or non-ASCII text) goes through
' the byte entry points above, which decode nothing.
Public Function VlaSha256HexOfAsciiText(ByVal s As String) As String
    Dim b() As Byte
    If Len(s) = 0 Then
        VlaSha256HexOfAsciiText = VlaSha256Hex(b, 0)
        Exit Function
    End If
    b = StrConv(s, vbFromUnicode)
    VlaSha256HexOfAsciiText = VlaSha256Hex(b, Len(s))
End Function

Public Function VlaSha256HexOfAsciiTextSkippingWhitespace(ByVal s As String, ByRef counted As Long) As String
    Dim b() As Byte
    If Len(s) = 0 Then
        counted = 0
        VlaSha256HexOfAsciiTextSkippingWhitespace = VlaSha256Hex(b, 0)
        Exit Function
    End If
    b = StrConv(s, vbFromUnicode)
    VlaSha256HexOfAsciiTextSkippingWhitespace = VlaSha256HexSkippingWhitespace(b, Len(s), counted)
End Function

' --- OPTIMIZE.1, step 1: a canonical, injective key ------------------
'
'  WHY THIS EXISTS. OPTIMIZE's session memo (its standing decision 1)
'  is keyed on the CONTENT of a program's rules text and of every input
'  Table. The memo may only make a search faster and may never change
'  what it answers (decision 2), so the key must be INJECTIVE: two
'  different inputs may never share one key, or one workbook's answer
'  would be handed to another workbook's inputs. Everything below
'  exists to make that true, and nothing below knows what a spreadsheet
'  is - the whole section is arithmetic over bytes, like the digest
'  above it.
'
'  THE THREE PROPERTIES, and what each one buys:
'
'  1. TEXT AS WTF-8. VlaWtf8Encode is byte-identical to UTF-8 for
'     every well-formed string, and encodes a LONE surrogate as its own
'     three bytes rather than as U+FFFD. VLA_Loader.VlaUtf8Encode does
'     the opposite, correctly, because it writes FILES: no reader
'     accepts a raw surrogate sequence, so U+FFFD is the honest thing
'     to put on disk. For a KEY it is a collision - "a" + U+D800 and
'     "a" + U+FFFD would be one key - and a collision here is a wrong
'     answer. Injective on every string VBA can hold.
'
'  2. EVERY FIELD TYPED AND LENGTH-FRAMED. A field is one TAG byte,
'     then a four-byte big-endian LENGTH, then exactly that many bytes.
'     So 1 and "1" differ (the tag), blank and "" differ (the tag), and
'     ("ab", "c") and ("a", "bc") differ (the lengths).
'     VLA_Relation's own Chr$(31) TupleKey joins are NOT injective and
'     are deliberately not reused here: a cell holding that one
'     character would forge a boundary.
'
'  3. NUMBERS BY THEIR EXACT BITS. A number is its Double's eight
'     bytes, never its text, so no locale, no rounding and no cell
'     format enters the key. Dates are already Doubles in Value2, so
'     they arrive by that same path.
'
'  WHAT IS DELIBERATELY NOT DISTINGUISHED. Each one is a MISS (a
'  needless re-solve, which costs time and changes no answer) and never
'  a collision:
'    - two different error values share the error tag. An error value
'      inside a table argument is already refused upstream -
'      VLA_Relation.RelFromRange's own all-blank test calls CStr on
'      every cell, and CStr raises 13 on an error value - so the case
'      cannot reach a memoized answer. When a later item gives an error
'      cell a meaning, the tag gains a payload and every key changes,
'      which is what the version tag in the key's first field is for.
'    - minus zero and zero have different bit patterns and answer
'      identically.
'    - an object or an array cannot sit in a Relation tuple at all (the
'      same CStr above), and so gets the OTHER tag with its VarType as
'      the payload rather than a refusal. That is on purpose: it keeps
'      this module depending on NOTHING, not even on VLA_Messages,
'      which is the property SEC.11's own header is built on.
'
'  BYTE ORDER. Lengths and counts are BIG-endian, chosen so a key's
'  bytes can be read left to right by hand. A Double's eight bytes are
'  whatever LSet lays down, which is little-endian on every host VBA
'  runs on; the pins hold those exact bytes, so the layout is pinned
'  rather than assumed.

' The WTF-8 bytes of s into b (0-based); returns the count. A surrogate
' PAIR is one code point and four bytes, exactly as UTF-8 has it. A
' LONE surrogate keeps its own value and gets the three bytes any code
' point of that value would get (U+D800 is ED A0 80): the WTF-8
' convention, and the one difference from VLA_Loader.VlaUtf8Encode,
' whose own header says why a file needs the other behaviour.
Public Function VlaWtf8Encode(ByVal s As String, ByRef b() As Byte) As Long
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
            End If
            ' No else: an unpaired high surrogate keeps cp as itself and
            ' falls through to the three-byte arm below. That single
            ' absent branch IS the difference from VlaUtf8Encode.
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
    VlaWtf8Encode = k
End Function

' The eight bytes of d into b (0-based); returns 8. See the two Types
' in the declarations section for why this needs no Declare.
Public Function VlaDoubleBytes(ByVal d As Double, ByRef b() As Byte) As Long
    Dim dv As VlaKeyDouble
    Dim bv As VlaKeyBytes8
    dv.d = d
    LSet bv = dv
    ReDim b(0 To 7)
    Dim i As Long
    For i = 0 To 7
        b(i) = bv.b(i)
    Next i
    VlaDoubleBytes = 8
End Function

' --- the key buffer --------------------------------------------------
'
' A key is built into a caller-held Byte array plus a count, rather
' than into module state: a key is built once inside one call, and
' module state is exactly what a VBA project reset wipes (LESSONS.md
' XXXII, and OPTIMIZE.0's probe C7, which found a reset re-running
' every UDF in the workbook). Nothing here is a hot loop either -
' one key per OPTIMIZE call, not one per search node.

' buf grown geometrically so a long rules text or a wide Table does not
' pay a ReDim Preserve per byte (OPTIMIZE.0.C's "arrays grown
' geometrically, never ReDim Preserve per element").
Private Sub KeyEnsure(ByRef buf() As Byte, ByVal n As Long, ByVal extra As Long)
    Dim have As Long
    have = 0
    On Error Resume Next
    have = UBound(buf) + 1
    On Error GoTo 0
    If n + extra <= have Then Exit Sub
    Dim want As Long
    want = have
    If want < 64 Then want = 64
    Do While want < n + extra
        want = want * 2
    Loop
    If have = 0 Then
        ReDim buf(0 To want - 1)
    Else
        ReDim Preserve buf(0 To want - 1)
    End If
End Sub

' One field: the tag byte, the four-byte big-endian length, the bytes.
' payload may be unallocated when payloadLen is 0 - the copy loop then
' runs zero times, which is why nothing indexes it before the bound.
Private Sub KeyAddField(ByRef buf() As Byte, ByRef n As Long, ByVal fieldTag As Long, _
                        ByRef payload() As Byte, ByVal payloadLen As Long)
    KeyEnsure buf, n, 5 + payloadLen
    Dim d As Double
    Dim i As Long
    buf(n) = CByte(fieldTag)
    n = n + 1
    d = CDbl(payloadLen)
    For i = 3 To 0 Step -1
        buf(n + i) = CByte(d - Int(d / 256) * 256)
        d = Int(d / 256)
    Next i
    n = n + 4
    For i = 0 To payloadLen - 1
        buf(n + i) = payload(i)
    Next i
    n = n + payloadLen
End Sub

' Start (or restart) a key.
Public Sub VlaKeyBegin(ByRef buf() As Byte, ByRef n As Long)
    Erase buf
    n = 0
End Sub

' A text field: the WTF-8 bytes of s, length-framed.
Public Sub VlaKeyAddText(ByRef buf() As Byte, ByRef n As Long, ByVal s As String)
    Dim p() As Byte
    Dim pn As Long
    pn = VlaWtf8Encode(s, p)
    KeyAddField buf, n, VLA_KEY_TAG_TEXT, p, pn
End Sub

' A count field: four big-endian bytes under its own tag, so a row
' count can never be read as a cell and vice versa.
Public Sub VlaKeyAddCount(ByRef buf() As Byte, ByRef n As Long, ByVal v As Long)
    Dim p(0 To 3) As Byte
    Dim d As Double
    Dim i As Long
    d = CDbl(v)
    If d < 0 Then d = d + 4294967296#
    For i = 3 To 0 Step -1
        p(i) = CByte(d - Int(d / 256) * 256)
        d = Int(d / 256)
    Next i
    KeyAddField buf, n, VLA_KEY_TAG_COUNT, p, 4
End Sub

' One cell of one tuple, tagged by what it IS. Every VarType a Value2
' read or a hand-built test tuple can hold has an arm; see the section
' header for what the last arm deliberately does not distinguish.
' IsError/VarType only - never CStr, which raises 13 on an error value
' (DATALOG.15's own lesson, VLA_Relation.SpillHeaderCheck).
Public Sub VlaKeyAddCell(ByRef buf() As Byte, ByRef n As Long, ByVal v As Variant)
    Dim p() As Byte
    Dim pn As Long
    Select Case VarType(v)
    Case vbEmpty
        KeyAddField buf, n, VLA_KEY_TAG_BLANK, p, 0
    Case vbNull
        KeyAddField buf, n, VLA_KEY_TAG_NULL, p, 0
    Case vbError
        KeyAddField buf, n, VLA_KEY_TAG_ERROR, p, 0
    Case vbBoolean
        ReDim p(0 To 0)
        If CBool(v) Then p(0) = 1 Else p(0) = 0
        KeyAddField buf, n, VLA_KEY_TAG_BOOLEAN, p, 1
    Case vbString
        VlaKeyAddText buf, n, CStr(v)
    Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDate, vbDecimal
        pn = VlaDoubleBytes(CDbl(v), p)
        KeyAddField buf, n, VLA_KEY_TAG_NUMBER, p, pn
    Case Else
        ReDim p(0 To 3)
        p(0) = 0
        p(1) = 0
        p(2) = CByte((VarType(v) \ 256) And 255)
        p(3) = CByte(VarType(v) And 255)
        KeyAddField buf, n, VLA_KEY_TAG_OTHER, p, 4
    End Select
End Sub

' The key itself: SHA-256 over exactly the bytes framed so far, as 64
' uppercase hex digits. Nothing is skipped here - unlike
' VlaSha256HexSkippingWhitespace, whose filter exists for a FILE's
' portability across line endings. A key's bytes are constructed, not
' read off a disk, so every one of them counts.
Public Function VlaKeyHex(ByRef buf() As Byte, ByVal n As Long) As String
    VlaKeyHex = VlaSha256Hex(buf, n)
End Function
