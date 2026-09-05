# INTRINSICS — the VBA behaviors a non-VBA port of the translator must match

*PORT.3: written 2026-08-31, alongside PORT.1 (`VlaSetPreludeOverride`,
`VLA.bas`) and PORT.2 (`tools/check_translate_purity.ps1`) — the three
cheap preparations named in the web-runtime discussion this session
had (see `VENTURE.md` §12, `README.md`'s "phrasebooks all the way
down," `SUBSTRATE.md`'s watch-list). Together they convert "a web port
is architecturally possible" into "a web port has a defined API, a
purity guarantee, and this page" before a line of the port exists.*

## Scope — read this before anything below

This document covers exactly one thing: **the semantics the
*translator itself* relies on while turning English text into VLA text
and VLA text into VBA text** — the string, array, and comparison
behaviors `VLA_English.bas`/`VLA.bas` lean on internally. It does
**not** cover VBA-the-target-language's own runtime semantics. A port
that only translates (the "Google Translate" web page: English in,
VLA/VBA text out, nothing executed) never needs to reimplement
`Range`, `Worksheet`, VBA's `Variant` coercion rules, or anything else
the *emitted* code would need to actually run — that machinery stays
exactly where `PORT.1`/`PORT.2` already fenced it: on the far side of
the object model, untouched by translation. If a future port ever
wants to *execute* generated code rather than merely display it, that
is a second, much larger document — this one is deliberately not it.

Also out of scope, and for the same reason: file I/O. `VocabReadFile`
(`VLA_English.bas`) and `VLA_Loader.VlaReadFile` both read text via
`CreateObject("ADODB.Stream")` — a real host dependency, but one that
never enters the translate path `PORT.2`'s ratchet scans, because
`VLA_Browser.bas`'s pure entry points take **already-read text**, not
paths. A port's own file loading (a browser's `File` API, `fetch`
against a bundled asset, Tauri's filesystem API) is the port's own
business and has no VBA behavior to match.

## The intrinsics, one at a time

### 1. Case folding is hand-rolled ASCII, not `LCase`

**VBA behavior, verified at `VLA_Identity.bas`, `Fold`:**

```vba
Public Function Fold(ByVal s As String) As String
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
```

**Why it matters.** Every identifier fold in the system — the head
table, the macro table, the doc table (SD-8's own standing decision)
— goes through this function, and deliberately *not* through VBA's
built-in `LCase`. `LCase` is locale-aware (`BETA_ROADMAP1.md`'s own
audit table: "in a Turkish locale, `I`/`İ` fold inconsistently"), so
identity would silently depend on the end user's Windows regional
settings. `Fold` sidesteps the entire problem by only ever touching
codepoints 65–90 (`A`–`Z`) and passing everything else — accented
letters, non-Latin scripts, punctuation — through completely
unchanged. It is not an approximation of `LCase`; it is a narrower,
simpler, and more predictable operation than `LCase` ever was.

**Port mapping.** Exact, not approximate — this is the one entry on
this page where the JS port can be *more* correct than a naive
translation would suggest, because the VBA original already avoided
every locale trap:

```js
function fold(s) {
  return s.replace(/[A-Z]/g, c => String.fromCharCode(c.charCodeAt(0) + 32));
}
```

Do **not** use `s.toLowerCase()` — it is Unicode-aware and will fold
codepoints (Turkish `İ`, German `ß`-adjacent forms, various
diacritics) that `Fold` deliberately leaves untouched, reintroducing
exactly the divergence `Fold` exists to prevent. The whole point is
narrowness; match the narrowness, not the intent.

### 2. Numeric literals go through `Val`/`Str$`, never `CDbl`/`CStr`

**VBA behavior, verified at `VLA.bas` (comment above
`EvalArithExpand`, and its earlier restatement ~L93):**

> "VBA's `IsNumeric`/`CDbl`/`CStr` were never safe to reuse here — all
> three respect `Application.International`/regional Windows
> settings, which would let identical source text fold to a different
> literal result on two machines. `Val()`/`Str$()` instead
> (Microsoft-documented locale-invariant, always a period decimal
> separator)."

**Why it matters.** The whole promise behind SD-4 ("a shipped
spelling keeps its meaning") depends on `5.5` meaning the same number
on every machine, regardless of the local decimal-separator
convention. `CDbl("5,5")` and `CStr(5.5)` both bend to Windows'
regional settings; `Val`/`Str$` never do.

**Port mapping.** This is the second entry where JS's *default*
behavior already matches the *correct* VBA behavior, not the common
one: `parseFloat("5.5")` and `Number("5.5").toString()` are
locale-invariant by construction — JS has no ambient regional-settings
coupling at all. **The risk runs the opposite direction from the VBA
case**: a JS port is safe by default and stays safe *only* as long as
nobody "improves" number formatting with `toLocaleString()` or an
Intl-aware formatter, which would reintroduce the exact bug `Val`/
`Str$` was chosen to prevent. Treat any `toLocaleString`/`Intl.*` call
appearing anywhere near numeric-literal handling as a regression, not
a feature.

### 3. String indexing is 1-based — except `Split`, which is 0-based

**VBA behavior, verified by usage:** `Mid$`, `InStr`, `Left$`, `Right$`
throughout the tokenizer and matcher (`VLA_English.bas`) are 1-based —
`Mid$(s, 1, 1)` is the *first* character. `Split`, used identically
throughout (e.g. `Split(Replace(text, vbCrLf, vbLf), vbLf)`,
`EnglishToVla`'s own first line), returns a **0-based array** — VBA's
one genuine internal inconsistency here, not a porting artifact.

**Why it matters.** A port that mechanically "subtracts one from every
index because VBA is 1-based" will get every `Split` result
*off-by-one in the wrong direction*, since `Split` was never 1-based
to begin with.

**Port mapping.** Convenient, not coincidental: JS strings and arrays
are uniformly 0-based, and `String.prototype.split` behaves like
VBA's `Split` exactly (same semantics, same edge cases on empty
strings and empty delimiters). The only work is at 1-based call sites:
every `Mid$(s, i, len)` becomes `s.substr(i - 1, len)` (or the
`.slice(i - 1, i - 1 + len)` equivalent), and every `InStr` result
needs a `-1` if it is ever used as a JS index rather than compared
against `0` for "not found" (VBA's `InStr` already returns `0` for
"not found," which happens to equal JS's own `-1`-is-"not found"
convention numerically inverted — don't let the coincidence at `0`
hide the offset everywhere else).

### 4. `Collection` — 1-based, append-only-by-position, insertion-ordered

**VBA behavior:** the matcher's rule tables (`mPatItems`, `mPatForms`,
`mPatTexts`, `mPatSigs`, `mPatSources`, and the many `m*` state
collections `EnglishToVla`/`EnglishResetGrammar` reset) are all
`Collection`, not `Dictionary` — confirmed throughout
`VLA_English.bas`; there is no `Scripting.Dictionary` reference
anywhere in the engine. `Collection.Add` appends; `.Item(n)` is
1-based positional access; `.Remove(n)` removes by position (or by
string key, if one was given at `Add` time — the engine does not use
keyed access); iteration via `For Each` walks in insertion order and
stays stable across `Remove` calls on later elements.

**Why it matters.** `EnglishResetGrammar`'s own pop-back-to-prelude
loop (`Do While mPatItems.Count > mPreludeCount: mPatItems.Remove
mPatItems.Count: ...`) depends on stable insertion order and on
`.Count` tracking removals exactly — the precise property that made a
stale-index bug real and catchable (the same function's own comment
documents a live bug this ordering guarantee was needed to fix).

**Port mapping.** A plain JS `Array` matches every property above
except one: `Collection` is 1-based, `Array` is 0-based, so every
`.Item(n)` becomes `arr[n - 1]` and every `.Count` becomes
`arr.length`. Insertion order and stability under removal both
transfer for free — `Array.prototype.splice` removes by position with
the same "later elements shift down, order preserved" behavior
`Collection.Remove` has. No `Map`/`Set` is needed anywhere on the
translate path; the engine never does keyed lookup through these
tables, only positional walk and positional pop.

### 5. Default string comparison is exact (`Option Compare Binary`) — already a JS-safe default

**VBA behavior:** no scanned module in the translate path declares
`Option Compare Text`; VBA's own module default is `Option Compare
Binary` — plain `=` and `Select Case` on strings are case-sensitive,
codepoint-exact comparisons. The one explicit case-insensitive
comparison in the file (`StrComp(outPath, programPath, vbTextCompare)`)
lives in the **file-based** `EnglishTranslateToVla`/`ToVba` — outside
every function `PORT.2`'s ratchet scans, and outside what a text-only
port needs at all (it compares filesystem paths).

**Why it matters, stated as reassurance rather than a warning:** this
is the one entry on this page that requires **no port work**. JS's
`===` on strings is exact and case-sensitive by default, identical to
VBA's `Option Compare Binary`. A porter who reads "VBA has
locale/case gotchas" (true, per entries 1 and 2 above) and defensively
lowercases or normalizes every string comparison in the matcher would
be *introducing* a divergence, not preventing one — match `===`,
verbatim, and move on.

### 6. Refusals are `Err.Raise`/`Err.Description` — already adapted at the seam

**VBA behavior:** every refusal in the translate path is
`VLA_Messages.RaiseMsg id, k1, v1, k2, v2, ...`, which looks the id up
in a catalogue, substitutes the named slots into a template string,
and raises via `Err.Raise CLng(rec(0)), CStr(rec(1)),
SubstituteSlots(...)`. Callers catch it with a bare `On Error GoTo
failed` and read `Err.Description` for the finished, human-readable
teaching text.

**Why it matters for a port:** this is *not* a gap — `VLA_Browser.bas`
(`PORT.1`) already adapts it: `EnglishTranslateTextToVla`/`ToVba`
catch the raise internally and return `refusal = Err.Description` as
a plain string via `ByRef`, plus a `Boolean` success flag. A JS port's
own equivalent needs only to decide its own idiom (throw an `Error`
whose `.message` is the same text; or return a `{ok, value, refusal}`
result object) — the *text* itself is already finished, catalogue-id
and slot-substitution work done entirely on the VBA side before the
port would ever see it. Nothing about the catalogue, the ids, or the
substitution mechanism needs to be reimplemented; only its two
possible outcomes (a translated string, or a finished refusal string)
cross the boundary.

## What this list is not

Not exhaustive of VBA-the-language — exhaustive of what this
*specific* translator actually leans on, each entry earned by a real
citation above, the same discipline `LESSONS.md` applies to every
other rule on this shelf. When the port is underway and a new
divergence surfaces that isn't listed here, the fix is the same
shape every other gap on this shelf gets: name it, cite the real VBA
behavior, cite the real usage site, state the mapping — add a numbered
entry, don't guess one in from memory.
