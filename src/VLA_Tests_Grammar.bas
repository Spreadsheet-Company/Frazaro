Attribute VB_Name = "VLA_Tests_Grammar"
Option Explicit
Public Const VLA_TESTS_GRAMMAR_VERSION As String = "G6.0"
' G6.0: TestG6 - list-valued slots ({name:cat-list}). Two item
' categories pinned (text-shaped field names, column-shaped via
' RefShapeOk) to prove the mechanism is generic, not hardcoded to one
' shape; the Oxford-comma disambiguation is pinned against the real
' motivating two-clause sentence from pareto.txt section 10 ("with rows
' of A, B and columns of C"), not just a single-list sentence, since
' that clause-boundary case is the one this design had to get right
' with zero lookahead. A dropped Oxford comma is pinned as a genuine
' refusal (translation fails), not a silent misparse. See
' VLA_English.bas's own G6.0 header note for the full mechanism.
' GROWLOOP.0: TestGRowLoop - G-ROWLOOP's row loop. Checked against the
' live grammar before writing anything: "Count <name> [down] from <a>
' to <b>:" (VLA_English.bas Case "count") already covers ascending and
' the safe bottom-up descending shape - a first draft of this item
' invented a parallel "For each row from ... to ...:" surface before
' that check caught the duplication, reverted whole. The one shape
' actually missing was a custom step for skip-N row loops ("every
' other row"), so this pass only adds an optional "step <s>" clause to
' the existing Count form - a "down" step is negated in the emitted
' form, (- 0 s), since this grammar has no way to say a negative
' number in words. Five pins, all against plain "Set x to row."
' bodies (no vocab needed, matching G8's style): plain ascending and
' plain descending as no-regression checks, ascending-with-step,
' descending-with-step (proving the auto-negation), and one proving
' the loop variable composes with the pre-existing "cell in column C
' row {n:expr}" expression grammar for free - "row" is just another
' {n:expr} there, no new plumbing.
' G8.0: TestG8 - number words and ordinals. Three pins against the
' core "Set v to {expr}." sentence: a cardinal past twelve, a bare
' ordinal, and - the one that actually justifies scoping ordinals to
' ParsePrimCore instead of the tokenizer - "first of X" still resolving
' to vlafirst rather than being shadowed by ordinal "first" -> 1.
' G7.0: TestG7 - slot defaults, {n:expr=1}, generalizing G1's optional-
' literal present/absent duality to typed slots. Two pins: the slot
' present (its value wins over the default) and the slot AND its
' leading [at] both absent (the default substitutes, matching to
' completion rather than a near-miss).
' G11r.0: TestG11r - the template-face remainder G11 deliberately left
' open (true/false-valued members and the numberformat family), shipped
' by exactly G11's fired-only lockstep method. Two pins: an antonym-pair
' functor (hide-row/unhide-row, proving the {d} substitution selects the
' right macro BODY, not just the right member name) and the three-way
' format-as functor (currency/percent/date).
' F8.0: new module, split out of VLA_Tests.bas (F.8 - that file crossed
' REBUILD.md R4's 1,200-line budget by 4x, at 4,945 lines). This half
' holds every grammar/phrasebook feature pin (the G*/L*/V*/PL* series
' plus TestAlonzoLib) - the largest single concern in the old file and
' the one most likely to keep growing as Grammar work continues.
' VLA_Tests.bas keeps VlaSelfTest (the dispatcher), the shared
' assertion/report infrastructure (now Public so this module can call
' it), VerifyReport, the goldens machinery, and the core-mechanics
' pins (emitter/loader/IDE/runtime/head-table/interpreter). Moving
' Private Test subs across a module boundary requires Public - VBA
' cannot call a Private Sub from another module - so every TestXxx
' below changed from Private to Public; nothing else about them
' changed. WriteTempLib (TestPL7's temp-file helper) stayed Private -
' it is used only within this module.
' This is a step toward REBUILD.md SS3's Layer 5 harness shape
' (VLAT_Sentences / VLAT_Phrasebooks), not that shape itself - the
' VLAT_ prefix is R3's manifest-derivation convention, which has no
' build/reload machinery behind it yet. Renaming into that convention
' is future work once R3 actually exists; this module uses the
' project's current VLA_ + descriptive-suffix pattern instead
' (VLA_Interpreter, VLA_HeadTable).
' AS6.0: TestAlonzoLib's FindDevFile("alonzo.vla") call was bare (no
' On Error guard) - a genuinely missing alonzo.vla raised uncaught and
' crashed the whole VlaSelfTest dispatch chain at test 53 of 59,
' silently skipping the six tests after it and the final summary
' line, the same hazard class AS.6 was fixing everywhere else in
' VLA_Tests.bas. Caught only by that pass's own adversarial
' verification, not by inspection. Fixed with the same On Error
' Resume Next / capture Err.Description / On Error GoTo 0 guard used
' there; the pre-existing Dir$ recheck stays as a second, harmless
' condition (FindDevFile never actually returns "" on success, so it
' was already dead code for the missing-file case, but removing it
' was not this fix's job).

' =====================================================================
'  VLA_Tests_Grammar - grammar/phrasebook feature pins: alternation and
'  optional literals (G1), typed ref slots (G2), override: (G3),
'  fail: negative proofs (G4), the EnglishTryRule scratchpad (G5),
'  bare surface tokens (G10), dot-hiding functors (G11), keyword-arg
'  ordering (G12), carried defmacros (L4), VlaTry history (L6/L16),
'  time/trace forms (L7/L17), doc/apropos (L8/L82... see individual
'  headers), template macro-stepper and related Lisp-bench items
'  (L11/L11_1/L12/L13/L14/L15/L18), dispatch/string-builder/cache
'  performance dials (V7/V8), the prelude-as-library series
'  (PL1-PL7), and the evergreen alonzo.vla import regression
'  (TestAlonzoLib). Dev-only, like VLA_Tests.bas: NOT shipped in the
'  add-in build.
'
'  LAYER:     Harness (dev-only; never ships - see REBUILD.md SS3
'             Layer 5)
'  MAY CALL:  VLA (transpiler), VLA_English (grammar engine),
'             VLA_Tests (shared assertion/report infrastructure:
'             Report, CheckFrags, CheckV, Norm, AssertVla,
'             AssertEnglish, TryTranspile, TryEnglish, AssertErrLine,
'             AssertClaim, AssertClaimRefusal, CountOcc, FindDevFile -
'             all promoted Private->Public in this pass for exactly
'             this call)
'  SHIPS:     nowhere. Dev-rig only, like VLA_DevRig and VLA_Tests.
'  PAYS INTO: F.8 (this split), R4 (the 1,200-line budget REBUILD.md
'             names), every Grammar-bucket roadmap item these pins
'             hold the machinery for.
'  REASON:    VLA_Tests.bas reached 4,945 lines - REBUILD.md's own
'             measured trigger for F.8. The grammar/phrasebook feature
'             pins were both the largest concern in the file (~2,500
'             of the 4,945 lines) and the fastest-growing one, so
'             splitting them out first buys the most headroom for the
'             least risk: a pure move, no logic touched, verified by
'             an empty VlaSelfTest delta (same pass/fail count,
'             same names, before and after).
' =====================================================================

' ---------------------------------------------------------------------
'  G1: alternation slots and optional literals in phrase patterns.
'  Engine-level pins here (the matcher, the load-time validation,
'  the expanded-signature audit); the vocabulary's own test: gates
'  prove the refactored and day-1 rules on every load, including the
'  pre-registered acceptance "Put 1 in cell A2." - so the corpus
'  proofs live in english.vla and these pins hold the
'  machinery those proofs stand on.
' ---------------------------------------------------------------------
Public Sub TestG1()
    ' Alternation matches each branch and binds the MATCHED literal
    ' into the template - {d:red|yellow} completing vb{d}.
    EnglishResetGrammar
    EnglishAddPhrase "tint cell {r:text} {d:red|yellow}", "(set! (. (range {r}) interior.color) vb{d})"
    AssertEnglish "g1: alternation matches branch one and binds it", _
                  "Tint cell B2 red.", "vbred"
    AssertEnglish "g1: alternation matches branch two and binds it", _
                  "Tint cell B2 yellow.", "vbyellow"

    ' A word outside the branches refuses, and the near-miss names
    ' the branches.
    Dim d As String
    On Error Resume Next
    Err.Clear
    Dim junk As String
    junk = EnglishToVla("Tint cell B2 green.")
    d = Err.Description
    On Error GoTo 0
    Report "g1: a word outside the alternation refuses naming the branches", _
           InStr(1, d, "'red'", vbTextCompare) > 0 And InStr(1, d, "'yellow'", vbTextCompare) > 0, _
           "got: " & d

    ' Optional literal: present is consumed, absent costs nothing,
    ' and both spellings reach the same template.
    EnglishResetGrammar
    EnglishAddPhrase "poke [new] cell {r:text}", "(debug-print {r})"
    AssertEnglish "g1: optional literal matched when present", _
                  "Poke new cell B2.", "(debug-print ""b2"")"
    AssertEnglish "g1: optional literal free when absent", _
                  "Poke cell B2.", "(debug-print ""b2"")"

    ' The composed idiom: alternation + optional covers a multi-word
    ' spelling pair - at / at key / under / under key, one rule.
    EnglishResetGrammar
    EnglishAddPhrase "stash {e:expr} {p:at|under} [key] {k:expr} in {v:var}", "(vladictset {v} {k} {e})"
    AssertEnglish "g1: composed at-key spelling translates", _
                  "Stash 5 at key 3 in prices.", "(vladictset prices 3 5)"
    AssertEnglish "g1: composed bare-under spelling translates", _
                  "Stash 5 under 3 in prices.", "(vladictset prices 3 5)"

    ' G1 upgraded shape errors to LOAD time: an unknown category and
    ' an empty alternation branch both refuse before any Check.
    EnglishResetGrammar
    Dim v As String
    v = "(english-vla ""warm cell {r:txet}"" (set! (range {r}) 1))"
    On Error Resume Next
    Err.Clear
    EnglishLoadVocabularyText v, "selftest-vocab"
    d = Err.Description
    On Error GoTo 0
    Report "g1: unknown slot category refuses at load", _
           InStr(1, d, "unknown slot category", vbTextCompare) > 0, "got: " & d
    EnglishResetGrammar
    v = "(english-vla ""warm cell {r:text} {d:left|}"" (set! (range {r}) 1))"
    On Error Resume Next
    Err.Clear
    EnglishLoadVocabularyText v, "selftest-vocab"
    d = Err.Description
    On Error GoTo 0
    Report "g1: empty alternation branch refuses at load", _
           InStr(1, d, "empty branch", vbTextCompare) > 0, "got: " & d

    ' The audit expands shapes: an alternation branch colliding with
    ' a plain rule is flagged; an omitted optional colliding with the
    ' bare rule is flagged; a lone collapsed rule stays clean.
    Dim aud As String
    aud = EnglishAuditText( _
        "(english-vla ""fix cell {r:text}"" (debug-print {r}))" & vbLf & _
        "(english-vla ""fix {w:cell|range} {r:text}"" (debug-print {r}))" & vbLf & _
        "(test-success ""Fix cell B2."" (debug-print ""b2""))")
    Report "g1: audit flags an alternation branch shadowed by a plain rule", _
           InStr(1, aud, "duplicates", vbTextCompare) > 0, "report was: " & Left$(aud, 160)
    aud = EnglishAuditText( _
        "(english-vla ""warm cell {r:text}"" (debug-print {r}))" & vbLf & _
        "(english-vla ""warm cell {r:text} [up]"" (debug-print {r}))" & vbLf & _
        "(test-success ""Warm cell B2."" (debug-print ""b2""))")
    Report "g1: audit flags an omitted optional duplicating the bare rule", _
           InStr(1, aud, "duplicates", vbTextCompare) > 0, "report was: " & Left$(aud, 160)
    aud = EnglishAuditText( _
        "(english-vla ""fix {w:cell|range} {r:text}"" (debug-print {r}))" & vbLf & _
        "(test-success ""Fix range A1:B2."" (debug-print ""a1:b2""))")
    Report "g1: audit passes a lone collapsed rule", Len(aud) = 0, "unexpected: " & Left$(aud, 160)

    ' ------------------------------------------------------------------
    ' G1.1 (owner design): stem/suffix surfaces. Either surface
    ' matches; the STEM binds - the invariant that keeps xl{d} a
    ' working constant whichever spelling the writer chose.
    ' ------------------------------------------------------------------
    EnglishResetGrammar
    EnglishAddPhrase "shade cell {r:text} {d:left|center/ed}", "(set! (range {r}) xl{d})"
    AssertEnglish "g1.1: stem surface matches and binds the stem", _
                  "Shade cell B2 center.", "xlcenter"
    AssertEnglish "g1.1: suffixed surface matches and still binds the stem", _
                  "Shade cell B2 centered.", "xlcenter"
    On Error Resume Next
    Err.Clear
    junk = EnglishToVla("Shade cell B2 middled.")
    d = Err.Description
    On Error GoTo 0
    Report "g1.1: the near-miss lists both surfaces as words", _
           InStr(1, d, "'center'", vbTextCompare) > 0 And InStr(1, d, "'centered'", vbTextCompare) > 0, _
           "got: " & d

    ' G1.3 (maiden-run fix of the fix): this block once sat ABOVE the
    ' near-miss pin, and its EnglishResetGrammar wiped the shade rule
    ' that pin depends on - the suite caught the insertion-order
    ' defect ("No loaded sentence starts with 'shade'", 12 = prelude
    ' + settle). A reset inside a section is a fence: everything
    ' depending on earlier registrations must run before it.

    ' G1.2 (maiden-run fix): a SINGLE-branch surface form - no pipe,
    ' one stem, two surfaces - is legal, which is the documented
    ' idiom for a bare word with two spellings. The original
    ' classifier ("alternation = contains |") routed this to the
    ' unknown-category refusal; the failing malformed-slash pin
    ' was the suite catching that, and IsAltCat is the fix.
    EnglishResetGrammar
    EnglishAddPhrase "settle cell {r:text} {d:center/ed}", "(set! (range {r}) xl{d})"
    AssertEnglish "g1.2: a single-branch surface form matches its stem", _
                  "Settle cell B2 center.", "xlcenter"
    AssertEnglish "g1.2: a single-branch surface form matches its suffix and binds the stem", _
                  "Settle cell B2 centered.", "xlcenter"

    EnglishResetGrammar
    EnglishAddPhrase "nudge [gentl/y] cell {r:text}", "(debug-print {r})"
    AssertEnglish "g1.1: suffixed optional consumed on its long surface", _
                  "Nudge gently cell B2.", "(debug-print ""b2"")"
    AssertEnglish "g1.1: suffixed optional consumed on its stem surface", _
                  "Nudge gentl cell B2.", "(debug-print ""b2"")"

    EnglishResetGrammar
    v = "(english-vla ""warm cell {r:text} {d:up/ward/s}"" (debug-print {r}))"
    On Error Resume Next
    Err.Clear
    EnglishLoadVocabularyText v, "selftest-vocab"
    d = Err.Description
    On Error GoTo 0
    Report "g1.1: a second slash in a branch refuses at load", _
           InStr(1, d, "more than one '/'", vbTextCompare) > 0, "got: " & d
    EnglishResetGrammar
    v = "(english-vla ""warm cell {r:text} {d:up|/ed}"" (debug-print {r}))"
    On Error Resume Next
    Err.Clear
    EnglishLoadVocabularyText v, "selftest-vocab"
    d = Err.Description
    On Error GoTo 0
    Report "g1.1: an empty side around '/' refuses at load", _
           InStr(1, d, "empty side", vbTextCompare) > 0, "got: " & d

    ' The audit sees every SURFACE as a shape: a plain rule ending in
    ' the suffixed spelling collides with the stem/suffix branch.
    aud = EnglishAuditText( _
        "(english-vla ""tilt cell {r:text} centered"" (debug-print {r}))" & vbLf & _
        "(english-vla ""tilt cell {r:text} {d:left|center/ed}"" (debug-print xl{d}))" & vbLf & _
        "(test-success ""Tilt cell B2 centered."" (debug-print ""b2""))")
    Report "g1.1: audit flags a suffixed surface shadowed by a plain rule", _
           InStr(1, aud, "duplicates", vbTextCompare) > 0, "report was: " & Left$(aud, 160)
    EnglishResetGrammar
End Sub

' ---------------------------------------------------------------------
'  G2: typed reference slots. Engine pins on the validators, the
'  quoted-assert door, the fall-through semantics, and the U1 touch
'  surface; the migrated vocabulary's own 197 proofs are the
'  positive corpus evidence on every load, and the goldens diffing
'  empty is the whole-corpus witness that migration changed nothing
'  emitted. Check-time judges SHAPE, never existence.
' ---------------------------------------------------------------------
Public Sub TestG2()
    Dim d As String
    Dim junk As String
    EnglishResetGrammar
    EnglishAddPhrase "probe cell {r:cell}", "(debug-print {r})"
    AssertEnglish "g2: cell accepts a bare address", _
                  "Probe cell B22.", "(debug-print ""b22"")"
    AssertEnglish "g2: cell accepts a bang-qualified address", _
                  "Probe cell Data!B2.", "(debug-print ""data!b2"")"
    AssertEnglish "g2: a quoted reference asserts past the shape check", _
                  "Probe cell ""Totals"".", "(debug-print ""Totals"")"
    On Error Resume Next
    Err.Clear
    junk = EnglishToVla("Probe cell banana.")
    d = Err.Description
    On Error GoTo 0
    Report "g2: cell refuses a bare word with the teaching shape", _
           InStr(1, d, "a cell (like", vbTextCompare) > 0 And InStr(1, d, "'banana'", vbTextCompare) > 0, _
           "got: " & d

    EnglishResetGrammar
    EnglishAddPhrase "probe range {r:range}", "(debug-print {r})"
    AssertEnglish "g2: range accepts a span", _
                  "Probe range A1:C50.", "(debug-print ""a1:c50"")"
    AssertEnglish "g2: range accepts a single cell", _
                  "Probe range B2.", "(debug-print ""b2"")"
    On Error Resume Next
    Err.Clear
    junk = EnglishToVla("Probe range banana.")
    d = Err.Description
    On Error GoTo 0
    Report "g2: range refuses a bare word with the teaching shape", _
           InStr(1, d, "a range (like", vbTextCompare) > 0 And InStr(1, d, "'banana'", vbTextCompare) > 0, _
           "got: " & d

    EnglishResetGrammar
    EnglishAddPhrase "trim column {c:column}", "(debug-print {c})"
    AssertEnglish "g2: column accepts letters", _
                  "Trim column AA.", "(debug-print ""aa"")"
    On Error Resume Next
    Err.Clear
    junk = EnglishToVla("Trim column banana.")
    d = Err.Description
    On Error GoTo 0
    Report "g2: column refuses a long word with the teaching shape", _
           InStr(1, d, "column letter", vbTextCompare) > 0 And InStr(1, d, "'banana'", vbTextCompare) > 0, _
           "got: " & d

    EnglishResetGrammar
    EnglishAddPhrase "visit sheet {s:sheet}", "(debug-print {s})"
    AssertEnglish "g2: sheet accepts any word including hyphenated names", _
                  "Visit sheet Nowhere-Land.", "(debug-print ""nowhere-land"")"

    EnglishResetGrammar
    EnglishAddPhrase "paint cell {r:cell} {c:color}", "(debug-print {c})"
    AssertEnglish "g2: color accepts a named color", _
                  "Paint cell B2 red.", "(debug-print ""red"")"
    On Error Resume Next
    Err.Clear
    junk = EnglishToVla("Paint cell B2 banana.")
    d = Err.Description
    On Error GoTo 0
    Report "g2: color refuses a non-color with the teaching shape", _
           InStr(1, d, "a color (like", vbTextCompare) > 0, "got: " & d

    ' A shape failure NoteFails and FALLS THROUGH: a later rule can
    ' still claim the sentence - first-match semantics survive typing.
    EnglishResetGrammar
    EnglishAddPhrase "grade cell {r:cell}", "(debug-print {r})"
    EnglishAddPhrase "grade cell {w:name}", "(debug-print 999)"
    AssertEnglish "g2: a typed refusal falls through to a later claiming rule", _
                  "Grade cell banana.", "(debug-print 999)"

    ' The U1 surface: a rule with typed slots is reported with its
    ' slotname:category pairs.
    EnglishResetGrammar
    EnglishAddPhrase "probe cell {r:cell} of sheet {s:sheet}", "(debug-print {r})"
    Report "g2: EnglishListRefSlots reports the typed touches", _
           InStr(1, EnglishListRefSlots(), "r:cell", vbTextCompare) > 0 And _
           InStr(1, EnglishListRefSlots(), "s:sheet", vbTextCompare) > 0, _
           Left$(EnglishListRefSlots(), 160)
    EnglishResetGrammar
End Sub

' ---------------------------------------------------------------------
'  LE.2: DidYouMean's tie list. NoteFail used to keep a single scalar
'  (earliest-registered rule wins silently); it now grows past that
'  one leader whenever a LATER rule reaches the exact same furthest
'  token position - a genuine tie, not a guess - so a refusal can
'  offer every rule that got equally far, not just whichever happened
'  to be registered first.
' ---------------------------------------------------------------------
Public Sub TestLE2()
    Dim junk As Variant, d As String

    ' Two rules sharing an identical prefix ("grow cell {r:cell}") that
    ' diverge on the very next literal ("by" vs "to") both fail at the
    ' exact same token position for a sentence that gets that far and
    ' no further - both should surface, not just the first-registered.
    EnglishResetGrammar
    EnglishAddPhrase "grow cell {r:cell} by {n:expr}", "(debug-print {r})"
    EnglishAddPhrase "grow cell {r:cell} to {n:expr}", "(debug-print {r})"
    On Error Resume Next
    Err.Clear
    junk = EnglishToVla("Grow cell B2 sideways.")
    d = Err.Description
    On Error GoTo 0
    Report "le.2: a genuine tie surfaces both rules, not just the earlier-registered one", _
           InStr(1, d, "grow cell {r:cell} by {n:expr}", vbTextCompare) > 0 And _
           InStr(1, d, "grow cell {r:cell} to {n:expr}", vbTextCompare) > 0, _
           "got: " & d

    ' A rule under a different first word never becomes a "real"
    ' dispatch candidate for this sentence, so it cannot pollute the
    ' tie - the earlier fix (leading with the sole best-progress rule)
    ' still holds when nothing else reaches the furthest position.
    EnglishResetGrammar
    EnglishAddPhrase "grow cell {r:cell} by {n:expr}", "(debug-print {r})"
    EnglishAddPhrase "shrink cell {r:cell} by {n:expr}", "(debug-print {r})"
    On Error Resume Next
    Err.Clear
    junk = EnglishToVla("Grow cell B2 sideways.")
    d = Err.Description
    On Error GoTo 0
    Report "le.2: a lone best-progress rule still leads with exactly one suggestion", _
           InStr(1, d, "grow cell {r:cell} by {n:expr}", vbTextCompare) > 0 And _
           InStr(1, d, "shrink cell", vbTextCompare) = 0, _
           "got: " & d

    EnglishResetGrammar
End Sub

' ---------------------------------------------------------------------
'  G4: the fail: directive - negative proofs in phrasebooks. Loader-
'  meta pins in the house style ("a failing test: refuses a load"):
'  the gates themselves are proven, and the vocabulary's own five
'  fail: lines are the corpus evidence on every load.
' ---------------------------------------------------------------------
Public Sub TestG4()
    Dim v As String
    Dim d As String
    ' A passing negative proof loads: the sentence refuses and the
    ' fragment rides the refusal.
    EnglishResetGrammar
    v = "(english-vla ""warm cell {r:cell}"" (debug-print {r}))" & vbLf & _
        "(test-success ""Warm cell B2."" (debug-print ""b2""))" & vbLf & _
        "(test-fail ""Warm cell banana."" ""a cell (like"")"
    On Error Resume Next
    Err.Clear
    EnglishLoadVocabularyText v, "selftest-vocab"
    Report "g4: a passing fail: proof loads", Err.Number = 0, Err.Description
    On Error GoTo 0

    ' A sentence that TRANSLATES fails its fail: proof and the load.
    EnglishResetGrammar
    v = "(english-vla ""warm cell {r:cell}"" (debug-print {r}))" & vbLf & _
        "(test-fail ""Warm cell B2."" ""anything"")"
    On Error Resume Next
    Err.Clear
    EnglishLoadVocabularyText v, "selftest-vocab"
    d = Err.Description
    On Error GoTo 0
    Report "g4: a translating sentence refuses the load", _
           InStr(1, d, "translated instead of refusing", vbTextCompare) > 0, "got: " & d

    ' A drifted message fails the proof naming both texts.
    EnglishResetGrammar
    v = "(english-vla ""warm cell {r:cell}"" (debug-print {r}))" & vbLf & _
        "(test-fail ""Warm cell banana."" ""the wrong wording entirely"")"
    On Error Resume Next
    Err.Clear
    EnglishLoadVocabularyText v, "selftest-vocab"
    d = Err.Description
    On Error GoTo 0
    Report "g4: a drifted refusal message refuses the load", _
           InStr(1, d, "drifted", vbTextCompare) > 0 And _
           InStr(1, d, "the wrong wording entirely", vbTextCompare) > 0, "got: " & d

    ' A fragment-less fail: refuses with the teaching. (The old
    ' paren-gather-exemption half of this pin no longer applies - the
    ' new grammar has no gathering to exempt anything from; VLA.bas's
    ' own reader handles nesting natively.)
    EnglishResetGrammar
    v = "(test-fail ""Warm cell banana."" """")"
    On Error Resume Next
    Err.Clear
    EnglishLoadVocabularyText v, "selftest-vocab"
    d = Err.Description
    On Error GoTo 0
    Report "g4: a fragment-less fail: teaches the doctrine", _
           InStr(1, d, "non-empty fragment", vbTextCompare) > 0, "got: " & d
    EnglishResetGrammar
End Sub

' ---------------------------------------------------------------------
'  G3: override semantics. The contract in eight pins: replacement
'  wins the match with provenance recorded, the unmarked shadow
'  refuses the load teaching the marker, the marker must be earned
'  (nothing / several / a built-in all refuse), the audit still
'  reports rather than dying, and cross-file layering carries the
'  file names. The ambiguity vector uses SAME categories on both
'  earlier rules - shapes include the category, so {r:cell} vs
'  {r:range} would never collide (the simulation caught exactly
'  that vector error before the machine could).
' ---------------------------------------------------------------------
Public Sub TestG3()
    Dim d As String
    ' Cross-file layering: the dialect's override wins the sentence,
    ' and its provenance names both files.
    EnglishResetGrammar
    EnglishLoadVocabularyText "(english-vla ""warm cell {r:cell}"" (debug-print 1))" & vbLf & _
        "(test-success ""Warm cell B2."" (debug-print 1))", "base-vocab"
    EnglishLoadVocabularyText "(english-vla-override ""warm cell {r:cell}"" (debug-print 2))" & vbLf & _
        "(test-success ""Warm cell B2."" (debug-print 2))", "dialect-vocab"
    AssertEnglish "g3: the dialect's override wins the sentence", _
                  "Warm cell B2.", "(debug-print 2)"

    ' The unmarked shadow refuses the load and teaches the marker.
    EnglishResetGrammar
    EnglishLoadVocabularyText "(english-vla ""warm cell {r:cell}"" (debug-print 1))", "base-vocab"
    On Error Resume Next
    Err.Clear
    EnglishLoadVocabularyText "(english-vla ""warm cell {r:cell}"" (debug-print 2))", "dialect-vocab"
    d = Err.Description
    On Error GoTo 0
    Report "g3: an unmarked same-shape rule refuses the load teaching the override directive", _
           InStr(1, d, "duplicates", vbTextCompare) > 0 And _
           InStr(1, d, "-vla-override", vbTextCompare) > 0, "got: " & d

    ' The marker must be earned, three ways.
    EnglishResetGrammar
    On Error Resume Next
    Err.Clear
    EnglishLoadVocabularyText "(english-vla-override ""warm cell {r:cell}"" (debug-print 2))", "dialect-vocab"
    d = Err.Description
    On Error GoTo 0
    Report "g3: an override matching nothing refuses", _
           InStr(1, d, "matches no earlier rule", vbTextCompare) > 0, "got: " & d

    EnglishResetGrammar
    EnglishLoadVocabularyText "(english-vla ""fix cell {r:range}"" (debug-print 1))" & vbLf & _
        "(english-vla ""fix range {r:range}"" (debug-print 2))", "base-vocab"
    On Error Resume Next
    Err.Clear
    EnglishLoadVocabularyText "(english-vla-override ""fix {w:cell|range} {r:range}"" (debug-print 3))", "dialect-vocab"
    d = Err.Description
    On Error GoTo 0
    Report "g3: an override matching several refuses as ambiguous", _
           InStr(1, d, "ambiguous", vbTextCompare) > 0, "got: " & d

    EnglishResetGrammar
    On Error Resume Next
    Err.Clear
    EnglishLoadVocabularyText "(english-vla-override ""set {v:var} to {e:expr}"" (debug-print 9))", "dialect-vocab"
    d = Err.Description
    On Error GoTo 0
    Report "g3: the built-in core is not overridable", _
           InStr(1, d, "not overridable", vbTextCompare) > 0, "got: " & d

    ' The audit reports the unmarked shadow instead of dying on it.
    EnglishResetGrammar
    Dim aud As String
    aud = EnglishAuditText( _
        "(english-vla ""warm cell {r:cell}"" (debug-print 1))" & vbLf & _
        "(english-vla ""warm cell {r:cell}"" (debug-print 2))" & vbLf & _
        "(test-success ""Warm cell B2."" (debug-print 1))")
    Report "g3: the audit reports a duplicate a load would refuse", _
           InStr(1, aud, "duplicates", vbTextCompare) > 0, "report was: " & Left$(aud, 160)

    ' An in-file override is legal too (a one-file dialect atop its
    ' own base block), and replacement is in place.
    EnglishResetGrammar
    EnglishLoadVocabularyText "(english-vla ""warm cell {r:cell}"" (debug-print 1))" & vbLf & _
        "(english-vla-override ""warm cell {r:cell}"" (debug-print 5))" & vbLf & _
        "(test-success ""Warm cell B2."" (debug-print 5))", "one-file-vocab"
    AssertEnglish "g3: an in-file override replaces its own base", _
                  "Warm cell B2.", "(debug-print 5)"

    ' Provenance is machine-readable through the sources: the
    ' overridden slot names both files.
    EnglishResetGrammar
    EnglishLoadVocabularyText "(english-vla ""warm cell {r:cell}"" (debug-print 1))", "base-vocab"
    EnglishLoadVocabularyText "(english-vla-override ""warm cell {r:cell}"" (debug-print 2))", "dialect-vocab"
    Dim junk As String
    junk = EnglishToVla("Warm cell B2.")
    Report "g3: provenance names the winner and the overridden file", _
           InStr(1, EnglishRuleSource(), "dialect-vocab (overrides base-vocab)", vbTextCompare) > 0, _
           "got: " & EnglishRuleSource()
    EnglishResetGrammar
End Sub

' ---------------------------------------------------------------------
'  L4: the macro: directive. The composition pass - phrasebooks meet
'  the macro system - so the pins deliberately span layers: the
'  loader's gather and three validations, the prepend, and one
'  full-stack pin driving a sentence through phrasebook -> macro
'  call -> transpile -> expanded VBA. The L5 composition pin is the
'  jewel: a rest-parameter violation inside a macro: form refuses
'  at VOCABULARY LOAD with the transpiler's pinned wording.
' ---------------------------------------------------------------------
Public Sub TestL4()
    Dim d As String
    Dim v As String
    Dim out As String
    ' The full stack: macro carried, rule emits the call, translation
    ' prepends the defmacro, transpile expands to both effects.
    EnglishResetGrammar
    v = "(defmacro (vla-mark r) (begin (set! (range r) 1) (set! (. (range r) font.bold) true)))" & vbLf & _
        "(english-vla ""mark cell {r:cell}"" (vla-mark {r}))" & vbLf & _
        "(test-success ""Mark cell B2."" (vla-mark ""b2""))"
    EnglishLoadVocabularyText v, "selftest-vocab"
    out = EnglishToVla("Mark cell B2.")
    ' L4.2: the carried block rides at the BOTTOM - source-map
    ' neutral, so the .vba golden's vla:N tags never shift when a
    ' phrasebook's macros change (the owner's +4 golden diff).
    Report "l4: the translation carries the defmacro, appended after the program", _
           InStr(1, out, "(defmacro (vla-mark", vbTextCompare) > 0 And _
           InStr(1, out, "(defmacro (vla-mark", vbTextCompare) > InStr(1, out, "(sub", vbTextCompare), _
           Left$(Norm(out), 160)
    Dim vba As String
    vba = TryTranspile("l4: the full stack expands the phrasebook macro to VBA", out)
    If Len(vba) > 0 Then
        CheckFrags "l4: the full stack expands the phrasebook macro to VBA", vba, _
                   Array("font.bold", "= 1")
    End If

    ' L4.1 (maiden-run incident, owner-caught): the resolve check
    ' predated carried macros, so the first phrasebook macro drew the
    ' missing-helper dialog at Reload Sentences - the vla-prefixed
    ' call resolved neither in the manifest nor in pass 1's
    ' sub/function scan. The regression is the incident's exact
    ' shape: the same translation that expands cleanly must also
    ' RESOLVE cleanly.
    Dim ml As Long
    Dim missing As String
    missing = EnglishResolveCheck(out, ml)
    Report "l4.1: a carried macro's call resolves (the Reload incident's shape)", _
           Len(missing) = 0, "reported missing: " & missing
    missing = EnglishResolveCheck( _
        "(defmacro (vla-mark r) (begin (set! (range r) 1)))" & vbCrLf & _
        "(sub main () (vla-mrak ""b2""))", ml)
    Report "l4.1: a genuinely missing helper is still caught beside a macro", _
           missing = "vla-mrak", "reported: " & missing

    ' Multi-line forms parse fine (VLA.bas's own reader is line-
    ' agnostic); two files' macros accumulate in order.
    EnglishResetGrammar
    EnglishLoadVocabularyText "(defmacro (vla-a x)" & vbLf & "    (set! x 1))", "base-vocab"
    EnglishLoadVocabularyText "(defmacro (vla-b x) (set! x 2))", "dialect-vocab"
    out = EnglishToVla("Show 1.")
    Report "l4: multi-line gather works and two files' macros accumulate", _
           InStr(1, out, "vla-a", vbTextCompare) > 0 And InStr(1, out, "vla-b", vbTextCompare) > 0, _
           Left$(Norm(out), 160)

    ' Reset clears carried macros.
    EnglishResetGrammar
    out = EnglishToVla("Show 1.")
    Report "l4: reset clears carried macros", _
           InStr(1, out, "defmacro", vbTextCompare) = 0, Left$(Norm(out), 160)

    ' Same-name macro refuses with directions (override undefined).
    EnglishResetGrammar
    EnglishLoadVocabularyText "(defmacro (vla-a x) (set! x 1))", "base-vocab"
    On Error Resume Next
    Err.Clear
    EnglishLoadVocabularyText "(defmacro (vla-a x) (set! x 2))", "dialect-vocab"
    d = Err.Description
    On Error GoTo 0
    Report "l4: a same-name macro refuses naming the carrying file", _
           InStr(1, d, "already carried by base-vocab", vbTextCompare) > 0, "got: " & d

    ' The L5 composition: a rest-parameter violation refuses at LOAD
    ' with the transpiler's own pinned wording.
    EnglishResetGrammar
    On Error Resume Next
    Err.Clear
    EnglishLoadVocabularyText "(defmacro (vla-c & body :rescue & h) (begin body))", "selftest-vocab"
    d = Err.Description
    On Error GoTo 0
    Report "l4: the engine's rest-parameter refusal fires at vocabulary load", _
           InStr(1, d, "rest parameter must be last", vbTextCompare) > 0, "got: " & d

    ' A non-directive top-level form refuses (F.13: there is no more
    ' "macro: carries defmacro forms only" check to make - the
    ' dispatch itself only ever calls RegisterVocabMacro when the
    ' form's own head symbol already is "defmacro"; anything else
    ' falls to the generic unrecognized-directive refusal).
    EnglishResetGrammar
    On Error Resume Next
    Err.Clear
    EnglishLoadVocabularyText "(set! x 1)", "selftest-vocab"
    d = Err.Description
    On Error GoTo 0
    Report "l4: a non-directive top-level form refuses", _
           InStr(1, d, "unrecognized top-level directive", vbTextCompare) > 0, "got: " & d
    ' An unclosed form refuses (VLA.bas's own reader's job now, not a
    ' hand-rolled paren scanner's - exact wording not pinned here,
    ' only that it refuses).
    EnglishResetGrammar
    On Error Resume Next
    Err.Clear
    EnglishLoadVocabularyText "(defmacro (vla-d x)" & vbLf & "    (set! x 1)", "selftest-vocab"
    Report "l4: an unclosed form refuses", Err.Number <> 0, "load succeeded but should not have"
    On Error GoTo 0
    EnglishResetGrammar
End Sub

' ---------------------------------------------------------------------
'  VOCABDIFF: structural diff between two vocabulary sources. State-
'  independent (no EnglishResetGrammar needed around these calls - the
'  diff never touches loaded grammar), so this can run beside any
'  other section without a fence.
Public Sub TestVocabDiff()
    Dim a As String, b As String, out As String

    ' Identical rules diff clean even when their literal whitespace
    ' differs - the comparison is VLA.VlaWriteForm's canonical text,
    ' never the source bytes.
    a = "(english-vla ""warm cell {r:cell}"" (debug-print {r}))" & vbLf & _
        "(test-success ""Warm cell B2."" (debug-print ""b2""))"
    b = "(english-vla   ""warm cell {r:cell}""   (debug-print {r}))" & vbLf & _
        "(test-success ""Warm cell B2."" (debug-print ""b2""))"
    out = EnglishDiffVocabularyText(a, b, "a", "b")
    Report "vocabdiff: identical rules diff clean despite whitespace", _
           InStr(1, out, "0 changed, 0 added, 0 removed", vbTextCompare) > 0, out

    ' Reordering alone reports no changes - proves this is a keyed
    ' diff by rule identity, not a line-by-line text diff.
    a = "(english-vla ""warm cell {r:cell}"" (debug-print {r}))" & vbLf & _
        "(english-vla ""cool cell {r:cell}"" (debug-print {r}))"
    b = "(english-vla ""cool cell {r:cell}"" (debug-print {r}))" & vbLf & _
        "(english-vla ""warm cell {r:cell}"" (debug-print {r}))"
    out = EnglishDiffVocabularyText(a, b, "a", "b")
    Report "vocabdiff: reordering alone reports no changes", _
           InStr(1, out, "0 changed, 0 added, 0 removed", vbTextCompare) > 0, out

    ' A changed template reports CHANGED, keyed by the pattern text.
    a = "(english-vla ""warm cell {r:cell}"" (debug-print {r}))"
    b = "(english-vla ""warm cell {r:cell}"" (msgbox {r}))"
    out = EnglishDiffVocabularyText(a, b, "a", "b")
    Report "vocabdiff: a changed template reports CHANGED", _
           InStr(1, out, "CHANGED rule ""warm cell {r:cell}""", vbTextCompare) > 0 And _
           InStr(1, out, "1 changed", vbTextCompare) > 0, out

    ' A rule only in B reports ADDED; the same pair reversed reports
    ' REMOVED - same machinery, opposite direction.
    a = "(english-vla ""warm cell {r:cell}"" (debug-print {r}))"
    b = "(english-vla ""warm cell {r:cell}"" (debug-print {r}))" & vbLf & _
        "(english-vla ""cool cell {r:cell}"" (debug-print {r}))"
    out = EnglishDiffVocabularyText(a, b, "a", "b")
    Report "vocabdiff: a rule only in B reports ADDED", _
           InStr(1, out, "ADDED rule", vbTextCompare) > 0 And _
           InStr(1, out, "cool cell", vbTextCompare) > 0 And _
           InStr(1, out, "1 added", vbTextCompare) > 0, out
    out = EnglishDiffVocabularyText(b, a, "b", "a")
    Report "vocabdiff: the same pair reversed reports REMOVED", _
           InStr(1, out, "REMOVED rule", vbTextCompare) > 0 And _
           InStr(1, out, "cool cell", vbTextCompare) > 0 And _
           InStr(1, out, "1 removed", vbTextCompare) > 0, out

    ' A macro's own name is its key, independent of any rule that
    ' calls it - a body-only change reports the macro changed and
    ' leaves the untouched rule out of the report entirely.
    a = "(defmacro (vla-mark r) (set! (range r) 1))" & vbLf & _
        "(english-vla ""mark cell {r:cell}"" (vla-mark {r}))"
    b = "(defmacro (vla-mark r) (set! (range r) 2))" & vbLf & _
        "(english-vla ""mark cell {r:cell}"" (vla-mark {r}))"
    out = EnglishDiffVocabularyText(a, b, "a", "b")
    Report "vocabdiff: a changed macro body reports CHANGED macro, rule untouched", _
           InStr(1, out, "CHANGED macro ""vla-mark""", vbTextCompare) > 0 And _
           InStr(1, out, "CHANGED rule", vbTextCompare) = 0, out
End Sub

' ---------------------------------------------------------------------
'  METAVOCAB: a generator macro's expansion, spliced back through the
'  loader's own dispatch, is indistinguishable from rules a human
'  typed by hand - including running through the same test-success
'  proof machinery. State-heavy (each case needs its own carried
'  macro), so every case resets first.
' ---------------------------------------------------------------------
Public Sub TestMetaVocab()
    Dim n As Long, d As String

    ' A generator expands into TWO english-vla rules plus their own
    ' test-success proofs, all in one load - if either generated rule
    ' didn't actually work, the load itself would raise (standing
    ' discipline: a vocabulary refuses to load on a failing proof).
    EnglishResetGrammar
    On Error Resume Next
    Err.Clear
    n = EnglishLoadVocabularyText( _
        "(defmacro (vocab-antonym-pair pat1 call1 pat2 call2)" & vbLf & _
        "    (begin (english-vla pat1 call1) (english-vla pat2 call2)))" & vbLf & _
        "(vocab-antonym-pair" & vbLf & _
        "    ""warm cell {r:cell}"" (debug-print {r})" & vbLf & _
        "    ""cool cell {r:cell}"" (debug-print {r}))" & vbLf & _
        "(test-success ""Warm cell B2."" (debug-print ""b2""))" & vbLf & _
        "(test-success ""Cool cell B3."" (debug-print ""b3""))", "metavocab-vocab")
    Report "metavocab: a generator's two rules load and pass their own tests", _
           Err.Number = 0 And n = 2, "err: " & Err.Description & " / count: " & n
    On Error GoTo 0

    ' Both generated rules generalize beyond their own literal test
    ' sentence - proof they work, not merely that they parsed.
    Report "metavocab: the first generated rule translates a NEW sentence", _
           InStr(1, EnglishToVla("Warm cell B4."), "debug-print", vbTextCompare) > 0, _
           Left$(EnglishToVla("Warm cell B4."), 160)
    Report "metavocab: the second generated rule translates a NEW sentence", _
           InStr(1, EnglishToVla("Cool cell B5."), "debug-print", vbTextCompare) > 0, _
           Left$(EnglishToVla("Cool cell B5."), 160)

    ' A genuine typo, unrelated to any carried macro, still refuses
    ' with the EXACT original wording - this pass changes nothing
    ' about what a plain mistake looks like.
    EnglishResetGrammar
    On Error Resume Next
    Err.Clear
    EnglishLoadVocabularyText "(zzz-not-a-thing ""x"")", "metavocab-vocab"
    d = Err.Description
    On Error GoTo 0
    Report "metavocab: a genuine typo still refuses with the original wording", _
           InStr(1, d, "unrecognized top-level directive 'zzz-not-a-thing'", vbTextCompare) > 0, "got: " & d

    ' A macro that DOES match but expands to something that is not a
    ' vocabulary directive refuses in words, naming the shape it got.
    EnglishResetGrammar
    On Error Resume Next
    Err.Clear
    EnglishLoadVocabularyText "(defmacro (oops x) (+ x 1))" & vbLf & "(oops 5)", "metavocab-vocab"
    d = Err.Description
    On Error GoTo 0
    Report "metavocab: an expansion that isn't a directive refuses, naming the shape", _
           InStr(1, d, "is not a vocabulary directive", vbTextCompare) > 0, "got: " & d

    ' A macro that DOES match but is called with the wrong arity names
    ' the real problem (the macro engine's own message) at the real
    ' file line (3 - the blank line between defmacro and the call is
    ' deliberate), not a synthetic line from the expansion's own blob.
    EnglishResetGrammar
    On Error Resume Next
    Err.Clear
    EnglishLoadVocabularyText _
        "(defmacro (needs-two a b) (english-vla a b))" & vbLf & vbLf & "(needs-two ""only one"")", "metavocab-vocab"
    d = Err.Description
    On Error GoTo 0
    Report "metavocab: a matched macro called with the wrong arity names the real problem and the real line", _
           InStr(1, d, "line 3", vbTextCompare) > 0 And _
           InStr(1, d, "needs-two", vbTextCompare) > 0 And _
           InStr(1, d, "expects 2 argument", vbTextCompare) > 0, "got: " & d

    ' One generator calling another - VlaExpandText's own fixpoint
    ' expansion flattens the inner call automatically, so this mostly
    ' exercises composability, not the splice recursion itself.
    EnglishResetGrammar
    On Error Resume Next
    Err.Clear
    n = EnglishLoadVocabularyText( _
        "(defmacro (one-rule pat call) (english-vla pat call))" & vbLf & _
        "(defmacro (two-rules pat1 call1 pat2 call2)" & vbLf & _
        "    (begin (one-rule pat1 call1) (one-rule pat2 call2)))" & vbLf & _
        "(two-rules" & vbLf & _
        "    ""glow cell {r:cell}"" (debug-print {r})" & vbLf & _
        "    ""fade cell {r:cell}"" (debug-print {r}))" & vbLf & _
        "(test-success ""Glow cell B2."" (debug-print ""b2""))" & vbLf & _
        "(test-success ""Fade cell B3."" (debug-print ""b3""))", "metavocab-vocab")
    Report "metavocab: one generator calling another still splices correctly", _
           Err.Number = 0 And n = 2, "err: " & Err.Description & " / count: " & n
    On Error GoTo 0

    ' A generator whose OWN body contains a literal nested (begin
    ' (begin ...) ...) - the one shape that actually exercises the
    ' splice recursion's depth parameter rather than relying on
    ' VlaExpandText's fixpoint expansion to flatten things first.
    EnglishResetGrammar
    On Error Resume Next
    Err.Clear
    n = EnglishLoadVocabularyText( _
        "(defmacro (nested-begin-rule)" & vbLf & _
        "    (begin (begin (english-vla ""spark cell {r:cell}"" (debug-print {r})))" & vbLf & _
        "           (test-success ""Spark cell B2."" (debug-print ""b2""))))" & vbLf & _
        "(nested-begin-rule)", "metavocab-vocab")
    Report "metavocab: a generator body with a literal nested (begin ...) still splices", _
           Err.Number = 0 And n = 1, "err: " & Err.Description & " / count: " & n
    On Error GoTo 0

    EnglishResetGrammar
End Sub

' ---------------------------------------------------------------------
'  P-PROBE: RegisterVocabMacro's own probe (VLA_English.bas) used to
'  re-transpile the ENTIRE accumulated mVocabMacros text plus the new
'  macro's own text on EVERY registration - O(macros^2) in the number
'  of vocab macros carried, confirmed the dominant cost of a real
'  corpus load (BETA_ROADMAP1.md's own P-PROBE entry; scripts/
'  instructions_golden.vla's own 134 defmacro forms is the real corpus
'  this surfaced against). Fixed to probe only the NEW macro's own
'  text: correct per DefineMacro's own contract (VLA.bas) - Pass 1
'  collects every top-level defmacro into a FRESH Collection on every
'  VlaTranspile call regardless of what else is fed in, and allows
'  silent redefinition (`mMacros.Remove` then `Add`, no collision
'  check), so no macro already carried is ever cross-checked against
'  another inside the probe; name collisions are already caught
'  earlier in RegisterVocabMacro (mVocabMacroNames), before the probe
'  ever runs. Two things pinned here, neither covered before this
'  pass: (1) a malformed macro registered THIRD, after two good ones,
'  still fails its own probe - the position-independence a
'  first-macro-only or last-macro-only fix could get wrong; (2) at
'  scale (mRunScaleTests-gated, matching real-corpus order of
'  magnitude), 150 valid macros register cleanly in one load, and a
'  malformed 151st macro loaded AFTER them - the corpus at its
'  largest, the exact case the old full-corpus probe fed itself and
'  this fix no longer does - still fails correctly. The printed
'  elapsed time is informational only (Timer output is not a pinnable
'  value, L7's own recorded reason above) - VlaSelfTestScale is where
'  a human reads the real number, same discipline as VlaTimeIt's own
'  dial (VLA_DevRig.bas).
' ---------------------------------------------------------------------
Public Sub TestVocabMacroProbe()
    Dim d As String

    ' (1) fast, always-on: a malformed macro registered third still
    ' fails on its own malformed form, not silently absorbed because
    ' two valid macros already came before it.
    EnglishResetGrammar
    On Error Resume Next
    Err.Clear
    EnglishLoadVocabularyText _
        "(defmacro (probe-ok-a x) (set! x 1))" & vbLf & _
        "(defmacro (probe-ok-b x) (set! x 2))" & vbLf & _
        "(defmacro (quote x) x)", "probe-vocab"
    d = Err.Description
    On Error GoTo 0
    Report "p-probe: a malformed macro registered THIRD still fails its own probe", _
           InStr(1, d, "does not stand", vbTextCompare) > 0 And _
           InStr(1, d, "quote", vbTextCompare) > 0, "got: " & d
    EnglishResetGrammar

    ' (2) scale: 150 macros (scripts/instructions_golden.vla's own real
    ' order of magnitude, 134 defmacros) register in one load, gated on
    ' mRunScaleTests the same way TestListopsDepthSafety/
    ' TestTablespecDepthSafety gate their own large cases (VLA_Tests.
    ' bas) - the point is to make the OLD O(macros^2) cost visible at
    ' real scale, not to run on every ordinary VlaSelfTest pass.
    If VLA_Tests.mRunScaleTests Then
        EnglishResetGrammar
        Dim parts() As String, i As Long
        ReDim parts(1 To 150)
        For i = 1 To 150
            parts(i) = "(defmacro (probe-scale-" & i & " x) (set! x " & i & "))"
        Next
        Dim t0 As Double, msElapsed As Double
        Dim loadErr As String
        t0 = Timer
        On Error Resume Next
        Err.Clear
        EnglishLoadVocabularyText Join(parts, vbLf), "probe-scale-vocab"
        loadErr = Err.Description
        On Error GoTo 0
        msElapsed = (Timer - t0) * 1000#
        Debug.Print "  p-probe scale: 150 macros registered in " & Format$(msElapsed, "0") & " ms"
        Report "p-probe: 150 macros register cleanly in one load", _
               loadErr = "", "err: " & loadErr

        ' the 151st macro, loaded AFTER 150 already carried - the
        ' corpus at its largest here - still fails correctly when it
        ' is malformed. This is the exact case the old full-corpus
        ' probe used to feed itself and this fix no longer does.
        On Error Resume Next
        Err.Clear
        EnglishLoadVocabularyText "(defmacro (quote x) x)", "probe-scale-vocab"
        d = Err.Description
        On Error GoTo 0
        Report "p-probe: a malformed 151st macro still fails its own probe with 150 already carried", _
               InStr(1, d, "does not stand", vbTextCompare) > 0 And _
               InStr(1, d, "quote", vbTextCompare) > 0, "got: " & d

        EnglishResetGrammar
    End If
End Sub

' ---------------------------------------------------------------------
'  TABLE-FAMILY: table-property-family (a real, scripts/english.vla-
'  shaped generator, not a synthetic stand-in like TestMetaVocab's own
'  cases) generates a (defmacro (name n extras) doc (set! (. (active-
'  sheet.listobjects n) property) value)) / (english-vla ...) /
'  (test-success ...) triple. Pinned here independent of the real
'  corpus (a self-contained synthetic vocab, distinct macro names) so
'  this proves the GENERATOR mechanism, not just that english.vla
'  happens to load - EnglishToVba (EnglishToVla + a real VlaTranspile)
'  proves generated macros compile to real VBA, not just that their
'  outer call shape parses. The two shapes this exercises: a generated
'  macro with a real second REQUIRED param threaded in via the
'  generator's own `& extras` rest splice (vocab-table-style, one
'  extra name, "s"), and one with none at all (vocab-table-totals-on/
'  -off, zero extras - the same generator, not a second one).
' ---------------------------------------------------------------------
Public Sub TestTableFamily()
    Dim n As Long

    EnglishResetGrammar
    On Error Resume Next
    Err.Clear
    n = EnglishLoadVocabularyText( _
        "(defmacro" & vbLf & _
        "    (table-property-family name property doc pattern call testsentence testcall value & extras)" & vbLf & _
        "    ""generate a table-name/value-setting macro, its english-vla rule, and its test-success proof""" & vbLf & _
        "    (begin" & vbLf & _
        "        (defmacro (name n extras) doc (set! (. (activesheet.listobjects n) property) value))" & vbLf & _
        "        (english-vla pattern call)" & vbLf & _
        "        (test-success testsentence testcall)))" & vbLf & _
        "(table-property-family vocab-table-style vocabstyle" & vbLf & _
        "    ""set a demo table's style""" & vbLf & _
        "    ""set demo style of table {n:text} to {s:text}"" (vocab-table-style {n} {s})" & vbLf & _
        "    ""Set demo style of table Sales to Foo."" (vocab-table-style ""sales"" ""foo"")" & vbLf & _
        "    s s)" & vbLf & _
        "(table-property-family vocab-table-totals-on vocabtotals" & vbLf & _
        "    ""show a demo table's totals""" & vbLf & _
        "    ""show demo totals of table {n:text}"" (vocab-table-totals-on {n})" & vbLf & _
        "    ""Show demo totals of table Sales."" (vocab-table-totals-on ""sales"")" & vbLf & _
        "    true)" & vbLf & _
        "(table-property-family vocab-table-totals-off vocabtotals" & vbLf & _
        "    ""hide a demo table's totals""" & vbLf & _
        "    ""hide demo totals of table {n:text}"" (vocab-table-totals-off {n})" & vbLf & _
        "    ""Hide demo totals of table Sales."" (vocab-table-totals-off ""sales"")" & vbLf & _
        "    false)", "tablefamily-vocab")
    Report "table-family: one generator's three calls (two arities) load and pass their own tests", _
           Err.Number = 0 And n = 3, "err: " & Err.Description & " / count: " & n
    On Error GoTo 0

    ' Generalizes beyond the literal test sentence - proof the rule
    ' works, not merely that it parsed (same discipline as METAVOCAB's
    ' own generalization pins).
    Report "table-family: the property-setting macro generalizes to a new table/style", _
           InStr(1, EnglishToVba("Set demo style of table Regions to Bar."), "vocabstyle", vbTextCompare) > 0, _
           Left$(EnglishToVba("Set demo style of table Regions to Bar."), 200)

    ' The two boolean-toggle macros share one property name (vocabtotals)
    ' but bake DIFFERENT literals into their own generated macro body -
    ' proof that `value` is bound at the GENERATOR's own call, never
    ' exposed as a sentence slot, and that zero `extras` really does
    ' vanish cleanly (both loaded with a plain single-param signature,
    ' or the earlier Report would have failed the load already).
    Report "table-family: the totals-on macro's baked-in TRUE reaches the emitted VBA", _
           InStr(1, EnglishToVba("Show demo totals of table Sales."), "True", vbTextCompare) > 0, _
           Left$(EnglishToVba("Show demo totals of table Sales."), 200)
    Report "table-family: the totals-off macro's baked-in FALSE reaches the emitted VBA", _
           InStr(1, EnglishToVba("Hide demo totals of table Sales."), "False", vbTextCompare) > 0, _
           Left$(EnglishToVba("Hide demo totals of table Sales."), 200)

    EnglishResetGrammar
End Sub

' ---------------------------------------------------------------------
'  ANTONYM-SWEEP: the five hand-copied verb-antonym macro pairs
'  (hide|unhide column, wrap|unwrap text, merge|unmerge range,
'  protect|unprotect sheet, hide|unhide row) refactored through two
'  real generators - bool-antonym-family (a boolean-property flip: 3 of
'  5) and method-antonym-family (two distinct method names, no boolean
'  at all: 1 of 5); protect-sheet/unprotect-sheet's keyword-argument-
'  plus-extra-parameter shape stays hand-written on purpose, same call
'  table-to-range's own precedent already made. Pinned here with
'  synthetic demo names (distinct from the real hide-column/merge-range
'  etc., same convention TestTableFamily's own vocab-table-* names
'  already use), same mechanism TestTableFamily itself uses -
'  EnglishToVba (EnglishToVla + a real VlaTranspile), not a bare
'  VlaTranspile: a generated macro lives in the vocabulary's own
'  accumulated macro text, a separate registry from a standalone
'  transpile's, so only the assembled-program path a real English
'  sentence takes can see it (caught before this shipped - a first
'  draft called VlaTranspile directly on a bare snippet, which cannot
'  possibly resolve a vocabulary-registered macro name at all).
'  Deliberately chosen demo names (hide-thing/lock-thing) so a plain
'  english-vla rule using G11r's own existing {d:hide|unhide}-style
'  glued-identifier dispatch reaches the generated pair exactly the way
'  the real corpus's rules already do - this test therefore also
'  proves the two mechanisms compose, not just that each works alone.
' ---------------------------------------------------------------------
Public Sub TestAntonymSweep()
    Dim n As Long

    ' bool-antonym-family: target given as a whole wrapped-call form
    ' ((things t)) - hide-column/hide-row's own shape, not wrap-text's
    ' bare-param shape (already exercised live by wrap-text itself).
    EnglishResetGrammar
    On Error Resume Next
    Err.Clear
    n = EnglishLoadVocabularyText( _
        "(defmacro" & vbLf & _
        "    (bool-antonym-family name param target property doc-pos doc-neg)" & vbLf & _
        "    ""generate a boolean-property on/off macro pair sharing one target-and-property shape""" & vbLf & _
        "    (begin" & vbLf & _
        "        (defmacro (name param) doc-pos (set! (. target property) true))" & vbLf & _
        "        (defmacro ((symbol ""un"" name) param) doc-neg (set! (. target property) false))))" & vbLf & _
        "(bool-antonym-family hide-thing t (things t) hidden" & vbLf & _
        "    ""hide a demo thing""" & vbLf & _
        "    ""show a hidden demo thing"")" & vbLf & _
        "(english-vla ""{d:hide|unhide} demo thing {t:text}"" ({d}-thing {t}))" & vbLf & _
        "(test-success ""Hide demo thing Foo."" (hide-thing ""foo""))" & vbLf & _
        "(test-success ""Unhide demo thing Foo."" (unhide-thing ""foo""))", "antonymsweep-vocab-bool")
    Report "antonym-sweep: bool-antonym-family generates both directions, both pass their own tests", _
           Err.Number = 0 And n = 1, "err: " & Err.Description & " / count: " & n
    On Error GoTo 0

    Report "antonym-sweep: bool-antonym-family's positive macro bakes in TRUE, targets the wrapped form", _
           InStr(1, EnglishToVba("Hide demo thing Foo."), "hidden = True", vbTextCompare) > 0, _
           Left$(EnglishToVba("Hide demo thing Foo."), 200)

    ' (symbol "un" name) derives the negative macro's own name - no
    ' second name spelled out anywhere in the generator's own call.
    Report "antonym-sweep: bool-antonym-family's (symbol ""un"" name)-derived negative macro bakes in FALSE", _
           InStr(1, EnglishToVba("Unhide demo thing Foo."), "hidden = False", vbTextCompare) > 0, _
           Left$(EnglishToVba("Unhide demo thing Foo."), 200)

    ' method-antonym-family: two distinct method names, no boolean or
    ' property at all - merge-range's own shape.
    EnglishResetGrammar
    On Error Resume Next
    Err.Clear
    n = EnglishLoadVocabularyText( _
        "(defmacro" & vbLf & _
        "    (method-antonym-family name param target pos-method neg-method doc-pos doc-neg)" & vbLf & _
        "    ""generate a method-call antonym pair sharing one target-and-method shape, no extra arguments""" & vbLf & _
        "    (begin" & vbLf & _
        "        (defmacro (name param) doc-pos (. target pos-method))" & vbLf & _
        "        (defmacro ((symbol ""un"" name) param) doc-neg (. target neg-method))))" & vbLf & _
        "(method-antonym-family lock-thing r r lock unlock" & vbLf & _
        "    ""lock a demo thing""" & vbLf & _
        "    ""unlock a demo thing"")" & vbLf & _
        "(english-vla ""{d:lock|unlock} demo thing {r:text}"" ({d}-thing {r}))" & vbLf & _
        "(test-success ""Lock demo thing Foo."" (lock-thing ""foo""))" & vbLf & _
        "(test-success ""Unlock demo thing Foo."" (unlock-thing ""foo""))", "antonymsweep-vocab-method")
    Report "antonym-sweep: method-antonym-family generates both directions, both pass their own tests", _
           Err.Number = 0 And n = 1, "err: " & Err.Description & " / count: " & n
    On Error GoTo 0

    Report "antonym-sweep: method-antonym-family's positive macro calls the positive method", _
           InStr(1, EnglishToVba("Lock demo thing Foo."), "lock", vbTextCompare) > 0, _
           Left$(EnglishToVba("Lock demo thing Foo."), 200)

    Report "antonym-sweep: method-antonym-family's negative macro calls the negative method, same derived-name convention", _
           InStr(1, EnglishToVba("Unlock demo thing Foo."), "unlock", vbTextCompare) > 0, _
           Left$(EnglishToVba("Unlock demo thing Foo."), 200)

    EnglishResetGrammar
End Sub

' ---------------------------------------------------------------------
'  TABLESPEC: tablespec-row/tablespec (scripts/english.vla, real
'  generators, not synthetic stand-ins) - one call over a walked
'  3-row spec, in place of three hand-typed table-property-family-style
'  calls. Redefined inline here with synthetic demo names (distinct
'  from any real corpus macro), same convention TestTableFamily/
'  TestAntonymSweep already use, so this proves the GENERATOR mechanism
'  composing with the real vocabulary loader end to end, not just that
'  english.vla happens to load. TestTablespecDepthSafety (VLA_Tests.bas)
'  separately proves the walk's own depth-safety in isolation via bare
'  VlaExpandText - this test's job is proving the WALK correctly
'  destructures each row and reaches the real English/test-success/
'  runtime path, not depth safety.
' ---------------------------------------------------------------------
Public Sub TestTablespec()
    Dim n As Long

    ' Built in three pieces, not one - a single EnglishLoadVocabularyText
    ' call with this much vocab text as one statement hit VBA's own hard
    ' "too many line continuations" ceiling (24 consecutive `_`
    ' continuations per logical line) on the first live import: caught
    ' live, not by tracing, since nothing in this codebase's own house
    ' style flags line-continuation COUNT as a risk the way it does
    ' depth/budget counters. Each piece below stays well under the
    ' ceiling; excelatfinance.com's own writeup on the limit names the
    ' identical fix - build the string in a number of steps.
    Dim vocabRowMacro As String
    vocabRowMacro = _
        "(defmacro" & vbLf & _
        "    (tablespec-row name doc body pattern call testsentence testcall)" & vbLf & _
        "    ""emit one single-parameter defmacro/english-vla/test-success triple""" & vbLf & _
        "    (begin" & vbLf & _
        "        (defmacro (name n) doc body)" & vbLf & _
        "        (english-vla pattern call)" & vbLf & _
        "        (test-success testsentence testcall)))" & vbLf

    Dim vocabWalkMacro As String
    vocabWalkMacro = _
        "(defmacro" & vbLf & _
        "    (tablespec spec)" & vbLf & _
        "    ""walk a quoted list of rows, calling tablespec-row once per row""" & vbLf & _
        "    (quote-if (null? spec)" & vbLf & _
        "        (begin)" & vbLf & _
        "        (begin (tablespec-row (car (car spec))" & vbLf & _
        "                              (cadr (car spec))" & vbLf & _
        "                              (car (cddr (car spec)))" & vbLf & _
        "                              (cadr (cddr (car spec)))" & vbLf & _
        "                              (caddr (cddr (car spec)))" & vbLf & _
        "                              (cadr (cddr (cddr (car spec))))" & vbLf & _
        "                              (caddr (cddr (cddr (car spec)))))" & vbLf & _
        "               (tablespec (cdr spec)))))" & vbLf

    Dim vocabRows As String
    vocabRows = _
        "(tablespec (quote (" & vbLf & _
        "    (vocab-ts-alpha ""set demo alpha"" (set! (. (range n) value) 1)" & vbLf & _
        "        ""set demo alpha of {n:text}"" (vocab-ts-alpha {n})" & vbLf & _
        "        ""Set demo alpha of Foo."" (vocab-ts-alpha ""foo""))" & vbLf & _
        "    (vocab-ts-beta ""set demo beta"" (set! (. (range n) value) 2)" & vbLf & _
        "        ""set demo beta of {n:text}"" (vocab-ts-beta {n})" & vbLf & _
        "        ""Set demo beta of Foo."" (vocab-ts-beta ""foo""))" & vbLf & _
        "    (vocab-ts-gamma ""set demo gamma"" (set! (. (range n) value) 3)" & vbLf & _
        "        ""set demo gamma of {n:text}"" (vocab-ts-gamma {n})" & vbLf & _
        "        ""Set demo gamma of Foo."" (vocab-ts-gamma ""foo""))" & vbLf & _
        ")))"

    EnglishResetGrammar
    On Error Resume Next
    Err.Clear
    n = EnglishLoadVocabularyText(vocabRowMacro & vocabWalkMacro & vocabRows, "tablespec-vocab")
    Report "tablespec: one call over a 3-row spec loads and passes all three generated tests", _
           Err.Number = 0 And n = 3, "err: " & Err.Description & " / count: " & n
    On Error GoTo 0

    Report "tablespec: row 1 generates a macro that reaches real VBA", _
           InStr(1, EnglishToVba("Set demo alpha of Foo."), "= 1", vbTextCompare) > 0, _
           Left$(EnglishToVba("Set demo alpha of Foo."), 200)

    Report "tablespec: row 2's own value is distinct from row 1's - not all rows collapsing to the same body", _
           InStr(1, EnglishToVba("Set demo beta of Foo."), "= 2", vbTextCompare) > 0, _
           Left$(EnglishToVba("Set demo beta of Foo."), 200)

    Report "tablespec: row 3, the LAST row, also reaches real VBA - proof the walk completes to the end, not just the first row", _
           InStr(1, EnglishToVba("Set demo gamma of Foo."), "= 3", vbTextCompare) > 0, _
           Left$(EnglishToVba("Set demo gamma of Foo."), 200)

    EnglishResetGrammar
End Sub

' ---------------------------------------------------------------------
'  LISTOPS-PROVENANCE: (at-row label form), a transparent wrapper
'  recognized by DispatchVocabForm alongside "begin". A rule/test/macro
'  a generator produces from row N of a table must be traceable to row
'  N, not just to the generator's own call site - today every spliced
'  form inherits the call site's line only, fine for a five-line
'  antonym pair, not fine for fifty rules from one call. Not
'  LISTOPS-only: usable today, by hand, in any large hand-written
'  block - and the generator-emitted case (last pin below) proves the
'  mechanism a future LISTOPS table-walker would actually need, without
'  needing LISTOPS itself to exist yet.
' ---------------------------------------------------------------------
Public Sub TestAtRow()
    Dim d As String, n As Long

    ' at-row is transparent to a correctly-loading rule/proof - no
    ' regression on the ordinary path.
    EnglishResetGrammar
    On Error Resume Next
    Err.Clear
    n = EnglishLoadVocabularyText( _
        "(english-vla ""warm cell {r:cell}"" (debug-print {r}))" & vbLf & _
        "(at-row ""row A"" (test-success ""Warm cell B2."" (debug-print ""b2"")))", "atrow-vocab")
    Report "at-row: a correctly-tagged proof loads clean", _
           Err.Number = 0 And n = 1, "err: " & Err.Description & " / count: " & n
    On Error GoTo 0

    ' Two at-row blocks, one wrong - the failure names the WRONG one's
    ' label, not the correct sibling's.
    EnglishResetGrammar
    On Error Resume Next
    Err.Clear
    EnglishLoadVocabularyText _
        "(english-vla ""warm cell {r:cell}"" (debug-print {r}))" & vbLf & _
        "(at-row ""row A"" (test-success ""Warm cell B2."" (debug-print ""b2"")))" & vbLf & _
        "(at-row ""row B"" (test-success ""Warm cell B3."" (debug-print ""WRONG"")))", "atrow-vocab"
    d = Err.Description
    On Error GoTo 0
    Report "at-row: a failing proof names the wrong sibling's row, not the correct one's", _
           InStr(1, d, "row B", vbTextCompare) > 0 And InStr(1, d, "row A", vbTextCompare) = 0, "got: " & d

    ' ExpandVocabMacroCall's own arity-mismatch path names the row too -
    ' not just the test-success path above.
    EnglishResetGrammar
    On Error Resume Next
    Err.Clear
    EnglishLoadVocabularyText _
        "(defmacro (needs-two a b) (english-vla a b))" & vbLf & vbLf & _
        "(at-row ""row 9"" (needs-two ""only one""))", "atrow-vocab"
    d = Err.Description
    On Error GoTo 0
    Report "at-row: an arity-mismatched macro call names its own row", _
           InStr(1, d, "row 9", vbTextCompare) > 0 And InStr(1, d, "needs-two", vbTextCompare) > 0, "got: " & d

    ' RegisterVocabMacro's own probe-failure path names the row too -
    ' the reserved-name refusal, distinct from an arity mismatch.
    EnglishResetGrammar
    On Error Resume Next
    Err.Clear
    EnglishLoadVocabularyText "(at-row ""row 3"" (defmacro (quote x) x))", "atrow-vocab"
    d = Err.Description
    On Error GoTo 0
    Report "at-row: a reserved macro-name refusal names its own row", _
           InStr(1, d, "row 3", vbTextCompare) > 0, "got: " & d

    ' The actual future case: a GENERATOR's own expansion emits several
    ' at-row-wrapped children (nested begin inside each row, exactly
    ' the shape a LISTOPS table-walker would produce) - rowTag survives
    ' begin's splice, at-row's own override, and a SECOND nested begin,
    ' all in one dispatch, and still names the right row.
    EnglishResetGrammar
    On Error Resume Next
    Err.Clear
    EnglishLoadVocabularyText _
        "(defmacro (two-tagged-rules)" & vbLf & _
        "    (begin" & vbLf & _
        "        (at-row ""row 1"" (begin (english-vla ""glow cell {r:cell}"" (debug-print {r}))" & vbLf & _
        "                              (test-success ""Glow cell B2."" (debug-print ""b2""))))" & vbLf & _
        "        (at-row ""row 2"" (begin (english-vla ""fade cell {r:cell}"" (debug-print {r}))" & vbLf & _
        "                              (test-success ""Fade cell B3."" (debug-print ""WRONG""))))))" & vbLf & _
        "(two-tagged-rules)", "atrow-vocab"
    d = Err.Description
    On Error GoTo 0
    Report "at-row: a generator's own emitted rows stay distinguishable", _
           InStr(1, d, "row 2", vbTextCompare) > 0 And InStr(1, d, "row 1", vbTextCompare) = 0, "got: " & d

    ' at-row's own shape is enforced - wrong arity refuses in words,
    ' not a raw VBA subscript error.
    EnglishResetGrammar
    On Error Resume Next
    Err.Clear
    EnglishLoadVocabularyText "(at-row ""onlyonearg"")", "atrow-vocab"
    d = Err.Description
    On Error GoTo 0
    Report "at-row: its own wrong arity refuses in words", _
           InStr(1, d, "takes exactly two arguments", vbTextCompare) > 0, "got: " & d

    ' Regression guard: an ordinary, untagged failure is byte-identical
    ' to before this pass - no stray "(" suffix leaks in when nothing
    ' was ever wrapped in at-row.
    EnglishResetGrammar
    On Error Resume Next
    Err.Clear
    EnglishLoadVocabularyText "(zzz-not-a-thing ""x"")", "atrow-vocab"
    d = Err.Description
    On Error GoTo 0
    Report "at-row: an untagged failure's wording is unchanged", _
           InStr(1, d, "unrecognized top-level directive 'zzz-not-a-thing'", vbTextCompare) > 0 And _
           InStr(1, d, "(", vbBinaryCompare) = 0, "got: " & d

    ' LISTOPS-PROVENANCE, the emitted-code half: an at-row-tagged
    ' defmacro's own template gets rewrapped in (gen-row ...) by
    ' RegisterVocabMacro, so EVERY call to the generated macro - here,
    ' one real sentence translating through it - carries the row it
    ' was generated from in the COMPILED VBA's own trailing comment,
    ' not just in a load-time error message.
    EnglishResetGrammar
    On Error Resume Next
    Err.Clear
    n = EnglishLoadVocabularyText( _
        "(at-row ""42"" (defmacro (tagged-set-style n s) ""doc"" (set! (. (activesheet.listobjects n) tablestyle) s)))" & vbLf & _
        "(english-vla ""tag style of table {n:text} to {s:text}"" (tagged-set-style {n} {s}))" & vbLf & _
        "(test-success ""Tag style of table Sales to Foo."" (tagged-set-style ""sales"" ""foo""))", "atrow-vocab")
    Report "at-row: an at-row-tagged macro's compiled VBA carries vla-row on every call", _
           Err.Number = 0 And InStr(1, EnglishToVba("Tag style of table Sales to Foo."), "vla-row:42", vbTextCompare) > 0, _
           "err: " & Err.Description & " / vba: " & Left$(EnglishToVba("Tag style of table Sales to Foo."), 200)
    On Error GoTo 0

    ' Regression guard, the emitted-code half: an ORDINARY macro (never
    ' at-row-wrapped) compiles with no vla-row anywhere - gen-row
    ' wrapping is opt-in only, never applied by accident.
    EnglishResetGrammar
    On Error Resume Next
    Err.Clear
    EnglishLoadVocabularyText _
        "(defmacro (plain-set-style n s) ""doc"" (set! (. (activesheet.listobjects n) tablestyle) s))" & vbLf & _
        "(english-vla ""plain style of table {n:text} to {s:text}"" (plain-set-style {n} {s}))" & vbLf & _
        "(test-success ""Plain style of table Sales to Foo."" (plain-set-style ""sales"" ""foo""))", "atrow-vocab"
    Report "at-row: an ordinary, untagged macro's compiled VBA carries no vla-row at all", _
           Err.Number = 0 And InStr(1, EnglishToVba("Plain style of table Sales to Foo."), "vla-row", vbTextCompare) = 0, _
           "err: " & Err.Description & " / vba: " & Left$(EnglishToVba("Plain style of table Sales to Foo."), 200)
    On Error GoTo 0

    EnglishResetGrammar
End Sub

' ---------------------------------------------------------------------
'  GEXPANDER.1: the fully macro-expanded form of a vocabulary, on
'  demand (EnglishExpandedVocabularyText) rather than automatically on
'  every load (GEXPANDER.0's own reversed design - see
'  BETA_ROADMAP.md). A live temp phrasebook exercises every real
'  directive kind DispatchVocabForm recognizes (a plain english-vla
'  rule, a defmacro whose call expands to a (begin ...) pair, an
'  at-row-tagged test-success) so the pins check the SAME real content
'  GEXPANDER.0's own pins did - only the trigger (a direct function
'  call, not a side effect of loading) and the staleness stamp (size
'  only, no mtime - unchanged from GEXPANDER.0's own second correction)
'  differ. No file-write pins here anymore: EnglishExpandedVocabularyText
'  returns text only; the actual Save-As write only happens from
'  VLA_IDE.bas's own dialog-driven EnglishIdeExportExpandedVocabulary,
'  which needs a live host's GetSaveAsFilename to exercise at all. The
'  generator's own shape - pattern/call/testsentence/testcall each
'  substituted as a WHOLE value, {r:cell}/{r} slot syntax riding
'  through untouched inside them - is table-property-family's own
'  real, already-shipped shape (scripts/english.vla), not an invented
'  one, so this pin is trusted the same way TABLE-FAMILY's own
'  production use already is.
' ---------------------------------------------------------------------
Public Sub TestGExpander()
    Dim vocabPath As String
    vocabPath = WriteTempLib("vla_gexpander_vocab.vla", _
        "(defmacro (glow-pair pattern call testsentence testcall)" & vbCrLf & _
        "    (begin" & vbCrLf & _
        "        (english-vla pattern call)" & vbCrLf & _
        "        (test-success testsentence testcall)))" & vbCrLf & vbCrLf & _
        "(english-vla ""warm cell {r:cell}"" (debug-print {r}))" & vbCrLf & _
        "(at-row ""row 7"" (test-success ""Warm cell B2."" (debug-print ""b2"")))" & vbCrLf & _
        "(glow-pair ""glow cell {r:cell}"" (debug-print {r}) ""Glow cell B2."" (debug-print ""b2""))")

    Dim expText As String
    On Error Resume Next
    Err.Clear
    expText = EnglishExpandedVocabularyText(vocabPath)
    Report "gexpander: EnglishExpandedVocabularyText loads a phrasebook exercising every directive kind and returns its expanded text clean", _
           Err.Number = 0 And Len(expText) > 0, "err: " & Err.Description
    On Error GoTo 0

    ' GEXPANDERLINT.0: expText is now VLA_Lint.VlaLintFormat's own
    ' house-style output, not GEXPANDER.0's flat one-line-per-directive
    ' text - english-vla/test-success are both IsDirectiveShaped
    ' (VLA_Lint.bas), so both ALWAYS verticalize, regardless of width.
    ' Owner correction to this pin, this session: PpVerticalize now puts
    ' EVERY argument (including the first) on its own line 4 past the
    ' form's own column, functor alone on the opening line - "both
    ' arguments as one aligned list beneath the functor" - not the
    ' original shape (first argument inline with the functor, second
    ' aligned under IT). Both these two top-level forms render the same
    ' way whether they're directive-shaped-forced or reached PpVerticalize
    ' via the width gate - one shape now, so this pin does not need to
    ' distinguish which path got them there. Hand-derived against
    ' PpForm/PpVerticalize directly - not yet live-verified, needs an
    ' owner run per this file's own bar for ✅.
    Report "gexpander: a generator's (begin ...) splice is fully flattened, never the one-line call", _
           InStr(1, expText, "(glow-pair """, vbBinaryCompare) = 0 And _
           InStr(1, expText, "(english-vla ""glow cell {r:cell}"" (debug-print {r}))", vbBinaryCompare) = 0 And _
           InStr(1, expText, "(english-vla" & vbCrLf & _
                              Space$(4) & """glow cell {r:cell}""" & vbCrLf & _
                              Space$(4) & "(debug-print {r}))", vbBinaryCompare) > 0 And _
           InStr(1, expText, "(test-success ""Glow cell B2."" (debug-print ""b2""))", vbBinaryCompare) = 0 And _
           InStr(1, expText, "(test-success" & vbCrLf & _
                              Space$(4) & """Glow cell B2.""" & vbCrLf & _
                              Space$(4) & "(debug-print ""b2""))", vbBinaryCompare) > 0, _
           "got: " & expText

    Report "gexpander: registration order is preserved (hand-written rule before the generator's own output)", _
           InStr(1, expText, "warm cell") > 0 And InStr(1, expText, "glow cell") > 0 And _
           InStr(1, expText, "warm cell") < InStr(1, expText, "glow cell"), "got: " & expText

    Report "gexpander: an at-row label prints as its own provenance comment directly above its form", _
           InStr(1, expText, "; row: row 7" & vbCrLf & "(test-success", vbBinaryCompare) > 0, _
           "got: " & expText

    Dim rowCommentCount As Long, scanPos As Long
    scanPos = 1
    Do
        scanPos = InStr(scanPos, expText, "; row:", vbBinaryCompare)
        If scanPos = 0 Then Exit Do
        rowCommentCount = rowCommentCount + 1
        scanPos = scanPos + 1
    Loop
    Report "gexpander: only the at-row-tagged form gets a provenance comment - exactly one, not zero or two", _
           rowCommentCount = 1, "count: " & rowCommentCount

    Report "gexpander: the artifact stamps a whitespace-insensitive hash of the source for a downstream staleness check", _
           InStr(1, expText, "source-hash:", vbTextCompare) > 0 And _
           InStr(1, expText, "source-size:", vbTextCompare) = 0, "got: " & Left$(expText, 200)

    ' 0.5.1: the stamp ignores whitespace by construction - re-indenting the
    ' source (what Lint VLA does) must leave it unchanged, and a one-byte
    ' token edit must not. Pinned directly on EnglishSourceHash so the
    ' property is proven, not just described in the header.
    Dim hashProbe As String, h1 As String
    hashProbe = WriteTempLib("vla_gexpander_hash_probe.vla", "(defmacro (a b)" & vbCrLf & "  (set! b 1))")
    h1 = EnglishSourceHash(hashProbe)
    hashProbe = WriteTempLib("vla_gexpander_hash_probe.vla", "(defmacro" & vbTab & "(a b)   (set! b 1))" & vbLf & vbLf)
    Report "gexpander: source-hash ignores indentation, tabs and line endings", _
           EnglishSourceHash(hashProbe) = h1, h1 & " vs " & EnglishSourceHash(hashProbe)
    hashProbe = WriteTempLib("vla_gexpander_hash_probe.vla", "(defmacro (a b)" & vbCrLf & "  (set! b 2))")
    Report "gexpander: source-hash changes on a one-byte token edit", _
           EnglishSourceHash(hashProbe) <> h1, h1
    ' 0.5.3 (SEC.11): the stamp is SHA-256, prefixed. The prefix is
    ' load-bearing rather than decorative - it is how
    ' check_rule_coverage.ps1 tells a 0.5.3 stamp from the 0.5.1 one it
    ' still reads, instead of inferring the generation from digest
    ' length. The retired shape is asserted ABSENT beside it, so a
    ' regression to the polynomial fails loudly here rather than
    ' silently downgrading every consent key.
    Dim hexPart As String
    hexPart = Mid$(h1, 8, 64)
    Report "gexpander: source-hash carries the explicit sha256: prefix", _
           Left$(h1, 7) = "sha256:", h1
    Report "gexpander: source-hash has the stamped shape (64 hex digits over N bytes)", _
           Len(h1) = 71 + Len(" over 22 non-whitespace bytes") And _
           Mid$(h1, 72) = " over 22 non-whitespace bytes" And _
           Not (hexPart Like "*[!0-9A-F]*"), h1
    Report "gexpander: source-hash no longer has the retired 8-hex shape", _
           Not (h1 Like "[0-9A-F][0-9A-F][0-9A-F][0-9A-F][0-9A-F][0-9A-F][0-9A-F][0-9A-F] over*"), h1
    On Error Resume Next
    Kill hashProbe
    On Error GoTo 0

    Report "gexpander: the header names the source by its bare filename, never a machine-specific absolute path", _
           InStr(1, expText, "from vla_gexpander_vocab.vla", vbTextCompare) > 0 And _
           InStr(1, expText, "source-mtime", vbTextCompare) = 0 And _
           InStr(1, expText, Environ$("TEMP"), vbTextCompare) = 0, _
           "got: " & Left$(expText, 200)

    EnglishResetGrammar
    Kill vocabPath
End Sub

' AS.1/RULECOVERAGE.0: EnglishRuleCoverageReport, VBA-native and firing-
' based - a rule with zero real test-success firings (never merely zero
' POSITIONALLY-nearby ones, tools/check_rule_coverage.ps1's own weaker
' substitute). Three rules, deliberately zero/one/two firings so all
' three buckets are exercised in one fixture: "warm cell" gets no test
' at all; "glow cell" gets exactly one; "cool cell" gets two, so it must
' be COUNTED but never individually LISTED (2+ rules are a tally only,
' mirroring AS.8's own tiered detail). "cool cell" is asserted absent
' from the whole report text for exactly that reason - if it ever shows
' up anywhere, the 2+ bucket stopped being count-only.
Public Sub TestRuleCoverage()
    Dim vocabPath As String
    vocabPath = WriteTempLib("vla_rulecoverage_vocab.vla", _
        "(english-vla ""warm cell {r:cell}"" (debug-print {r}))" & vbCrLf & vbCrLf & _
        "(english-vla ""glow cell {r:cell}"" (debug-print {r}))" & vbCrLf & _
        "(test-success ""Glow cell B2."" (debug-print ""b2""))" & vbCrLf & vbCrLf & _
        "(english-vla ""cool cell {r:cell}"" (debug-print {r}))" & vbCrLf & _
        "(test-success ""Cool cell B2."" (debug-print ""b2""))" & vbCrLf & _
        "(test-success ""Cool cell C3."" (debug-print ""c3""))" & vbCrLf & vbCrLf & _
        "(test-fail ""Bad thing."" ""Don't understand"")")

    Dim rpt As String
    On Error Resume Next
    Err.Clear
    rpt = EnglishRuleCoverageReport(vocabPath)
    Report "rulecoverage: EnglishRuleCoverageReport loads a phrasebook exercising all three buckets and returns clean", _
           Err.Number = 0 And Len(rpt) > 0, "err: " & Err.Description
    On Error GoTo 0

    Report "rulecoverage: rule count is correct", InStr(1, rpt, "Rules: 3", vbBinaryCompare) > 0, "got: " & rpt

    Report "rulecoverage: the untested rule lands in the zero-test bucket, listed by pattern", _
           InStr(1, rpt, "Zero tests, 1 rule(s)", vbBinaryCompare) > 0 And _
           InStr(1, rpt, "warm cell {r:cell}", vbBinaryCompare) > 0 And _
           InStr(1, rpt, "warm cell") < InStr(1, rpt, "Exactly one test"), "got: " & rpt

    Report "rulecoverage: the once-tested rule lands in the one-test bucket, listed by pattern", _
           InStr(1, rpt, "Exactly one test, 1 rule(s)", vbBinaryCompare) > 0 And _
           InStr(1, rpt, "Exactly one test") < InStr(1, rpt, "glow cell {r:cell}"), "got: " & rpt

    Report "rulecoverage: the twice-tested rule is counted in 2+ but never individually listed", _
           InStr(1, rpt, "cool cell", vbBinaryCompare) = 0, "got: " & rpt

    Report "rulecoverage: test-fail is counted at the file level, never attributed to a rule", _
           InStr(1, rpt, "test-fail proofs (file-level, never attributed to a rule): 1", vbBinaryCompare) > 0, _
           "got: " & rpt

    Report "rulecoverage: every test-success in this fixture fired a real phrase rule, so the structural bucket is zero", _
           InStr(1, rpt, "test-success proofs for a structural form, not a phrase rule: 0", vbBinaryCompare) > 0, _
           "got: " & rpt

    Report "rulecoverage: the summary line's own arithmetic is correct", _
           InStr(1, rpt, "SUMMARY: 1/3 rules have 2+ tests; 1 have exactly one; 1 have zero", vbBinaryCompare) > 0, _
           "got: " & rpt

    EnglishResetGrammar
    Kill vocabPath
End Sub

' SEC.2: raw behind explicit, per-phrasebook consent - see
' VLA_SentenceEngine.bas's own header just above EnglishLoadVocabularyText
' for the full design (gated at EnglishLoadVocabulary, the file-path
' loader, specifically NOT at EnglishLoadVocabularyText - that shared
' primitive is what VLA_Browser.bas's already-shipped, host-free
' translate API calls directly and documents as never showing a
' dialog), and especially why no test-bypass toggle exists anywhere in
' this mechanism. This proves the one safely automatable DEVICE-scope
' path: content already consented at device scope (SaveSetting
' pre-seeded with the exact EnglishSourceHash a real "Yes, remember for
' every workbook on this device" click would have stored) skips both
' prompts and loads normally - the real production code path, no
' shortcut. TestRawConsentWorkbookScope, just below, proves the sibling
' WORKBOOK-scope path the same way. The live prompt-and-decline/accept
' interaction, including the second (which-scope) dialog, is
' owner-verified manually, DI.1's own precedent for an interactive
' security dialog in this codebase - deliberately never exercised
' here, since doing so would either hang on a real MsgBox or require
' the very kind of bypass toggle this design has none of.
Public Sub TestRawConsentDeviceScope()
    Dim vocabPath As String
    vocabPath = WriteTempLib("vla_rawconsent_device_vocab.vla", _
        "(english-vla ""scorch cell {r:cell}"" (raw ""Debug.Print 1""))")

    Dim contentHash As String
    contentHash = EnglishSourceHash(vocabPath)

    ' The literal "SEC2RawConsentV2" here (and at every other seed/clean
    ' site in this module) MIRRORS VLA_SentenceEngine's own
    ' RAW_CONSENT_SECTION, which is Private and so cannot be referenced
    ' from here. It must move whenever that constant moves. SEC.11
    ' renamed it from "SEC2RawConsent" precisely so pre-0.5.3 grants
    ' could not collide with SHA-256-keyed ones - and these seeds went
    ' stale in that same edit: a seed written under the OLD name leaves
    ' the engine finding no grant, which in a test run means a live
    ' consent MsgBox with nobody to click it. If this suite ever hangs
    ' on a dialog here, that is the first thing to check.
    On Error Resume Next
    DeleteSetting "Frazaro", "SEC2RawConsentV2", contentHash
    On Error GoTo 0
    SaveSetting "Frazaro", "SEC2RawConsentV2", contentHash, "granted"

    EnglishResetGrammar
    Dim n As Long
    On Error Resume Next
    Err.Clear
    n = EnglishLoadVocabulary(vocabPath)
    Report "sec2: a raw-bearing phrasebook already consented at DEVICE scope loads with no prompt", _
           Err.Number = 0 And n = 1, "err: " & Err.Description & " n=" & n
    On Error GoTo 0

    EnglishResetGrammar
    On Error Resume Next
    DeleteSetting "Frazaro", "SEC2RawConsentV2", contentHash
    On Error GoTo 0
    Kill vocabPath
End Sub

' SEC.2, continued: the WORKBOOK-scope sibling of TestRawConsentDeviceScope
' above - content already consented at workbook scope (a
' CustomDocumentProperty on ActiveWorkbook, keyed the same way a real
' "Yes, remember for this workbook only" click would have stored it)
' also skips both prompts and loads normally. Proves EnglishLoadVocabulary
' checks ActiveWorkbook, not ThisWorkbook - this test runs from
' whatever workbook is active when VlaSelfTest runs, exactly the
' production shape (VLA_IDE.bas's own CaptureHost precedent: "add-ins
' never appear [as ActiveWorkbook], which is why [it] was never
' ThisWorkbook").
Public Sub TestRawConsentWorkbookScope()
    Dim vocabPath As String
    vocabPath = WriteTempLib("vla_rawconsent_workbook_vocab.vla", _
        "(english-vla ""char cell {r:cell}"" (raw ""Debug.Print 3""))")

    Dim contentHash As String
    contentHash = EnglishSourceHash(vocabPath)
    Dim propName As String
    propName = "SEC2RawConsentV2 " & contentHash

    On Error Resume Next
    ActiveWorkbook.CustomDocumentProperties(propName).Delete
    On Error GoTo 0
    ActiveWorkbook.CustomDocumentProperties.Add Name:=propName, LinkToContent:=False, Type:=4, Value:="granted"

    EnglishResetGrammar
    Dim n As Long
    On Error Resume Next
    Err.Clear
    n = EnglishLoadVocabulary(vocabPath)
    Report "sec2: a raw-bearing phrasebook already consented at WORKBOOK scope loads with no prompt", _
           Err.Number = 0 And n = 1, "err: " & Err.Description & " n=" & n
    On Error GoTo 0

    EnglishResetGrammar
    On Error Resume Next
    ActiveWorkbook.CustomDocumentProperties(propName).Delete
    On Error GoTo 0
    Kill vocabPath
End Sub

' GO.6: VLA_IDE.PersistPhrasebookPath/LoadedPhrasebookPaths/
' PhrasebookAlreadyLoaded, proven the same way SEC.2's own workbook-
' scope consent tests are above - real ActiveWorkbook.CustomDocument-
' Properties round-tripped through the real functions, no bypass
' toggle. Fake paths are enough here (these three never touch the
' filesystem, only the property store); ReplayPersistedPhrasebooks
' (which DOES read real files) gets its own test just below.
Public Sub TestPhrasebookPersistence()
    On Error Resume Next
    ActiveWorkbook.CustomDocumentProperties("VLA_LoadedPhrasebooks").Delete
    On Error GoTo 0

    PersistPhrasebookPath ActiveWorkbook, "C:\fake\one_go6.vla"
    PersistPhrasebookPath ActiveWorkbook, "C:\fake\two_go6.vla"

    Dim paths As Collection
    Set paths = LoadedPhrasebookPaths(ActiveWorkbook)
    Report "go6: two distinct paths persist in load order", _
           paths.Count = 2 And CollItemIs(paths, 1, "C:\fake\one_go6.vla") And CollItemIs(paths, 2, "C:\fake\two_go6.vla"), _
           "count=" & paths.Count

    Report "go6: PhrasebookAlreadyLoaded finds an already-persisted path", _
           PhrasebookAlreadyLoaded(ActiveWorkbook, "C:\fake\one_go6.vla"), "n/a"
    Report "go6: PhrasebookAlreadyLoaded is case-insensitive, matching Windows paths", _
           PhrasebookAlreadyLoaded(ActiveWorkbook, "C:\FAKE\ONE_GO6.vla"), "n/a"
    Report "go6: an unrelated path is correctly reported as not loaded", _
           Not PhrasebookAlreadyLoaded(ActiveWorkbook, "C:\fake\three_go6.vla"), "n/a"

    PersistPhrasebookPath ActiveWorkbook, "C:\fake\one_go6.vla"   ' re-persisting a known path must not add a duplicate entry
    Set paths = LoadedPhrasebookPaths(ActiveWorkbook)
    Report "go6: persisting an already-known path is a no-op, not a new entry", _
           paths.Count = 2, "count=" & paths.Count

    ' GO.1's own "no cap on how many sources" correction: a dozen
    ' distinct paths all persist and read back, unbounded - the same
    ' shape as a source file's own import list, no ceiling anywhere.
    Dim n As Long
    For n = 3 To 14
        PersistPhrasebookPath ActiveWorkbook, "C:\fake\many_go6_" & n & ".vla"
    Next n
    Set paths = LoadedPhrasebookPaths(ActiveWorkbook)
    Report "go6: persisting well beyond a handful of phrasebooks is not capped", _
           paths.Count = 14, "count=" & paths.Count

    On Error Resume Next
    ActiveWorkbook.CustomDocumentProperties("VLA_LoadedPhrasebooks").Delete
    On Error GoTo 0
End Sub

' GO.6, continued: ReplayPersistedPhrasebooks (VLA_IDE.bas), reached
' by every real command through IdeLoadVocab, proven end to end - a
' real temp phrasebook file persisted exactly as EnglishIdeLoadPhrasebook
' would leave it, plus a second remembered path naming a file that no
' longer exists (the moved/deleted case the header comment says must
' be skipped, never fatal). Confirms the ADD semantics GO.1 ratified: a
' rule already loaded from elsewhere (test-base, loaded directly here
' to stand in for "whatever IdeLoadVocab's base-corpus block already
' loaded") survives the replay side by side with the replayed
' phrasebook's own rule - neither one resets the other.
Public Sub TestPhrasebookReplayAddsNotReplaces()
    Dim vocabPath As String
    vocabPath = WriteTempLib("vla_go6_replay_vocab.vla", _
        "(english-vla ""frobnicate cell {r:cell}"" (raw ""Debug.Print 9""))")

    Dim contentHash As String
    contentHash = EnglishSourceHash(vocabPath)
    On Error Resume Next
    DeleteSetting "Frazaro", "SEC2RawConsentV2", contentHash
    On Error GoTo 0
    SaveSetting "Frazaro", "SEC2RawConsentV2", contentHash, "granted"

    ' SEC.9: the replay is now gated on a DEVICE-side approval per path,
    ' so both remembered paths need a record here or this test stops on a
    ' live MsgBox with nobody to answer it - fatal for an automated suite,
    ' exactly the hazard the SEC.2 note at the bottom of this module
    ' warns about. Seeding is not a bypass: this is byte-for-byte what a
    ' real "yes" leaves behind, the same way the SEC2RawConsentV2 seeds
    ' above stand in for a real raw-consent click. The ghost path gets a
    ' record too - without one it would be the thing that prompts, and
    ' its whole purpose is to prove a MISSING file is skipped quietly.
    On Error Resume Next
    DeleteSetting "Frazaro", "SEC9PhrasebookPath", LCase$(vocabPath)
    DeleteSetting "Frazaro", "SEC9PhrasebookPath", LCase$("C:\nonexistent\ghost_go6.vla")
    On Error GoTo 0
    SaveSetting "Frazaro", "SEC9PhrasebookPath", LCase$(vocabPath), contentHash & "|granted"
    SaveSetting "Frazaro", "SEC9PhrasebookPath", LCase$("C:\nonexistent\ghost_go6.vla"), "|granted"

    On Error Resume Next
    ActiveWorkbook.CustomDocumentProperties("VLA_LoadedPhrasebooks").Delete
    On Error GoTo 0
    ' pre-seeds two remembered paths directly (bypassing PersistPhrasebookPath's
    ' own dedupe) - the real one this test loads plus a second, already-gone
    ' one, exactly what a workbook that remembers a moved file would hold
    ActiveWorkbook.CustomDocumentProperties.Add Name:="VLA_LoadedPhrasebooks", LinkToContent:=False, _
        Type:=4, Value:=vocabPath & vbLf & "C:\nonexistent\ghost_go6.vla"

    EnglishResetGrammar
    EnglishLoadVocabularyText "(english-vla ""already here cell {r:cell}"" (set! (range {r}) 1))", "test-base"

    Dim errNum As Long
    On Error Resume Next
    Err.Clear
    ReplayPersistedPhrasebooks
    errNum = Err.Number
    On Error GoTo 0
    Report "go6: replaying a moved/deleted phrasebook slot does not raise", errNum = 0, "err=" & errNum

    Dim rpt As String
    rpt = EnglishLoadedSourcesReport()
    Report "go6: replay ADDS the phrasebook onto the already-loaded base - both provenances visible", _
           InStr(rpt, "test-base") > 0 And InStr(rpt, vocabPath) > 0, rpt

    Dim got1 As String, got2 As String
    got1 = EnglishToVla("Already here cell A1.")
    got2 = EnglishToVla("Frobnicate cell A1.")
    Report "go6: the base rule loaded before replay still resolves after it", InStr(got1, "range") > 0, got1
    Report "go6: the replayed phrasebook's own rule resolves too", Len(got2) > 0, got2

    EnglishResetGrammar
    On Error Resume Next
    ActiveWorkbook.CustomDocumentProperties("VLA_LoadedPhrasebooks").Delete
    On Error GoTo 0
    On Error Resume Next
    DeleteSetting "Frazaro", "SEC2RawConsentV2", contentHash
    DeleteSetting "Frazaro", "SEC9PhrasebookPath", LCase$(vocabPath)
    DeleteSetting "Frazaro", "SEC9PhrasebookPath", LCase$("C:\nonexistent\ghost_go6.vla")
    On Error GoTo 0
    Kill vocabPath
End Sub

' SEC.9: the gate actually gates. The sibling test above proves an
' APPROVED path still replays; this one proves a DENIED one does not,
' through the same real loader, with no dialog anywhere - a recorded
' "no" is exactly what makes that possible, and is why declining is
' persisted rather than merely acted on once.
'
' Designed so it would fail if the gate were removed: the phrasebook is
' real, present, and readable, and its rule is one nothing else defines.
' With the gate gone, ReplayPersistedPhrasebooks loads it and the
' sentence resolves. The pass condition is therefore a POSITIVE
' observation - "this specific sentence does not resolve, and the report
' does not name this file" - rather than the absence of a prompt, which
' is not evidence of anything on a machine that would not have prompted.
Public Sub TestSec9DeniedPhrasebookIsNotReplayed()
    Dim vocabPath As String
    vocabPath = WriteTempLib("vla_sec9_denied_vocab.vla", _
        "(english-vla ""zorblat cell {r:cell}"" (set! (range {r}) 7))")

    Dim contentHash As String
    contentHash = EnglishSourceHash(vocabPath)
    On Error Resume Next
    DeleteSetting "Frazaro", "SEC9PhrasebookPath", LCase$(vocabPath)
    On Error GoTo 0
    ' Exactly what clicking "No" leaves behind.
    SaveSetting "Frazaro", "SEC9PhrasebookPath", LCase$(vocabPath), contentHash & "|denied"

    On Error Resume Next
    ActiveWorkbook.CustomDocumentProperties("VLA_LoadedPhrasebooks").Delete
    On Error GoTo 0
    ActiveWorkbook.CustomDocumentProperties.Add Name:="VLA_LoadedPhrasebooks", LinkToContent:=False, _
        Type:=4, Value:=vocabPath

    EnglishResetGrammar
    EnglishLoadVocabularyText "(english-vla ""already here cell {r:cell}"" (set! (range {r}) 1))", "test-base"

    Dim errNum As Long
    On Error Resume Next
    Err.Clear
    ReplayPersistedPhrasebooks
    errNum = Err.Number
    On Error GoTo 0
    Report "sec9: a denied phrasebook is skipped without raising (no error, no dialog)", _
           errNum = 0, "err=" & errNum

    Dim rpt As String
    rpt = EnglishLoadedSourcesReport()
    Report "sec9: the denied phrasebook does not appear in the loaded-sources report", _
           InStr(rpt, vocabPath) = 0, rpt
    Report "sec9: the base grammar loaded before the replay is untouched", _
           InStr(rpt, "test-base") > 0, rpt

    ' The load-bearing one: its rule must NOT be reachable.
    Dim got As String
    On Error Resume Next
    Err.Clear
    got = EnglishToVla("Zorblat cell A1.")
    Dim resolveErr As Long
    resolveErr = Err.Number
    On Error GoTo 0
    Report "sec9: a sentence that needs the denied phrasebook's rule does not resolve", _
           resolveErr <> 0 Or InStr(got, "range") = 0, "got: " & got & " err=" & resolveErr

    EnglishResetGrammar
    On Error Resume Next
    ActiveWorkbook.CustomDocumentProperties("VLA_LoadedPhrasebooks").Delete
    DeleteSetting "Frazaro", "SEC9PhrasebookPath", LCase$(vocabPath)
    On Error GoTo 0
    Kill vocabPath
End Sub

' SEC.2, continued: EnglishLoadVocabularyText itself (the primitive
' VLA_Browser.bas's host-free EnglishTranslateTextToVla/ToVba call
' directly) must NEVER show the raw-consent prompt or touch
' SaveSetting/GetSetting at all - the gate lives one level up, in the
' file-based EnglishLoadVocabulary, specifically so this stays true.
' Proven by loading a raw-bearing phrasebook as TEXT (no file, no
' consent record seeded anywhere) and confirming it registers with no
' error - if the gate had leaked into this function, this call would
' either hang on a live MsgBox (fatal for an automated suite) or raise
' the decline message, neither of which happens.
Public Sub TestRawConsentTextPathUngated()
    EnglishResetGrammar
    Dim n As Long
    On Error Resume Next
    Err.Clear
    n = EnglishLoadVocabularyText( _
        "(english-vla ""char cell {r:cell}"" (raw ""Debug.Print 2""))", _
        "rawconsent-textpath-vocab")
    Report "sec2: EnglishLoadVocabularyText stays host-free - a raw-bearing TEXT load never prompts, matching VLA_Browser.bas's documented contract", _
           Err.Number = 0 And n = 1, "err: " & Err.Description & " n=" & n
    On Error GoTo 0
    EnglishResetGrammar
End Sub

' ---------------------------------------------------------------------
'  G5: the rule scratchpad. TryRule PRINTS its audition (Immediate
'  window - the authoring channel), which no pin can read, so these
'  pins hold the half that matters more: the RESTORATION contract.
'  Every observable state - rule count, existing translations, an
'  overridden rule's original template, usage counters, error
'  propagation - must be byte-for-byte as before the try.
' ---------------------------------------------------------------------
Public Sub TestG5()
    Dim d As String
    Dim junk As String
    Dim n0 As Long
    ' The candidate is discarded: its probe refuses again after the
    ' try, and the rule count is unchanged.
    EnglishResetGrammar
    n0 = EnglishRuleCount()
    EnglishTryRule "zap cell {r:cell}", "(debug-print {r})", "Zap cell B2."
    Report "g5: rule count unchanged after a try", EnglishRuleCount() = n0, _
           "count " & EnglishRuleCount() & " vs " & n0
    On Error Resume Next
    Err.Clear
    junk = EnglishToVla("Zap cell B2.")
    d = Err.Description
    On Error GoTo 0
    Report "g5: the candidate is discarded - its probe refuses again", _
           Len(d) > 0, "translated: " & Left$(Norm(junk), 120)

    ' Existing behavior is untouched by a try.
    AssertEnglish "g5: an existing sentence translates identically after a try", _
                  "Set x to 5.", "(set! x 5)"

    ' An override audition restores the ORIGINAL template.
    EnglishResetGrammar
    EnglishLoadVocabularyText "(english-vla ""warm cell {r:cell}"" (debug-print 1))" & vbLf & _
        "(test-success ""Warm cell B2."" (debug-print 1))", "base-vocab"
    EnglishTryRule "warm cell {r:cell}", "(debug-print 2)", "Warm cell B2.", True
    AssertEnglish "g5: an override audition restores the original rule", _
                  "Warm cell B2.", "(debug-print 1)"

    ' A registration the load would refuse is a report, not a crash.
    On Error Resume Next
    Err.Clear
    EnglishTryRule "warm cell {r:cell}", "(debug-print 3)", "Warm cell B2."
    Report "g5: a refused registration reports instead of raising", Err.Number = 0, Err.Description
    On Error GoTo 0

    ' A try is not usage: a shadowed probe that matched the EARLIER
    ' rule during the audition bumps nothing.
    Dim u0 As Long
    u0 = EnglishRuleUsageCount("rule: warm cell {r:cell}")
    EnglishTryRule "warm range {r:range}", "(debug-print 9)", "Warm cell B2."
    Report "g5: a try never counts as usage, even for the shadowing rule", _
           EnglishRuleUsageCount("rule: warm cell {r:cell}") = u0, _
           "usage " & EnglishRuleUsageCount("rule: warm cell {r:cell}") & " vs " & u0
    EnglishResetGrammar
End Sub

' ---------------------------------------------------------------------
'  G10 (owner design): bare surface tokens. Braces now mean exactly
'  one thing - capture - and a bare in|into just matches. The pins:
'  both branches, bare surfaces, the refusal naming branches, the
'  alternation optional both ways, load validation, the bare/braced
'  shape equivalence through the audit, and the set-formula template
'  face expanding to the same dot the language keeps first-class.
' ---------------------------------------------------------------------
Public Sub TestG10()
    Dim d As String
    Dim junk As String
    EnglishResetGrammar
    EnglishAddPhrase "push {e:expr} onto|into cell {r:cell}", "(debug-print {e})"
    AssertEnglish "g10: bare alternation matches branch one, binds nothing", _
                  "Push 5 onto cell B2.", "(debug-print 5)"
    AssertEnglish "g10: bare alternation matches branch two", _
                  "Push 5 into cell B2.", "(debug-print 5)"
    On Error Resume Next
    Err.Clear
    junk = EnglishToVla("Push 5 near cell B2.")
    d = Err.Description
    On Error GoTo 0
    Report "g10: a miss names the bare branches", _
           InStr(1, d, "'onto'", vbTextCompare) > 0 And InStr(1, d, "'into'", vbTextCompare) > 0, _
           "got: " & d

    EnglishResetGrammar
    EnglishAddPhrase "tilt cell {r:cell} flat|center/ed", "(debug-print 1)"
    AssertEnglish "g10: a bare branch carries surfaces", _
                  "Tilt cell B2 centered.", "(debug-print 1)"

    EnglishResetGrammar
    EnglishAddPhrase "poke [in|into] cell {r:cell}", "(debug-print {r})"
    AssertEnglish "g10: an alternation optional consumed when present", _
                  "Poke into cell B2.", "(debug-print ""b2"")"
    AssertEnglish "g10: an alternation optional free when absent", _
                  "Poke cell B2.", "(debug-print ""b2"")"

    EnglishResetGrammar
    On Error Resume Next
    Err.Clear
    EnglishLoadVocabularyText "(english-vla ""warm cell {r:cell} up||down"" (debug-print 1))", "selftest-vocab"
    d = Err.Description
    On Error GoTo 0
    Report "g10: an empty bare branch refuses at load", _
           InStr(1, d, "empty branch", vbTextCompare) > 0, "got: " & d

    ' Bare and braced twins are the SAME shape - the audit says so.
    EnglishResetGrammar
    Dim aud As String
    aud = EnglishAuditText( _
        "(english-vla ""fix cell|range {r:range}"" (debug-print {r}))" & vbLf & _
        "(english-vla ""fix {w:cell|range} {r:range}"" (debug-print {w}))" & vbLf & _
        "(test-success ""Fix cell B2."" (debug-print ""b2""))")
    Report "g10: bare and braced twins collide as duplicates in the audit", _
           InStr(1, aud, "duplicates", vbTextCompare) > 0, "report was: " & Left$(aud, 160)

    ' The set-formula template face: the friendly call expands to the
    ' first-class dot, and the emitted VBA is a .Formula2 line (not
    ' .Formula - Range.Formula auto-inserts "@" implicit intersection on
    ' anything that could spill, silently breaking every query engine's
    ' own "returns a spilled array" promise; VLA.bas's own "set!" case
    ' special-cases exactly this (. obj formula) shape to Formula2).
    EnglishResetGrammar
    Dim q As String
    q = Chr$(34)
    EnglishLoadVocabularyText "(defmacro (set-formula r f) (set! (. r formula) f))" & vbLf & _
        "(english-vla ""write formula {f:text} into|in cell {r:cell}""" & vbLf & _
        "    (set-formula (range {r}) {f}))" & vbLf & _
        "(test-success ""Write formula \" & q & "=A1\" & q & " in cell B3.""" & vbLf & _
        "    (set-formula (range ""b3"") ""=A1""))", "selftest-vocab"
    Dim vbaOut As String
    vbaOut = TryTranspile("g10: set-formula expands to the first-class dot", _
                          EnglishToVla("Write formula ""=A1"" in cell B3."))
    If Len(vbaOut) > 0 Then
        CheckFrags "g10: set-formula expands to the first-class dot", vbaOut, _
                   Array(".formula2 =")
    End If
    EnglishResetGrammar
End Sub

' ---------------------------------------------------------------------
'  G11: the template face. Three full-stack pins on the functor
'  pattern: the name-substitution trick (the ENGLISH layer computes
'  the macro name - (make-{d} ...) -> (make-bold ...) - before the
'  reader sees it), a nullary method functor, and a keyword-arg
'  functor - each expanding to the same dot the language keeps
'  first-class, emitted VBA unmoved.
' ---------------------------------------------------------------------
Public Sub TestG11()
    ' (Note the pattern MUST use the braced {d:bold|italic} - a bare
    ' alternation binds nothing, so a bare spelling would leave the
    ' template's {d} unsubstituted. Braces capture; that is the rule.)
    EnglishResetGrammar
    EnglishLoadVocabularyText _
        "(defmacro (make-bold r) (set! (. r font.bold) true))" & vbLf & _
        "(defmacro (make-italic r) (set! (. r font.italic) true))" & vbLf & _
        "(defmacro (merge-range r) (. r merge))" & vbLf & _
        "(defmacro (protect-sheet ws pw) (. ws protect :password pw))" & vbLf & _
        "(english-vla ""style cell {r:cell} {d:bold|italic}""" & vbLf & _
        "    (make-{d} (range {r})))" & vbLf & _
        "(english-vla ""seal range {r:range}""" & vbLf & _
        "    (merge-range (range {r})))" & vbLf & _
        "(english-vla ""guard this sheet with password {p:expr}""" & vbLf & _
        "    (protect-sheet activesheet {p}))" & vbLf & _
        "(test-success ""Style cell A1 italic.""" & vbLf & _
        "    (make-italic (range ""a1"")))", "selftest-vocab"
    Dim vbaOut As String
    vbaOut = TryTranspile("g11: the name-substitution functor expands to the member", _
                          EnglishToVla("Style cell A1 italic."))
    If Len(vbaOut) > 0 Then
        CheckFrags "g11: the name-substitution functor expands to the member", vbaOut, _
                   Array("font.italic")
    End If
    vbaOut = TryTranspile("g11: a nullary method functor expands to the call", _
                          EnglishToVla("Seal range A1:D1."))
    If Len(vbaOut) > 0 Then
        CheckFrags "g11: a nullary method functor expands to the call", vbaOut, _
                   Array(".merge")
    End If
    vbaOut = TryTranspile("g11: a keyword-arg functor carries its keyword through", _
                          EnglishToVla("Guard this sheet with password ""abc""."))
    If Len(vbaOut) > 0 Then
        CheckFrags "g11: a keyword-arg functor carries its keyword through", vbaOut, _
                   Array("Password:=")
    End If
    EnglishResetGrammar
End Sub

' ---------------------------------------------------------------------
'  G11r: the template-face remainder G11 deliberately left open - true/
'  false-valued members (wraptext, hidden) and the numberformat family,
'  by exactly G11's fired-only lockstep method. Two pins: the
'  antonym-pair functor (one alternation rule minting two macros whose
'  bodies bake in opposite literals - hide-row/unhide-row - proving the
'  {d} substitution actually selects the right macro body, not just the
'  right member name) and the three-way format-as functor (currency/
'  percent/date sharing one rule).
' ---------------------------------------------------------------------
Public Sub TestG11r()
    EnglishResetGrammar
    EnglishLoadVocabularyText _
        "(defmacro (hide-row n) (set! (. (rows n) hidden) true))" & vbLf & _
        "(defmacro (unhide-row n) (set! (. (rows n) hidden) false))" & vbLf & _
        "(defmacro (format-as-currency r) (set! (. r numberformat) ""$#,##0.00""))" & vbLf & _
        "(defmacro (format-as-percent r) (set! (. r numberformat) ""0.0%""))" & vbLf & _
        "(english-vla ""{d:hide|unhide} row {n:expr}""" & vbLf & _
        "    ({d}-row {n}))" & vbLf & _
        "(english-vla ""format cell {r:cell} as {d:currency|percent}""" & vbLf & _
        "    (format-as-{d} (range {r})))" & vbLf & _
        "(test-success ""Hide row 5.""" & vbLf & _
        "    (hide-row 5))", "selftest-vocab"
    Dim vbaOut As String
    vbaOut = TryTranspile("g11r: the hide branch bakes in True", _
                          EnglishToVla("Hide row 5."))
    If Len(vbaOut) > 0 Then
        CheckFrags "g11r: the hide branch bakes in True", vbaOut, _
                   Array("Rows(5)", "Hidden", "True")
    End If
    vbaOut = TryTranspile("g11r: the unhide branch bakes in False, not True", _
                          EnglishToVla("Unhide row 5."))
    If Len(vbaOut) > 0 Then
        CheckFrags "g11r: the unhide branch bakes in False, not True", vbaOut, _
                   Array("Rows(5)", "Hidden", "False")
    End If
    vbaOut = TryTranspile("g11r: the three-way format-as functor picks its branch", _
                          EnglishToVla("Format cell B2 as percent."))
    If Len(vbaOut) > 0 Then
        CheckFrags "g11r: the three-way format-as functor picks its branch", vbaOut, _
                   Array("NumberFormat", "0.0%")
    End If
    EnglishResetGrammar
End Sub

' ---------------------------------------------------------------------
'  G6: list-valued slots. {name:cat-list} (cat one of text/range/cell/
'  column/sheet/color) parses one or more MatchRefToken items (VLA_
'  English.bas - the extracted, shared body of every singular typed
'  slot's own token-consumption logic), comma-separated, lowering to a
'  new runtime primitive, `(array item1 item2 ...)` (EmitExpr/VLA.bas,
'  EvalExpr/VLA_Interpreter.bas) - deliberately NOT a LISTOPS primitive,
'  since this is a real value the emitted/interpreted PROGRAM uses (a
'  pivot's field names), never expand-time compiler data.
'  Oxford comma REQUIRED, by construction rather than by a special
'  rejection rule: "and" is only ever consumed as a silent no-op
'  immediately after an already-consumed comma, never as a standalone
'  separator - so "rows of A, B and columns of C" (the scoping pass's
'  own motivating sentence, pareto.txt section 10) stops the rows list
'  cleanly at "B" with zero lookahead against the surrounding pattern,
'  while a genuinely dropped Oxford comma on a real 3-item list produces
'  the ordinary "I understood '...' - then expected X but found Y"
'  refusal every other slot failure already gives, not a silent
'  misparse - proven below, not just claimed.
'  Two item categories tested, not one, to prove the mechanism is
'  actually generic: text-shaped (field names, no shape check) and
'  column-shaped (RefShapeOk's own IsColLetters, {c:column}'s existing
'  check, reused unchanged - a bad item still refuses per-item, so the
'  generic list mechanism doesn't cost {c:column}'s own early-refusal
'  quality).
' ---------------------------------------------------------------------
Public Sub TestG6()
    EnglishResetGrammar
    EnglishLoadVocabularyText _
        "(english-vla ""list fields {f:text-list}"" (debug-print {f}))" & vbLf & _
        "(english-vla ""sort by columns {c:column-list}"" (debug-print {c}))" & vbLf & _
        "(english-vla ""with rows of {a:text-list} and columns of {b:text-list}"" (debug-print {a} {b}))", "selftest-vocab"

    AssertEnglish "g6: a two-item text-list, comma only, no trailing 'and'", _
                  "List fields Region, Product.", _
                  "(debug-print (array ""region"" ""product""))"
    AssertEnglish "g6: a three-item text-list with the required Oxford comma", _
                  "List fields Region, Product, and Date.", _
                  "(debug-print (array ""region"" ""product"" ""date""))"
    AssertEnglish "g6: a single item is still a one-element list", _
                  "List fields Region.", _
                  "(debug-print (array ""region""))"
    AssertEnglish "g6: a column-list reuses column shape-checking via RefShapeOk, not text's no-op check", _
                  "Sort by columns B, C, and F.", _
                  "(debug-print (array ""b"" ""c"" ""f""))"

    ' The motivating disambiguation case itself (pareto.txt section 10's
    ' own illustrative sentence, verbatim): "and columns of" must never
    ' be swallowed as more rows-list items, with zero pattern lookahead
    ' - the rows list stops at "B" purely because no comma precedes
    ' "and", and "and columns of" is left for the sentence's own second
    ' clause to match normally.
    AssertEnglish "g6: 'and' with no preceding comma ends the list and hands off to the next clause, not a false extension", _
                  "With rows of Region, Product and columns of Segment.", _
                  "(debug-print (array ""region"" ""product"") (array ""segment""))"

    ' A dropped Oxford comma on a real 3-item list is the deliberate
    ' teaching moment: the list correctly stops at "Product" (no comma
    ' precedes "and"), so "and Date" is left over and the whole sentence
    ' fails to match to completion - a real, honest refusal, never a
    ' silent two-item misparse that quietly drops "Date" on the floor.
    Dim t As String, d As String
    On Error Resume Next
    Err.Clear
    t = EnglishToVla("List fields Region, Product and Date.")
    d = Err.Description
    On Error GoTo 0
    Report "g6: a missing Oxford comma refuses instead of silently dropping the last item", _
           Len(t) = 0, "vla: " & t & " / err: " & d

    ' A bad item in a column-list still shape-checks per item - proof
    ' the generic <category>-list mechanism didn't lose {c:column}'s
    ' own early refusal quality for a garbage column name.
    On Error Resume Next
    Err.Clear
    t = EnglishToVla("Sort by columns B, 123, and F.")
    d = Err.Description
    On Error GoTo 0
    Report "g6: a column-list item is still shape-checked per item, same as a bare {c:column} would be", _
           Len(t) = 0, "vla: " & t & " / err: " & d

    EnglishResetGrammar
End Sub

' ---------------------------------------------------------------------
'  G7: slot defaults. {n:expr=1} generalizes G1's optional-literal
'  present/absent duality from literals to typed slots: present, the
'  parsed value binds; absent, the text after '=' substitutes instead
'  (pure substitution - never re-parsed). Composes with an ordinary
'  [optional] literal riding ahead of it, so a whole trailing clause
'  ("at 5") can vanish together, not just the number. Two pins: the
'  slot present (its value wins, not the default) and the slot AND
'  its leading optional both absent (the default substitutes and the
'  rule still matches to completion, not a near-miss).
' ---------------------------------------------------------------------
Public Sub TestG7()
    EnglishResetGrammar
    EnglishLoadVocabularyText _
        "(defmacro (insert-row-at n) (. (rows n) insert))" & vbLf & _
        "(english-vla ""insert row [at] {n:expr=1}""" & vbLf & _
        "    (insert-row-at {n}))" & vbLf & _
        "(test-success ""Insert row at 5.""" & vbLf & _
        "    (insert-row-at 5))" & vbLf & _
        "(test-success ""Insert a row.""" & vbLf & _
        "    (insert-row-at 1))", "selftest-vocab"
    Dim vbaOut As String
    vbaOut = TryTranspile("g7: a present slot's value wins over its default", _
                          EnglishToVla("Insert row at 5."))
    If Len(vbaOut) > 0 Then
        CheckFrags "g7: a present slot's value wins over its default", vbaOut, _
                   Array("Rows(5)", "Insert")
    End If
    vbaOut = TryTranspile("g7: an absent slot and its leading optional both vanish", _
                          EnglishToVla("Insert a row."))
    If Len(vbaOut) > 0 Then
        CheckFrags "g7: an absent slot and its leading optional both vanish", vbaOut, _
                   Array("Rows(1)", "Insert")
    End If
    EnglishResetGrammar
End Sub

' ---------------------------------------------------------------------
'  G8: number words and ordinals. Cardinals thirteen..twenty extend
'  NumberWord's existing zero..twelve table at the tokenizer - safe
'  unconditionally, since none of those thirteen words is claimed
'  anywhere else. Ordinals (first..twentieth) are NOT at the tokenizer:
'  "first" is already reserved by mFnOf's "first of X" list-accessor
'  idiom (first of found-items -> vlafirst), so OrdinalWord is checked
'  only inside ParsePrimCore's expr grammar, and only after that idiom
'  check has already failed. Three pins, all against the core "Set v
'  to {expr}." sentence (needs no vocab): a cardinal past twelve, a
'  bare ordinal, and - the one that actually justifies the design -
'  "first of X" still resolving to vlafirst, not the digit 1.
' ---------------------------------------------------------------------
Public Sub TestG8()
    EnglishResetGrammar
    AssertEnglish "g8: a cardinal past twelve resolves at the tokenizer", _
                  "Set x to fifteen.", "(set! x 15)"
    AssertEnglish "g8: a bare ordinal resolves in the expr grammar", _
                  "Set x to third.", "(set! x 3)"
    AssertEnglish "g8: first-of-list is unshadowed by ordinal 'first'", _
                  "Set x to first of found-items.", "(set! x (vlafirst found-items))"
    EnglishResetGrammar
End Sub

' ---------------------------------------------------------------------
'  G-ROWLOOP: the row-loop family. "Count <name> [down] from <a> to
'  <b>:" already covered ascending and safe bottom-up descending (both
'  pinned here as no-regression checks); the one missing shape was a
'  custom step for skip-N loops, added to that same form rather than a
'  new one - see the Case "count" header note in VLA_English.bas.
' ---------------------------------------------------------------------
Public Sub TestGRowLoop()
    EnglishResetGrammar
    AssertEnglish "growloop: plain ascending is unchanged", _
                  "Count row from 2 to 9, set x to row.", "(for (row 2 9) (set! x row))"
    AssertEnglish "growloop: plain descending is unchanged", _
                  "Count row down from 9 to 2, set x to row.", "(for (row 9 2 -1) (set! x row))"
    AssertEnglish "growloop: ascending with an explicit step, one-line", _
                  "Count row from 2 to 20 step 2, set x to row.", "(for (row 2 20 2) (set! x row))"
    AssertEnglish "growloop: descending with an explicit step negates automatically, block", _
                  "Count row down from 20 to 2 step 2:" & vbLf & _
                  "  Set x to row." & vbLf & _
                  "Done.", "(for (row 20 2 (- 0 2))"
    AssertEnglish "growloop: the loop variable composes with column-row cell addressing", _
                  "Count row from 2 to 9, set x to cell in column A row row.", _
                  "(for (row 2 9) (set! x (cells row ""a"")))"
    EnglishResetGrammar
End Sub

' ---------------------------------------------------------------------
'  V7: first-token rule dispatch. Behavior-invariant by construction -
'  the goldens and the vocabulary proofs are the whole-corpus witness;
'  these pins hold the construction itself: first-match order across
'  buckets (an earlier slot-first rule beats a later bucketed one) and
'  within one (earliest wins; a shape fall-through still reaches the
'  later rule), surface expansion into buckets (bare branches and
'  stem/suffix, via the same BareSurfaces the signatures ride),
'  article-skipped keying, the skipped-rule near-miss simulated to the
'  exact old wording, and the two invalidations a count check cannot
'  see - an in-place override that rewords the pattern's head, and
'  G5's restore (a stale bucket entry would index past the restored
'  store, so the discarded candidate must leave a clean near-miss,
'  never a subscript error).
' ---------------------------------------------------------------------
Public Sub TestV7()
    Dim d As String
    Dim junk As String

    ' Order ACROSS buckets: the earlier universal (slot-first) rule
    ' must win a sentence the later bucketed rule also matches.
    EnglishResetGrammar
    EnglishAddPhrase "{v:var} wobbles", "(wob {v})"
    EnglishAddPhrase "gizmo wobbles", "(gizmo-rule)"
    AssertEnglish "v7: an earlier slot-first rule beats a later bucketed rule", _
                  "Gizmo wobbles.", "(wob gizmo)"

    ' Order WITHIN a bucket, and G2's shape fall-through inside it.
    EnglishResetGrammar
    EnglishAddPhrase "zap {r:cell}", "(zap-cell {r})"
    EnglishAddPhrase "zap {t:text}", "(zap-text {t})"
    AssertEnglish "v7: within a bucket the earlier rule wins", _
                  "Zap B2.", "(zap-cell ""b2"")"
    AssertEnglish "v7: a shape fall-through still reaches the later rule in the bucket", _
                  "Zap kettle.", "(zap-text ""kettle"")"
    ' Articles are skipped before the dispatch key is read, exactly
    ' as TryPhrase skips them before a first literal.
    AssertEnglish "v7: the dispatch key is read past leading articles", _
                  "A zap kettle.", "(zap-text ""kettle"")"

    ' Every surface of every branch heads its own bucket.
    EnglishResetGrammar
    EnglishAddPhrase "nudge|shove cell {r:cell}", "(debug-print {r})"
    AssertEnglish "v7: a bare-alternation branch reaches its own bucket", _
                  "Shove cell B2.", "(debug-print ""b2"")"
    EnglishResetGrammar
    EnglishAddPhrase "tilt/ed cell {r:cell}", "(debug-print {r})"
    AssertEnglish "v7: a stem/suffix surface reaches its own bucket", _
                  "Tilted cell B2.", "(debug-print ""b2"")"

    ' The skipped-rule near-miss is simulated exactly: with the
    ' prelude only, "A zzz." heads no bucket, so every rule is
    ' skipped - and the recorded near-miss must still be rule 1's
    ' ('set', at the article-skipped position), word for word what
    ' the linear scan produced.
    EnglishResetGrammar
    On Error Resume Next
    Err.Clear
    junk = EnglishToVla("A zzz.")
    d = Err.Description
    On Error GoTo 0
    Report "v7: a skipped rule's near-miss is simulated exactly", _
           InStr(1, d, "I understood 'a'", vbTextCompare) > 0 And _
           InStr(1, d, "expected 'set' but found 'zzz'", vbTextCompare) > 0, _
           "got: " & d

    ' An override can reword the pattern's head IN PLACE (the rule
    ' count never moves), so the index must rebuild on the flag: the
    ' new head's bucket has to exist before its first sentence.
    EnglishResetGrammar
    EnglishLoadVocabularyText "(english-vla ""shove cell {r:cell}"" (debug-print 1))" & vbLf & _
        "(test-success ""Shove cell B2."" (debug-print 1))", "base-vocab"
    EnglishLoadVocabularyText "(english-vla-override ""nudge|shove cell {r:cell}""" & vbLf & _
        "    (debug-print 2))", "dialect-vocab"
    AssertEnglish "v7: an in-place override rebuilds the dispatch index", _
                  "Nudge cell B2.", "(debug-print 2)"

    ' G5's restore pops the audited candidate back off - same count
    ' before and after, so only the flag protects the next dispatch
    ' from a bucket entry pointing past the restored store. The
    ' discarded candidate's sentence must refuse with the teaching
    ' parse error, never a subscript error.
    EnglishResetGrammar
    EnglishTryRule "wiggle cell {r:cell}", "(debug-print 7)", "Wiggle cell B2."
    On Error Resume Next
    Err.Clear
    junk = EnglishToVla("Wiggle cell B2.")
    d = Err.Description
    On Error GoTo 0
    Report "v7: a discarded try leaves a clean near-miss, not a stale index", _
           InStr(1, d, "understand", vbTextCompare) > 0, "got: " & d

    ' The sharpest restore case: an OVERRIDE audition replaces in
    ' place and the restore puts the original back - the count never
    ' moves at any step, and the audition's index was built over the
    ' candidate's narrower head. After the restore, the original's
    ' other branch must still find its rule.
    EnglishResetGrammar
    EnglishLoadVocabularyText "(english-vla ""twist|wiggle cell {r:cell}""" & vbLf & _
        "    (debug-print 1))" & vbLf & _
        "(test-success ""Wiggle cell B2."" (debug-print 1))", "base-vocab"
    EnglishTryRule "wiggle cell {r:cell}", "(debug-print 2)", "Wiggle cell B2.", True
    AssertEnglish "v7: an override audition's restore re-arms the original head's buckets", _
                  "Twist cell B2.", "(debug-print 1)"
    EnglishResetGrammar
End Sub

' ---------------------------------------------------------------------
'  V8 pins: the string-building pass. The builder is Private in two
'  modules, so every pin drives it through the public surface it
'  serves - transpile, expand-text pretty printing, sub assembly -
'  plus the manifest cache behind the resolve check. The REAL
'  invariance witness is the pair of empty golden diffs; these pins
'  hold the edges goldens cannot see (empty input, the doubling
'  boundary, cache transparency). The ADODB stream cache has no pin:
'  VocabReadFile is Private and file-path plumbing belongs to the
'  owner's Reload, where any wedge would be loud; recorded on the
'  ledger.
' ---------------------------------------------------------------------
Public Sub TestV8()
    ' Empty program: the builder's zero-length truncation - output is
    ' exactly the header, no padding leaked.
    Dim t As String
    t = VlaTranspile("")
    Report "v8: empty program transpiles to exactly the header", _
           t = "Option Explicit" & vbCrLf & vbCrLf, "got: [" & Left$(t, 60) & "]"

    ' A 200-statement main crosses the doubling boundary many times.
    Dim big As String, i As Long
    big = "(sub main ("
    big = big & ")"                          ' (sub main () ...)
    For i = 1 To 200
        big = big & " (set! x " & i & ")"
    Next
    big = big & ")"
    Dim r1 As String, r2 As String
    r1 = VlaTranspile(big)
    r2 = VlaTranspile(big)
    Report "v8: a large transpile is deterministic byte for byte", r1 = r2, _
           "lengths " & Len(r1) & " vs " & Len(r2)
    Dim cnt As Long, p As Long
    p = 1
    Do
        p = InStr(p, r1, "x = ")
        If p = 0 Then Exit Do
        cnt = cnt + 1
        p = p + 4
    Loop
    Report "v8: all 200 statements emitted once each", cnt = 200, "count " & cnt
    Report "v8: no padding leaks past the final line break", _
           Right$(r1, 2) = vbCrLf And Right$(RTrim$(Replace(r1, vbCrLf, "|")), 1) <> " ", _
           "tail: [" & Right$(r1, 12) & "]"

    ' The flat writer path: separator logic under the builder.
    Dim fired As Long
    t = VlaExpandText("(alpha beta gamma)", True, fired)
    Report "v8: flat pretty output is byte-exact", _
           t = "(alpha beta gamma)" & vbCrLf, "got: [" & Left$(t, 60) & "]"

    ' The multi-line writer path: a form too wide for one line opens
    ' with its head and indents each element - byte-exact.
    Dim a As String, b As String, c As String
    a = String$(30, "a")
    b = String$(30, "b")
    c = String$(30, "c")
    t = VlaExpandText("(seq " & a & " " & b & " " & c & ")", True, fired)
    Report "v8: multi-line pretty output is byte-exact", _
           t = "(seq" & vbCrLf & "    " & a & vbCrLf & "    " & b & vbCrLf & _
               "    " & c & ")" & vbCrLf, "got: [" & Left$(t, 80) & "]"

    ' Sub assembly and the statement join keep order under the builder.
    Dim ev As String
    On Error Resume Next
    ev = EnglishToVla("Set alpha to 1. Set beta to 2.")
    On Error GoTo 0
    Dim pa As Long, pb As Long
    pa = InStr(1, ev, "(set! alpha 1)")
    pb = InStr(1, ev, "(set! beta 2)")
    Report "v8: translated statements keep their order through the join", _
           pa > 0 And pb > pa, "positions " & pa & ", " & pb

    ' The manifest cache is transparent: a typo'd helper still
    ' refuses by name, a real one still passes, and a second check
    ' answers identically to the first.
    Dim ml As Long
    Dim m1 As String, m2 As String
    m1 = EnglishResolveCheck("(sub main () (vlanosuchhelper 1))", ml)
    Report "v8: the cached manifest still refuses a typo'd helper by name", _
           m1 = "vlanosuchhelper", "got '" & m1 & "'"
    m2 = EnglishResolveCheck("(sub main () (vlanosuchhelper 1))", ml)
    Report "v8: a second resolve check answers identically from the cache", _
           m1 = m2, "'" & m1 & "' vs '" & m2 & "'"
    m1 = EnglishResolveCheck("(sub main () (vlashowerror ""x""))", ml)
    Report "v8: a real runtime helper still passes the cached check", _
           m1 = "", "got '" & m1 & "'"
End Sub

' ---------------------------------------------------------------------
'  L6 pins: VlaTryValue's pure halves (VlaTryBuild's precedent - the
'  inject/run/delete half mutates the live project, so only the
'  Immediate-window smoke in the verification loop judges it; these
'  pin the build, the star substitution, and the value formatter,
'  all Public precisely so this suite can reach them). History
'  storage itself (the Names) is impure and unpinned for the same
'  recorded reason.
' ---------------------------------------------------------------------
Public Sub TestL6()
    Dim b As String, t As String, prob As String, d As String
    Dim ok As Boolean

    b = VlaTryValueBuild("(+ 1 2)")
    CheckFrags "l6: a bare expression wraps in the value function", b, _
               Array("(function vla-scratch-value () Variant", "(return (+ 1 2))")
    t = TryTranspile("l6: the value scratch transpiles runnable", b)
    If Len(t) > 0 Then
        CheckFrags "l6: the value scratch transpiles runnable", t, _
                   Array("Function vla_scratch_value", "Exit Function", "End Function")
    End If

    b = VlaTryValueBuild("(set! x 3) (* x x)")
    Dim pStmt As Long, pRet As Long
    pStmt = InStr(1, b, "(set! x 3)", vbTextCompare)
    pRet = InStr(1, b, "(return (* x x))", vbTextCompare)
    Report "l6: leading chunks are statements; the LAST chunk is the value", _
           pStmt > 0 And pRet > pStmt, Left$(Norm(b), 160)

    b = VlaTryValueBuild("(defmacro (tw x) (* x 2)) (tw 4)")
    Dim pMac As Long, pFn As Long
    pMac = InStr(1, b, "(defmacro (tw x)", vbTextCompare)
    pFn = InStr(1, b, "(function vla-scratch-value", vbTextCompare)
    Report "l6: definitions hoist above the value function", _
           pMac > 0 And pFn > pMac And InStr(1, b, "(return (tw 4))", vbTextCompare) > pFn, _
           Left$(Norm(b), 160)
    t = TryTranspile("l6: a hoisted macro expands inside the returned value", b)
    If Len(t) > 0 Then
        CheckFrags "l6: a hoisted macro expands inside the returned value", t, _
                   Array("(4 * 2)")
    End If

    CheckFrags "l6: a bare atom is a returnable value", VlaTryValueBuild("x"), _
               Array("(return x)")

    d = ""
    On Error Resume Next
    b = VlaTryValueBuild("(defmacro (tw x) (* x 2))")
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    Report "l6: definitions alone refuse with words", _
           InStr(1, d, "expression", vbTextCompare) > 0, "got: " & d

    Report "l6: *1 substitutes at a token boundary", _
           VlaTryExpandStars("(* *1 2)", "3", "", "", prob) = "(* 3 2)" And Len(prob) = 0, _
           "got '" & VlaTryExpandStars("(* *1 2)", "3", "", "", prob) & "' / " & prob

    Report "l6: *2 and *3 substitute; *12 is not a slot", _
           VlaTryExpandStars("(+ *2 *3 *12)", "9", "4", "5", prob) = "(+ 4 5 *12)" And Len(prob) = 0, _
           "got '" & VlaTryExpandStars("(+ *2 *3 *12)", "9", "4", "5", prob) & "' / " & prob

    Report "l6: stars inside strings and comments are never touched", _
           VlaTryExpandStars("(debug-print ""*1"") ; note *2", "3", "4", "", prob) = _
           "(debug-print ""*1"") ; note *2" And Len(prob) = 0, _
           "got '" & VlaTryExpandStars("(debug-print ""*1"") ; note *2", "3", "4", "", prob) & "'"

    b = VlaTryExpandStars("(f *1)", "", "", "", prob)
    Report "l6: an empty slot refuses with words and leaves the input untouched", _
           InStr(1, prob, "no *1 yet", vbTextCompare) > 0 And b = "(f *1)", _
           "prob '" & prob & "' text '" & b & "'"
    b = VlaTryExpandStars("(f *1)", "#<empty>", "", "", prob)
    Report "l6: a #<...> marker refuses substitution with words", _
           InStr(1, prob, "could not be stored", vbTextCompare) > 0, "prob '" & prob & "'"

    Dim lit As String
    lit = VlaTryValueLit(3.5, ok)
    Report "l6: numbers store as period-decimal literals", _
           ok And lit = "3.5" And VlaTryValueLit(CLng(3), ok) = "3", "got '" & lit & "'"
    lit = VlaTryValueLit("a" & Chr$(34) & "b\c", ok)
    Report "l6: strings store re-quoted with the tokenizer's escapes", _
           ok And lit = Chr$(34) & "a\" & Chr$(34) & "b\\c" & Chr$(34), "got '" & lit & "'"
    ' L6.1: the original compound pin read ok AFTER an in-expression
    ' ByRef mutation (... VlaTryValueLit(True, ok) = "True" And ok) and
    ' failed on-machine while its atomic siblings passed - the write
    ' was not observed by the trailing read. In-expression ByRef
    ' mutation order is NOT a contract; pins take one side-effecting
    ' call per STATEMENT and snapshot ok, so a real failure names
    ' itself with full state.
    Dim ev As Variant
    lit = VlaTryValueLit(ev, ok)
    Report "l6: an unstorable value becomes a refusing marker", _
           (Not ok) And lit = "#<empty>", _
           "got '" & lit & "' ok=" & ok
    Dim lit2 As String
    lit2 = VlaTryValueLit(True, ok)
    Report "l6: True stores as a recallable literal", _
           ok And lit2 = "True", "got '" & lit2 & "' ok=" & ok

    ' Round trip: yesterday's value rides into today's expression.
    b = VlaTryValueBuild(VlaTryExpandStars("(* *1 *1)", "7", "", "", prob))
    CheckFrags "l6: a recalled value rides into the next build", b, _
               Array("(return (* 7 7))")
End Sub

' ---------------------------------------------------------------------
'  REPLEVAL.0 pins: VlaEvalDisplay, eval's own display formatter
'  (VLA_DevRig.bas), pinned the same way VlaTryValueLit's L6 pins
'  just above hold that formatter. The eval Sub itself is a
'  Debug.Print wrapper over VlaEvalExpression (already pinned by
'  TestExprParity, host suite) - interactive verification is its
'  path, expand's own precedent; the formatter is the part a pin can
'  hold. No ByRef-out parameter here, so compound assertions carry
'  none of L6.1's in-expression-mutation hazard.
' ---------------------------------------------------------------------
Public Sub TestEvalDisplay()
    Report "repl-eval: numbers print with the locale-proof period", _
           VlaEvalDisplay(3.5) = "3.5" And VlaEvalDisplay(CLng(4)) = "4", _
           "got '" & VlaEvalDisplay(3.5) & "' / '" & VlaEvalDisplay(CLng(4)) & "'"
    Report "repl-eval: booleans print as VLA's own literals, lowercase", _
           VlaEvalDisplay(True) = "true" And VlaEvalDisplay(False) = "false", _
           "got '" & VlaEvalDisplay(True) & "' / '" & VlaEvalDisplay(False) & "'"
    Report "repl-eval: strings re-quote with the tokenizer's escapes", _
           VlaEvalDisplay("a" & Chr$(34) & "b\c") = Chr$(34) & "a\" & Chr$(34) & "b\\c" & Chr$(34), _
           "got '" & VlaEvalDisplay("a" & Chr$(34) & "b\c") & "'"
    Report "repl-eval: a numeric-looking string stays quoted, never a number", _
           VlaEvalDisplay("42") = Chr$(34) & "42" & Chr$(34), _
           "got '" & VlaEvalDisplay("42") & "'"
    Report "repl-eval: an array prints as a parenthesized list, recursively", _
           VlaEvalDisplay(Array(1, "x", Array(2, 3))) = "(1 ""x"" (2 3))", _
           "got '" & VlaEvalDisplay(Array(1, "x", Array(2, 3))) & "'"
    Report "repl-eval: Empty and Null print by name, objects as #<...>", _
           VlaEvalDisplay(Empty) = "Empty" And VlaEvalDisplay(Null) = "Null" And _
           VlaEvalDisplay(ThisWorkbook) = "#<Workbook>", _
           "got '" & VlaEvalDisplay(Empty) & "' / '" & VlaEvalDisplay(Null) & "' / '" & VlaEvalDisplay(ThisWorkbook) & "'"
End Sub

' ---------------------------------------------------------------------
'  L16 pins: the small-verbs batch. Expansion exactness through
'  VlaExpandText (the goldens-through-VlaExpand rule from the
'  definition), emission order full-stack through TryTranspile, and
'  the returning-in-a-Sub refusal BY ITS WORDS - the design leans on
'  EmitReturn's trappable teaching error, so the pin holds that door.
'  (dec! is not pinned here: L3 shipped it and its pin has lived in
'  TestPrelude since - L16 discovered the verb already on the bench.)
' ---------------------------------------------------------------------
Public Sub TestL16()
    Dim t As String, d As String
    Dim fired As Long

    t = VlaExpandText("(swap! p q t1)", True, fired)
    Report "l16: swap! expands to the classic three moves, dim included", _
           t = "(begin (dim t1) (set! t1 p) (set! p q) (set! q t1))" & vbCrLf, _
           "got: [" & Left$(t, 80) & "]"

    t = TryTranspile("l16: swap! emits the three moves in order", "(sub t () (swap! x y tmp))")
    If Len(t) > 0 Then
        Dim p1 As Long, p2 As Long, p3 As Long, p4 As Long
        p1 = InStr(1, t, "Dim tmp", vbTextCompare)
        p2 = InStr(1, t, "tmp = x", vbTextCompare)
        p3 = InStr(1, t, "x = y", vbTextCompare)
        p4 = InStr(1, t, "y = tmp", vbTextCompare)
        Report "l16: swap! emits the three moves in order", _
               p1 > 0 And p2 > p1 And p3 > p2 And p4 > p3, Left$(Norm(t), 160)
    End If

    t = VlaExpandText("(returning t1 (+ 1 2) (debug-print 1) (debug-print 2))", True, fired)
    Report "l16: returning saves first, splices the body, returns the saved", _
           t = "(begin (dim t1) (set! t1 (+ 1 2)) (debug-print 1) (debug-print 2) (return t1))" & vbCrLf, _
           "got: [" & Left$(t, 100) & "]"

    t = TryTranspile("l16: returning in a function computes, tidies, then returns", _
                     "(function f () Variant (returning tmp (+ 1 2) (debug-print 9)))")
    If Len(t) > 0 Then
        Dim q1 As Long, q2 As Long, q3 As Long, q4 As Long, q5 As Long
        q1 = InStr(1, t, "Dim tmp", vbTextCompare)
        q2 = InStr(1, t, "tmp = (1 + 2)", vbTextCompare)
        q3 = InStr(1, t, "Debug.Print 9", vbTextCompare)
        q4 = InStr(1, t, "f = tmp", vbTextCompare)
        q5 = InStr(1, t, "Exit Function", vbTextCompare)
        Report "l16: returning in a function computes, tidies, then returns", _
               q1 > 0 And q2 > q1 And q3 > q2 And q4 > q3 And q5 > q4, Left$(Norm(t), 160)
    End If

    d = ""
    On Error Resume Next
    t = VlaTranspile("(sub t () (returning tmp 1 (debug-print 1)))")
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    Report "l16: returning inside a Sub refuses with the emitter's words", _
           InStr(1, d, "only valid inside a function", vbTextCompare) > 0, "got: " & d

    t = TryTranspile("l16: swap! expands inside another template's spliced body", _
                     "(sub t () (when (= 1 1) (swap! x y tmp)))")
    If Len(t) > 0 Then
        CheckFrags "l16: swap! expands inside another template's spliced body", t, _
                   Array("(1 = 1)", "tmp = x")
    End If
End Sub

' ---------------------------------------------------------------------
'  L7 pins: (time ...) / (trace-form ...). Expansion through
'  VlaExpandText per the definition's goldens-through-VlaExpand rule
'  (CheckFrags rather than byte-exact - both expansions exceed the
'  pretty-printer's one-line width, and the frags plus order pins
'  hold the contract without pinning indentation); emission full-
'  stack through TryTranspile. The printed seconds themselves are
'  unpinnable for V6's recorded reason - Timer output is not a
'  constant - so the pins hold the scaffolding and the smoke in the
'  verification loop holds the dial.
' ---------------------------------------------------------------------
Public Sub TestL7()
    Dim t As String
    Dim fired As Long

    t = VlaExpandText("(time ""work"" t1 (debug-print 9))", True, fired)
    CheckFrags "l7: time expands to stamp, body, and the labeled print", t, _
               Array("(dim t1)", "(set! t1 (timer))", "(debug-print 9)", "(- (timer) t1)")
    Dim p1 As Long, p2 As Long, p3 As Long
    p1 = InStr(1, t, "(set! t1 (timer))", vbTextCompare)
    p2 = InStr(1, t, "(debug-print 9)", vbTextCompare)
    p3 = InStr(1, t, "(- (timer) t1)", vbTextCompare)
    Report "l7: time stamps BEFORE the body and reads AFTER it", _
           p1 > 0 And p2 > p1 And p3 > p2, Left$(Norm(t), 160)

    t = TryTranspile("l7: time emits the full VBA scaffolding", _
                     "(sub t () (time ""w"" t1 (debug-print 9)))")
    If Len(t) > 0 Then
        CheckFrags "l7: time emits the full VBA scaffolding", t, _
                   Array("Dim t1", "t1 = timer()", "Debug.Print 9", "(timer() - t1)")
    End If

    t = VlaExpandText("(trace-form ""step"" (debug-print 9))", True, fired)
    CheckFrags "l7: trace-form brackets the body with entry and exit", t, _
               Array(""">> """, "(debug-print 9)", """<< """)
    Dim q1 As Long, q2 As Long, q3 As Long
    q1 = InStr(1, t, ">> ", vbTextCompare)
    q2 = InStr(1, t, "(debug-print 9)", vbTextCompare)
    q3 = InStr(1, t, "<< ", vbTextCompare)
    Report "l7: trace-form announces entry before the body, exit after", _
           q1 > 0 And q2 > q1 And q3 > q2, Left$(Norm(t), 160)

    t = TryTranspile("l7: trace-form emits both bracket prints", _
                     "(sub t () (trace-form ""step"" (debug-print 9)))")
    If Len(t) > 0 Then
        CheckFrags "l7: trace-form emits both bracket prints", t, _
                   Array("("">> "" & ""step"")", "Debug.Print 9", "(""<< "" & ""step"")")
    End If

    t = VlaExpandText("(time ""w"" t2 (trace-form ""x"" (debug-print 1)))", True, fired)
    CheckFrags "l7: the littermates compose - trace inside time", t, _
               Array("(dim t2)", """>> """, "(debug-print 1)", "(- (timer) t2)")
End Sub

' ---------------------------------------------------------------------
'  L17 pins: defmacro docstrings. The doc must change NOTHING about
'  expansion (the invariance half) and must be readable through
'  VlaMacroDoc after the parse that defined it (the seam half - the
'  query flow L11 will ride). The lone-string rule is pinned from
'  both sides: with a following form the string is doc; alone it is
'  the template and the doc reads "".
' ---------------------------------------------------------------------
Public Sub TestL17()
    Dim t As String
    Dim fired As Long

    t = VlaExpandText("(defmacro (tw x) ""doubles x"" (* x 2))" & vbCrLf & "(tw 4)", True, fired)
    Report "l17: a documented macro expands exactly as an undocumented one", _
           t = "(* 4 2)" & vbCrLf, "got: [" & Left$(t, 60) & "]"
    Report "l17: the docstring reads back through VlaMacroDoc", _
           VlaMacroDoc("tw") = "doubles x", "got '" & VlaMacroDoc("tw") & "'"
    Report "l17: the doc lookup is case-insensitive like the macro table", _
           VlaMacroDoc("TW") = "doubles x", "got '" & VlaMacroDoc("TW") & "'"

    t = VlaExpandText("(defmacro (greet) ""hello"")" & vbCrLf & "(greet)", True, fired)
    Report "l17: a lone string is the TEMPLATE, not documentation", _
           InStr(1, t, "hello", vbTextCompare) > 0 And VlaMacroDoc("greet") = "", _
           "expansion [" & Left$(Norm(t), 60) & "] doc '" & VlaMacroDoc("greet") & "'"

    t = VlaExpandText("(defmacro (tw2 x) ""d"" (debug-print x) (debug-print x))" & vbCrLf & "(tw2 7)", True, fired)
    Report "l17: a doc plus a multi-form template begin-wraps the forms after the doc", _
           InStr(1, t, "(begin", vbTextCompare) > 0 And _
           InStr(1, t, "(debug-print 7)", vbTextCompare) > 0 And VlaMacroDoc("tw2") = "d", _
           "got: [" & Left$(Norm(t), 100) & "]"

    ' L11.2: this pin once borrowed "when" as its undocumented
    ' specimen - a premise, not a contract, and L11.1's prelude docs
    ' invalidated it (the pin failed by finding documentation - the
    ' right failure). The pin now defines its own bare macro.
    t = VlaExpandText("(defmacro (bare-mac x) (* x 1))" & vbCrLf & "(bare-mac 1)", True, fired)
    Report "l17: undocumented and unknown names both read as empty", _
           VlaMacroDoc("bare-mac") = "" And VlaMacroDoc("no-such-macro-zzz") = "", _
           "'" & VlaMacroDoc("bare-mac") & "' / '" & VlaMacroDoc("no-such-macro-zzz") & "'"

    Dim n As Long
    n = EnglishLoadVocabularyText("(defmacro (vt-doc x) ""vocab doc"" (* x 3))", "l17-vocab")
    Report "l17: a vocabulary-carried docstring survives the probe and reads back", _
           n = 0 And VlaMacroDoc("vt-doc") = "vocab doc", _
           "rules " & n & " doc '" & VlaMacroDoc("vt-doc") & "'"
    EnglishResetGrammar
End Sub

' ---------------------------------------------------------------------
'  L11 pins: (doc ...) / apropos. The doc form is pinned through its
'  EMISSION (the baked Debug.Print is the contract; the runtime print
'  is just VBA executing a literal); apropos through its pure half
'  VlaAproposText (VlaApropos only prints it - the VlaTimeIt
'  precedent for print-shaped tools). The manifest leg leans on the
'  dev workbook's live manifest, the same environmental fact the
'  resolve-check pins have leaned on since S4.
' ---------------------------------------------------------------------
Public Sub TestL11()
    Dim t As String
    Dim fired As Long

    t = TryTranspile("l11: (doc) bakes a documented macro's line into the program", _
                     "(defmacro (tw x) ""doubles x"" (* x 2))" & vbCrLf & "(sub t () (doc tw))")
    If Len(t) > 0 Then
        CheckFrags "l11: (doc) bakes a documented macro's line into the program", t, _
                   Array("Debug.Print ""tw: doubles x""")
    End If

    ' L11.2: "when" carries a doc since L11.1 - the pin supplies its
    ' own undocumented macro (the failed run's output was the doc
    ' form working perfectly on the newly documented prelude).
    t = TryTranspile("l11: (doc) says so plainly for an undocumented macro", _
                     "(defmacro (bare2 x) (* x 1))" & vbCrLf & "(sub t () (doc bare2))")
    If Len(t) > 0 Then
        CheckFrags "l11: (doc) says so plainly for an undocumented macro", t, _
                   Array("Debug.Print ""bare2: (no documentation)""")
    End If

    t = TryTranspile("l11: (doc) teaches parse scope for an unknown name", _
                     "(sub t () (doc zzz-nope))")
    If Len(t) > 0 Then
        CheckFrags "l11: (doc) teaches parse scope for an unknown name", t, _
                   Array("not a known macro in this parse")
    End If

    ' L11.1 contract: apropos answers the CARRIED bench, so the
    ' searchable macro arrives the way a user's would - by load.
    Dim nA As Long
    nA = EnglishLoadVocabularyText("(defmacro (aprop-target x) ""a searchable pearl"" (* x 1))", "l11-vocab-a")
    Report "l11: apropos matches on the NAME", _
           InStr(1, VlaAproposText("aprop-target"), "aprop-target - a searchable pearl", vbTextCompare) > 0, _
           Left$(VlaAproposText("aprop-target"), 120)
    Report "l11: apropos matches on the DOCSTRING too", _
           InStr(1, VlaAproposText("pearl"), "aprop-target", vbTextCompare) > 0, _
           Left$(VlaAproposText("pearl"), 120)
    Report "l11: an empty pattern lists the whole bench, prelude included", _
           InStr(1, VlaAproposText(""), "when - ", vbTextCompare) > 0 And _
           InStr(1, VlaAproposText(""), "swap! - ", vbTextCompare) > 0, _
           Left$(VlaAproposText(""), 160)
    Report "l11: apropos reaches the runtime helper manifest", _
           InStr(1, VlaAproposText("vlashowerror"), "vlashowerror - (runtime helper)", vbTextCompare) > 0, _
           Left$(VlaAproposText("vlashowerror"), 120)
    Report "l11: a miss answers with words, not silence", _
           InStr(1, VlaAproposText("qqqq-zzzz-none"), "nothing matches", vbTextCompare) > 0, _
           Left$(VlaAproposText("qqqq-zzzz-none"), 120)

    Dim n As Long
    n = EnglishLoadVocabularyText("(defmacro (vt3 x) ""vocab bench entry"" (* x 1))", "l11-vocab")
    Report "l11: load-then-ask sees a vocabulary-carried macro on the bench", _
           n = 0 And InStr(1, VlaAproposText("vt3"), "vt3 - vocab bench entry", vbTextCompare) > 0, _
           Left$(VlaAproposText("vt3"), 120)
    EnglishResetGrammar
End Sub

' ---------------------------------------------------------------------
'  L11.1 pins: the fix-pass the maiden smokes ordered. Apropos is now
'  state-INDEPENDENT (it re-parses prelude + the carried vocabulary
'  per ask), the prelude carries docs, and the two riders the ledger
'  promised to the next English pass are aboard: U.9's same-path
'  teaching and U.10's counters line.
' ---------------------------------------------------------------------
Public Sub TestL11_1()
    Dim n As Long, d As String, fired As Long, t As String

    n = EnglishLoadVocabularyText("(defmacro (vt4 x) ""fresh pearl"" (* x 1))", "l111-vocab")
    t = VlaExpandText("(zzz)", True, fired)   ' a STALE unrelated parse
    Report "l111: apropos survives an unrelated parse - the carry holds the bench", _
           InStr(1, VlaAproposText("vt4"), "vt4 - fresh pearl", vbTextCompare) > 0, _
           Left$(VlaAproposText("vt4"), 120)
    EnglishResetGrammar
    Report "l111: reset clears the carry - the bench is prelude-only again", _
           InStr(1, VlaAproposText("vt4"), "nothing matches", vbTextCompare) > 0, _
           Left$(VlaAproposText("vt4"), 120)

    Report "l111: the prelude now documents itself", _
           InStr(1, VlaAproposText("swap!"), "swap two values", vbTextCompare) > 0 And _
           InStr(1, VlaAproposText(""), "when - run the body when the test is true", vbTextCompare) > 0, _
           Left$(VlaAproposText("swap!"), 120)

    d = ""
    n = EnglishLoadVocabularyText("(defmacro (u9m x) (* x 1))", "u9-file")
    On Error Resume Next
    n = EnglishLoadVocabularyText("(defmacro (u9m x) (* x 1))", "u9-file")
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    Report "u9: a same-path re-carry teaches reset-or-Reload, not rename", _
           InStr(1, d, "already carries this file", vbTextCompare) > 0 And _
           InStr(1, d, "Reload", vbTextCompare) > 0, "got: " & d
    d = ""
    On Error Resume Next
    n = EnglishLoadVocabularyText("(defmacro (u9m x) (* x 1))", "u9-other")
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    Report "u9: a cross-file collision keeps the pinned rename wording", _
           InStr(1, d, "already carried by u9-file", vbTextCompare) > 0 And _
           InStr(1, d, "rename this one", vbTextCompare) > 0, "got: " & d
    EnglishResetGrammar

    n = EnglishLoadVocabularyText("(english-vla ""hop"" (debug-print 1))" & vbCrLf & _
                                  "(test-success ""Hop."" (debug-print 1))" & vbCrLf & _
                                  "(defmacro (u10m x) (* x 1))", "stats-test")
    Report "u10: the counters line is one honest copy", _
           EnglishVocabStats() = "loaded: 1 rule, 1 macro, 1 test (0 expected fails) from stats-test", _
           "got: " & EnglishVocabStats()

    ' L11.3: the miss line distinguishes "no match" from "nothing
    ' loaded" (the S5.1 wipe finding from the on-machine smokes).
    Report "l113: a miss while a vocabulary is carried does NOT claim emptiness", _
           InStr(1, VlaAproposText("qq-none-zz"), "no vocabulary is carried", vbTextCompare) = 0 And _
           InStr(1, VlaAproposText("qq-none-zz"), "nothing matches", vbTextCompare) > 0, _
           Left$(VlaAproposText("qq-none-zz"), 160)
    EnglishResetGrammar
    Report "l113: a miss on an empty bench teaches the wipe and the Reload", _
           InStr(1, VlaAproposText("qq-none-zz"), "no vocabulary is carried", vbTextCompare) > 0, _
           Left$(VlaAproposText("qq-none-zz"), 200)
End Sub

' ---------------------------------------------------------------------
'  APROPOSPLUS: the whole-stack tier - a loaded rule's own English
'  pattern is now part of what apropos searches, alongside its macro
'  (L11's own tier). "aprop-glow"/"a searchable pearl, second edition"
'  never appear in "illuminate cell {r:cell}", and "illuminate" never
'  appears in the macro name or its docstring - the fixture is built
'  so each assertion can only pass through the specific direction it
'  claims to prove.
' ---------------------------------------------------------------------
Public Sub TestAproposPlus()
    Dim n As Long
    n = EnglishLoadVocabularyText( _
        "(defmacro (aprop-glow r) ""a searchable pearl, second edition"" (set! (range r) 1))" & vbCrLf & _
        "(english-vla ""illuminate cell {r:cell}"" (aprop-glow {r}))", "aproposplus-vocab")

    Report "aproposplus: a search hits the ENGLISH SENTENCE, not just the macro", _
           InStr(1, VlaAproposText("illuminate"), "illuminate cell {r:cell}", vbTextCompare) > 0, _
           Left$(VlaAproposText("illuminate"), 160)

    Report "aproposplus: searching the macro name also surfaces the sentence that reaches it", _
           InStr(1, VlaAproposText("aprop-glow"), "illuminate cell {r:cell}", vbTextCompare) > 0, _
           Left$(VlaAproposText("aprop-glow"), 200)

    Report "aproposplus: a sentence hit shows what it calls, tagged as a sentence", _
           InStr(1, VlaAproposText("illuminate"), "aprop-glow", vbTextCompare) > 0 And _
           InStr(1, VlaAproposText("illuminate"), "(sentence)", vbTextCompare) > 0, _
           Left$(VlaAproposText("illuminate"), 200)

    EnglishResetGrammar
    Report "aproposplus: reset clears the sentence tier the same call clears the macro tier", _
           InStr(1, VlaAproposText("illuminate"), "illuminate cell", vbTextCompare) = 0, _
           Left$(VlaAproposText("illuminate"), 160)
End Sub

' ---------------------------------------------------------------------
'  L8 pins: the check family (expansion + full-stack emission - the
'  RUN of a failing check is the smoke's job, since raising in-suite
'  would require injection) and the operator-statement guard, the
'  owner-found rider: (+ 1 2) as a statement now refuses at
'  transpile time with words instead of meeting VBA's modal.
' ---------------------------------------------------------------------
Public Sub TestL8()
    Dim t As String, d As String
    Dim fired As Long

    t = VlaExpandText("(check ""ok"" (= 1 1))", True, fired)
    CheckFrags "l8: check expands to if-not-raise with the label", t, _
               Array("(if (not (= 1 1))", "(err.raise 5 ""check""", "check failed: ")

    t = VlaExpandText("(check= ""math"" 4 (+ 2 2))", True, fired)
    CheckFrags "l8: check= compares and carries both sides into the message", t, _
               Array("(<> 4 (+ 2 2))", "expected", "but got")
    Dim p1 As Long, p2 As Long
    p1 = InStr(1, t, "expected", vbTextCompare)
    p2 = InStr(1, t, "but got", vbTextCompare)
    Report "l8: the failure message reads expected-then-got", _
           p1 > 0 And p2 > p1, Left$(Norm(t), 160)

    ' L8.1: the raise rides the generic call path - the emitter
    ' renders (err.raise ...) as "Call err.raise(5, ...)", legal VBA
    ' that raises identically (with-fast-excel has emitted it so all
    ' along). The original frag pinned an ASSUMED statement-style
    ' emission; the lesson joins the ledger: pin frags come from
    ' verified emission, not from VBA idiom.
    t = TryTranspile("l8: a check= transpiles to a guarded raise", _
                     "(sub t () (check= ""math"" 4 (+ 2 2)))")
    If Len(t) > 0 Then
        CheckFrags "l8: a check= transpiles to a guarded raise", t, _
                   Array("(4 <> (2 + 2))", "err.raise", "check failed: ")
    End If

    Report "l8: the check family sits documented on the bench", _
           InStr(1, VlaAproposText("check"), "check - raise unless the condition holds", vbTextCompare) > 0 And _
           InStr(1, VlaAproposText("check"), "check= - raise unless expected equals", vbTextCompare) > 0, _
           Left$(VlaAproposText("check"), 160)

    d = ""
    On Error Resume Next
    t = VlaTranspile("(sub t () (+ 1 2))")
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    Report "l8r: an operator in statement position refuses with directions", _
           InStr(1, d, "expression, not a statement", vbTextCompare) > 0 And _
           InStr(1, d, "VlaTryValue", vbTextCompare) > 0, "got: " & d

    t = TryTranspile("l8r: ordinary call statements still emit Call", _
                     "(sub t () (vlacount x))")
    If Len(t) > 0 Then
        CheckFrags "l8r: ordinary call statements still emit Call", t, _
                   Array("Call vlacount(x)")
    End If
End Sub

' ---------------------------------------------------------------------
'  L8.2 pins: the scratch self-handler (owner-found at the check
'  smoke - a failing check met VBA's dialog because runtime errors
'  never cross Application.Run; the scratch now catches and speaks
'  its own). Pins hold the pure builders' new shape and its runnable
'  transpile; the dialog's retirement is the smoke's job.
' ---------------------------------------------------------------------
Public Sub TestL82()
    Dim b As String, t As String
    Dim p1 As Long, p2 As Long, p3 As Long, p4 As Long

    b = VlaTryBuild("(debug-print 1)")
    p1 = InStr(1, b, "(on-error goto vla-try-oops)", vbTextCompare)
    p2 = InStr(1, b, "(debug-print 1)", vbTextCompare)
    p3 = InStr(1, b, "(exit-sub)", vbTextCompare)
    p4 = InStr(1, b, "(label vla-try-oops)", vbTextCompare)
    Report "l82: the scratch arms its handler before the body and exits before the label", _
           p1 > 0 And p2 > p1 And p3 > p2 And p4 > p3 And _
           InStr(1, b, "scratch error: ", vbTextCompare) > p4, Left$(Norm(b), 160)

    Report "l82: an empty scratch stays bare - nothing can error in it", _
           InStr(1, VlaTryBuild(""), "(sub vla-scratch ())", vbTextCompare) > 0 And _
           InStr(1, VlaTryBuild(""), "vla-try-oops", vbTextCompare) = 0, _
           Left$(Norm(VlaTryBuild("")), 80)

    b = VlaTryValueBuild("(+ 1 2)")
    p1 = InStr(1, b, "(on-error goto vla-try-oops)", vbTextCompare)
    p2 = InStr(1, b, "(return (+ 1 2))", vbTextCompare)
    p3 = InStr(1, b, "(label vla-try-oops)", vbTextCompare)
    Report "l82: the value scratch arms, returns on success, and speaks on error", _
           p1 > 0 And p2 > p1 And p3 > p2, Left$(Norm(b), 160)

    t = TryTranspile("l82: the wrapped scratch transpiles runnable", VlaTryBuild("(check ""x"" (= 1 1))"))
    If Len(t) > 0 Then
        CheckFrags "l82: the wrapped scratch transpiles runnable", t, _
                   Array("On Error GoTo vla_try_oops", "Exit Sub", "vla_try_oops:", "scratch error: ")
    End If
End Sub

' ---------------------------------------------------------------------
'  L12 pins: VlaFormat (the L1 writer as canonical formatter - flat
'  normalization, the width-90 split, and above all NO EXPANSION)
'  and the balance hints (the pure count in both directions with the
'  tokenizer's own escape rules, the pre-probe raw-validation hint,
'  and the L0 capture's own count - shipped back then, pinned now).
' ---------------------------------------------------------------------
Public Sub TestL12()
    Dim t As String, d As String

    Report "l12: format normalizes spacing to the canonical flat form", _
           VlaFormat("(set!   x    1)") = "(set! x 1)" & vbCrLf, _
           "got: [" & Left$(VlaFormat("(set!   x    1)"), 60) & "]"

    Report "l12: format writes one top form per line", _
           VlaFormat("(a 1) (b 2)") = "(a 1)" & vbCrLf & "(b 2)" & vbCrLf, _
           "got: [" & Left$(Norm(VlaFormat("(a 1) (b 2)")), 80) & "]"

    Dim a As String, b As String, c As String
    a = String$(30, "a")
    b = String$(30, "b")
    c = String$(30, "c")
    t = VlaFormat("(seq " & a & " " & b & " " & c & ")")
    Report "l12: format opens a wide form exactly as the L1 writer does", _
           t = "(seq" & vbCrLf & "    " & a & vbCrLf & "    " & b & vbCrLf & _
               "    " & c & ")" & vbCrLf, "got: [" & Left$(t, 80) & "]"

    Report "l12: format NEVER expands - what you wrote is what formats", _
           InStr(1, VlaFormat("(when (= 1 1) (debug-print 1))"), "(when", vbTextCompare) > 0 And _
           InStr(1, VlaFormat("(when (= 1 1) (debug-print 1))"), "(if", vbTextCompare) = 0, _
           "got: [" & Left$(Norm(VlaFormat("(when (= 1 1) (debug-print 1))")), 80) & "]"

    d = ""
    On Error Resume Next
    t = VlaFormat("(f (g 1)")
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    Report "l12: a malformed source formats into the hinted refusal", _
           InStr(1, d, "missing ')'", vbTextCompare) > 0 And _
           InStr(1, d, "opened at", vbTextCompare) > 0, "got: " & d

    Report "l12: the balance hint counts both directions and stays quiet when balanced", _
           VlaBalanceHint("(a (b") = "this row opens 2 forms it never closes" And _
           VlaBalanceHint("(a) )") = "this row closes 1 form it never opened" And _
           VlaBalanceHint("(a (b 1))") = "", _
           "'" & VlaBalanceHint("(a (b") & "' / '" & VlaBalanceHint("(a) )") & "'"

    Report "l12: the hint honors string escapes and comments like the tokenizer", _
           VlaBalanceHint("(f ""a\"")b("")" & " ; (open") = "" And _
           VlaBalanceHint("(f ""("")") = "", _
           "'" & VlaBalanceHint("(f ""("")") & "'"

    d = ""
    On Error Resume Next
    t = EnglishToVla("(zz (yy 1)")
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    Report "l12: an unbalanced raw row refuses with the count (the L0 hint, pinned at last)", _
           InStr(1, d, "still open", vbTextCompare) > 0 Or InStr(1, d, "never closes", vbTextCompare) > 0, _
           "got: " & d
End Sub

' ---------------------------------------------------------------------
'  L15 pins: the macro-stepper. Determinism is the load-bearing
'  claim - the walk order is fixed, so the first N applications of
'  any two runs are the SAME N - and the pins hold it frame by
'  frame on a two-application source whose intermediate state is
'  fully predictable: (when ...) wrapping (inc! ...).
' ---------------------------------------------------------------------
Public Sub TestL15()
    Dim t As String
    Dim total As Long, f As Long
    Dim srcW As String
    srcW = "(when (= 1 1) (inc! x))"

    t = VlaExpandStepText(srcW, 99, total)
    Report "l15: the odometer totals two applications for when-around-inc!", _
           total = 2, "total " & total

    t = VlaExpandStepText(srcW, 0, total)
    Report "l15: step 0 is the parse as written - nothing applied", _
           InStr(1, t, "(when", vbTextCompare) > 0 And InStr(1, t, "(inc!", vbTextCompare) > 0, _
           Left$(Norm(t), 120)

    t = VlaExpandStepText(srcW, 1, total)
    Report "l15: step 1 applies the outer macro and ONLY the outer", _
           InStr(1, t, "(if", vbTextCompare) > 0 And InStr(1, t, "(when", vbTextCompare) = 0 And _
           InStr(1, t, "(inc!", vbTextCompare) > 0, Left$(Norm(t), 120)

    t = VlaExpandStepText(srcW, 2, total)
    Report "l15: step 2 is the fixpoint, byte-equal to the full expansion", _
           t = VlaExpandText(srcW, True, f) And InStr(1, t, "(inc!", vbTextCompare) = 0 And _
           InStr(1, t, "(+ x 1)", vbTextCompare) > 0, Left$(Norm(t), 120)

    Report "l15: a step past the total is the fixpoint, and the total still reports", _
           VlaExpandStepText(srcW, 99, total) = VlaExpandText(srcW, True, f) And total = 2, _
           "total " & total

    t = VlaExpandStepText("(inc! a) (dec! b)", 1, total)
    Report "l15: the budget spends in walk order - first form first", _
           InStr(1, t, "(+ a 1)", vbTextCompare) > 0 And InStr(1, t, "(dec! b)", vbTextCompare) > 0, _
           Left$(Norm(t), 120)

    t = VlaExpandText(srcW, True, f)
    Report "l15: the stepper leaves no budget residue - a plain expansion runs full", _
           f = 2 And InStr(1, t, "(inc!", vbTextCompare) = 0, "fired " & f & " [" & Left$(Norm(t), 80) & "]"
End Sub

' ---------------------------------------------------------------------
'  P.L1 pins: the predicate prelude (zero? positive? negative? even?
'  odd? empty? blank?). Each expansion fits the pretty-printer's one
'  line, so the seven are pinned byte-exact through VlaExpandText
'  (the goldens-through-VlaExpand rule); emission is full-stack
'  through TryTranspile with frags derived from the emitter's own
'  paths (EmitChain parenthesizes and mod displays as Mod; generic
'  calls render name(args) - the L8.1 rule honored by reading the
'  code paths, since the assistant writes blind); and apropos holds
'  the docstrings from birth - a name-hit and a docstring-hit, on
'  this pass's OWN specimens per the borrowed-specimen lesson. The
'  condition-position pin is the load-bearing one: predicates are
'  the prelude's first expression-shaped macros, and the pin holds
'  the walk's rebuild-every-element behavior that makes them legal.
' ---------------------------------------------------------------------
Public Sub TestPL1()
    Dim t As String
    Dim fired As Long

    t = VlaExpandText("(zero? p)", True, fired)
    Report "pl1: zero? expands to the = comparison", _
           t = "(= p 0)" & vbCrLf, "got: [" & Left$(t, 80) & "]"

    t = VlaExpandText("(positive? p)", True, fired)
    Report "pl1: positive? expands to the > comparison", _
           t = "(> p 0)" & vbCrLf, "got: [" & Left$(t, 80) & "]"

    t = VlaExpandText("(negative? p)", True, fired)
    Report "pl1: negative? expands to the < comparison", _
           t = "(< p 0)" & vbCrLf, "got: [" & Left$(t, 80) & "]"

    t = VlaExpandText("(even? p)", True, fired)
    Report "pl1: even? expands through mod", _
           t = "(= (mod p 2) 0)" & vbCrLf, "got: [" & Left$(t, 80) & "]"

    t = VlaExpandText("(odd? p)", True, fired)
    Report "pl1: odd? expands through mod, negated", _
           t = "(<> (mod p 2) 0)" & vbCrLf, "got: [" & Left$(t, 80) & "]"

    t = VlaExpandText("(empty? p)", True, fired)
    Report "pl1: empty? expands to isempty", _
           t = "(isempty p)" & vbCrLf, "got: [" & Left$(t, 80) & "]"

    t = VlaExpandText("(blank? p)", True, fired)
    Report "pl1: blank? expands to the len-trim-coerce test", _
           t = "(= (len (trim (& p """"))) 0)" & vbCrLf, "got: [" & Left$(t, 100) & "]"

    t = TryTranspile("pl1: a predicate expands in condition position", _
                     "(sub t () (if (even? x) (then (debug-print 1))))")
    If Len(t) > 0 Then
        CheckFrags "pl1: a predicate expands in condition position", t, _
                   Array("If ((x Mod 2) = 0) Then")
    End If

    t = TryTranspile("pl1: a predicate composes inside another template's slot", _
                     "(sub t () (when (blank? v) (debug-print 1)))")
    If Len(t) > 0 Then
        CheckFrags "pl1: a predicate composes inside another template's slot", t, _
                   Array("If (len(trim((v & """"))) = 0) Then")
    End If

    Report "pl1: the predicates carry their docstrings on the bench", _
           InStr(1, VlaAproposText("zero?"), "zero? - true when the value is 0", vbTextCompare) > 0 And _
           InStr(1, VlaAproposText("whitespace"), "blank?", vbTextCompare) > 0, _
           Left$(VlaAproposText("zero?"), 120)
End Sub

' ---------------------------------------------------------------------
'  P.L2 pins: when-let / if-let, the binding-conditionals. Fixpoint
'  expansions pinned byte-exact through VlaExpandText on short
'  specimens (each fits the printer's one line) - the fixpoint also
'  proves the P-layer's own layering, since both templates consume
'  P.L1's empty? and the pinned text must show isempty, one
'  expansion deeper. Emission full-stack through TryTranspile with
'  InStr position ordering (the L16 pattern); frags from verified
'  emission paths: (dim x) emits Dim x As Variant (the let-one pin's
'  own frag), (not (isempty w)) emits (Not isempty(w)).
' ---------------------------------------------------------------------
Public Sub TestPL2()
    Dim t As String
    Dim fired As Long

    t = VlaExpandText("(when-let w q (debug-print w))", True, fired)
    Report "pl2: when-let expands to dim, set, and the guarded body", _
           t = "(begin (dim w) (set! w q) (if (not (isempty w)) (then (debug-print w))))" & vbCrLf, _
           "got: [" & Left$(t, 110) & "]"

    t = VlaExpandText("(when-let w q (f) (g))", True, fired)
    Report "pl2: when-let splices a multi-statement body", _
           t = "(begin (dim w) (set! w q) (if (not (isempty w)) (then (f) (g))))" & vbCrLf, _
           "got: [" & Left$(t, 110) & "]"

    t = TryTranspile("pl2: when-let emits dim, set, guard, and body in order", _
                     "(sub t () (when-let w q (debug-print w)))")
    If Len(t) > 0 Then
        Dim p1 As Long, p2 As Long, p3 As Long, p4 As Long, p5 As Long
        p1 = InStr(1, t, "Dim w As Variant", vbTextCompare)
        p2 = InStr(1, t, "w = q", vbTextCompare)
        p3 = InStr(1, t, "If (Not isempty(w)) Then", vbTextCompare)
        p4 = InStr(1, t, "Debug.Print w", vbTextCompare)
        p5 = InStr(1, t, "End If", vbTextCompare)
        Report "pl2: when-let emits dim, set, guard, and body in order", _
               p1 > 0 And p2 > p1 And p3 > p2 And p4 > p3 And p5 > p4, Left$(Norm(t), 160)
    End If

    t = VlaExpandText("(if-let w q (f) (g))", True, fired)
    Report "pl2: if-let expands with both branches in place", _
           t = "(begin (dim w) (set! w q) (if (not (isempty w)) (then (f)) (else (g))))" & vbCrLf, _
           "got: [" & Left$(t, 110) & "]"

    t = TryTranspile("pl2: if-let emits If, Else, End If with the branches in order", _
                     "(sub t () (if-let w q (debug-print 1) (debug-print 2)))")
    If Len(t) > 0 Then
        Dim q1 As Long, q2 As Long, q3 As Long, q4 As Long, q5 As Long
        q1 = InStr(1, t, "If (Not isempty(w)) Then", vbTextCompare)
        q2 = InStr(1, t, "Debug.Print 1", vbTextCompare)
        q3 = InStr(1, t, "Else", vbTextCompare)
        q4 = InStr(1, t, "Debug.Print 2", vbTextCompare)
        q5 = InStr(1, t, "End If", vbTextCompare)
        Report "pl2: if-let emits If, Else, End If with the branches in order", _
               q1 > 0 And q2 > q1 And q3 > q2 And q4 > q3 And q5 > q4, Left$(Norm(t), 160)
    End If

    Report "pl2: the binding-conditionals carry their docstrings on the bench", _
           InStr(1, VlaAproposText("when-let"), "when-let - bind a value and run the body only when it is not empty", vbTextCompare) > 0 And _
           InStr(1, VlaAproposText("if-let"), "if-let - bind a value and pick the branch by whether it is empty", vbTextCompare) > 0, _
           Left$(VlaAproposText("when-let"), 120)
End Sub

' ---------------------------------------------------------------------
'  P.L3 pins: try-else, the ignore-errors-with-default binding. The
'  full expansion exceeds the printer's one-line width, so the
'  expansion pin is frags-plus-order (the L7 precedent, reason
'  recorded) - the inner if-form DOES fit one child line at its
'  indent, so it pins as a single flat frag, clear-then-fall-back
'  order inside it and all. Emission full-stack with InStr position
'  ordering; frags from verified paths: (on-error resume-next) ->
'  "On Error Resume Next", (on-error goto 0) -> "On Error GoTo 0",
'  (err.clear) rides the generic call path -> "Call err.clear"
'  (zero args, no parens - the L8.1 rule honored by reading
'  EmitCallStmt, not VBA idiom). The two-try-else pin holds the
'  re-arm behavior that makes the form repeatable in one procedure.
' ---------------------------------------------------------------------
Public Sub TestPL3()
    Dim t As String
    Dim fired As Long

    t = VlaExpandText("(try-else w q 0)", True, fired)
    Dim e1 As Long, e2 As Long, e3 As Long, e4 As Long, e5 As Long
    e1 = InStr(1, t, "(dim w)", vbTextCompare)
    e2 = InStr(1, t, "(on-error resume-next)", vbTextCompare)
    e3 = InStr(1, t, "(set! w q)", vbTextCompare)
    e4 = InStr(1, t, "(if (<> err.number 0) (then (err.clear) (set! w 0)))", vbTextCompare)
    e5 = InStr(1, t, "(on-error goto 0)", vbTextCompare)
    Report "pl3: try-else expands to dim, arm, set, check-and-fall-back, disarm in order", _
           e1 > 0 And e2 > e1 And e3 > e2 And e4 > e3 And e5 > e4, _
           "got: [" & Left$(Norm(t), 160) & "]"

    t = TryTranspile("pl3: try-else emits the six moves in order", _
                     "(sub t () (try-else w q 0))")
    If Len(t) > 0 Then
        Dim p1 As Long, p2 As Long, p3 As Long, p4 As Long, p5 As Long, p6 As Long, p7 As Long
        p1 = InStr(1, t, "Dim w As Variant", vbTextCompare)
        p2 = InStr(1, t, "On Error Resume Next", vbTextCompare)
        p3 = InStr(1, t, "w = q", vbTextCompare)
        p4 = InStr(1, t, "If (err.number <> 0) Then", vbTextCompare)
        p5 = InStr(1, t, "Call err.clear", vbTextCompare)
        p6 = InStr(1, t, "w = 0", vbTextCompare)
        p7 = InStr(1, t, "On Error GoTo 0", vbTextCompare)
        Report "pl3: try-else emits the six moves in order", _
               p1 > 0 And p2 > p1 And p3 > p2 And p4 > p3 And p5 > p4 And p6 > p5 And p7 > p6, _
               Left$(Norm(t), 200)
    End If

    t = TryTranspile("pl3: both slots substitute whole forms - the P-layer rides the expr", _
                     "(sub t () (try-else w (even? (zz)) (+ 1 2)))")
    If Len(t) > 0 Then
        CheckFrags "pl3: both slots substitute whole forms - the P-layer rides the expr", t, _
                   Array("w = ((zz() Mod 2) = 0)", "w = (1 + 2)")
    End If

    t = TryTranspile("pl3: a second try-else re-arms after the first disarms", _
                     "(sub t () (try-else a (f) 1) (try-else b (g) 2))")
    If Len(t) > 0 Then
        Dim d1 As Long, r2 As Long, d2 As Long
        d1 = InStr(1, t, "On Error GoTo 0", vbTextCompare)
        r2 = InStr(d1 + 1, t, "On Error Resume Next", vbTextCompare)
        d2 = InStr(r2 + 1, t, "On Error GoTo 0", vbTextCompare)
        Report "pl3: a second try-else re-arms after the first disarms", _
               d1 > 0 And r2 > d1 And d2 > r2, Left$(Norm(t), 200)
    End If

    Report "pl3: try-else carries its docstring on the bench", _
           InStr(1, VlaAproposText("try-else"), "try-else - bind an expression's value, or a fallback when it errors", vbTextCompare) > 0, _
           Left$(VlaAproposText("try-else"), 120)
End Sub

' ---------------------------------------------------------------------
'  P.L4 pins: keyword arguments. The general :kw-value mechanism
'  already lived in EmitArgs (the honest-scope finding on the
'  ledger) with exactly ONE pin (the dot-member :type case in
'  TestEmitter) - these pins complete the coverage across the other
'  two call paths, the manglings, and the template ride-through,
'  then hold the NEW guard family: a keyword token outside an
'  argument list refuses at transpile time with words instead of
'  emitting broken VBA into the untrappable compile modal (the L8
'  rider's pattern, applied to the ':' shape). The missing-value
'  refusal predates this pass unpinned - pinned now, L12's
'  "shipped then, pinned now" precedent.
' ---------------------------------------------------------------------
Public Sub TestPL4()
    Dim t As String, d As String

    t = TryTranspile("pl4: keyword args reach expression-position calls", _
                     "(sub t () (set! x (f :color 1)))")
    If Len(t) > 0 Then
        CheckFrags "pl4: keyword args reach expression-position calls", t, _
                   Array("x = f(color:=1)")
    End If

    t = TryTranspile("pl4: keyword args reach statement-position calls", _
                     "(sub t () (f :color 1))")
    If Len(t) > 0 Then
        CheckFrags "pl4: keyword args reach statement-position calls", t, _
                   Array("Call f(color:=1)")
    End If

    t = TryTranspile("pl4: positional and keyword arguments mix in one call", _
                     "(sub t () (g 1 :b 2))")
    If Len(t) > 0 Then
        CheckFrags "pl4: positional and keyword arguments mix in one call", t, _
                   Array("Call g(1, b:=2)")
    End If

    t = TryTranspile("pl4: keyword names mangle like every identifier", _
                     "(sub t () (f :fill-color 3))")
    If Len(t) > 0 Then
        CheckFrags "pl4: keyword names mangle like every identifier", t, _
                   Array("fill_color:=3")
    End If

    t = TryTranspile("pl4: a keyword rides a template parameter into named emission", _
                     "(sub t () (when-let w (f :k 1) (g)))")
    If Len(t) > 0 Then
        CheckFrags "pl4: a keyword rides a template parameter into named emission", t, _
                   Array("w = f(k:=1)")
    End If

    d = ""
    On Error Resume Next
    t = VlaTranspile("(sub t () (f :color))")
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    Report "pl4: a keyword missing its value refuses with words", _
           InStr(1, d, "missing its value", vbTextCompare) > 0, "got: " & d

    d = ""
    On Error Resume Next
    t = VlaTranspile("(sub t () (set! x :color))")
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    Report "pl4: a keyword standing alone as a value refuses with words", _
           InStr(1, d, "keyword argument token", vbTextCompare) > 0 And _
           InStr(1, d, "a value on its own", vbTextCompare) > 0, "got: " & d

    d = ""
    On Error Resume Next
    t = VlaTranspile("(sub t () (:color 1))")
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    Report "pl4: a keyword heading a statement refuses with words", _
           InStr(1, d, "the head of a statement", vbTextCompare) > 0, "got: " & d

    d = ""
    On Error Resume Next
    t = VlaTranspile("(sub t () (set! x (:color 1)))")
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    Report "pl4: a keyword heading an expression refuses with words", _
           InStr(1, d, "the head of an expression", vbTextCompare) > 0, "got: " & d

    d = ""
    On Error Resume Next
    t = VlaTranspile("(sub t () (. ws :color 1))")
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    Report "pl4: a keyword in the dot form's member slot refuses with words", _
           InStr(1, d, "a member name after '.'", vbTextCompare) > 0, "got: " & d
End Sub

' ---------------------------------------------------------------------
'  P.L5 pins: (quote ...) - the data literal. Emission full-stack
'  through TryTranspile (frags are exact substrings of the one-line
'  Array literal); the two refusals and the reserved-name guard by
'  their words; and the two expansion-boundary facts that make
'  quote a real data context: the walkers never expand a quote's
'  interior (a quoted macro NAME stays a string), while template
'  parameters DO substitute inside a quoted template body (the
'  data-parameterization privilege). Expansion-side pins ride
'  VlaExpandText; the suppression pin's specimen deliberately
'  quotes "when" - the prelude's own most-expanded name.
' ---------------------------------------------------------------------
Public Sub TestPL5()
    Dim t As String, d As String
    Dim fired As Long

    t = TryTranspile("pl5: a quoted symbol list becomes an Array of strings", _
                     "(sub t () (set! x (quote (red green blue))))")
    If Len(t) > 0 Then
        CheckFrags "pl5: a quoted symbol list becomes an Array of strings", t, _
                   Array("x = Array(""red"", ""green"", ""blue"")")
    End If

    t = TryTranspile("pl5: numbers and strings keep their reader classes", _
                     "(sub t () (set! x (quote (1 2.5 ""txt""))))")
    If Len(t) > 0 Then
        CheckFrags "pl5: numbers and strings keep their reader classes", t, _
                   Array("x = Array(1, 2.5, ""txt"")")
    End If

    t = TryTranspile("pl5: nested lists nest Arrays", _
                     "(sub t () (set! x (quote ((1 2) (3 4)))))")
    If Len(t) > 0 Then
        CheckFrags "pl5: nested lists nest Arrays", t, _
                   Array("x = Array(Array(1, 2), Array(3, 4))")
    End If

    t = TryTranspile("pl5: bare atoms quote by class - symbol to string, number to number", _
                     "(sub t () (set! x (quote red)) (set! y (quote 5)))")
    If Len(t) > 0 Then
        CheckFrags "pl5: bare atoms quote by class - symbol to string, number to number", t, _
                   Array("x = ""red""", "y = 5")
    End If

    t = TryTranspile("pl5: quoted data is not identifiers - hyphens and case survive", _
                     "(sub t () (set! x (quote (hot-pink Red))))")
    If Len(t) > 0 Then
        CheckFrags "pl5: quoted data is not identifiers - hyphens and case survive", t, _
                   Array("x = Array(""hot-pink"", ""Red"")")
    End If

    t = TryTranspile("pl5: keyword tokens are data inside a quote", _
                     "(sub t () (set! x (quote (:a 1))))")
    If Len(t) > 0 Then
        CheckFrags "pl5: keyword tokens are data inside a quote", t, _
                   Array("x = Array("":a"", 1)")
    End If

    t = TryTranspile("pl5: the empty list quotes to an empty Array", _
                     "(sub t () (set! x (quote ())))")
    If Len(t) > 0 Then
        CheckFrags "pl5: the empty list quotes to an empty Array", t, _
                   Array("x = Array()")
    End If

    t = TryTranspile("pl5: expansion never enters a quote - a macro name stays a string", _
                     "(sub t () (set! x (quote (when 1 (f)))))")
    If Len(t) > 0 Then
        CheckFrags "pl5: expansion never enters a quote - a macro name stays a string", t, _
                   Array("x = Array(""when"", 1, Array(""f""))")
    End If

    t = TryTranspile("pl5: template parameters substitute inside a quoted template body", _
                     "(defmacro (pal c) (set! x (quote (c blue)))) (sub t () (pal red))")
    If Len(t) > 0 Then
        CheckFrags "pl5: template parameters substitute inside a quoted template body", t, _
                   Array("x = Array(""red"", ""blue"")")
    End If

    d = ""
    On Error Resume Next
    t = VlaTranspile("(sub t () (set! x (quote a b)))")
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    Report "pl5: a multi-datum quote refuses with words", _
           InStr(1, d, "exactly one datum", vbTextCompare) > 0, "got: " & d

    d = ""
    On Error Resume Next
    t = VlaTranspile("(sub t () (quote (a)))")
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    Report "pl5: quote in statement position refuses with words and directions", _
           InStr(1, d, "'(quote ...)' is an expression, not a statement", vbTextCompare) > 0, "got: " & d

    d = ""
    On Error Resume Next
    t = VlaTranspile("(defmacro (quote x) (f x)) (sub t () (debug-print 1))")
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    Report "pl5: defmacro refuses the name quote with words", _
           InStr(1, d, "cannot be a macro name", vbTextCompare) > 0, "got: " & d
End Sub

' ---------------------------------------------------------------------
'  P.L6 pins: procedure docstrings - L17's machinery extended to
'  (sub ...) / (function ...). The invariance pin is byte-EQUALITY
'  of a documented and an undocumented transpile (one-line
'  specimens, so source-map tags agree and the comparison is
'  exact); the function pins hold the return-type sniff fix (a
'  string at position 4 is never a type - before this pass it was
'  silently mangled into one); the lone-string pin holds behavior
'  PRESERVATION (the emitter's existing worded refusal, not a new
'  path); the lifetime pin holds last-transpile-wins; and the
'  apropos pins prove the tier lists program procedures by kind
'  AND that apropos's own prelude refresh cannot wipe the table
'  it is about to read (the refresh runs inside every call).
' ---------------------------------------------------------------------
Public Sub TestPL6()
    Dim t As String, t2 As String, d As String

    t = TryTranspile("pl6: a docstring never emits - documented equals undocumented byte for byte", _
                     "(sub t () (debug-print 1))")
    t2 = TryTranspile("pl6: a docstring never emits - documented equals undocumented byte for byte", _
                      "(sub t () ""says one"" (debug-print 1))")
    If Len(t) > 0 And Len(t2) > 0 Then
        Report "pl6: a docstring never emits - documented equals undocumented byte for byte", _
               t = t2, "undoc [" & Left$(Norm(t), 80) & "] doc [" & Left$(Norm(t2), 80) & "]"
    End If

    Report "pl6: the doc reads back, case-insensitively", _
           VlaSubDoc("t") = "says one" And VlaSubDoc("T") = "says one", _
           "got: [" & VlaSubDoc("t") & "] / [" & VlaSubDoc("T") & "]"

    t = TryTranspile("pl6: a function carries type and doc together", _
                     "(function f () Long ""answers"" (return 1)) (sub plain () (debug-print 2))")
    If Len(t) > 0 Then
        CheckFrags "pl6: a function carries type and doc together", t, _
                   Array("As Long", "f = 1")
    End If

    Report "pl6: documented, undocumented, and unknown names each answer honestly", _
           VlaSubDoc("f") = "answers" And VlaSubDoc("plain") = "" And VlaSubDoc("zz-none") = "", _
           "got: [" & VlaSubDoc("f") & "] / [" & VlaSubDoc("plain") & "] / [" & VlaSubDoc("zz-none") & "]"

    t = TryTranspile("pl6: a doc with no return type is never eaten as the type", _
                     "(function g () ""gives"" (return 2))")
    If Len(t) > 0 Then
        CheckFrags "pl6: a doc with no return type is never eaten as the type", t, _
                   Array("As Variant", "g = 2")
    End If

    Report "pl6: the untyped function's doc reads back", _
           VlaSubDoc("g") = "gives", "got: [" & VlaSubDoc("g") & "]"

    d = ""
    On Error Resume Next
    t = VlaTranspile("(sub t () ""only"")")
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    ' PL6.1: the frag now matches the VERIFIED refusal - EmitStmt's
    ' own atom guard ("bare atom used as a statement"), which fires
    ' before any HeadSym/SymText path could; the original pin quoted
    ' SymText's message from an inferred raise path, the L8.1
    ' mistake repeated. The guard predates P.L6 (present in the
    ' pre-pass source), so the behavior-preservation claim stands -
    ' only the quoted words were wrong.
    Report "pl6: a lone string is the body and keeps the existing worded refusal", _
           InStr(1, d, "bare atom used as a statement", vbTextCompare) > 0, "got: " & d

    t = TryTranspile("pl6: the table reflects the last transpile only", _
                     "(sub alpha () ""first"" (debug-print 1))")
    d = "before: [" & VlaSubDoc("alpha") & "]"
    t2 = TryTranspile("pl6: the table reflects the last transpile only", _
                      "(sub beta () (debug-print 2))")
    If Len(t) > 0 And Len(t2) > 0 Then
        Report "pl6: the table reflects the last transpile only", _
               VlaSubDoc("alpha") = "" And VlaSubDoc("beta") = "", _
               d & " after: [" & VlaSubDoc("alpha") & "]"
    End If

    t = TryTranspile("pl6: emission mangles the name, the table keeps it as written", _
                     "(sub tidy-report () ""sort and re-band"" (debug-print 1))")
    If Len(t) > 0 Then
        CheckFrags "pl6: emission mangles the name, the table keeps it as written", t, _
                   Array("Sub tidy_report(")
    End If

    Report "pl6: apropos lists a program sub by its written name, doc, and kind", _
           InStr(1, VlaAproposText("tidy"), "tidy-report - sort and re-band (program sub)", vbTextCompare) > 0, _
           Left$(VlaAproposText("tidy"), 160)

    t = TryTranspile("pl6: kinds and the no-documentation wording render on the tier", _
                     "(sub bare () (debug-print 1)) (function calc () ""adds"" (return 1))")
    If Len(t) > 0 Then
        Report "pl6: kinds and the no-documentation wording render on the tier", _
               InStr(1, VlaAproposText("bare"), "bare - (no documentation) (program sub)", vbTextCompare) > 0 And _
               InStr(1, VlaAproposText("calc"), "calc - adds (program function)", vbTextCompare) > 0, _
               Left$(VlaAproposText("bare"), 120) & " / " & Left$(VlaAproposText("calc"), 120)
    End If
End Sub

' ---------------------------------------------------------------------
'  P.L7 pins: (include "lib.vla"). The suite writes its own fixture
'  files under %TEMP% and includes them by ABSOLUTE path, so no pin
'  depends on ActiveWorkbook.Path. Coverage: the splice emits the
'  lib's code and never the include line; transpile-time errors name
'  the included FILE and its LOCAL line at both error surfaces
'  (emit-time via emitfail, reader-time via SrcLineTag - the
'  definition's own "unbalanced parens at lib.vla line N" example,
'  delivered); main-file lines AFTER an include still report their
'  own numbers (the resumption arithmetic); the no-include baseline
'  wording is byte-stable (the fast path); nesting works and deep
'  errors name the innermost file; missing files, self-inclusion,
'  and misplaced include forms each refuse with words; and included
'  statements' RUNTIME tags anchor to the include line (' vla:N of
'  the main file), keeping the runtime map main-file-coherent.
' ---------------------------------------------------------------------
Private Function WriteTempLib(ByVal name As String, ByVal text As String) As String
    Dim p As String
    p = Environ$("TEMP") & Application.PathSeparator & name
    Dim fnum As Integer
    fnum = FreeFile
    Open p For Output As #fnum
    Print #fnum, text;
    Close #fnum
    WriteTempLib = p
End Function

Public Sub TestPL7()
    Dim t As String, d As String
    Dim p1 As String, pBad As String, pA As String, pB As String, pBBad As String, pCyc As String, pUnb As String

    ' Two lines ON PURPOSE: the debug-print sits at lib line 2, so
    ' its runtime tag reading vla:1 (the include's main-file line)
    ' proves the ANCHOR override - a one-line lib could not tell
    ' anchor from raw coincidence.
    p1 = WriteTempLib("vla_pl7_lib.vla", "(sub lib-one ()" & vbCrLf & "(debug-print 77))")

    t = TryTranspile("pl7: an included file's code transpiles into the program", _
                     "(include """ & p1 & """)" & vbCrLf & "(sub t () (lib-one))")
    If Len(t) > 0 Then
        CheckFrags "pl7: an included file's code transpiles into the program", t, _
                   Array("Sub lib_one(", "Debug.Print 77", "Call lib_one")
    End If

    Report "pl7: the include line itself never emits", _
           Len(t) > 0 And InStr(1, t, "include", vbTextCompare) = 0, _
           "got: [" & Left$(Norm(t), 160) & "]"

    Report "pl7: an included statement's runtime tag anchors to the include line", _
           Len(t) > 0 And InStr(1, t, "Debug.Print 77 ' vla:1", vbTextCompare) > 0, _
           "got: [" & Left$(Norm(t), 200) & "]"

    pBad = WriteTempLib("vla_pl7_bad.vla", "(sub lib-fail ()" & vbCrLf & "(+ 1 2)" & vbCrLf & ")")
    d = ""
    On Error Resume Next
    t = VlaTranspile("(include """ & pBad & """)")
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    Report "pl7: an emit-time error names the included file and its local line", _
           InStr(1, d, pBad & " line 2", vbTextCompare) > 0, "got: " & d

    d = ""
    On Error Resume Next
    t = VlaTranspile("(include """ & p1 & """)" & vbCrLf & "(sub t () (+ 1 2))")
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    Report "pl7: a main-file error after an include still names the main line", _
           InStr(1, d, "vla line 2", vbTextCompare) > 0 And InStr(1, d, ".vla line", vbTextCompare) = 0, _
           "got: " & d

    d = ""
    On Error Resume Next
    t = VlaTranspile("(sub t ()" & vbCrLf & "(+ 1 2)" & vbCrLf & ")")
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    Report "pl7: the no-include baseline wording is untouched", _
           InStr(1, d, "(near vla line 2)", vbTextCompare) > 0, "got: " & d

    pB = WriteTempLib("vla_pl7_libb.vla", "(sub lib-two () (debug-print 88))")
    pA = WriteTempLib("vla_pl7_liba.vla", "(include """ & pB & """)" & vbCrLf & "(sub lib-outer () (lib-two))")
    t = TryTranspile("pl7: includes nest - the inner file's code arrives", _
                     "(include """ & pA & """)" & vbCrLf & "(sub t () (lib-outer))")
    If Len(t) > 0 Then
        CheckFrags "pl7: includes nest - the inner file's code arrives", t, _
                   Array("Debug.Print 88", "Sub lib_outer(")
    End If

    pBBad = WriteTempLib("vla_pl7_libb_bad.vla", "(sub deep-fail ()" & vbCrLf & "(+ 3 4)" & vbCrLf & ")")
    pA = WriteTempLib("vla_pl7_liba_bad.vla", "(include """ & pBBad & """)")
    d = ""
    On Error Resume Next
    t = VlaTranspile("(include """ & pA & """)")
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    Report "pl7: a deep error names the innermost file and line", _
           InStr(1, d, pBBad & " line 2", vbTextCompare) > 0, "got: " & d

    d = ""
    On Error Resume Next
    t = VlaTranspile("(include """ & Environ$("TEMP") & Application.PathSeparator & "vla_pl7_missing.vla"")")
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    Report "pl7: a missing file refuses with words", _
           InStr(1, d, "include: cannot read", vbTextCompare) > 0, "got: " & d

    pCyc = Environ$("TEMP") & Application.PathSeparator & "vla_pl7_cycle.vla"
    pCyc = WriteTempLib("vla_pl7_cycle.vla", "(include """ & pCyc & """)")
    d = ""
    On Error Resume Next
    t = VlaTranspile("(include """ & pCyc & """)")
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    Report "pl7: a self-including file refuses with the depth message", _
           InStr(1, d, "nesting deeper than 16", vbTextCompare) > 0, "got: " & d

    d = ""
    On Error Resume Next
    t = VlaTranspile("(sub t () (include ""x.vla""))")
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    Report "pl7: an include inside a procedure refuses with the whole-line rule", _
           InStr(1, d, "must stand alone on its own line", vbTextCompare) > 0, "got: " & d

    pUnb = WriteTempLib("vla_pl7_unbal.vla", "(sub broken (")
    d = ""
    On Error Resume Next
    t = VlaTranspile("(include """ & pUnb & """)")
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    Report "pl7: a reader-time error names the included file - unbalanced parens included", _
           InStr(1, d, pUnb & " line 1", vbTextCompare) > 0, "got: " & d
End Sub

' ---------------------------------------------------------------------
'  L13 pins: the with- family growth (with-screen-off, with-no-
'  alerts, with-protected-sheet), each on with-fast-excel's proven
'  restore-label pattern. Expansion pins are frags-plus-order (the
'  multi-statement expansions exceed one printed line - L7's
'  precedent); emission pins are full-stack positional order, the
'  restore-on-both-exits contract held as: arm BEFORE body, label
'  AFTER body, restore AFTER label, disarm LAST. Frags come from
'  verified paths only: set!/label/on-error shapes are with-fast-
'  excel's own, and password:= is EmitArgs' pinned rendering. The
'  composition pin proves P-layer binding forms ride inside bracket
'  bodies (the recorded design line: they compose in bodies, not in
'  bracket plumbing); the two-bracket pin holds the C1 promise that
'  distinct caller-named labels coexist in one procedure.
' ---------------------------------------------------------------------
Public Sub TestL13()
    Dim t As String
    Dim fired As Long

    t = VlaExpandText("(with-screen-off done (debug-print 1))", True, fired)
    Dim e1 As Long, e2 As Long, e3 As Long, e4 As Long, e5 As Long, e6 As Long, e7 As Long
    e1 = InStr(1, t, "(set! application.screenupdating false)", vbTextCompare)
    e2 = InStr(1, t, "(on-error goto done)", vbTextCompare)
    e3 = InStr(1, t, "(debug-print 1)", vbTextCompare)
    e4 = InStr(1, t, "(label done)", vbTextCompare)
    e5 = InStr(1, t, "(set! application.screenupdating true)", vbTextCompare)
    e6 = InStr(1, t, "(err.raise err.number", vbTextCompare)
    e7 = InStr(1, t, "(on-error goto 0)", vbTextCompare)
    Report "l13: with-screen-off expands arm, body, label, restore, re-raise, disarm in order", _
           e1 > 0 And e2 > e1 And e3 > e2 And e4 > e3 And e5 > e4 And e6 > e5 And e7 > e6, _
           "got: [" & Left$(Norm(t), 200) & "]"

    t = TryTranspile("l13: with-screen-off emits the restore contract in order", _
                     "(sub t () (with-screen-off done (debug-print 1)))")
    If Len(t) > 0 Then
        Dim p1 As Long, p2 As Long, p3 As Long, p4 As Long, p5 As Long, p6 As Long
        p1 = InStr(1, t, "application.screenupdating = False", vbTextCompare)
        p2 = InStr(1, t, "On Error GoTo done", vbTextCompare)
        p3 = InStr(1, t, "Debug.Print 1", vbTextCompare)
        p4 = InStr(1, t, "done:", vbTextCompare)
        p5 = InStr(1, t, "application.screenupdating = True", vbTextCompare)
        p6 = InStr(1, t, "On Error GoTo 0", vbTextCompare)
        Report "l13: with-screen-off emits the restore contract in order", _
               p1 > 0 And p2 > p1 And p3 > p2 And p4 > p3 And p5 > p4 And p6 > p5, _
               Left$(Norm(t), 200)
    End If

    t = TryTranspile("l13: with-no-alerts brackets DisplayAlerts and restores True", _
                     "(sub t () (with-no-alerts done (debug-print 2)))")
    If Len(t) > 0 Then
        Dim q1 As Long, q2 As Long, q3 As Long
        q1 = InStr(1, t, "application.displayalerts = False", vbTextCompare)
        q2 = InStr(1, t, "Debug.Print 2", vbTextCompare)
        q3 = InStr(1, t, "application.displayalerts = True", vbTextCompare)
        Report "l13: with-no-alerts brackets DisplayAlerts and restores True", _
               q1 > 0 And q2 > q1 And q3 > q2, Left$(Norm(t), 200)
    End If

    t = VlaExpandText("(with-protected-sheet ws pw done (debug-print 9))", True, fired)
    Dim x1 As Long, x2 As Long, x3 As Long, x4 As Long
    x1 = InStr(1, t, "(. ws unprotect :password pw)", vbTextCompare)
    x2 = InStr(1, t, "(debug-print 9)", vbTextCompare)
    x3 = InStr(1, t, "(label done)", vbTextCompare)
    x4 = InStr(1, t, "(. ws protect :password pw)", vbTextCompare)
    Report "l13: with-protected-sheet expands unprotect, body, label, re-protect in order", _
           x1 > 0 And x2 > x1 And x3 > x2 And x4 > x3, "got: [" & Left$(Norm(t), 200) & "]"

    t = TryTranspile("l13: with-protected-sheet emits the keyword at both call sites in order", _
                     "(sub t () (with-protected-sheet ws pw done (debug-print 9)))")
    If Len(t) > 0 Then
        Dim y1 As Long, y2 As Long, y3 As Long, y4 As Long, y5 As Long
        y1 = InStr(1, t, "ws.unprotect", vbTextCompare)
        y2 = InStr(1, t, "Debug.Print 9", vbTextCompare)
        y3 = InStr(1, t, "ws.protect", vbTextCompare)
        y4 = InStr(1, t, "password:=pw", vbTextCompare)
        ' InStr raises on a start of 0, and VBA And does not short-
        ' circuit - guard the second-occurrence search explicitly.
        y5 = 0
        If y3 > 0 Then y5 = InStr(y3, t, "password:=pw", vbTextCompare)
        Report "l13: with-protected-sheet emits the keyword at both call sites in order", _
               y1 > 0 And y2 > y1 And y3 > y2 And y4 > 0 And y4 < y3 And y5 > y3, _
               Left$(Norm(t), 220)
    End If

    t = TryTranspile("l13: the P-layer's binding forms compose inside a bracket body", _
                     "(sub t () (with-screen-off done (when-let w 5 (debug-print w))))")
    If Len(t) > 0 Then
        Dim z1 As Long, z2 As Long, z3 As Long
        z1 = InStr(1, t, "On Error GoTo done", vbTextCompare)
        z2 = InStr(1, t, "If (Not isempty(w)) Then", vbTextCompare)
        z3 = InStr(1, t, "done:", vbTextCompare)
        Report "l13: the P-layer's binding forms compose inside a bracket body", _
               z1 > 0 And z2 > z1 And z3 > z2, Left$(Norm(t), 200)
    End If

    t = TryTranspile("l13: two brackets with distinct labels coexist in one procedure", _
                     "(sub t () (with-screen-off d1 (f)) (with-no-alerts d2 (g)))")
    If Len(t) > 0 Then
        Dim w1 As Long, w2 As Long
        w1 = InStr(1, t, "d1:", vbTextCompare)
        w2 = InStr(1, t, "d2:", vbTextCompare)
        Report "l13: two brackets with distinct labels coexist in one procedure", _
               w1 > 0 And w2 > w1, Left$(Norm(t), 200)
    End If

    Report "l13: the three brackets carry their docstrings on the bench", _
           InStr(1, VlaAproposText("with-screen-off"), "restored on both exits", vbTextCompare) > 0 And _
           InStr(1, VlaAproposText("with-no-alerts"), "display alerts off", vbTextCompare) > 0 And _
           InStr(1, VlaAproposText("with-protected-sheet"), "re-protected on both exits", vbTextCompare) > 0, _
           Left$(VlaAproposText("with-"), 200)
End Sub

' ---------------------------------------------------------------------
'  G12 pins: "Use library" - the English entry point for (include).
'  All full-stack through EnglishToVla; the placement pin is
'  positional (the include line must sit ABOVE (sub main - the
'  top-matter channel working); dedup is an occurrence count with
'  deliberately case-varied spellings; the refusal pin holds the
'  quoted-name teaching by its words. The Check-transpile hook is
'  IDE machinery needing sheets and lives with the on-machine
'  smokes, recorded on the ledger - the suite pins the language,
'  the smokes pin the panel (the standing rig/suite boundary).
' ---------------------------------------------------------------------
Public Sub TestG12()
    Dim t As String, d As String

    On Error Resume Next
    t = EnglishToVla("Use library ""helpers.vla""." & vbCrLf & "Log ""x"".")
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    If Len(d) > 0 Then
        Report "g12: Use library emits the include below the program", False, "translate error: " & d
    Else
        Dim p1 As Long, p2 As Long
        p1 = InStr(1, t, "(include ""helpers.vla"")", vbTextCompare)
        p2 = InStr(1, t, "(sub main", vbTextCompare)
        ' G12.1: the pin FLIPPED with the placement - bottom-matter,
        ' so spliced library procedures land below the program and
        ' its module-level declarations stay above the first sub.
        Report "g12: Use library emits the include below the program", _
               p1 > 0 And p2 > 0 And p1 > p2, Left$(Norm(t), 200)
    End If

    AssertEnglish "g12: Import code from is the same sentence", _
                  "Import code from ""kit.vla""." & vbCrLf & "Log ""x"".", "(include ""kit.vla"")"

    t = EnglishToVla("Use library ""Helpers.vla""." & vbCrLf & "Import library ""helpers.vla""." & vbCrLf & "Log ""x"".")
    Dim c As Long, q As Long
    c = 0
    q = 1
    Do
        q = InStr(q, t, "(include ", vbTextCompare)
        If q = 0 Then Exit Do
        c = c + 1
        q = q + 1
    Loop
    Report "g12: a repeated library splices once, case-insensitively", _
           c = 1, "include count " & c & " in: " & Left$(Norm(t), 160)

    d = ""
    On Error Resume Next
    t = EnglishToVla("Use library helpers." & vbCrLf & "Log ""x"".")
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    Report "g12: an unquoted library name refuses teaching the period rule", _
           InStr(1, d, "must be quoted", vbTextCompare) > 0, "got: " & d

    t = EnglishToVla("Use library ""a.vla""." & vbCrLf & "Use code ""b.vla""." & vbCrLf & "Log ""x"".")
    Dim a1 As Long, b1 As Long
    a1 = InStr(1, t, "(include ""a.vla"")", vbTextCompare)
    b1 = InStr(1, t, "(include ""b.vla"")", vbTextCompare)
    Report "g12: two libraries emit in first-mention order", _
           a1 > 0 And b1 > a1, Left$(Norm(t), 160)

    d = ""
    On Error Resume Next
    t = VlaTranspile("(sub t () (debug-print 1))" & vbCrLf & "(dim late Long)")
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    Report "g12: a module-level declaration after a procedure refuses with words", _
           InStr(1, d, "must come before the first", vbTextCompare) > 0, "got: " & d
End Sub

' ---------------------------------------------------------------------
'  L14 pins: deflambda - the two-Lisps handshake. All full-stack
'  through TryTranspile; the RefersTo frags carry the exact formula
'  text with VBA-doubled quotes (built by hand against the formula
'  dialect's stated rules, then held here byte-for-byte). The
'  bench-consumption pin is the load-bearing one: the SOURCE says
'  (even? n) and the SHEET receives MOD(n, 2) - the P-layer
'  reaching a worksheet formula through ordinary macro expansion.
'  Refusals each by their words: statement bodies, the statement-
'  if's (then ...) blocks, nested quote, top-level placement.
' ---------------------------------------------------------------------
Public Sub TestL14()
    Dim t As String, d As String

    t = TryTranspile("l14: deflambda registers a Name carrying the LAMBDA", _
                     "(sub t () (deflambda double (x) (* x 2)))")
    If Len(t) > 0 Then
        CheckFrags "l14: deflambda registers a Name carrying the LAMBDA", t, _
                   Array("ThisWorkbook.Names.Add Name:=""double"", RefersTo:=""=LAMBDA(x, (x*2))""")
    End If

    t = TryTranspile("l14: parameters join comma-first and comparisons render infix", _
                     "(sub t () (deflambda bigger (a b) (> a b)))")
    If Len(t) > 0 Then
        CheckFrags "l14: parameters join comma-first and comparisons render infix", t, _
                   Array("RefersTo:=""=LAMBDA(a, b, (a>b))""")
    End If

    t = TryTranspile("l14: the P-layer reaches the sheet - even? becomes MOD", _
                     "(sub t () (deflambda parity (n) (if (even? n) ""even"" ""odd"")))")
    If Len(t) > 0 Then
        CheckFrags "l14: the P-layer reaches the sheet - even? becomes MOD", t, _
                   Array("RefersTo:=""=LAMBDA(n, IF((MOD(n, 2)=0), """"even"""", """"odd""""))""")
    End If

    t = TryTranspile("l14: the docstring rides the Name's Comment and the doc table", _
                     "(sub t () (deflambda tax (x) ""the vat"" (* x 0.2)))")
    If Len(t) > 0 Then
        CheckFrags "l14: the docstring rides the Name's Comment and the doc table", t, _
                   Array(").Comment = ""the vat""", "RefersTo:=""=LAMBDA(x, (x*0.2))""")
    End If

    Report "l14: apropos lists the lambda as a worksheet function", _
           VlaSubDoc("tax") = "the vat" And _
           InStr(1, VlaAproposText("vat"), "tax - the vat (worksheet function)", vbTextCompare) > 0, _
           "doc: [" & VlaSubDoc("tax") & "] " & Left$(VlaAproposText("vat"), 120)

    t = TryTranspile("l14: quote data becomes an array constant", _
                     "(sub t () (deflambda picks () (quote (1 2 3))))")
    If Len(t) > 0 Then
        CheckFrags "l14: quote data becomes an array constant", t, _
                   Array("RefersTo:=""=LAMBDA({1,2,3})""")
    End If

    t = TryTranspile("l14: hyphenated names mangle for Excel, stay as written on the bench", _
                     "(sub t () (deflambda tax-rate () ""flat fifth"" (+ 0.1 0.1)))")
    If Len(t) > 0 Then
        Report "l14: hyphenated names mangle for Excel, stay as written on the bench", _
               InStr(1, t, "Name:=""tax_rate""", vbTextCompare) > 0 And VlaSubDoc("tax-rate") = "flat fifth", _
               Left$(Norm(t), 160) & " doc: [" & VlaSubDoc("tax-rate") & "]"
    End If

    d = ""
    On Error Resume Next
    t = VlaTranspile("(sub t () (deflambda bad (x) (set! x 1)))")
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    Report "l14: a statement body refuses with words", _
           InStr(1, d, "is a statement", vbTextCompare) > 0, "got: " & d

    d = ""
    On Error Resume Next
    t = VlaTranspile("(sub t () (deflambda bad (x) (if (> x 1) (then 2))))")
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    Report "l14: the statement-if's then-blocks refuse teaching the three-arg spelling", _
           InStr(1, d, "three-arg spelling", vbTextCompare) > 0, "got: " & d

    d = ""
    On Error Resume Next
    t = VlaTranspile("(sub t () (deflambda bad () (quote ((1 2) 3))))")
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    Report "l14: nested quote refuses - Excel arrays do not nest", _
           InStr(1, d, "do not nest", vbTextCompare) > 0, "got: " & d

    d = ""
    On Error Resume Next
    t = VlaTranspile("(deflambda toplevel (x) (* x 2))")
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    Report "l14: top-level deflambda refuses with placement words", _
           InStr(1, d, "runs when the program runs", vbTextCompare) > 0, "got: " & d
End Sub

' ---------------------------------------------------------------------
'  L18 pins: TCO for VLA functions. A (return (self args...)) is a
'  tail call by construction, and under an armed transform it
'  compiles to rebind-and-GoTo instead of a call - recursion in the
'  language, a loop in the metal. Pins hold: the loop shape and the
'  ABSENCE of a call-shaped self-reference; the single-parameter
'  fast path (no temps); the evaluate-all-temps-BEFORE-any-rebind
'  order under cross-referencing arguments; that a NON-tail self-
'  call keeps true recursion (and arms nothing); that a ByRef
'  parameter disarms the transform entirely (rebinding ByRef would
'  mutate the original caller's variable - correctness over
'  optimization, stated not smoothed); the arity refusal by its
'  words; and the P.L6 docstring coexisting with the transform.
' ---------------------------------------------------------------------
Public Sub TestL18()
    Dim t As String, d As String

    t = TryTranspile("l18: a self-tail-call compiles to rebind and jump", _
                     "(function fact-iter ((byval n Long) (byval acc Long)) Long ""iterative factorial"" (if (<= n 1) (then (return acc)) (else (return (fact-iter (- n 1) (* n acc))))))")
    If Len(t) > 0 Then
        Dim p1 As Long, p2 As Long, p3 As Long, p4 As Long, p5 As Long, p6 As Long
        p1 = InStr(1, t, "vla_tco:", vbTextCompare)
        p2 = InStr(1, t, "vla_tco_1 = (n - 1)", vbTextCompare)
        p3 = InStr(1, t, "vla_tco_2 = (n * acc)", vbTextCompare)
        p4 = InStr(1, t, "n = vla_tco_1", vbTextCompare)
        p5 = InStr(1, t, "acc = vla_tco_2", vbTextCompare)
        p6 = InStr(1, t, "GoTo vla_tco", vbTextCompare)
        Report "l18: a self-tail-call compiles to rebind and jump", _
               p1 > 0 And p2 > p1 And p3 > p2 And p4 > p3 And p5 > p4 And p6 > p5 And _
               InStr(1, t, "= fact_iter(", vbTextCompare) = 0, _
               Left$(Norm(t), 220)
    End If

    Report "l18: the docstring rides the transform untouched", _
           VlaSubDoc("fact-iter") = "iterative factorial", "got: [" & VlaSubDoc("fact-iter") & "]"

    t = TryTranspile("l18: one parameter rebinds directly - no temps", _
                     "(function count-down ((byval n Long)) Long (if (zero? n) (then (return 0)) (else (return (count-down (- n 1))))))")
    If Len(t) > 0 Then
        Report "l18: one parameter rebinds directly - no temps", _
               InStr(1, t, "n = (n - 1)", vbTextCompare) > 0 And _
               InStr(1, t, "GoTo vla_tco", vbTextCompare) > 0 And _
               InStr(1, t, "vla_tco_1", vbTextCompare) = 0, _
               Left$(Norm(t), 200)
    End If

    t = TryTranspile("l18: a non-tail self-call keeps true recursion", _
                     "(function depth ((byval n Long)) Long (if (zero? n) (then (return 0)) (else (return (+ 1 (depth (- n 1)))))))")
    If Len(t) > 0 Then
        Report "l18: a non-tail self-call keeps true recursion", _
               InStr(1, t, "depth((n - 1))", vbTextCompare) > 0 And _
               InStr(1, t, "vla_tco", vbTextCompare) = 0, _
               Left$(Norm(t), 200)
    End If

    t = TryTranspile("l18: a ByRef parameter disarms the transform", _
                     "(function echo ((byref n Long)) Long (if (zero? n) (then (return 0)) (else (return (echo n)))))")
    If Len(t) > 0 Then
        Report "l18: a ByRef parameter disarms the transform", _
               InStr(1, t, "echo = echo(n)", vbTextCompare) > 0 And _
               InStr(1, t, "vla_tco", vbTextCompare) = 0, _
               Left$(Norm(t), 200)
    End If

    d = ""
    On Error Resume Next
    t = VlaTranspile("(function f ((byval a Long) (byval b Long)) Long (return (f a)))")
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    Report "l18: an arity-mismatched tail call refuses with words", _
           InStr(1, d, "passes 1 argument(s) but it takes 2", vbTextCompare) > 0, "got: " & d
End Sub

' ---------------------------------------------------------------------
'  TestAlonzoLib: the evergreen import-regression library. alonzo.vla
'  (a function loader - renamed from hello.vla, whose old name didn't
'  say what it held) joins instructions.txt and english.vla as the
'  third corpus file (located by the same FindDevFile prober; its
'  ABSENCE is a failure, the corpus-family contract). The pins run one
'  include-and-register program through the full stack and hold: the
'  bricks' Add lines with the REDUCE/SEQUENCE/inner-LAMBDA formula
'  text; the P-layer brick reaching the sheet ((v>0) from
'  positive?); an imported brick's docstring on the bench with the
'  worksheet-function kind; and multi-line library forms parsing -
'  the file is deliberately multi-line, so this whole section is
'  also the standing proof that VLA files never required one-line
'  forms (only the IDE's raw rows do).
' ---------------------------------------------------------------------
Public Sub TestAlonzoLib()
    ' AS.6 (verifier-caught, both independent passes): this call was
    ' bare - FindDevFile raises loudly on a miss, it never returns ""
    ' - so the Dir$ recheck below was dead code for the actual missing-
    ' file case, and an unguarded raise here would escape this Sub
    ' (TestAlonzoLib at the time) and kill VlaSelfTest mid-run at test
    ' 53 of 59, the exact hazard TestCorpusFamily/TestDotCount were
    ' just hardened against.
    Dim libPath As String
    Dim d As String
    On Error Resume Next
    libPath = FindDevFile("alonzo.vla")
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    ' VBA's Or does not short-circuit, so the Dir$ recheck only runs
    ' once d is known empty - Dir$ on an unset libPath after a failed
    ' FindDevFile call is exactly the kind of untested edge this fix
    ' exists to close, not reopen.
    If Len(d) = 0 Then
        If Len(Dir$(libPath)) = 0 Then d = "path resolved but Dir$ could not confirm it: " & libPath
    End If
    If Len(d) > 0 Then
        Report "alonzolib: the evergreen library is present beside the corpus", False, _
               "alonzo.vla not found - it is part of the corpus family now (" & d & ")"
        Exit Sub
    End If

    Dim t As String
    t = TryTranspile("alonzolib: the bricks register through include", _
                     "(include """ & libPath & """)" & vbCrLf & "(sub t () (register-bricks))")
    If Len(t) = 0 Then Exit Sub

    CheckFrags "alonzolib: the bricks register through include", t, _
               Array("Name:=""sum_to""", "reduce(0, sequence(n), LAMBDA(acc, i, (acc+i)))", "Name:=""clamp""")

    CheckFrags "alonzolib: the P-layer brick reaches the sheet", t, _
               Array("IF((v>0), 1, 0)", "Name:=""count_positive""")

    Report "alonzolib: an imported brick's docstring lands on the bench", _
           VlaSubDoc("count-positive") = "how many cells are positive - the P-layer answering from a sheet" And _
           InStr(1, VlaAproposText("pinned"), "clamp - x pinned into the closed range lo..hi (worksheet function)", vbTextCompare) > 0, _
           "doc: [" & VlaSubDoc("count-positive") & "] " & Left$(VlaAproposText("pinned"), 160)
End Sub

' LX.10: the falsification test - "a non-English market exists and is
' reachable," the belief BETA_ROADMAP.md's own entry says this item
' exists to be cheaply wrong about. scripts/espanol.vla already carries
' 19 rules and 20 test-success proofs in Spanish (session evidence
' logged against this item, six sibling dialect files besides), written
' under the same blind-authoring-then-engine-verifies discipline
' english.vla's own header names - but never wired into VlaSelfTest, run
' only informally, by hand. That is this item's actual gap, not a
' missing phrasebook: loading it here, right after english.vla (its own
' comment's stated load order - the base layer's macros must exist
' first), exercises every one of its sentences through the real
' engine. EnglishLoadVocabularyText's own loader raises
' (english-test-failed/english-test-failed-to-translate) the instant any
' proof's claimed translation drifts, so a clean load IS the
' falsification test passing, not a weaker proxy for it. Guarded the
' same FindDevFile/On-Error idiom TestAlonzoLib/TestDotCount/
' TestCorpusFamily already use (their own comments name the shared
' hazard: a self-test must never be able to die on the corpus it
' judges).
' LX5.3 (this session): three of espanol.vla's proofs are no longer
' phrase-vocabulary-only - they exercise real Spanish control-flow
' keywords (si/repetir/mientras) through LX5.1's keyword-alias seam,
' the first non-English sentences in this project to reach ParseStmt's
' own structural dispatch rather than only the DCG phrase matcher.
' Their expected VLA text is hand-derived against VLA_SentenceEngine.
' bas's own ParseStmt/ParseCondSimple source, not yet confirmed
' against a live run - if any is wrong, THIS test fails with the
' engine's own real error text, not a silent pass; that failure text
' is the correction. Owner verification with VlaSelfTest is still
' needed for the whole fixture, keyword-alias proofs included.
Public Sub TestLx10NonEnglishFixture()
    Dim d As String
    On Error Resume Next
    EnglishResetGrammar
    EnglishLoadVocabulary FindDevFile("english.vla")
    EnglishLoadVocabulary FindDevFile("espanol.vla")
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    Report "lx10: espanol.vla's non-English test-success proofs all pass (LX5.3: incl. 3 real control-flow proofs)", _
           Len(d) = 0, "error: " & d
End Sub

' F.4 (shape 2, shipped this pass): a noise word ("a"/"an"/"the"/
' "please") is stripped from every PATTERN at registration
' (IsNoiseWord, AddPhraseRule) before the matcher ever sees it. When
' one of those words sits immediately before a {...} slot in the raw
' pattern, the slot silently swallows whatever token the input leaves
' behind instead of its intended value - confirmed live in
' espanol.vla's "envia un correo a {who:expr} ..." (closing LX.10) and,
' confirmed live this session (owner's own VlaSelfTest run) in
' francais.vla:151's "envoie un courriel a {who:expr} ...". The two
' confirmed instances surface through DIFFERENT gates, both honest:
' espanol's slot swallowed the leftover "a" silently and still produced
' a translation (no test failure - the family's original "no error, no
' warning" shape); francais's swallow is loud enough to blow the whole
' rule's own test-success proof (the {who:expr} slot greedily accepts
' bare "a" as a one-letter column reference, then the match fails
' looking for the next literal "avec" instead of the address it should
' have parsed) - EnglishLoadVocabulary raises immediately, before
' EnglishAuditPhrasebook ever gets to return a warnings string. The
' lint warning still fires at registration time either way (LintRule
' runs before any test: line is evaluated) - EnglishLintReport() below
' confirms it independently of which gate the load then hits. The
' check itself is purely lexical (scan the raw pattern string before
' stripping), so the synthetic pins below lock in both the positive
' and negative case without needing the real matcher. This is NOT the
' other F.4 shape -
' a general rule's {:expr}/{:cond} slot swallowing a more specific
' rule's own literal continuation (english.vla:517's "show cell in
' column {c:column} row {n:expr}", shadowed by the core "show
' {e:expr}" rule) - that needs real grammar-aware reasoning and is not
' yet built; it will not be flagged here.
Public Sub TestF4NoiseWordBeforeSlot()
    Dim aud As String
    aud = EnglishAuditText( _
        "(english-vla ""send an email a {who:expr} with subject {s:expr}"" (vlasendmail {who} {s}))")
    Report "f4: audit flags a noise word immediately before a slot", _
           InStr(1, aud, "noise word", vbTextCompare) > 0 And InStr(1, aud, "{who:expr}", vbTextCompare) > 0, _
           "report was: " & Left$(aud, 200)

    aud = EnglishAuditText( _
        "(english-vla ""send an email to {who:expr} with subject {s:expr}"" (vlasendmail {who} {s}))")
    Report "f4: audit stays clean when a real word (not stripped as noise) precedes the slot", _
           Len(aud) = 0, "unexpected: " & Left$(aud, 200)

    ' Real acceptance target: francais.vla:151's own rule, the second
    ' live instance BETA_ROADMAP.md's F.4 entry names. Confirmed live
    ' (owner's VlaSelfTest run): loading raises - the rule's own
    ' test-success proof fails to translate, because the swallowed "a"
    ' gets parsed as a one-letter column reference and the match then
    ' derails looking for the literal "avec" - so EnglishAuditPhrasebook
    ' never gets to return a warnings string for THIS rule. That raise
    ' is the expected shape, not a test bug: assert it by name (the
    ' "avec"/"courriel a" fingerprint), then confirm independently, via
    ' EnglishLintReport(), that the lint warning fired at registration
    ' time regardless of how the subsequent test-success line fared.
    Dim d As String
    On Error Resume Next
    EnglishResetGrammar
    EnglishLoadVocabulary FindDevFile("english.vla")
    EnglishLoadVocabulary FindDevFile("francais.vla")
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    Report "f4: francais.vla:151's swallowed 'a' breaks its own test-success proof, as predicted", _
           InStr(1, d, "courriel a", vbTextCompare) > 0 And InStr(1, d, "avec", vbTextCompare) > 0, _
           "error: " & d
    Dim lint As String
    lint = EnglishLintReport()
    Report "f4: the lint warning fired at registration regardless of the later test failure", _
           InStr(1, lint, "noise word", vbTextCompare) > 0 And InStr(1, lint, "{who:expr}", vbTextCompare) > 0, _
           "report was: " & Left$(lint, 300)
    EnglishResetGrammar
End Sub

' F.4 (shape 1, this pass): does an earlier-registered rule's slot
' swallow this rule's own literal continuation, for every shape it can
' present? Automates EnglishTryRule's (G5) own one-at-a-time verdict
' ("an EARLIER rule matched - ... - the candidate never fired") for
' the whole corpus, via AuditCrossRuleShadow (VLA_English.bas) - a
' post-load pass run only from EnglishAuditText/EnglishAuditPhrasebook
' (not a per-registration hook: an earlier version hooked it into
' AddPhraseRule directly and quadrupled VlaSelfTest's own runtime,
' since a per-rule dispatch probe is O(corpus size so far) and that
' cost was being paid by every ordinary test's own reload, not just
' audits - see BETA_ROADMAP.md's F.4 entry for the measured numbers).
' The real target is english.vla:517's own "show cell in column
' {c:column} row {n:expr}", shadowed by the BUILT-IN prelude rule
' "show {e:expr}" -> (msgbox {e}) (AddPhraseRule, ~line 1198 - always
' registered first, before any phrasebook loads, so it survives every
' EnglishResetGrammar), because "cell in column X row Y" is itself a
' recognized :expr production. The synthetic pin below needs no
' companion rule at all for exactly that reason - the shadowing rule
' is already the built-in. Split in two, deliberately: this Sub (cheap
' synthetic pins, wired into VlaSelfTest, pins the DETECTOR's own
' logic) and TestF4RealCorpusShadow below (the real english.vla
' confirmation, NOT wired in - see its own header for why).
Public Sub TestF4CrossRuleShadow()
    Dim aud As String
    aud = EnglishAuditText( _
        "(english-vla ""show cell in column {c:column} row {n:expr}"" (debug-print (cells {n} {c})))")
    Report "f4: audit flags a rule shadowed by an earlier :expr slot", _
           InStr(1, aud, "claimed by an earlier rule", vbTextCompare) > 0 And _
           InStr(1, aud, "show {e:expr}", vbTextCompare) > 0, _
           "report was: " & Left$(aud, 300)

    aud = EnglishAuditText( _
        "(english-vla ""zap the widget {n:expr}"" (debug-print {n}))")
    Report "f4: audit stays clean when no earlier rule can claim the shape", _
           Len(aud) = 0, "unexpected: " & Left$(aud, 200)
    EnglishResetGrammar
End Sub

' F.4 shape 1's real-corpus confirmation. Deliberately NOT wired into
' VlaSelfTest, unlike TestF4CrossRuleShadow above: AuditCrossRuleShadow's
' own dispatch probe over the real ~100-rule corpus measured at ~25-30s
' by itself (owner's own timed VlaSelfTest runs, before/after comparison
' isolated it to this one call) - correct architecturally (production
' loads never pay it, see AuditCrossRuleShadow's own header) but too
' slow to pay on every routine test pass for what is, once shape 1's
' DETECTOR logic is already pinned by the cheap synthetic cases above,
' a confirmation run rather than a regression guard. Call directly when
' only THIS check matters, or via VlaSelfTestsAll (VLA_Tests_Host.bas)
' to run it alongside everything else excluded purely for cost (the
' scale tests included) without having to remember its name - after a
' phrasebook change, before a release, or after touching
' AuditCrossRuleShadow itself, not on every routine VlaSelfTest/
' VlaSelfTests pass.
'
' THIS TEST'S JOB CHANGED, once, live: it used to assert the four real
' shadows F.4 shape 1 found in english.vla (rule 517's own "show cell
' in column..." plus three "set {v:var} to cell..." siblings) were
' still detected - a real acceptance target while those rules existed.
' Once fixed (all four removed as dead code, byte-identical output
' already provided by the shadowing prelude rules - see english.vla's
' own comments at each removed rule's old location), there is nothing
' left in the real corpus for this test to find, so the assertion
' flipped to its natural next job: prove the corpus STAYS clean. This
' is now a regression guard against a future phrasebook edit
' reintroducing a cross-rule shadow, not proof the detector works -
' that proof still lives entirely in TestF4CrossRuleShadow's own
' synthetic pins above, unaffected by whatever english.vla currently
' contains. Prints the full audit text unconditionally (not just on
' failure) so a live run shows anything found directly.
Public Sub TestF4RealCorpusShadow()
    Dim d As String, aud As String
    On Error Resume Next
    EnglishResetGrammar
    aud = EnglishAuditPhrasebook(FindDevFile("english.vla"))
    If Err.Number <> 0 Then d = Err.Description
    On Error GoTo 0
    Debug.Print "--- TestF4RealCorpusShadow: english.vla audit ---"
    Debug.Print IIf(Len(aud) = 0, "(clean - no warnings)", aud)

    Report "f4: english.vla's real cross-rule shadows are fixed and stay fixed", _
           Len(d) = 0 And Len(aud) = 0, _
           "error: " & d & " report: " & aud
    EnglishResetGrammar
End Sub

' G-PATH: the new {:path} slot category (pareto.txt SS15's own
' preamble - "needs a new {p:path} slot type (quoted, period-safe) -
' the G12 rule already solved exactly this problem for library names;
' reuse that decision"). Reuses G12's POLICY, not its bespoke top-
' level-statement MECHANISM (see MatchPathToken's own header note,
' VLA_English.bas, for the full reasoning) - rides the ordinary G2
' typed-ref-slot machinery instead, since SS15 has ~13 ordinary verbs
' that need this, not one special form.
Public Sub TestGPath()
    Dim junk As String, d As String

    EnglishResetGrammar
    EnglishAddPhrase "probe path {p:path}", "(debug-print {p})"

    AssertEnglish "gpath: a quoted literal path is accepted, backslashes intact", _
                  "Probe path ""C:\Reports\file.xlsx"".", """C:\\Reports\\file.xlsx"""

    ' Deliberately NOT VlaStringLit(t) the way cell/range/sheet's bare-
    ' token branch works (MatchRefToken) - {:path} matches {:expr}'s
    ' own "Open workbook report-path." ergonomics instead, so an
    ' unquoted bareword must emit as a bare identifier (a variable
    ' reference), never as a quoted string.
    AssertEnglish "gpath: an unquoted bareword is a variable reference, not a literal", _
                  "Probe path report-path.", "(debug-print report-path)"

    ' "C:." - a bareword immediately followed by its own free-standing
    ' ':' token (not glued, since IsWordChar has no '.'; not a block
    ' colon either, since that requires trailing WHITESPACE and this
    ' one runs straight into the sentence's own '.') - the narrower,
    ' backslash-free shape MatchPathToken's own colon-lookahead can
    ' actually catch (a real "C:\..." dies one layer earlier, at
    ' tokenization, before any slot matching runs at all - see the
    ' next two checks).
    On Error Resume Next
    Err.Clear
    junk = EnglishToVla("Probe path C:.")
    d = Err.Description
    On Error GoTo 0
    Report "gpath: an unquoted colon-shaped path refuses with the teaching message", _
           InStr(1, d, "must be quoted", vbTextCompare) > 0 And _
           InStr(1, d, "period ends a sentence", vbTextCompare) > 0, _
           "got: " & d

    ' The realistic failure mode: a real Windows path's backslash dies
    ' at EnTokenize itself, before {:path} or any other slot ever sees
    ' a token - proven here with NO {:path} rule registered at all
    ' (grammar reset, no probe phrase), so this is purely the
    ' tokenizer's own StrayCharHint fix, not MatchPathToken's.
    EnglishResetGrammar
    d = ""
    On Error Resume Next
    Err.Clear
    junk = EnglishToVla("Log C:\Reports.")
    d = Err.Description
    On Error GoTo 0
    Report "gpath: an unquoted backslash path is refused with a path-specific hint", _
           InStr(1, d, "file path", vbTextCompare) > 0 And _
           InStr(1, d, "quotes", vbTextCompare) > 0, _
           "got: " & d

    EnglishResetGrammar
End Sub
