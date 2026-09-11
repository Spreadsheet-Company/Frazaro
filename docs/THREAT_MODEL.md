# THREAT MODEL — SEC.0

*What a program may reach; what a phrasebook may reach; who is trusted at
each layer; what "zero-trust runtime" actually promises today versus what
an IT reviewer will assume it promises. Written from reading the real
dispatch code (`VLA_Interpreter.bas`, `VLA.bas`), not from the roadmap's
own secondhand description of it — every claim below is cited by file and,
where useful, by line, and marked CONFIRMED (read directly, or reproduced)
or OPEN QUESTION (named, not yet resolved) rather than asserted either way.
`SEC.1`–`SEC.3`, `SEC.6` all depend on this document existing; nothing here
is itself a fix — this is the artifact `SEC.0` asks for, no more.*

---

## 0. The one thing to get right before anything else

**"Zero-trust runtime" (`IN.9`, `BETA_ROADMAP1.md`) does not mean what an
IT reviewer will assume it means.** IN.9's own text: the interpreter is the
default runtime specifically because it needs *"no VBProject trust"* — it
runs VLA source directly against a live workbook, generating no VBA code,
so the Trust Center's "Trust access to the VBA project object model"
setting is never required. That is the *entire* claim: no macro-trust
prompt at install/first-run.

An IT reviewer hearing "zero-trust runtime" will very plausibly assume the
modern security-architecture sense instead — nothing is trusted by
default, every capability is explicitly granted, least privilege
throughout. **That is not true today**, per every finding below. This gap
— between "doesn't need a Trust Center prompt" and "verifies every
capability before granting it" — is this document's own single most
important finding, because it is the one a procurement conversation is
most likely to walk into blind.

---

## 1. What a program may reach — the real object-model surface

Confirmed by reading `VLA_Interpreter.bas` directly. Four tiers, in
increasing order of how ungoverned they are.

### 1.1 Native dispatch (bounded, enumerated, safe)

A fixed, hand-written list of Excel object-model members and six "place
helpers" (`Range`/`Cells`/`Rows`/`Columns`/`Worksheets`/`Workbooks`),
called directly as native VBA, never through reflection:

- The six place helpers (`EvalDynamicHead`, ~line 1690) — called
  unqualified (`Range(...)`, not `Application.Range(...)`) because
  `CallByName(Application, "range", ...)` is documented to fail against
  Excel's parameterized default members (a real, live-found gap, not a
  guess — see the comment at `VLA_Interpreter.bas:1677`).
- `Application.WorksheetFunction`'s seven most-used members, dispatched
  natively for the identical reason (`~line 1756`) — CallByName against a
  parameterized member here doesn't raise, it silently returns garbage
  (found live: `sum-of` returned a raw-pointer-shaped number instead of a
  real sum, no error at all).
- `DynamicSet`'s own ~20 named `.`-settable members (`Value`, `Size`,
  `Color`, `Formula`, `Bold`, `Italic`, `HorizontalAlignment`,
  `ColumnWidth`, `ColorIndex`, `NumberFormat`, `LineStyle`, `Hidden`,
  `RowHeight`, `Name`, `FreezePanes`, `WrapText`, plus the five
  bare-dotted-global `Application` members `Calculation`/`CutCopyMode`/
  `DisplayAlerts`/`ScreenUpdating`/`StatusBar` — `VLA_Interpreter.bas`
  ~2301 onward) — census-built from real corpus usage, not designed
  up front (`IN.3`'s own history note names the exact live incidents that
  grew this list one member at a time). `G-FORMAT` slice 1 added five,
  each a formatting property with no effect beyond the cell it styles:
  `Underline`, `Strikethrough`, `VerticalAlignment`, `IndentLevel`,
  `Orientation` (mirrored on the read side), and taught the existing
  `Borders` read to honour its one index argument instead of ignoring
  it. No method was added — the outline border is four edge writes, not
  `Range.BorderAround`. `G-SORTFILTER` added one settable member,
  `Worksheet.AutoFilterMode`, which Excel lets a program set to `False`
  only — it takes a sheet's filter buttons away and shows every row, and
  changes no cell — plus three reads, each navigation to cells already on
  the sheet: `Range.Columns(n)` (a sort key), `Worksheet.UsedRange`
  ("Sort this sheet") and `Range.SpecialCells(type)` ("Copy only the
  visible cells of"). It widened two existing named-argument calls
  rather than adding methods: `Range.Sort` takes an optional second key
  (`Key2`/`Order2`) and `Range.AutoFilter` an optional second condition
  (`Operator`/`Criteria2`). Turning filter buttons on is a runtime
  helper (`VlaAddFilters`, tier 1.3), not a new member, because
  `Range.AutoFilter` with no arguments is a toggle and the helper exists
  to make "add" never take them away.
- `NeutralizeFormulaInjection` (`SEC.4`, shipped) now guards the one
  member in this tier with a real injection risk (`Value`) against a
  leading `=`/`+`/`-`/`@`.

This tier is safe *by construction*: every reachable member is a line of
VBA someone wrote and can audit. It is also, honestly, the smallest tier.

### 1.2 The `CallByName` fallback (unbounded — the real finding)

**STATUS UPDATE (SEC.1 Tier 2, this session, awaiting the owner's own
live verification before this finding is called closed):** the fallback
described below has been removed from `DynamicGet`/`DynamicCall`/
`DynamicSet` (`VLA_Interpreter.bas`). A full-repo census (every
`(. obj member...)` and bare dotted-global shape across `scripts/*.vla`,
`scripts/polyglotta/*.vla`, and the `src/*.bas` test suites) found 25
distinct members still reached only through it — not only the G-TABLES
surface SEC.1's own text anticipated, but also plain cell-`.value` reads
and cross-sheet `.range` lookups the census almost missed because they
appear mainly in the host-test suite, not the shipped corpus text. Each
was promoted to its own fixed, audited native `Select Case` arm (Tier 0
by the same definition as every member already there); anything still
unmatched now refuses in words instead of reaching arbitrary late-bound
dispatch. The audit narrative below is preserved as-is — it is still an
accurate account of how this hole was found and why detection-based
alternatives don't work here, and everything in it up to the removal
itself remains true history.

**CONFIRMED, this is the audit finding the roadmap's own `SEC` section
intro names, verified by reading the code rather than trusted on the
intro's own word:** any member name reaching `DynamicGet`/`DynamicSet`/
`DynamicCall` that is *not* in the native list above (§1.1) falls through
to VBA's own `CallByName(obj, member, ...)` (`VLA_Interpreter.bas:2398`,
`2427`). This is a **heuristic, not a capability gate** — the code's own
comment at `~2057` says exactly that. There is no allowlist and no
denylist: if VBA's own late-bound reflection can resolve `member` against
whatever object `obj` currently is (a `Range`, or `Application` itself,
reached via a bare dotted-global like `application.somemember`), it runs.
A name that doesn't resolve raises VBA error 1004, caught and reported as
"not handled" — a *availability* failure, not a *security* one.

**What this concretely means:** a phrasebook rule that can construct a
dotted-member reference to `Application` (already reachable — see the
five bare-dotted-globals above, which are exactly this shape, just
pre-enumerated) can attempt to reach *any* `Application` member VBA's own
`CallByName` can resolve, not just the ~20 already-known ones. Nothing
in the interpreter distinguishes "a member we've audited and expect" from
"a member nobody has looked at yet." The list in §1.1 grew by live
incident, not by security review — three separate corpus runs each found
a new gap the hard way (`Value`, `Size`/`Color`, `screenupdating`).

**OPEN QUESTION, not resolved here:** the practical ceiling of what
`Application`'s own member surface actually allows (file-system access via
`Application.FileDialog`, `Shell`-adjacent members, etc.) has not been
enumerated. `SEC.1` is the item that turns this heuristic into an actual
capability gate; this document only establishes that no such gate exists
today.

### 1.3 `Application.Run` against `VLA_Runtime.bas` (bounded to one module,
confirmed — narrower than it first looks)

`TryRuntimeHelper` (`VLA_Interpreter.bas:2020`) dispatches "nearly all
`VLA_Runtime.bas` helpers... through ONE mechanism" via
`Application.Run("'" & ThisWorkbook.Name & "'!VLA_Runtime." & MangleIdent(h))`
(`~line 2037`). **Read carefully, not assumed:** the module prefix
(`VLA_Runtime.`) is a hardcoded string literal, never derived from user or
phrasebook input — only `h` (the head symbol, run through `MangleIdent`)
is program-controlled. This means the reachable surface through this
specific mechanism is bounded to `VLA_Runtime.bas`'s own `Public`
procedures, by name, with program-author-controlled arguments — **not**
an arbitrary module or an arbitrary project, and not a path to code
outside this codebase's own already-published surface. Real, but
narrower than "Application.Run" read in isolation suggests.

### 1.4 `raw` — the ungoverned ceiling

**CONFIRMED, `VLA.bas:3288`:** `Case "raw": EmitTop = StrLitContent(Nth(lst, 2)) & vbCrLf`.
A `(raw "...")` form's own string content is spliced **verbatim, with zero
validation or transformation**, directly into the emitted VBA module. Once
that module runs, its contents run with full VBA privilege — file I/O,
`Shell`, COM automation of other applications, anything VBA itself can do.
This is real arbitrary-code-execution capability, and it is reachable from
any phrasebook rule whose expansion happens to produce a `(raw ...)` form,
exactly as reachable as any other form in the language.

**No static analysis can ever bound this, by design** — this project's own
already-adjudicated position (`SEC.1`'s own roadmap text): `raw`'s payload
is opaque text, the same way `eval(base64_decode(...))` defeats a source
scanner. This is not a gap to be closed by cleverer detection; it is a
capability that needs **consent**, not analysis. That is exactly `SEC.2`'s
own scope, and exactly why it is scoped as "behind explicit per-phrasebook
consent" rather than "detected and refused."

---

## 2. What a phrasebook may reach, and who is trusted at each layer

**CONFIRMED, the honest finding this section exists to state plainly:**
today, there are no distinct trust layers. A `.vla` file is loaded the
same way regardless of where it came from — the base corpus shipped with
the product, an organization's own internal phrasebook, or a phrasebook
downloaded from a community source all load through the identical
mechanism (`VLA_English.bas`'s vocabulary loader, now `VLA_SentenceEngine.bas`
post-`LX.5`), with no provenance tag distinguishing them and no consent
prompt gating what any of them may contain. A rule using `raw` from a
community-sourced phrasebook loads exactly as silently as one from the
base corpus.

The roadmap's own SEC tranche already names four layers as the *intended*
model, worth recording here as the target this document's findings are
measured against, not as something already built:

1. **Base corpus** — shipped with the product, presumably the
   highest-trust tier by construction (the owner wrote it).
2. **Org phrasebook** — authored or vetted by the organization running
   Frazaro.
3. **Community phrasebook** — sourced from outside the organization,
   lowest presumptive trust.
4. **A `raw`-bearing rule specifically** — a cross-cutting concern
   regardless of which of the above three tiers it appears in, since
   `raw`'s own ceiling (§1.4) is the same no matter who wrote it.

`GO.3` already calls the phrasebook registry "a supply chain" with no
mechanism (per `SEC.3`'s own citation) — this section is that same
observation, grounded in the actual loader rather than the registry
concept.

---

## 3. What this document is not

Not a resolved architecture, and not a vulnerability report — the
roadmap's own `SEC` section intro already draws this line and it is worth
repeating: the `Application`-reachability and `raw`-no-consent findings
above are **audit findings, not confirmed live exploits.** Nothing here
was exploited against a real workbook; everything here was read directly
from the dispatch code and, where the code's own comments already
documented a live incident (`Value`, `WorksheetFunction`, `screenupdating`),
cited to that incident rather than re-asserted from scratch.

Not a complete enumeration of `Application`'s own member surface (§1.2's
own open question) — that is real, uncompleted work, not an oversight.

Not `SEC.1`/`SEC.2`/`SEC.3` themselves — this document is what makes
those items buildable against real findings instead of a restated
intuition; it fixes nothing on its own.

---

## 4. What each downstream `SEC.*` item closes, precisely

- **`SEC.1`** turns §1.2's heuristic (whatever `CallByName` can resolve,
  runs) into an actual capability gate — the roadmap's own adjudication
  already on record: signature/pattern detection over *forms* was
  considered and rejected, because a form's own danger depends on what its
  member-name argument resolves to *at runtime*, undecidable from the
  form's shape alone.
- **`SEC.2`** closes §1.4's consent gap specifically — `raw` behind an
  explicit `requires: capability:raw` declaration and a trust dialog,
  shippable ahead of `SEC.1`'s own full generalization.
- **`SEC.3`** answers §2's own "who do I hold accountable" question once
  `SEC.1`'s tiers exist to attribute an effect to.
- **`SEC.6`** is the point at which this document's own findings get a
  real external security review — gated specifically on §1.2 becoming
  default-deny and §1.4 becoming consent-gated, not before, since
  reviewing a surface that is about to change is reviewing the wrong one.
