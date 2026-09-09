Attribute VB_Name = "VLA_Messages"
Option Explicit
Public Const VLA_MESSAGES_VERSION As String = "LINTERPOLATE.0"
' LINTERPOLATE.0: twelve new entries (six vla-interpolate-*/six
' interp-interpolate-*) for L-INTERPOLATE's new
' (interpolate tpl :key val ...) primitive - VLA.bas's EmitExpr and
' VLA_Interpreter.bas's EvalExpr each own an "interpolate" Case; see
' either module's own LINTERPOLATE.0 header note for the full design.
' Named "interpolate", not CL's "format" - this codebase's own
' format-as-currency/-percent/-date (english.vla) already own "format"
' for visual cell styling; a second, unrelated "format" primitive doing
' string substitution would collide in READING even with no symbol
' collision in the code (owner's own call, this session). A dangling
' ":key" with no value reuses the existing vla-keyword-arg-missing-
' value/interp-keyword-arg-missing-value instead of minting a 13th id -
' already generic, already accurate, no interpolate-specific wording
' needed.
' PF4C.0: one new entry, interp-for-each-row-needs-range - PF.4c's own
' interpreter-side refusal when for-each-row's range slot doesn't
' evaluate to a real Range.

' =====================================================================
'  VLA_Messages - SD-2/LX.2: every refusal goes through RaiseMsg with a
'  stable id and named parameters, instead of a hand-built string at
'  the call site. English is the only catalogue today - LX.2's own
'  "English is the default catalogue" scoping; a future non-English
'  catalogue would swap AddEntries' bodies without any call site
'  changing, since a call site only ever names an id plus values.
'
'  LAYER:     0
'  MAY CALL:  (nothing, with one named exception: AddEntries references
'             VLA.VLA_ERR_INTERPRETER_ONLY by its fully-qualified name,
'             for the one catalogue entry that must preserve that
'             specific non-5/53 error number - VLA_Tests_Host.bas
'             already references it the same fully-qualified way, so
'             this is an existing cross-module contract, not a new one)
'  SHIPS:     add-in
'  PAYS INTO: every refusal site migrated off raw Err.Raise (F.14 tracks
'             the remaining raw count per module, and its own held
'             ceiling drops as each module migrates).
'  REASON:    SD-2 declared this; LX.2 is the artifact. Ids are semantic
'             slugs (SD-9's never-reused-never-re-minted discipline, at
'             message-id scale rather than section-id scale) so a
'             stable id survives the English text next to it being
'             edited freely.
'
'  NOT cached at module level - Catalogue() rebuilds the Collection on
'  every RaiseMsg call. This deliberately does NOT match LX.2's own
'  scoping note, which described this as the "lazy, build-once shape
'  VlaHeadTableAliasMap already proved" - checked against that
'  function's actual body (VLA_HeadTable.bas) rather than trusted, and
'  it builds fresh on every call, on purpose: "a recompile drops module
'  state" (LESSONS.md XXXII) is that function's own stated reason, and
'  it applies here at least as strongly, since a refusal is exactly the
'  code path most likely to run once, unpredictably, in the middle of a
'  live dev-reload session where a stale cache would be hardest to
'  notice. A refusal is never a hot loop; rebuilding a few hundred
'  Collection.Add calls costs nothing a human would notice.
' =====================================================================

Private Function Catalogue() As Collection
    Dim m As New Collection
    AddEntries m
    Set Catalogue = m
End Function

' One AddMsg call per id: id, then the VBA error number and Err.Source
' tag exactly as the site used before migration (SD-2 adds the id
' alongside what a raise already carries, it does not change
' Err.Number/Err.Source semantics), then the English template.
' {slotName} is VLA_English.bas's own phrase-rule slot syntax
' (AddPhraseRule's "{v:var}", "{e:expr}", ...), reused on the
' substitution side rather than inventing a second templating syntax.
'
' Migrated so far (LX.2): VLA_Loader.bas (3), VLA_Lint.bas (2),
' VLA_IDE.bas (15 of 16 - line ~1978's `Err.Raise Err.Number, Err.Source,
' Err.Description` stays raw on purpose: a bare re-raise inside a
' cleanup block, propagating whatever error was already caught, not an
' origination of new English text - there is no id for an arbitrary
' caught error, so it is infrastructure like VLA_Messages's own three
' internal checks, not a refusal site in SD-2's sense). The remaining
' ~262 raw sites across VLA.bas/VLA_English.bas/VLA_Interpreter.bas/
' VLA_Runtime.bas are unmigrated - each still raises its own inline
' Err.Raise, tracked by F.14's ratchet.
Private Sub AddEntries(ByVal m As Collection)
    AddMsg m, "vla-source-not-found", 53, "VLA", "VLA source file not found: {path}"
    AddMsg m, "vla-protected-module-name", 5, "VLA", "'{name}' would overwrite the VLA transpiler itself; rename the file or pass an explicit module name"
    AddMsg m, "vla-file-not-found", 53, "VLA", "File not found: {path}"
    AddMsg m, "lint-trailing-tokens", 5, "VLA-Lint", "trailing tokens after top-level form: {text}"
    AddMsg m, "lint-not-a-list", 5, "VLA-Lint", "top-level form is not a list: {form}"
    AddMsg m, "ide-no-workbook", 5, "VLA-IDE", "Open a workbook first"
    AddMsg m, "ide-program-name-too-long", 5, "VLA-IDE", "That name is too long for a sheet tab - up to 20 characters, please"
    AddMsg m, "ide-sheet-name-bad-char", 5, "VLA-IDE", "Sheet names cannot contain {char} - pick a name without it"
    AddMsg m, "ide-program-name-too-similar", 5, "VLA-IDE", "'{new}' is too similar to the existing program '{existing}' (both shorten to '{tag}') - pick a more different name"
    AddMsg m, "ide-compile-needs-windows", 5, "VLA-IDE", "Compiling needs Windows Excel for now: compiling a program generates VBA into the workbook, and Mac Excel does not allow that kind of access. Check Instructions works here - write and check on Mac, or use Interpret Instructions, which needs no VBA access at all."
    AddMsg m, "ide-compile-needs-trust", 5, "VLA-IDE", "Compiling instructions needs ""Trust access to the VBA project object model"" (File > Options > Trust Center > Trust Center Settings > Macro Settings), which this Excel does not currently grant - use Interpret Instructions instead; it runs the same program without needing that access."
    AddMsg m, "ide-export-needs-trust", 5, "VLA-IDE", "Exporting needs ""Trust access to the VBA project object model"" (File > Options > Trust Center > Trust Center Settings > Macro Settings), which this Excel does not currently grant."
    AddMsg m, "ide-lint-folder-needs-windows", 5, "VLA-IDE", "Linting a whole folder needs Windows Excel's folder-picker dialog, which Mac Excel does not provide - use Lint VLA on one file at a time instead; it works the same on both."
    AddMsg m, "ide-vocab-not-found", 53, "VLA-IDE", "Vocabulary file not found: {path} - edit IdeVocabPath in module VLA_IDE, and this add-in carries no built-in copy either."
    AddMsg m, "ide-no-workspace-sheet", 5, "VLA-IDE", "No '{sheet}' sheet yet - run Set Up Workspace first"
    AddMsg m, "ide-workspace-ambiguous", 5, "VLA-IDE", "This workbook has {count} programs ({names}) - go to the program's own sheet and use its buttons"
    AddMsg m, "ide-programs-share-short-name", 5, "VLA-IDE", "The programs '{a}' and '{b}' would share the short name '{tag}' - rename one of them so each program keeps its own module and its own Undo"
    AddMsg m, "ide-autoload-no-slot", 5, "VLA-IDE", "Could not register for auto-load - no free OPEN slot found in 50 tries ({key})."
    AddMsg m, "ide-program-file-moved", 53, "VLA-IDE", "The program file has moved: {path} - use Import to pick it again"
    AddMsg m, "ide-word-read-failed", 5, "VLA-IDE", "Could not read the Word document (is Word installed?): {detail}"
    ' SEC.9 deliberately adds NO id here. Every refusal it can produce is
    ' a decision the person just made in a dialog, not a fault to report
    ' back to them - the gate skips and records rather than raising, so
    ' that a declined phrasebook cannot trap them in an error on every
    ' command (nothing in this codebase removes an entry from
    ' VLA_LoadedPhrasebooks). See VLA_IDE.bas's SEC.9 header for why that
    ' departs from SEC.2's raise-on-decline shape, and the roadmap entry
    ' for the reporting gap it accepts.
    ' LX2.0 correction, same session: 22 VLA_Runtime.bas entries that
    ' briefly lived here (runtime-not-a-color through
    ' runtime-pivot-sort-ambiguous-field) were removed after a live
    ' crash - see VLA_Runtime.bas's own header comment for the full
    ' story. That module's EN_RUNTIME INJECT BOUNDARY discipline means
    ' code above the boundary ships into user workbooks with no other
    ' add-in module present, VLA_Messages included, so those 22 sites
    ' went back to raw Err.Raise. Only the one call below the boundary
    ' (add-in-side only, never shipped standalone) stays migrated.
    AddMsg m, "runtime-helpers-unreadable", 5, "VLA-Runtime", "The runtime helpers could not be read (no live module and no VLAr_Source sheet) - rebuild the add-in with VlaBuildAddin"
    ' VLA_Interpreter.bas (60 of 61 - line ~1199's `Err.Raise num, src,
    ' desc` inside ExecStmtTrapped stays raw on purpose: a re-raise of an
    ' already-caught error escaping a local On Error Resume Next, not an
    ' origination of new English text, same category as VLA_IDE.bas's
    ' line ~1972). Several ids below are deliberately reused across more
    ' than one call site: same conceptual refusal, same exact English
    ' text, different place in the file that can trigger it (the
    ' goto/resume-label-not-found id fires from 3 sites, the two
    ' '.'-form-shape ids and the dispatch-arg-count id each from 2) -
    ' one id per REFUSAL, not one per call site, matching this project's
    ' own semantic-slug naming intent.
    AddMsg m, "interp-goto-label-not-found", 5, "VLA-Interpreter", "IN3.6: goto/resume target label '{label}' was never found"
    AddMsg m, "interp-entry-proc-not-found", 5, "VLA-Interpreter", "VlaInterpretEntry: no procedure named '{proc}' in this source"
    AddMsg m, "interp-expr-not-single", 5, "VLA-Interpreter", "VlaEvalExpression: expected exactly one expression, got {count}"
    AddMsg m, "interp-on-error-bad-shape", 5, "VLA-Interpreter", "IN3.6: on-error: expected (on-error resume-next) or (on-error goto <label|0>)"
    AddMsg m, "interp-resume-bare-unsupported", 5, "VLA-Interpreter", "IN3.6: bare (resume) is not supported - this interpreter's Try:-shaped on-error only ever generates (resume <label>)"
    AddMsg m, "interp-resume-next-unsupported", 5, "VLA-Interpreter", "IN3.6: (resume next) is not supported - this interpreter's Try:-shaped on-error only ever generates (resume <label>)"
    AddMsg m, "interp-dot-needs-object-member", 5, "VLA-Interpreter", "(. obj member) needs an object and a member"
    AddMsg m, "interp-dot-needs-object-left", 5, "VLA-Interpreter", "IN.2: '.' needs an object to the left of the member - got a plain value"
    AddMsg m, "interp-set-place-not-object", 5, "VLA-Interpreter", "IN.2: computed set! place did not evaluate to an object"
    AddMsg m, "interp-for-each-row-needs-range", 5, "VLA-Interpreter", "PF.4c: for-each-row's second header slot must evaluate to a Range"
    AddMsg m, "interp-select-bad-clause", 5, "VLA-Interpreter", "select: expected (case (values...) ...) or (case-else ...), got '({head} ...)'"
    AddMsg m, "interp-dot-call-needs-object-member", 5, "VLA-Interpreter", "(. obj member ...) needs an object and a member"
    AddMsg m, "interp-err-field-unsupported", 5, "VLA-Interpreter", "IN3.6: 'err.{field}' is not supported - this interpreter's caught-error shadow only tracks description/number/source"
    AddMsg m, "interp-new-unsupported-type", 5, "VLA-Interpreter", "IN.3: '(new {type})' - this interpreter only supports (new Collection) today"
    AddMsg m, "interp-quote-arity", 5, "VLA-Interpreter", "quote takes exactly one datum - wrap a list to quote several: (quote (a b c))"
    AddMsg m, "interp-range-arity", 5, "VLA-Interpreter", "IN.2: 'range' takes 1 or 2 arguments, got {n}"
    AddMsg m, "interp-cells-arity", 5, "VLA-Interpreter", "IN.2: 'cells' takes 0 or 2 arguments, got {n}"
    AddMsg m, "interp-rows-arity", 5, "VLA-Interpreter", "IN.2: 'rows' takes 0 or 1 argument, got {n}"
    AddMsg m, "interp-columns-arity", 5, "VLA-Interpreter", "IN.2: 'columns' takes 0 or 1 argument, got {n}"
    AddMsg m, "interp-worksheets-arity", 5, "VLA-Interpreter", "IN.2: 'worksheets' takes 0 or 1 argument, got {n}"
    AddMsg m, "interp-workbooks-arity", 5, "VLA-Interpreter", "IN.2: 'workbooks' takes 0 or 1 argument, got {n}"
    AddMsg m, "interp-make-button-arity", 5, "VLA-Interpreter", "IN.7: 'make-button' takes 2 arguments (caption, cell), got {n}"
    AddMsg m, "interp-sum-arity", 5, "VLA-Interpreter", "IN.3: 'sum' takes 1 argument, got {n}"
    AddMsg m, "interp-average-arity", 5, "VLA-Interpreter", "IN.3: 'average' takes 1 argument, got {n}"
    AddMsg m, "interp-max-arity", 5, "VLA-Interpreter", "IN.3: 'max' takes 1 argument, got {n}"
    AddMsg m, "interp-min-arity", 5, "VLA-Interpreter", "IN.3: 'min' takes 1 argument, got {n}"
    AddMsg m, "interp-countif-arity", 5, "VLA-Interpreter", "IN.3: 'countif' takes 2 arguments, got {n}"
    AddMsg m, "interp-sumif-arity", 5, "VLA-Interpreter", "IN.3: 'sumif' takes 3 arguments, got {n}"
    AddMsg m, "interp-vlookup-arity", 5, "VLA-Interpreter", "IN.3: 'vlookup' takes 4 arguments, got {n}"
    AddMsg m, "interp-head-unresolved", 5, "VLA-Interpreter", "'{head}' is not a form, place helper, dotted global, built-in, or VLA_Runtime helper this interpreter can reach yet - IN.2's remaining hand-work"
    ' IN.15: the OTHER half of the question interp-head-unresolved used to
    ' answer alone. TryRuntimeHelper now settles "is this a real helper"
    ' by name against VlaHelperManifest BEFORE calling anything, so an
    ' error after the call can no longer mean "unknown name" - it means a
    ' helper the manifest vouched for could not be called with these
    ' arguments. Saying that is the point: reporting it as "not a form"
    ' was a confident wrong answer, which this project holds to be worse
    ' than a crash.
    AddMsg m, "interp-runtime-helper-call-failed", 5, "VLA-Interpreter", "IN.15: '{head}' is a runtime helper, but this call to it could not be completed - Excel reported error {num}: {desc}. Check the number and kind of values being passed to it."
    AddMsg m, "interp-make-button-not-cell", 5, "VLA-Interpreter", "IN.7: 'make-button' needs a cell for its place, got {type}"
    AddMsg m, "interp-dynamic-member-refused", 5, "VLA-Interpreter", "SEC.1: '{member}' is not in this interpreter's native dynamic-dispatch allowlist - refused, not attempted via CallByName (Tier 2's own subtractive fix removed that fallback; a real, legitimate use belongs in VLA_Interpreter.bas's own DynamicGet/DynamicCall/DynamicSet, reviewed and added by name)"
    ' SEC.8: provenance, not verb. SEC.1's refusal above asks "is this
    ' member reachable at all"; this one asks "may the workbook CARRYING
    ' this program reach outside the workbook". The remedy named in the
    ' text is Windows' own Unblock deliberately - SEC.8 keeps no trust
    ' store of its own, so there is nothing for Frazaro to offer here
    ' and nothing a hostile workbook can pre-fill. See VLA_Provenance.
    AddMsg m, "sec8-untrusted-workbook-effect", 5, "VLA-Provenance", "SEC.8: '{verb}' would have an effect outside this workbook, and {why} - so it is refused. The workbook is {path}. If you trust it: close it, right-click the file in File Explorer, choose Properties, tick Unblock, then reopen it. (Frazaro deliberately has no button of its own for this - a workbook that arrives from outside must not be able to carry its own permission slip.)"
    AddMsg m, "interp-member-chain-not-object", 5, "VLA-Interpreter", "IN.2: '{segment}' in '{path}' did not return an object - cannot continue the member chain"
    AddMsg m, "interp-positional-args-too-many", 5, "VLA-Interpreter", "IN.2: this interpreter's dynamic dispatch supports up to 4 positional arguments; got {n}"
    AddMsg m, "interp-named-args-not-supported-here", 5, "VLA-Interpreter", "IN.2: named arguments ({tok} ...) are not supported here - a statement-position '.' call where EVERY argument is a keyword pair goes through IN2.5's named-argument dispatch instead; this one either mixes positional and named arguments, or is in expression position, neither of which is supported yet"
    AddMsg m, "interp-expected-keyword-arg", 5, "VLA-Interpreter", "IN.2: expected a keyword argument (':name value') at position {pos}, got a plain value - this call mixes positional and named arguments, which this interpreter's named-argument dispatch does not support yet"
    AddMsg m, "interp-keyword-arg-missing-value", 5, "VLA-Interpreter", "IN.2: keyword argument '{tok}' is missing its value"
    AddMsg m, "interp-interpolate-template-must-be-literal", 5, "VLA-Interpreter", "L-INTERPOLATE: interpolate's first argument must be a literal string template, not a computed value - got {got}"
    AddMsg m, "interp-interpolate-expected-keyword-arg", 5, "VLA-Interpreter", "L-INTERPOLATE: interpolate's arguments after the template must all be keyword pairs (a colon-prefixed name, then its value) - expected one at position {pos}, got a plain value"
    AddMsg m, "interp-interpolate-unclosed-hole", 5, "VLA-Interpreter", "L-INTERPOLATE: interpolate template has an unclosed hole - an opening brace needs a matching closing brace before the template ends; double a brace to use it as a literal character"
    AddMsg m, "interp-interpolate-empty-hole-name", 5, "VLA-Interpreter", "L-INTERPOLATE: interpolate template has an empty hole - every hole needs a name between its braces"
    AddMsg m, "interp-interpolate-unknown-key", 5, "VLA-Interpreter", "L-INTERPOLATE: interpolate template references '{key}', which was never given as a keyword argument"
    AddMsg m, "interp-interpolate-unused-argument", 5, "VLA-Interpreter", "L-INTERPOLATE: interpolate was given the keyword argument '{key}', but it is never used as a hole in the template"
    AddMsg m, "interp-missing-keyword-arg", 5, "VLA-Interpreter", "IN.2: missing required keyword argument '{key}'"
    AddMsg m, "interp-too-many-positional-args", 5, "VLA-Interpreter", "IN.10: call passes {given} positional argument(s) but only {declared} parameter(s) are declared"
    AddMsg m, "interp-missing-required-arg", 5, "VLA-Interpreter", "IN.10: call is missing required argument '{name}'"
    AddMsg m, "interp-missing-required-named-arg", 5, "VLA-Interpreter", "IN.10: call is missing required named argument '{name}'"
    AddMsg m, "interp-paramarray-unsupported", 5, "VLA-Interpreter", "IN.10: this interpreter's user-procedure call/return does not support (paramarray ...) parameters yet"
    AddMsg m, "interp-named-args-unsupported-member", 5, "VLA-Interpreter", "IN.2: named arguments to '{member}' are not supported yet - only add/protect/unprotect/close/printout/exportasfixedformat/copy/sort/pastespecial/replace/removeduplicates/autofilter, this corpus's own real uses"
    AddMsg m, "interp-dispatch-missing-arg", 5, "VLA-Interpreter", "IN.2: dynamic dispatch call is missing argument {n}"
    AddMsg m, "interp-opchain-unknown-operator", 5, "VLA-Interpreter", "IN2.7: '{op}' is not one of EvalOpChain's 18 operators"
    AddMsg m, "interp-form-missing-element", 5, "VLA-Interpreter", "form is missing required element {n}"
    AddMsg m, "interp-expected-symbol-got-list", 5, "VLA-Interpreter", "expected a symbol, got a list"
    AddMsg m, "interp-expected-symbol-got-string", 5, "VLA-Interpreter", "expected a symbol, got a string literal"
    AddMsg m, "interp-empty-form", 5, "VLA-Interpreter", "empty form ()"
    AddMsg m, "interp-expected-string-got-list", 5, "VLA-Interpreter", "expected a string literal, got a list"
    AddMsg m, "interp-expected-string-got-other", 5, "VLA-Interpreter", "expected a string literal, got '{value}'"
    AddMsg m, "interp-dev-file-not-found", 53, "VLA-Interpreter", "{filename} not found - looked beside the workbook ({beside}), in its scripts folder ({scripts}), and in scripts/polyglotta ({polyglotta})"
    ' VLA_English.bas (96 of 97 - line ~7648's `Err.Raise errNum, errSrc,
    ' errDesc` inside EnglishAuditText's cleanup stays raw on purpose,
    ' same re-raise category as VLA_IDE.bas/VLA_Interpreter.bas's own
    ' exceptions; its own comment already documents exactly why: Err's
    ' properties are captured to locals before re-raising because live
    ' Err access during Err.Raise's own argument evaluation can misbehave.
    ' Several sites here embed LITERAL curly braces in their static
    ' English text (this file's whole domain is slot syntax, which uses
    ' {name}/{name:category} literally) - those are passed as ordinary
    ' string VALUES (e.g. "example", "{name}"), never written as literal
    ' braces inside a template, so RaiseMsg's own {slot} scanner never
    ' misreads them as a substitution it must resolve.
    AddMsg m, "english-sheet-change-dup", 5, "VLA-English", "there is already a 'When the sheet changes:' handler in this program - one program, one handler{loc}"
    AddMsg m, "english-click-handler-dup", 5, "VLA-English", "there is already a 'When ""{caption}"" is clicked:' handler in this program - one handler per button{loc}"
    AddMsg m, "english-click-handler-slug-collision", 5, "VLA-English", "'When ""{a}"" is clicked:' and 'When ""{b}"" is clicked:' compile to the same internal name once punctuation and spacing fold away - Compile could not tell the two handlers apart; rename one of the two button captions{loc}"
    AddMsg m, "english-param-name-taken", 5, "VLA-English", "'{name}' already means something (a built-in or vocabulary word) - pick another parameter name{loc}"
    AddMsg m, "english-param-missing-default", 5, "VLA-English", "parameter '{name}' has no default, but an earlier parameter does - once one parameter has a default ('of ...'), all the ones after it need defaults too{loc}"
    AddMsg m, "english-param-required-after-optional", 5, "VLA-English", "parameter '{name}' is required, but an earlier parameter is optional - once one parameter is optional (or has a default), all the ones after it must be too{loc}"
    AddMsg m, "english-old-slot-spelling", 5, "VLA-English", "'{t}' uses the old question-mark slot spelling - slots are written in braces now: {suggested} in the pattern, and {example} in the template"
    AddMsg m, "english-slot-not-closed", 5, "VLA-English", "'{t}' looks like a slot but isn't closed - slots are {ex1} or {ex2}, no spaces"
    AddMsg m, "english-malformed-slot", 5, "VLA-English", "'{t}' is a malformed slot - write {ex1} or {ex2}"
    AddMsg m, "english-malformed-default", 5, "VLA-English", "'{t}' is a malformed default - write {ex}, both sides non-empty"
    AddMsg m, "english-empty-pattern", 5, "VLA-English", "empty phrase pattern"
    AddMsg m, "english-override-no-match", 5, "VLA-English", "the override '{pattern}' matches no earlier rule - use the plain (non-override) directive, or check that the rule it replaces still exists (and loads first)"
    AddMsg m, "english-override-ambiguous", 5, "VLA-English", "the override '{pattern}' is ambiguous - it matches {count} earlier rules ({list}) and an override replaces exactly one"
    AddMsg m, "english-override-matches-builtin", 5, "VLA-English", "the override '{pattern}' matches the built-in rule '{builtin}' - the built-in core is not overridable; write a differently-worded rule instead"
    AddMsg m, "english-rule-shadow", 5, "VLA-English", "{dupMsg}" & vbCrLf & "  If this replacement is intentional, use this rule's own <lingua>-vla-override directive instead of <lingua>-vla."
    AddMsg m, "english-template-not-well-formed", 5, "VLA-English", "template does not parse as well-formed VLA: {template}{detail}"
    AddMsg m, "english-slot-value-not-one-form", 5, "VLA-English", "F.2: bound value for slot '{slot}' did not read back as exactly one form ({count}) - {value}"
    AddMsg m, "english-slot-glued-to-identifier", 5, "VLA-English", "F.2: slot '{quoted}' is glued onto a larger identifier in its template - only a bare word can bind there, got: {value}"
    AddMsg m, "english-render-no-category-renderer", 5, "VLA-English", "G-RENDER: no renderer for slot category ':{cat}'"
    AddMsg m, "english-render-no-matching-rule", 5, "VLA-English", "G-RENDER: no loaded rule's template form matches {form}"
    AddMsg m, "english-render-not-single-form", 5, "VLA-English", "G-RENDER: EnglishRenderText expects exactly one top-level form, got {count}"
    AddMsg m, "english-alternation-empty-branch", 5, "VLA-English", "pattern '{pattern}': alternation '{token}' has an empty branch - write {example}, every branch a word"
    AddMsg m, "english-unknown-slot-category", 5, "VLA-English", "pattern '{pattern}': unknown slot category ':{cat}' - categories are name, var, text, expr, cond, range, cell, column, sheet, color, path (or a list of one: text-list, range-list, cell-list, column-list, sheet-list, color-list); or write an alternation {example}"
    AddMsg m, "english-optional-empty-branch", 5, "VLA-English", "pattern '{pattern}': optional '{token}' has an empty branch - write [word] or [word|word]"
    AddMsg m, "english-bare-alternation-empty-branch", 5, "VLA-English", "pattern '{pattern}': alternation '{token}' has an empty branch - write word|word, every branch a word"
    AddMsg m, "english-optional-not-closed", 5, "VLA-English", "'{token}' looks like an optional literal but isn't closed - optionals are [word], one word, no spaces"
    AddMsg m, "english-surface-spec-multi-slash", 5, "VLA-English", "pattern '{pattern}': {where} '{spec}' has more than one '/' - a form is word or stem/suffix; write further spellings as separate branches"
    AddMsg m, "english-surface-spec-empty-side", 5, "VLA-English", "pattern '{pattern}': {where} '{spec}' has an empty side around '/' - write stem/suffix, both parts non-empty"
    AddMsg m, "english-form-never-closes", 5, "VLA-English", "this VLA form never closes - {depth} '(' still open by the end of the program{loc}"
    AddMsg m, "english-unknown-character", 5, "VLA-English", "I don't understand the character '{char}'{hint}{loc}"
    AddMsg m, "english-library-name-not-quoted", 5, "VLA-English", "The library name must be quoted - file names contain periods, and a period ends a sentence. Write: Use library ""helpers.vla"". near: '{context}'{loc}"
    AddMsg m, "english-path-not-quoted", 5, "VLA-English", "A file path must be quoted - file names contain periods, and a period ends a sentence. Write: Open workbook ""C:\Reports\file.xlsx"". (or bind a variable first: Set report-path to ""C:\Reports\file.xlsx"". then Open workbook report-path.) near: '{context}'{loc}"
    AddMsg m, "english-expected-near", 5, "VLA-English", "Expected {what} near: '{context}'{loc}"
    AddMsg m, "english-name-has-colon", 5, "VLA-English", "A name cannot contain ':' - got '{word}' near: '{context}'{loc}"
    AddMsg m, "english-expected-word-near", 5, "VLA-English", "Expected '{word}' near: '{context}'{loc}"
    AddMsg m, "english-top-level-definition-in-row", 5, "VLA-English", "'({head} ...)' is a top-level VLA definition - a sheet row is a statement inside your program. Define subs, macros, and types in a module or vocabulary instead{loc}"
    AddMsg m, "english-form-balance-hint", 5, "VLA-English", "{hint}{loc}"
    AddMsg m, "english-form-doesnt-transpile", 5, "VLA-English", "this VLA form doesn't transpile - {detail}{loc}"
    AddMsg m, "english-expected-sentence-blank-line", 5, "VLA-English", "Expected a sentence but found a blank line"
    AddMsg m, "english-if-fails-misplaced", 5, "VLA-English", "'If that fails:' must start its own paragraph right after the Try block it rescues - add a blank line (or 'Done.') before it, and make sure a 'Try:' comes first - near: '{context}'{loc}"
    AddMsg m, "english-stop-loop-outside-loop", 5, "VLA-English", "'Stop the loop.' only makes sense inside a loop - near: '{context}'"
    AddMsg m, "english-percent-mixed-amount", 5, "VLA-English", "'{sentence}' mixes % into a longer amount - a plain share works ('Increase total by 10%.' grows it by 10%), and a plain number works, but a blend has two readings and I will not pick one. Compute the amount first ('Set step to ...') and then {suggestion}{loc}"
    AddMsg m, "english-standalone-get", 5, "VLA-English", "'Get {fn} using ...' is a value, not an instruction - put the value somewhere: 'Set fee to get {fn} using ...' (or use it inside a condition){loc}"
    AddMsg m, "english-give-back-outside-action", 5, "VLA-English", "'Give back' only makes sense inside a value-returning action ('To tax of amount: ... Give back ...') - to show a value use Show, to stop use Stop - near: '{context}'{loc}"
    AddMsg m, "english-when-it-is-first", 5, "VLA-English", "'When it is ...' continues a choice - the first sentence names the value being examined, like 'When region is ""North"":' - near: '{context}'{loc}"
    AddMsg m, "english-click-handler-needs-quotes", 5, "VLA-English", "a button click handler's name needs quotes - use 'When ""{tok}"" is clicked:' instead of 'When {tok} is clicked:'{loc}"
    AddMsg m, "english-clicked-misspelled", 5, "VLA-English", "'When {tok} is {misspelled}:' looks like a button click handler with 'clicked' misspelled as '{misspelled}' - use 'When {tok} is clicked:'{loc}"
    AddMsg m, "english-button-not-created-with-create", 5, "VLA-English", "a button isn't created with 'Create ... called ...' - use 'Make a button ""<caption>"" at cell <cell>.' instead{loc}"
    AddMsg m, "english-pivot-not-created-with-create", 5, "VLA-English", "a pivot table isn't created with 'Create ... called ...' - use 'Make a pivot table from <range> at <cell> called ""<name>"".' instead{loc}"
    AddMsg m, "english-create-unknown-kind", 5, "VLA-English", "Create what? Say 'a number', 'a text', 'a value', 'a list', or 'a lookup', not '{kind}'"
    AddMsg m, "english-to-not-top-level", 5, "VLA-English", "'To ...' definitions must be at the top level, not inside a block"
    AddMsg m, "english-define-not-top-level", 5, "VLA-English", "'Define ...' must be at the top level, not inside a block"
    AddMsg m, "english-unexpected-token-block", 5, "VLA-English", "Unexpected '{tok}' — is there a stray 'Done.' or a missing block?"
    AddMsg m, "english-parse-error", 5, "VLA-English", "{msg}"
    AddMsg m, "english-unknown-slot-category-runtime", 5, "VLA-English", "Unknown slot category ':{cat}' in a phrase pattern"
    AddMsg m, "english-action-unknown-param", 5, "VLA-English", "the action '{action}' has no parameter called '{param}'. Its parameters are: {list}{loc}"
    AddMsg m, "english-param-given-twice", 5, "VLA-English", "the parameter '{param}' was given twice in this call to '{action}'{loc}"
    AddMsg m, "english-call-missing-required-param", 5, "VLA-English", "the call to '{action}' is missing its required parameter '{param}'. Its parameters are: {list}{loc}"
    AddMsg m, "english-expected-value", 5, "VLA-English", "Expected a value near: '{context}'"
    AddMsg m, "english-expected-condition", 5, "VLA-English", "Expected a condition near: '{context}'"
    AddMsg m, "english-define-value-immutable", 5, "VLA-English", "'{name}' was given a fixed value by Define and cannot be changed - use a different name{loc}"
    AddMsg m, "english-reserved-word-name", 5, "VLA-English", "'{name}' is a reserved word in Excel's programming language and cannot be used as a name - a hyphenated name like '{name}-row' or 'my-{name}' always works"
    AddMsg m, "english-value-word-name", 5, "VLA-English", "'{name}' is a language word for a value and cannot be used as a name - try '{name}-stamp' or similar"
    AddMsg m, "english-action-name-taken", 5, "VLA-English", "there is already an action called '{name}' in this program - one name, one definition{loc}"
    AddMsg m, "english-action-name-means-something", 5, "VLA-English", "'{name} of ...' already means something (a built-in or vocabulary word) - pick another name for the action{loc}"
    AddMsg m, "english-vocab-file-not-found", 53, "VLA-English", "Vocabulary file not found: {path}"
    AddMsg m, "english-vocab-nesting-too-deep", 5, "VLA-English", "{loc}: a generator's own expansion nests (begin ...) more than 20 levels deep - almost certainly a mistake, not a real vocabulary"
    AddMsg m, "english-vocab-expected-directive", 5, "VLA-English", "{loc}: expected a directive (a parenthesized form), got a bare word '{word}'"
    AddMsg m, "english-vocab-at-row-arity", 5, "VLA-English", "{loc}: (at-row label form) takes exactly two arguments - a row label and one wrapped form (use (begin ...) to wrap several)"
    AddMsg m, "english-vocab-test-fail-empty-fragment", 5, "VLA-English", "{loc}: a test-fail proof protects the MESSAGE, not just the refusal - give a non-empty fragment"
    AddMsg m, "english-vocab-expansion-not-directive", 5, "VLA-English", "{loc}: a generator's expansion produced '{head}', which is not a vocabulary directive (english-vla/test-success/test-fail/defmacro/english-function, or an override) - check the generator macro's own body"
    AddMsg m, "english-vocab-unrecognized-directive", 5, "VLA-English", "{loc}: unrecognized top-level directive '{head}'"
    ' F.10: a phrasebook's own declared preconditions. Every one of
    ' these refuses BEFORE a single rule from the file registers, so a
    ' phrasebook is never half-loaded on the strength of a requirement
    ' that turned out not to hold.
    AddMsg m, "english-vocab-requires-version-unmet", 5, "VLA-English", "{loc}: this phrasebook needs Frazaro {wanted} or newer; this is {have} - upgrade Frazaro, or use a build of the phrasebook written for {have}"
    AddMsg m, "english-vocab-requires-version-malformed", 5, "VLA-English", "{loc}: '{wanted}' is not a version - a (requires-version ...) declaration takes a plain MAJOR.MINOR.PATCH number in quotes, like (requires-version ""0.5.2"")"
    AddMsg m, "english-vocab-requires-missing-value", 5, "VLA-English", "{loc}: this (requires-{namespace} ...) declaration names no value - it takes exactly one, in quotes, like (requires-{namespace} ""...""); a bare word is not one"
    AddMsg m, "english-vocab-requires-unknown-namespace", 5, "VLA-English", "{loc}: '{namespace}' is not a kind of requirement this build understands (it knows version, capability and form), so it cannot confirm the requirement was met and will not load the phrasebook - this phrasebook was most likely written for a newer Frazaro than {have}, and a (requires-version ...) declaration would have said exactly which one"
    AddMsg m, "english-vocab-requires-capability-ungranted", 5, "VLA-English", "{loc}: this phrasebook requires the '{capability}' capability, and no capability can be granted yet - the consent step that would let you approve one is not built (SEC.7). Nothing can load this phrasebook today; remove the declaration only if the phrasebook genuinely does not use that capability"
    AddMsg m, "english-vocab-requires-form-unsupported", 5, "VLA-English", "{loc}: (requires-form ""{form}"") is understood but not yet enforceable - the record of which release each form first worked in (docs/GRAMMAR_SINCE.md) is not carried inside Frazaro yet, so this build cannot check it and refuses rather than passing it silently. Use (requires-version ...) for now"
    AddMsg m, "english-vocab-requires-from-expansion", 5, "VLA-English", "{loc}: '{head}' was produced by a generator's expansion, not written in the file - a requirement is read from the source text before anything loads, so one that only appears after expansion is never actually checked. Write it literally, at the top of the phrasebook"
    AddMsg m, "english-vocab-macro-expansion-failed", 5, "VLA-English", "{loc}: this directive names a macro ('{head}') but its own expansion failed - {detail}"
    AddMsg m, "english-vocab-file-already-carried-same", 5, "VLA-English", "{loc}: this session already carries this file - reset first (EnglishResetGrammar) or use the IDE's Reload, which resets for you"
    AddMsg m, "english-vocab-macro-name-collision", 5, "VLA-English", "{loc}: a macro named '{name}' is already carried by {prev} - rename this one (macro override semantics are not defined)"
    AddMsg m, "english-vocab-macro-form-invalid", 5, "VLA-English", "{loc}: this macro form does not stand - {detail}"
    AddMsg m, "english-extra-words-after-statement", 5, "VLA-English", "extra words after the first statement"
    AddMsg m, "english-test-failed", 5, "VLA-English", "{loc}: test FAILED" & vbCrLf & "  sentence: {sentence}" & vbCrLf & "  expected: {expected}" & vbCrLf & "  got:      {got}"
    AddMsg m, "english-test-failed-to-translate", 5, "VLA-English", "{loc}: test FAILED to translate" & vbCrLf & "  sentence: {sentence}" & vbCrLf & "  error:    {error}"
    AddMsg m, "english-failtest-translated", 5, "VLA-English", "{loc}: fail: proof FAILED - the sentence translated instead of refusing" & vbCrLf & "  sentence: {sentence}" & vbCrLf & "  became:   {became}"
    AddMsg m, "english-failtest-message-drifted", 5, "VLA-English", "{loc}: fail: proof FAILED - the refusal message drifted" & vbCrLf & "  sentence: {sentence}" & vbCrLf & "  wanted:   {wanted}" & vbCrLf & "  message:  {message}"
    AddMsg m, "english-call-unknown-param", 5, "VLA-English", "'{call}': the action '{action}' has no parameter called '{param}'. Its parameters are: {list}{loc}"
    AddMsg m, "english-call-missing-param", 5, "VLA-English", "'{call}': the action '{action}' requires '{param}' - add: with {param} of ...{loc}"
    AddMsg m, "english-program-file-not-found", 53, "VLA-English", "Program file not found: {path}"
    AddMsg m, "english-translate-vla-overwrite", 5, "VLA-English", "Translate to VLA: refusing to overwrite the source file ({path}) - point this at a .txt/.en program, not an already-.vla file."
    AddMsg m, "english-translate-vba-overwrite", 5, "VLA-English", "Translate to VBA: refusing to overwrite the source file ({path}) - point this at a .txt/.en program, not an already-.vba file."
    AddMsg m, "english-define-needs-fixed-value", 5, "VLA-English", "Define needs a fixed value - quoted text (""#FF69B4""), a number, or an earlier alias - near: '{context}'"
    AddMsg m, "english-function-word-not-one-word", 5, "VLA-English", "{context}: a function word is one word, optionally followed by 'of'"
    ' SEC.2: raw behind explicit, per-phrasebook consent.
    AddMsg m, "english-vocab-raw-consent-declined", 5, "VLA-English", "'{source}' was not loaded - it contains a (raw ...) form, which runs unrestricted VBA once a program using it runs, and consent for it was declined"
    ' VLA.bas (134 of 141 - lines ~563/565/568/570 (VlaTranspile's
    ' emitfail handler), ~688 and ~1084 (VlaReadForms/a second parse
    ' path's own fail: handlers), and ~2186 (VlaExpandStepText's
    ' restoreBudget) all stay raw on purpose: every one re-raises an
    ' ALREADY-CAUGHT error (VlaTranspile's own emitfail additionally
    ' appends a location suffix to Err.Description before re-raising,
    ' but Err.Number/Source and the substance of Description still come
    ' from whatever inner call already failed - often itself already
    ' migrated, with its own id - so this is propagation/annotation, not
    ' origination of new English text, the same category as every other
    ' file's own re-raise exception). "vla-defmacro-reserved-name" alone
    ' covers 22 call sites - every reserved primitive name defmacro
    ' refuses - sharing one id with two named slots rather than 22
    ' near-identical entries; "vla-keyword-misuse" covers 4 sites that
    ' each just relay KwMisuseMsg's own return value verbatim.
    ' "vla-interpreter-only-handler" preserves VLA_ERR_INTERPRETER_ONLY
    ' (VLA.bas's own named non-5 error number, vbObjectError + 7001) via
    ' a fully-qualified cross-module reference - see this module's own
    ' header note on why that is the one deliberate exception to Layer 0.
    AddMsg m, "vla-pop-context-empty", 5, "VLA", "VlaPopContext: no matching VlaPushContext - the context stack is empty"
    AddMsg m, "vla-module-not-ours", 5, "VLA", "'{name}' already exists in this workbook's VBA project and was not created by Frazaro - overwriting it would destroy whatever code is there. Frazaro chooses this name itself, so the fix is on the module, not the program: open the VBA editor (Alt+F11) and rename or remove the '{name}' module, then try again."
    AddMsg m, "vla-prelude-not-found", 53, "VLA", "prelude.vla not found - looked beside this workbook and in its scripts folder, and this add-in carries no built-in copy either. Every VLA compile needs it."
    AddMsg m, "vla-include-too-deep", 5, "VLA", "include: nesting deeper than 16 files - is a file including itself? (while splicing '{label}')"
    AddMsg m, "vla-include-cannot-read", 5, "VLA", "include: cannot read '{name}' (looked at '{path}')"
    AddMsg m, "vla-unbalanced-parens", 5, "VLA", "unbalanced parentheses: missing ')'{loc}"
    AddMsg m, "vla-unexpected-close-paren", 5, "VLA", "unexpected ')'{loc}"
    AddMsg m, "vla-form-missing-element", 5, "VLA", "form is missing required element {n}"
    AddMsg m, "vla-expected-list-at-position", 5, "VLA", "expected a list at position {n}, got '{value}'"
    AddMsg m, "vla-expected-symbol-got-list", 5, "VLA", "expected a symbol, got a list"
    AddMsg m, "vla-expected-symbol-got-string", 5, "VLA", "expected a symbol, got a string literal"
    AddMsg m, "vla-empty-form", 5, "VLA", "empty form ()"
    AddMsg m, "vla-expected-string-got-list", 5, "VLA", "expected a string literal, got a list"
    AddMsg m, "vla-expected-string-got-other", 5, "VLA", "expected a string literal, got '{value}'"
    AddMsg m, "vla-defmacro-reserved-name", 5, "VLA", "defmacro '{word}': '{word}' is the {desc} and cannot be a macro name"
    AddMsg m, "vla-defmacro-rest-param-missing", 5, "VLA", "defmacro '{name}': '&' must be followed by a rest parameter"
    AddMsg m, "vla-defmacro-rest-param-not-last", 5, "VLA", "defmacro '{name}': rest parameter must be last"
    AddMsg m, "vla-defmacro-missing-template", 5, "VLA", "defmacro '{name}': missing template body"
    AddMsg m, "vla-doc-arity", 5, "VLA", "(doc <macro-name>) names one macro"
    AddMsg m, "vla-macro-expansion-too-deep", 5, "VLA", "macro expansion too deep (recursive macro?)"
    AddMsg m, "vla-quasiquote-outside-template", 5, "VLA", "'{head}' only has meaning inside a defmacro's own template - there is no substitution environment outside one"
    AddMsg m, "vla-macro-self-expansion-limit", 5, "VLA", "macro self-expansion exceeded 5000 applications in one chain (infinite macro?)"
    AddMsg m, "vla-macro-arity", 5, "VLA", "macro '{name}' expects {n} argument(s), got {given}"
    AddMsg m, "vla-macro-arity-min", 5, "VLA", "macro '{name}' expects at least {n} argument(s), got {given}"
    AddMsg m, "vla-quasiquote-nested", 5, "VLA", "quasiquote: nested quasiquote is not supported"
    AddMsg m, "vla-quasiquote-arity", 5, "VLA", "quasiquote expects exactly 1 argument, got {n}"
    AddMsg m, "vla-unquote-outside-quasiquote", 5, "VLA", "unquote used outside quasiquote"
    AddMsg m, "vla-unquote-arity", 5, "VLA", "unquote expects exactly 1 argument, got {n}"
    AddMsg m, "vla-unquote-splicing-outside-quasiquote", 5, "VLA", "unquote-splicing used outside quasiquote"
    AddMsg m, "vla-unquote-splicing-not-element", 5, "VLA", "unquote-splicing must appear as a list element, not as a value"
    AddMsg m, "vla-unquote-splicing-arity", 5, "VLA", "unquote-splicing expects exactly 1 argument, got {n}"
    AddMsg m, "vla-unquote-splicing-not-list", 5, "VLA", "unquote-splicing: operand did not resolve to a list"
    AddMsg m, "vla-symbol-arity", 5, "VLA", "symbol expects at least 1 argument, got 0"
    AddMsg m, "vla-symbol-arg-is-list", 5, "VLA", "symbol: argument {n} is a list, expected an atom or string literal"
    AddMsg m, "vla-quote-arity", 5, "VLA", "quote takes exactly one datum - wrap a list to quote several: (quote (a b c))"
    AddMsg m, "vla-car-arity", 5, "VLA", "car expects exactly 1 argument, got {n}"
    AddMsg m, "vla-car-expected-list", 5, "VLA", "car: expected a list, got '{value}'"
    AddMsg m, "vla-car-empty-list", 5, "VLA", "car: expected a non-empty list"
    AddMsg m, "vla-list-too-short", 5, "VLA", "{caller}: expected a list of at least {n} element(s), got {count}"
    AddMsg m, "vla-cdr-arity", 5, "VLA", "cdr expects exactly 1 argument, got {n}"
    AddMsg m, "vla-cdr-expected-list", 5, "VLA", "cdr: expected a list, got '{value}'"
    AddMsg m, "vla-cddr-arity", 5, "VLA", "cddr expects exactly 1 argument, got {n}"
    AddMsg m, "vla-cddr-expected-list", 5, "VLA", "cddr: expected a list, got '{value}'"
    AddMsg m, "vla-cons-arity", 5, "VLA", "cons expects exactly 2 arguments, got {n}"
    AddMsg m, "vla-cons-second-not-list", 5, "VLA", "cons: second argument must be a list, got '{value}'"
    AddMsg m, "vla-nullq-arity", 5, "VLA", "null? expects exactly 1 argument, got {n}"
    AddMsg m, "vla-eq-arity", 5, "VLA", "eq? expects exactly 2 arguments, got {n}"
    AddMsg m, "vla-eq-expected-atom", 5, "VLA", "eq?: expected an atom, got a list (use equal? for structural comparison)"
    AddMsg m, "vla-equal-arity", 5, "VLA", "equal? expects exactly 2 arguments, got {n}"
    AddMsg m, "vla-quote-if-arity", 5, "VLA", "quote-if expects exactly 3 arguments (test then-form else-form), got {n}"
    AddMsg m, "vla-quote-if-test-is-list", 5, "VLA", "quote-if: test must resolve to true or false, got a list"
    AddMsg m, "vla-quote-if-test-invalid", 5, "VLA", "quote-if: test must resolve to true or false, got '{value}'"
    AddMsg m, "vla-arith-expand-arity", 5, "VLA", "{op} expects exactly 2 arguments, got {n}"
    AddMsg m, "vla-arith-expand-not-number-list", 5, "VLA", "{op}: expected a number, got a list"
    AddMsg m, "vla-arith-expand-not-number-value", 5, "VLA", "{op}: expected a number, got '{value}'"
    AddMsg m, "vla-cond-clause-shape", 5, "VLA", "cond: each clause must be (test form) or (else form), got {n} element(s)"
    AddMsg m, "vla-cond-else-not-last", 5, "VLA", "cond: 'else' must be the last clause"
    AddMsg m, "vla-module-level-decl-after-proc", 5, "VLA", "a module-level ({what} ...) must come before the first (sub ...) or (function ...) - VBA ignores declarations placed after a procedure, and every use of the name then fails Option Explicit. Move the declaration above the procedures. (A library include splices whole procedures and belongs at the BOTTOM of the program; libraries should carry procedures, not module-level declarations.)"
    AddMsg m, "vla-top-level-not-list", 5, "VLA", "top-level form must be a list, got '{value}'"
    AddMsg m, "vla-deflambda-runs", 5, "VLA", "(deflambda ...) runs when the program runs - put it inside a sub (main is the usual place); it registers the worksheet function there"
    AddMsg m, "vla-include-must-stand-alone", 5, "VLA", "'(include ...)' splices a whole file and must stand alone on its own line at the top level - it cannot appear inside a procedure, an expression, or beside other forms"
    AddMsg m, "vla-unknown-top-level-form", 5, "VLA", "unknown top-level form '{head}' (executable statements must live inside a sub or function)"
    AddMsg m, "vla-interpreter-only-handler", VLA.VLA_ERR_INTERPRETER_ONLY, "VLA", "'{name}' (IN.7) is an interpreter-native event handler and cannot be compiled into VBA - Interpret this program instead of using Compile/Compile and Trace"
    AddMsg m, "vla-bare-atom-statement", 5, "VLA", "bare atom used as a statement: '{value}'"
    AddMsg m, "vla-type-enum-module-level-only", 5, "VLA", "({head} ...) is module-level only - move it outside the sub or function"
    AddMsg m, "vla-on-error-bad-shape", 5, "VLA", "on-error: expected (on-error resume-next) or (on-error goto <label|0>)"
    AddMsg m, "vla-quote-in-statement-position", 5, "VLA", "'(quote ...)' is an expression, not a statement - use its value ((set! x (quote ...))) or evaluate it with VlaTryValue"
    AddMsg m, "vla-at-line-arity", 5, "VLA", "at-line expects a line number and at least one statement"
    AddMsg m, "vla-at-line-missing-number", 5, "VLA", "at-line expects a line number"
    AddMsg m, "vla-at-line-not-a-number", 5, "VLA", "at-line expects a line number, got '{value}'"
    AddMsg m, "vla-gen-row-arity", 5, "VLA", "gen-row expects a row label and at least one statement"
    AddMsg m, "vla-clause-outside-parent", 5, "VLA", "clause '({head} ...)' used outside its parent form"
    AddMsg m, "vla-operator-in-statement-position", 5, "VLA", "'({op} ...)' is an expression, not a statement - use its value ((set! x ({op} ...))) or evaluate it with VlaTryValue"
    AddMsg m, "vla-keyword-misuse", 5, "VLA", "{msg}"
    AddMsg m, "vla-return-outside-function", 5, "VLA", "(return <value>) is only valid inside a function"
    AddMsg m, "vla-tco-arity", 5, "VLA", "tail call to '{name}' passes {given} argument(s) but it takes {declared}"
    AddMsg m, "vla-if-bad-clause", 5, "VLA", "if: expected (then ...), (elseif ...), or (else ...), got '({head} ...)'"
    AddMsg m, "vla-select-bad-clause", 5, "VLA", "select: expected (case (values...) ...) or (case-else ...), got '({head} ...)'"
    AddMsg m, "vla-decl-bad-type", 5, "VLA", "declaration of '{name}': expected a type name or (array ...), got '({head} ...)'"
    AddMsg m, "vla-array-bound-arity", 5, "VLA", "array bound: (to lower upper) needs exactly two bounds"
    AddMsg m, "vla-redim-missing-bound", 5, "VLA", "redim '{name}': at least one bound is required (ReDim a() is not legal VBA)"
    AddMsg m, "vla-type-no-members", 5, "VLA", "type '{name}': VBA requires at least one member"
    AddMsg m, "vla-enum-no-members", 5, "VLA", "enum '{name}': VBA requires at least one member"
    AddMsg m, "vla-visibility-wraps-unknown", 5, "VLA", "({vis} ...) wraps a sub, function, dim, const, type, or enum - got '({head} ...)'"
    AddMsg m, "vla-dot-needs-object-member", 5, "VLA", "(. obj member ...) needs an object and a member"
    AddMsg m, "vla-deflambda-arity", 5, "VLA", "deflambda: (deflambda name (params) [""doc""] body-expression)"
    AddMsg m, "vla-deflambda-body-not-one-formula", 5, "VLA", "deflambda '{name}': the body is ONE formula expression (got {n} forms) - statements cannot run inside a worksheet function"
    AddMsg m, "vla-deflambda-params-not-plain", 5, "VLA", "deflambda '{name}': parameters are plain names"
    AddMsg m, "vla-formula-no-named-args", 5, "VLA", "deflambda: '{tok}' - formulas have no named arguments"
    AddMsg m, "vla-formula-mod-arity", 5, "VLA", "deflambda: Excel's MOD takes two operands - (mod a b)"
    AddMsg m, "vla-formula-if-then-block", 5, "VLA", "deflambda: formula if is the three-arg spelling - (if test then-value else-value) - not the statement (then ...) blocks"
    AddMsg m, "vla-formula-if-arity", 5, "VLA", "deflambda: formula if is (if test then-value else-value)"
    AddMsg m, "vla-formula-lambda-arity", 5, "VLA", "deflambda: (lambda (params) body-expression)"
    AddMsg m, "vla-formula-lambda-params-not-plain", 5, "VLA", "deflambda: lambda parameters are plain names"
    AddMsg m, "vla-formula-is-statement", 5, "VLA", "deflambda: '({head} ...)' is a statement - a worksheet function's body is ONE formula expression"
    AddMsg m, "vla-formula-array-no-nest", 5, "VLA", "deflambda: Excel array constants do not nest - (quote ...) in a formula must be flat"
    AddMsg m, "vla-keyword-arg-missing-value", 5, "VLA", "keyword argument '{tok}' is missing its value"
    AddMsg m, "vla-interpolate-template-must-be-literal", 5, "VLA", "L-INTERPOLATE: interpolate's first argument must be a literal string template, not a computed value - got {got}"
    AddMsg m, "vla-interpolate-expected-keyword-arg", 5, "VLA", "L-INTERPOLATE: interpolate's arguments after the template must all be keyword pairs (a colon-prefixed name, then its value) - expected one at position {pos}, got a plain value"
    AddMsg m, "vla-interpolate-unclosed-hole", 5, "VLA", "L-INTERPOLATE: interpolate template has an unclosed hole - an opening brace needs a matching closing brace before the template ends; double a brace to use it as a literal character"
    AddMsg m, "vla-interpolate-empty-hole-name", 5, "VLA", "L-INTERPOLATE: interpolate template has an empty hole - every hole needs a name between its braces"
    AddMsg m, "vla-interpolate-unknown-key", 5, "VLA", "L-INTERPOLATE: interpolate template references '{key}', which was never given as a keyword argument"
    AddMsg m, "vla-interpolate-unused-argument", 5, "VLA", "L-INTERPOLATE: interpolate was given the keyword argument '{key}', but it is never used as a hole in the template"
    AddMsg m, "vla-operator-needs-two-operands", 5, "VLA", "operator '{op}' needs at least two operands"
    AddMsg m, "vla-name-no-plain-letters", 5, "VLA", "'{value}' has no plain letters or digits once its accented characters are converted - give it a name with at least one a-z letter in it"

    ' CO.4 - the one refusal grammar semantic versioning needs. A
    ' version that cannot be read is never treated as satisfied AND
    ' never quietly treated as unmet: either alone would hide the typo
    ' that caused it. Names the shape wanted, not just the value
    ' rejected (LX.8's doctrine).
    AddMsg m, "vla-version-malformed", 5, "VLA", "'{value}' is not a version this build can compare. A version is three whole numbers separated by dots, like 0.5.1 - MAJOR.MINOR.PATCH (SD-14). A suffix such as '-beta' is not read here. Fix the version text: a version that cannot be read is never treated as satisfied."

    ' PROLOG.1 - VLA_Unify.bas's own refusal: a malformed glued-slot
    ' atom (two embedded slots in one atom) is a template/program-
    ' authoring defect regardless of which DSL is calling, so this
    ' module owns the wording directly rather than deferring it to each
    ' caller the way VLA_Relation's own LAYER 0.5 functions do. Replaces
    ' the former "english-render-glue-multiple-slots" (VLA_English.bas's
    ' own UnifyForm raised it directly before this item hoisted the
    ' check into the shared walker) - no caller raises the old id
    ' anymore.
    AddMsg m, "unify-glue-multiple-slots", 5, "VLA-Unify", "'{atom}' glues more than one slot into a single atom - not supported yet"
    AddMsg m, "prolog-occurs-check", 5, "VLA-Unify", "variable {var} can't bind to a term that contains itself - that would build an infinite term, which can never render, spill to a worksheet, or be audited"

    ' PROLOG.3/PROLOG.4 - VLA_Prolog.bas's own refusals: parse-time shape
    ' checks for (fact ...)/(rule ...)/(query ...) forms, cross-
    ' definition arity consistency (DATALOG's own RecordArity precedent,
    ' repeated - see VLA_Prolog.bas's own header for why this is a
    ' deliberate small duplication, not a missed hoist), the resolution-
    ' step ceiling (PROLOG.4's own termination guarantee - PROLOG has
    ' none by construction, unlike DATALOG's provably-terminating
    ' fixpoint), and the worksheet-facing signature's own stated limit
    ' (table-sourced facts are PROLOG.6's job, not yet built - a table
    ' argument passed before then is refused by name, never silently
    ' ignored).
    AddMsg m, "prolog-top-form-not-a-list", 5, "VLA-Prolog", "expected a form like (fact ...), (rule ...), or (query ...), but found something else."
    AddMsg m, "prolog-top-form-empty", 5, "VLA-Prolog", "found an empty () where a top-level (fact ...)/(rule ...)/(query ...) form was expected."
    AddMsg m, "prolog-top-form-bad-head", 5, "VLA-Prolog", "a top-level form's own head must be a plain word (fact/rule/query), not a nested form or a quoted string."
    AddMsg m, "prolog-unknown-top-form", 5, "VLA-Prolog", "'{head}' isn't a form PROLOG recognizes at the top level - only (fact ...), (rule ...), and (query ...) are supported."
    AddMsg m, "prolog-atom-not-a-list", 5, "VLA-Prolog", "expected a predicate form like (name arg1 arg2 ...) in {context}, but found something else."
    AddMsg m, "prolog-atom-empty", 5, "VLA-Prolog", "found an empty () where a predicate form was expected, in {context}."
    AddMsg m, "prolog-predicate-name-not-symbol", 5, "VLA-Prolog", "a predicate's name must be a plain word, in {context} - not a nested form or a quoted string."
    AddMsg m, "prolog-fact-bad-shape", 5, "VLA-Prolog", "(fact ...) takes exactly one predicate form, like (fact (parent tom bob))."
    AddMsg m, "prolog-fact-has-variable", 5, "VLA-Prolog", "predicate '{predicate}' in a (fact ...) has a variable ({var}) - facts must be fully ground; a variable only ever belongs in a (query ...)."
    AddMsg m, "prolog-rule-needs-body", 5, "VLA-Prolog", "(rule ...) needs a head plus at least one body predicate, like (rule (grandparent X Z) (parent X Y) (parent Y Z)) - a head with no body at all is just a fact, so write it as (fact ...) instead."
    AddMsg m, "prolog-arity-mismatch", 5, "VLA-Prolog", "predicate '{predicate}' is used with {a} argument(s) in one place and {b} in another - every use of one predicate must have the same arity."
    AddMsg m, "prolog-query-bad-shape", 5, "VLA-Prolog", "(query ...) needs at least one predicate form, like (query (parent tom X))."
    AddMsg m, "prolog-query-missing", 5, "VLA-Prolog", "this program has no (query ...) - PROLOG needs exactly one, naming what to answer."
    AddMsg m, "prolog-query-ambiguous", 5, "VLA-Prolog", "this program has {count} (query ...) forms - PROLOG needs exactly one."
    AddMsg m, "prolog-step-ceiling", 5, "VLA-Prolog", "this query took more than {steps} resolution steps without finishing, which is almost certainly a rule that recurses without ever reaching a base case rather than a very large proof - check for a rule whose recursive call never gets closer to a fact."

    ' PROLOG.5.1 - `is`/arithmetic. `is`/`not`/`findall`/`!` are reserved
    ' words - forward-declared for PROLOG.5.2-5.4 too, even though only
    ' `is` is built this item, so a knowledge base written against 5.1
    ' alone can never be silently broken once the other three ship.
    ' PROLOG.7 extends the same list, and this text with it: the six
    ' comparison operators are reserved on the identical precedent. A
    ' message that still enumerated only four would be quietly wrong about
    ' which names it had just refused, which is the same class of
    ' confidently-wrong answer the {form} rewrite above exists to remove.
    AddMsg m, "prolog-reserved-predicate-name", 5, "VLA-Prolog", "'{name}' is a reserved word in PROLOG (is/not/findall/!/between, the six comparisons < > =< >= =:= =\=, the four term-matching goals = \= == \==, the six type tests var? nonvar? atom? number? atomic? compound?, the six list goals length member nth append reverse sum-list, and the six ISO spellings var nonvar atom number atomic compound - those last six reserved only so PROLOG can point you at the question-mark form instead of failing silently) and can't be used as a predicate name in a (fact ...) or (rule ...)."
    ' `prolog-cut-not-yet-supported` (PROLOG.5.1-5.3-era: "cut (!) isn't
    ' supported yet.") is retired at PROLOG.5.4 - cut is real now, no
    ' code path can raise it anymore, and this project's own precedent
    ' (`unify-glue-multiple-slots`, replacing the former
    ' `english-render-glue-multiple-slots`, this file's own header above)
    ' is to remove a superseded id from the catalogue outright rather
    ' than keep an unreachable entry around.

    ' PROLOG.5.2 - negation-as-failure (`not`). Only the malformed-shape
    ' refusal is new; a malformed inner Goal (a bare non-! atom - a bare
    ' ! is legal from PROLOG.5.4 onward, real Prolog's own opaque-cut
    ' rule) reuses prolog-atom-not-a-list directly via ValidateBodyItem's
    ' own recursive call, rather than inventing a not-specific twin.
    AddMsg m, "prolog-not-bad-shape", 5, "VLA-Prolog", "(not ...) needs exactly one argument - the goal to negate, like (not (likes bob pizza))."

    ' PROLOG.5.3 - findall. Only the malformed-shape refusal is new; a
    ' malformed inner Goal reuses whatever ValidateBodyItem's own
    ' recursive call would already raise for it (the same reuse `not`
    ' already established).
    AddMsg m, "prolog-findall-bad-shape", 5, "VLA-Prolog", "(findall ...) needs exactly three arguments - a template, a goal, and a target list, like (findall X (likes bob X) Bag)."
    AddMsg m, "prolog-is-bad-shape", 5, "VLA-Prolog", "(is ...) needs exactly two arguments - a target and an arithmetic expression, like (is Total (+ X Y))."

    ' PROLOG.7 - the six comparison goals. ONE id for all six, not a
    ' sibling each: the shape rule is identical, so six texts would be six
    ' chances to drift. It is {form}-templated for the same reason the
    ' arithmetic refusals above are - it serves six forms and cannot know
    ' which one is running - and is pinned by the same script, which reads
    ' its own second baseline list for ids that serve more than one form
    ' without being raised from inside the shared evaluator.
    ' `prolog-is-bad-shape` (above) and its not/findall siblings stay
    ' form-SPECIFIC and correctly so: each is raised by exactly one arm,
    ' about exactly one form, and can therefore name it.
    AddMsg m, "prolog-comparison-bad-shape", 5, "VLA-Prolog", "{form} needs exactly two arguments - the two numbers to compare, like (> Salary 80000)."
    ' PROLOG.8 - the four term-matching goals, on the identical terms: ONE
    ' id for all four, {form}-templated because it serves four forms and
    ' cannot know which is running, and carried in the same script's own
    ' multi-form baseline. Kept SEPARATE from the comparison id above even
    ' though both say "exactly two arguments": what the two arguments ARE
    ' differs, and that is the part a user who got the shape wrong needs.
    ' A comparison's operands are NUMBERS - arithmetic expressions the
    ' evaluator will reduce to two Doubles - whereas these four take any
    ' two TERMS at all, atoms and compounds included, and never evaluate
    ' them. Folding the two into one message would have to drop that
    ' distinction to stay true of both, which would make it useless to
    ' both.
    AddMsg m, "prolog-unification-bad-shape", 5, "VLA-Prolog", "{form} needs exactly two arguments - the two terms to match, like (= X 1)."
    ' PROLOG.9 - the six type tests, the widest fan-out of the three: ONE
    ' id for all six, {form}-templated for the same reason and carried in
    ' the same script's own multi-form baseline (added to it BEFORE this
    ' message existed, so the check failed on it until it did). Kept
    ' separate from both ids above on the identical reasoning: these take
    ' ONE argument rather than two, and what that argument is - any term
    ' at all, classified rather than compared or evaluated - is exactly
    ' what a user who wrote (atom X Y) needs told.
    AddMsg m, "prolog-type-test-bad-shape", 5, "VLA-Prolog", "{form} needs exactly one argument - the term to classify, like (number? Salary)."
    ' PROLOG.9 - the six BARE ISO spellings. This engine writes a type
    ' test with a trailing question mark, the convention VLA's own macro
    ' layer already uses for null?/eq?/equal?, so a Prolog author's first
    ' instinct - (atom X) - would otherwise be an unknown predicate, and
    ' an unknown predicate in PROLOG is a SILENT dead end: no rows, no
    ' explanation. One id for all six, {form}-templated for the same
    ' reason its five siblings above are, and carried in the same
    ' script's multi-form baseline. It is the only refusal in this module
    ' whose entire job is to name a spelling.
    AddMsg m, "prolog-type-test-iso-spelling", 5, "VLA-Prolog", "{form} isn't how PROLOG spells this type test - all six of them end in a question mark, so write {fixed} instead, like (number? Salary)."
    ' PROLOG.10 - the adjudication. A text cell reading eng becomes the
    ' term ""eng, and a bare eng written in a query is a DIFFERENT term;
    ' that stays true, and this message does not change it. What changes
    ' is that comparing the two no longer answers silently - (\= D eng)
    ' succeeding on every row of a table was a confidently wrong answer.
    ' Raised ONLY from the four explicit comparison goals, never from
    ' clause matching, so a query that found a real answer is never
    ' aborted by a near-miss against some other clause; VLA_Prolog.bas's
    ' RefuseIfQuotedVersusBare carries the four reasons for that scope.
    ' {kind} is "name" or "number" because the same near-miss happens
    ' between a TEXT 42 and a NUMERIC 42, where calling 42 a name would
    ' be wrong - PROLOG.10's own open sub-question, surfacing in the
    ' wording rather than being papered over.
    AddMsg m, "prolog-quoted-versus-bare", 5, "VLA-Prolog", "the text ""{text}"" and the {kind} {text} are different things in PROLOG - a text cell keeps its quotes, so if you meant the cell's own value write it as ""{text}""."
    ' PROLOG.12 - the same mistake, found the other way round. This one is
    ' raised POST HOC, only when the whole query found nothing at all, so
    ' it opens by saying so: the user is looking at an empty result and
    ' needs to know the empty result is the symptom, not the subject.
    ' Never raised while a query is still solving, and never when it found
    ' anything, so it can only ever turn an empty answer into an explained
    ' one - it cannot break a query that works.
    AddMsg m, "prolog-quoted-versus-bare-no-rows", 5, "VLA-Prolog", "this query found no rows at all, and the likeliest reason is a quoting mismatch: the text ""{text}"" and the {kind} {text} are different things in PROLOG. A text cell keeps its quotes, so if you meant the cell's own value write it as ""{text}""."
    ' PROLOG.9 - between's own four refusals. All four name "(between
    ' ...)" in their own text, and that is CORRECT rather than the defect
    ' the {form} rewrite above removes: this form is the only caller of
    ' any of them, so naming it can never tell a user about a form they
    ' did not write. The three ids above are templated because they serve
    ' six, four and six forms respectively; these serve one.
    ' PROLOG.13 - the six list goals' own shared SHAPE refusal. The
    ' widest fan-out of any id in this module: one raise site serving
    ' length, member, nth, append, reverse and sum-list, and the first
    ' family whose members do not even agree on an arity - length,
    ' reverse and sum-list take two arguments, nth and append three. So
    ' it takes {count} from the caller as well as {form}, and naming
    ' either in its own text would be wrong five times out of six.
    ' Carried in tools/check_prolog_form_attribution.ps1's multi-form
    ' baseline, added there BEFORE this message existed, and the check
    ' failed on it until it did.
    AddMsg m, "prolog-list-bad-shape", 5, "VLA-Prolog", "{form} needs exactly {count} arguments. The six list goals are (length L N), (member X L), (nth N L X), (append A B C), (reverse L R) and (sum-list L N)."
    ' PROLOG.13 - and their shared refusal for an argument that is not a
    ' PROPER LIST. Its whole job is to teach the spelling, because a
    ' list is the one term in PROLOG whose written form a user cannot
    ' guess: there is no [a,b,c] here, since the reader treats [ ] and |
    ' as ordinary letters. So the message shows a real list rather than
    ' describing one. {form}-templated for the same reason as its
    ' sibling above, and in the same baseline.
    '
    ' It fires for a partial list - (cons a T) with T unbound - as well
    ' as for an outright non-list, deliberately. Real Prolog would solve
    ' some of those; this engine says so instead, because the
    ' alternative is a query that quietly finds nothing and cannot be
    ' told apart from one that correctly found nothing.
    AddMsg m, "prolog-list-not-a-list", 5, "VLA-Prolog", "{form} needs a list, but found '{value}'. A list is built with cons and ends in nil - (cons a (cons b nil)) is the list a, b - and the empty list is written nil. A findall bag is already one."
    AddMsg m, "prolog-between-bad-shape", 5, "VLA-Prolog", "(between ...) needs exactly three arguments - a low bound, a high bound, and the value to generate or test, like (between 1 10 X)."
    ' Raised for a bound OR for a bound third argument, since the answer
    ' is the same in both cases and so is the fix: between counts, and a
    ' fractional value has no next value to count to.
    AddMsg m, "prolog-between-not-whole-number", 5, "VLA-Prolog", "(between ...) counts in whole numbers, but '{value}' isn't one."
    ' Deliberately a refusal and not a quiet failure. Answering FALSE here
    ' would be indistinguishable from an in-range miss, so a user could
    ' not tell 'out of range' from 'not the kind of thing between talks
    ' about' - and it is what keeps testing and generating the SAME
    ' relation, since testing then succeeds on exactly the values
    ' generating would produce.
    AddMsg m, "prolog-between-not-a-number", 5, "VLA-Prolog", "(between ...) generates or tests numbers, but its third argument is '{value}', which isn't one - note that a text cell reading 5 is not the number 5."
    ' The range ceiling IS the query's own step ceiling, not a second
    ' budget - see VLA_Prolog.bas's own PROLOG.9 header. This message
    ' exists because without it the same query stops with
    ' prolog-step-ceiling, which blames a runaway rule the user does not
    ' have.
    AddMsg m, "prolog-between-range-too-wide", 5, "VLA-Prolog", "(between {low} {high} ...) would generate {count} values, and a whole query gets {max} resolution steps - narrow the range, or bind the third argument to test one value instead of generating them all."
    ' PROLOG.7: {form}, not a hard-coded "(is ...)". These five refusals
    ' are raised by ValidateArithExpr/EvalArithTerm, which from PROLOG.7
    ' onward serve TWO callers - `(is Var Expr)` and each of the six
    ' comparison goals - so neither procedure can know which form the
    ' user actually wrote. Naming the wrong one is a confidently wrong
    ' answer, which this project holds to be worse than a crash (IN.15).
    ' The caller passes the form label it already knows; a {form}-
    ' templated text was chosen over generalized per-form siblings
    ' precisely so the two wordings cannot drift apart, the same
    ' single-source-of-truth reasoning ValidateArithExpr's own header
    ' already uses for not re-checking numeric-ness at parse time.
    ' tools/check_prolog_form_attribution.ps1 holds this mechanically:
    ' every refusal these two procedures raise must carry {form} and must
    ' not spell a form into its own text. DATALOG.6 inherits this shape
    ' for `datalog-sum-needs-one-value-variable`, whose text is
    ' `sum`-specific for the identical reason.
    AddMsg m, "prolog-arith-unknown-operator", 5, "VLA-Prolog", "'{op}' isn't an arithmetic operator PROLOG recognizes inside {form} - only +, -, *, and / are supported."
    AddMsg m, "prolog-arith-wrong-arity", 5, "VLA-Prolog", "'{op}' inside {form} needs exactly two operands, like (+ X Y)."
    AddMsg m, "prolog-arith-unbound-variable", 5, "VLA-Prolog", "{form} can't compute a value that still has an unbound variable ({var}) in it."
    AddMsg m, "prolog-arith-not-numeric", 5, "VLA-Prolog", "{form} expected a number but found '{value}', which isn't one."
    AddMsg m, "prolog-arith-divide-by-zero", 5, "VLA-Prolog", "{form} tried to divide by zero."

    ' `prolog-tables-not-yet-supported` (PROLOG.3-5.4-era: "table-sourced
    ' facts aren't supported yet...") is retired at PROLOG.6 - tables are
    ' real now, no code path can raise it anymore, the same retirement
    ' this file's own `unify-glue-multiple-slots` header already
    ' precedents.
    '
    ' PROLOG.6 - table-sourced and named-column facts. The first three
    ' mirror DATALOG's own table-argument wording exactly (VLA_Relation.
    ' TableArgResolve's own shared category codes, turned back into
    ' PROLOG's own wording by VLA_Prolog.bas's own thin TableArgName
    ' wrapper, the identical DATALOG precedent). The last three mirror
    ' DATALOG.5's own keyed-atom wording, adapted for PROLOG's own single
    ' unified "not consistently keyed" refusal - PROLOG has no separate
    ' pre-existing compound-term ban to fall back on the way DATALOG does
    ' (PROLOG allows compound-term arguments everywhere), so a malformed
    ' pair shape against a table-sourced predicate is refused here
    ' directly rather than left to a DIFFERENT check to catch.
    AddMsg m, "prolog-table-not-a-range", 5, "VLA-Prolog", "every PROLOG table argument must be a cell range - pass a reference like Employees, not a computed value."
    AddMsg m, "prolog-table-needs-a-name", 5, "VLA-Prolog", "this range has no name PROLOG can use as a predicate - make it an Excel Table (Ctrl+T) or give it a defined name, then reference that name in the formula."
    AddMsg m, "prolog-table-noncontiguous-columns", 5, "VLA-Prolog", "a table argument spanning multiple disjoint areas (a Ctrl-selected, non-contiguous range) isn't supported - select one contiguous block of the table's own columns instead."
    AddMsg m, "prolog-atom-mixed-keying", 5, "VLA-Prolog", "'{predicate}' is table-sourced, but its own arguments here aren't consistently keyed - either key every argument by its own column name, like (name Alice), or use plain positional arguments throughout, never a mix."
    AddMsg m, "prolog-unknown-column", 5, "VLA-Prolog", "'{column}' isn't a column of '{predicate}' - its own columns are: {columns}."
    AddMsg m, "prolog-keyed-column-repeated", 5, "VLA-Prolog", "'{column}' is keyed more than once in the same '{predicate}' atom - each column name may appear at most once per (predicate (col val) ...) form."

    ' DATALOG.0 - VLA_Datalog.bas's own refusals: parse-time shape
    ' checks (compound terms, malformed fact/rule/query forms), the
    ' Datalog safety condition (every head variable must appear in the
    ' body), cross-definition arity consistency, the runaway-recursion
    ' safety valve, and the worksheet-facing table-argument checks.
    AddMsg m, "datalog-compound-term", 5, "VLA-Datalog", "DATALOG doesn't allow one predicate nested inside another's argument, in {context} - write two rules instead of nesting them (this is what keeps every DATALOG() query guaranteed to finish)."
    AddMsg m, "datalog-atom-not-a-list", 5, "VLA-Datalog", "expected a predicate form like (name arg1 arg2 ...) in {context}, but found something else."
    AddMsg m, "datalog-atom-empty", 5, "VLA-Datalog", "found an empty () where a predicate form was expected, in {context}."
    AddMsg m, "datalog-predicate-name-not-symbol", 5, "VLA-Datalog", "a predicate's name must be a plain word, in {context} - not a nested form or a quoted string."
    AddMsg m, "datalog-predicate-needs-argument", 5, "VLA-Datalog", "'{predicate}' has no arguments - DATALOG predicates need at least one, e.g. (fact ({predicate} some-value)) rather than (fact ({predicate}))."
    AddMsg m, "datalog-top-form-not-a-list", 5, "VLA-Datalog", "every top-level line in a DATALOG program must be a (fact ...), (rule ...), or (query ...) form."
    AddMsg m, "datalog-top-form-empty", 5, "VLA-Datalog", "found an empty () at the top level - expected (fact ...), (rule ...), or (query ...)."
    AddMsg m, "datalog-top-form-bad-head", 5, "VLA-Datalog", "a top-level form's first word must be fact, rule, or query - not a nested form."
    AddMsg m, "datalog-fact-bad-shape", 5, "VLA-Datalog", "(fact ...) takes exactly one predicate form, like (fact (parent tom bob))."
    AddMsg m, "datalog-fact-has-variable", 5, "VLA-Datalog", "fact '{predicate}' uses '{arg}', which looks like a variable (it starts with a capital letter) - facts must be fully specific; did you mean to write a rule instead?"
    AddMsg m, "datalog-rule-needs-body", 5, "VLA-Datalog", "a (rule head ...) needs at least one body predicate after the head - a rule with no body is a fact; use (fact ...) instead."
    AddMsg m, "datalog-query-bad-shape", 5, "VLA-Datalog", "(query ...) takes exactly one predicate name, like (query indirect_report)."
    AddMsg m, "datalog-headless-bad-shape", 5, "VLA-Datalog", "(headless) takes no arguments - write it exactly as (headless) to skip the header row and return only data rows."
    AddMsg m, "datalog-query-not-a-symbol", 5, "VLA-Datalog", "(query ...) takes a plain predicate name, not a nested form."
    AddMsg m, "datalog-query-missing", 5, "VLA-Datalog", "add (query predicate-name) to say which relation DATALOG should return - for example (query indirect_report)."
    AddMsg m, "datalog-query-ambiguous", 5, "VLA-Datalog", "found {count} (query ...) forms - DATALOG needs exactly one, naming the single relation to return."
    AddMsg m, "datalog-unknown-top-form", 5, "VLA-Datalog", "'{head}' is not a DATALOG form - expected fact, rule, or query."
    AddMsg m, "datalog-unsafe-head-variable", 5, "VLA-Datalog", "in the rule deriving '{predicate}', the variable '{var}' appears in the head but never in the body - DATALOG can't know what values it should take. Every head variable must also appear in at least one body predicate."
    AddMsg m, "datalog-arity-mismatch", 5, "VLA-Datalog", "predicate '{predicate}' is used with {a} argument(s) in one place and {b} in another - every DATALOG predicate needs one fixed number of arguments everywhere it appears."
    AddMsg m, "datalog-round-ceiling", 5, "VLA-Datalog", "this DATALOG program is still deriving new facts after {rounds} rounds, which is almost certainly a mistake rather than a very large answer - check for a rule whose recursion never narrows."
    AddMsg m, "datalog-query-unknown-predicate", 5, "VLA-Datalog", "(query {predicate}) names a predicate with no facts, no rule, and no matching table - check the spelling, or that the table argument's name matches."
    AddMsg m, "datalog-table-not-a-range", 5, "VLA-Datalog", "every DATALOG table argument must be a cell range - pass a reference like Employees, not a computed value."
    AddMsg m, "datalog-table-needs-a-name", 5, "VLA-Datalog", "this range has no name DATALOG can use as a predicate - make it an Excel Table (Ctrl+T) or give it a defined name, then reference that name in the formula."
    AddMsg m, "datalog-not-bad-shape", 5, "VLA-Datalog", "(not ...) takes exactly one predicate form, like (not (excluded X)) - not zero, and not more than one."
    AddMsg m, "datalog-negation-unsafe-variable", 5, "VLA-Datalog", "in '{predicate}' under (not ...), the variable '{var}' hasn't been given a value by any body predicate written before it - DATALOG can't evaluate a negation over a value it doesn't already know; move a predicate that binds '{var}' earlier in the rule's body."
    AddMsg m, "datalog-negation-not-stratifiable", 5, "VLA-Datalog", "predicate '{predicate}' depends on itself through a (not ...), (count ...), or (sum ...) - directly, or through a chain of rules - and DATALOG requires every one of those to be safely stratifiable (the predicate being negated or aggregated must be fully known before the rule using it can run), so a cycle running back through one can't be evaluated."
    AddMsg m, "datalog-aggregate-bad-shape", 5, "VLA-Datalog", "({form} ...) takes exactly a result variable then a predicate form, like ({form} N (weight X W))."
    AddMsg m, "datalog-aggregate-result-not-a-variable", 5, "VLA-Datalog", "({form} ...)'s first argument must be a bare variable (a word starting with a capital letter) - the name that will hold the result, not the predicate itself."
    AddMsg m, "datalog-sum-needs-one-value-variable", 5, "VLA-Datalog", "(sum ...) over '{predicate}' needs EXACTLY ONE argument left unbound (the value to add up) - found {count}. Bind every other argument to a variable from earlier in the body (the group), or a constant, and leave only the value column free."
    AddMsg m, "datalog-aggregate-result-reused", 5, "VLA-Datalog", "'{var}' is already bound earlier in this rule's body - an aggregate's own result variable must be a brand-new name, not one already in use."
    AddMsg m, "datalog-table-noncontiguous-columns", 5, "VLA-Datalog", "a table argument spanning multiple disjoint areas (a Ctrl-selected, non-contiguous range) isn't supported - select one contiguous block of the table's own columns instead."

    ' DATALOG.4 - comparison filters ((> X 50000) and the rest of the
    ' </<=/>/>=/=/<> set) and the (let Z (+ X Y)) arithmetic binding
    ' form, both new rule-body shapes.
    AddMsg m, "datalog-builtin-needs-two-operands", 5, "VLA-Datalog", "'{operator}' needs exactly two operands, like (> X 50000) or (+ X Y) - found {count}."
    AddMsg m, "datalog-builtin-unsafe-variable", 5, "VLA-Datalog", "in '{operator}', the variable '{var}' hasn't been given a value by any body predicate written before it - a comparison or arithmetic built-in can't evaluate a value it doesn't already know; move a predicate that binds '{var}' earlier in the rule's body."
    AddMsg m, "datalog-unknown-arithmetic-operator", 5, "VLA-Datalog", "'{operator}' isn't one of DATALOG's own arithmetic built-ins (+, -, *, /) - refused by name rather than guessed."
    AddMsg m, "datalog-let-bad-shape", 5, "VLA-Datalog", "(let ...) takes exactly a result variable then an arithmetic expression, like (let Z (+ X Y))."
    AddMsg m, "datalog-let-result-not-a-variable", 5, "VLA-Datalog", "(let ...)'s first argument must be a bare variable (a word starting with a capital letter) - the name that will hold the result, not the expression itself."
    AddMsg m, "datalog-let-result-reused", 5, "VLA-Datalog", "'{var}' is already bound earlier in this rule's body - a (let ...) result variable must be a brand-new name, not one already in use."
    AddMsg m, "datalog-arithmetic-non-numeric-operand", 5, "VLA-Datalog", "(let ...) using '{operator}' needs two numeric operands - at least one side wasn't a number."
    AddMsg m, "datalog-division-by-zero", 5, "VLA-Datalog", "(let ...) using '/' would divide by zero."

    ' DATALOG.5 - named-column atoms: a rule-BODY atom's arguments may be
    ' (header var-or-const) pairs instead of bare positional tokens,
    ' resolved against the target predicate's own header row and
    ' desugared back to the positional atom it stands for.
    AddMsg m, "datalog-atom-mixed-keying", 5, "VLA-Datalog", "'{predicate}' mixes keyed arguments like (column value) with plain positional ones in the same atom - key every argument by its own column name, or none of them."
    AddMsg m, "datalog-keyed-atom-needs-header", 5, "VLA-Datalog", "'{predicate}' has no column headers to key against (it comes from a (fact ...) block, a plain named range, or a rule - none of those have column names) - use plain positional arguments instead, like ({predicate} X Y)."
    AddMsg m, "datalog-unknown-column", 5, "VLA-Datalog", "'{column}' isn't a column of '{predicate}' - its own columns are: {columns}."
    AddMsg m, "datalog-keyed-column-repeated", 5, "VLA-Datalog", "'{column}' is keyed more than once in the same '{predicate}' atom - each column name may appear at most once per (predicate (col val) ...) form."

    ' VLA_Relation.bas - the shared table substrate SQL/DATALOG/PROLOG all
    ' read ranges through. Each engine refuses a non-contiguous selection
    ' by its own name first (sql-/datalog-/prolog-table-noncontiguous-
    ' columns); this is the substrate's own last-line guard, for a caller
    ' that reached SourceToArray without one. Migrated from a raw Err.Raise
    ' at 0.5.1's pre-flight (2026-09-07), where F.14's ratchet caught it.
    AddMsg m, "relation-table-noncontiguous-areas", 5, "VLA-Relation", "a table argument spanning multiple disjoint areas is not supported - select one contiguous block of the table's own columns instead."

    AddMsg m, "sql-table-not-a-range", 5, "VLA-Sql", "SQL's table argument must be a cell range - pass a reference like Employees, not a computed value."
    AddMsg m, "sql-table-noncontiguous-columns", 5, "VLA-Sql", "a table argument spanning multiple disjoint areas (a Ctrl-selected, non-contiguous range) isn't supported - select one contiguous block of the table's own columns instead."
    AddMsg m, "sql-table-needs-a-name", 5, "VLA-Sql", "this range has no name SQL can use as a table - make it an Excel Table (Ctrl+T) or give it a defined name, then reference that name in the formula."
    AddMsg m, "sql-table-needs-real-table", 5, "VLA-Sql", "SQL needs a real Excel Table (select the range, Ctrl+T) so it has real column names to SELECT/WHERE by - a plain named range has no column names of its own."
    AddMsg m, "sql-needs-exactly-one-table", 5, "VLA-Sql", "SQL takes exactly one table for now (found {count}) - multi-table queries (JOIN) aren't built yet."
    AddMsg m, "sql-from-table-mismatch", 5, "VLA-Sql", "the query's own FROM {from} isn't among the table(s) actually passed to this call - FROM must name one of the passed table arguments, the same way a DATALOG rule must name its own table's real predicate."
    AddMsg m, "sql-unterminated-string", 5, "VLA-Sql", "this query has a string literal that's never closed with a matching quote character."
    AddMsg m, "sql-unterminated-comment", 5, "VLA-Sql", "this query has a /* block comment */ that's never closed with a matching */."
    AddMsg m, "sql-unexpected-character", 5, "VLA-Sql", "'{char}' isn't a character SQL.1 recognizes anywhere in a query."
    AddMsg m, "sql-expected-token", 5, "VLA-Sql", "expected {expected}, found {found}."
    AddMsg m, "sql-unsupported-keyword", 5, "VLA-Sql", "'{keyword}' isn't part of SQL's own supported subset yet (SELECT DISTINCT/FROM/WHERE, AS, AND/OR/NOT, = <> < <= > >=, + - * /, parentheses) - it's refused by name rather than silently ignored or misread."
    AddMsg m, "sql-unknown-column", 5, "VLA-Sql", "'{column}' isn't a column of any table this query reads - check the spelling against the table's own header row (or, once a JOIN is involved, qualify it as tablename.{column})."

    ' SQL.2 - computed SELECT expressions (arithmetic), AS aliases, DISTINCT.
    AddMsg m, "sql-computed-column-needs-alias", 5, "VLA-Sql", "a computed SELECT expression needs its own alias (AS name) - only a bare column reference can be selected without one, since GROUP BY/ORDER BY and the output header both need a real name to refer back to."
    AddMsg m, "sql-arithmetic-non-numeric-operand", 5, "VLA-Sql", "'{operator}' needs two numeric operands - at least one side wasn't a real number (a text column that merely looks numeric is never silently promoted)."
    AddMsg m, "sql-division-by-zero", 5, "VLA-Sql", "'/' would divide by zero."

    ' SQL.3 - INNER JOIN ... ON, multiple tables folded left-to-right.
    AddMsg m, "sql-needs-at-least-one-table", 5, "VLA-Sql", "SQL needs at least one table argument (found {count})."
    AddMsg m, "sql-join-table-not-passed", 5, "VLA-Sql", "the query's own JOIN {table} isn't among the table(s) actually passed to this call."
    AddMsg m, "sql-duplicate-join-table", 5, "VLA-Sql", "'{table}' is named more than once across FROM/JOIN - a self-join needs table aliasing (AS), which isn't supported yet."
    AddMsg m, "sql-outer-join-not-supported", 5, "VLA-Sql", "'{keyword}' isn't supported yet - only INNER JOIN (or a bare JOIN, which means the same thing) is built so far; LEFT/RIGHT/FULL/OUTER/CROSS JOIN are refused by name rather than silently misread."
    AddMsg m, "sql-ambiguous-column", 5, "VLA-Sql", "'{column}' exists in more than one joined table - qualify it (e.g. tablename.{column}) to say which one you mean."

    ' SQL.4 - GROUP BY / aggregates (COUNT/SUM/MIN/MAX/AVG) / HAVING.
    AddMsg m, "sql-aggregate-star-only-count", 5, "VLA-Sql", "{function}(*) isn't allowed - only COUNT(*) may use '*' as its own argument; SUM/MIN/MAX/AVG need a real column or expression."
    AddMsg m, "sql-aggregate-nested", 5, "VLA-Sql", "an aggregate function (COUNT/SUM/MIN/MAX/AVG) can't itself contain another aggregate call - aggregate over a plain column or expression instead."
    AddMsg m, "sql-aggregate-not-allowed-in-where", 5, "VLA-Sql", "COUNT/SUM/MIN/MAX/AVG can't appear in WHERE - WHERE filters rows before grouping ever happens; filter on an aggregate's own result with HAVING instead."
    AddMsg m, "sql-aggregate-not-allowed-in-on", 5, "VLA-Sql", "COUNT/SUM/MIN/MAX/AVG can't appear in a JOIN's own ON condition - ON filters rows before grouping ever happens; filter on an aggregate's own result with HAVING instead."
    AddMsg m, "sql-group-by-needs-column", 5, "VLA-Sql", "GROUP BY needs a plain column reference (optionally qualified, like tablename.column) - a computed expression or an aggregate call isn't a valid GROUP BY item."
    AddMsg m, "sql-aggregate-select-star-not-allowed", 5, "VLA-Sql", "SELECT * can't be combined with GROUP BY or an aggregate function - name the columns/aggregates actually wanted instead."
    AddMsg m, "sql-select-item-not-grouped", 5, "VLA-Sql", "SELECT item #{position} is neither a GROUP BY column nor wrapped in an aggregate (COUNT/SUM/MIN/MAX/AVG) - once a query groups or aggregates, every OTHER SELECT item must be one or the other."
    AddMsg m, "sql-aggregate-non-numeric-operand", 5, "VLA-Sql", "{function} needs a numeric operand - at least one grouped value wasn't a real number (a text column that merely looks numeric is never silently promoted)."
    AddMsg m, "sql-aggregate-empty-no-rows", 5, "VLA-Sql", "{function} over zero matching rows has no defined answer (unlike COUNT/SUM, which are 0 there) - the query matched no rows, and with no GROUP BY there's no group for {function} to summarize."

    ' SQL.5 - ORDER BY / LIMIT.
    AddMsg m, "sql-order-by-invalid-position", 5, "VLA-Sql", "ORDER BY position '{value}' isn't valid - a 1-based output position must be a positive whole number."
    AddMsg m, "sql-order-by-position-out-of-range", 5, "VLA-Sql", "ORDER BY {position} is out of range - this query's own SELECT list only has {count} output column(s)."
    AddMsg m, "sql-order-by-unknown-column", 5, "VLA-Sql", "ORDER BY '{column}' isn't one of this query's own output columns - order by a SELECT-list column's own name/alias, or its 1-based position."
    AddMsg m, "sql-order-by-ambiguous-column", 5, "VLA-Sql", "ORDER BY '{column}' matches more than one output column - use its 1-based position instead to say which one you mean."
    AddMsg m, "sql-limit-needs-nonneg-integer", 5, "VLA-Sql", "LIMIT '{value}' isn't a valid row count - it must be a non-negative whole number."

    ' SQL.6 - UNION / UNION ALL / INTERSECT / EXCEPT.
    AddMsg m, "sql-set-op-column-mismatch", 5, "VLA-Sql", "{operator}'s two sides SELECT a different number of columns - the left side has {left}, the right side has {right}; only the column COUNT has to match, never the names or types, but it does have to match."

    ' SQL.7 - WITH CTEs, then recursive WITH.
    AddMsg m, "sql-cte-unterminated-definition", 5, "VLA-Sql", "a WITH clause's own '(' definition is never closed with a matching ')'."
    AddMsg m, "sql-cte-duplicate-name", 5, "VLA-Sql", "'{name}' is named more than once in the same WITH clause - each CTE needs its own distinct name."
    AddMsg m, "sql-cte-self-reference-needs-recursive", 5, "VLA-Sql", "'{name}' references itself inside its own definition, but this WITH clause doesn't say RECURSIVE - add WITH RECURSIVE if a self-referencing CTE is really what's wanted."
    AddMsg m, "sql-cte-recursive-shape", 5, "VLA-Sql", "'{name}' is marked RECURSIVE (or references itself) but doesn't match the one shape this engine supports - exactly `base UNION ALL step`, with neither side itself a further UNION/INTERSECT/EXCEPT/UNION ALL, and the self-reference only in the step (the second) side."
    AddMsg m, "sql-cte-recursive-base-self-references", 5, "VLA-Sql", "'{name}''s own base (the FIRST side of its UNION ALL) references '{name}' itself - the base case must be non-recursive; only the step (the second side) may reference the CTE being defined."
    AddMsg m, "sql-cte-recursive-step-single-reference", 5, "VLA-Sql", "'{name}''s own step references '{name}' more than once (e.g. joining the recursive CTE against itself) - not supported; a recursive step may reference the CTE being defined exactly once."
    AddMsg m, "sql-cte-recursive-no-order-by-limit", 5, "VLA-Sql", "'{name}' is a recursive CTE and can't have its own ORDER BY/LIMIT inside its own definition - per-round ordering has no defined meaning here; ORDER BY/LIMIT the query that USES '{name}' instead."
    AddMsg m, "sql-cte-round-ceiling", 5, "VLA-Sql", "'{name}' didn't converge within {rounds} rounds - a self-referencing step over data with a cycle can keep producing new rows forever; check whether '{name}''s own step is really supposed to still be finding new rows this far in."

    ' Internal consistency checks, not user refusals: each names a state
    ' the SQL evaluator's own earlier passes are supposed to make
    ' unreachable. Catalogued anyway - SD-2 has no "internal" exemption,
    ' and F.14's ratchet caught all three raw at 0.5.1's pre-flight
    ' (2026-09-07) - so if one ever fires live, the report carries a
    ' stable id to search for rather than a one-off string.
    AddMsg m, "sql-internal-aggregate-unregistered", 5, "VLA-Sql", "internal: aggregate '{key}' not registered in the grouped colMap"
    AddMsg m, "sql-internal-scalar-node-shape", 5, "VLA-Sql", "internal: EvalScalar called on a non-scalar AST node"
    AddMsg m, "sql-internal-bool-node-shape", 5, "VLA-Sql", "internal: EvalBool called on a non-boolean AST node"
End Sub

Private Sub AddMsg(ByVal m As Collection, ByVal id As String, ByVal errNum As Long, _
                    ByVal source As String, ByVal template As String)
    On Error GoTo dup
    m.Add Array(errNum, source, template), id
    Exit Sub
dup:
    ' Raw by necessity, not oversight: the chokepoint that would report
    ' this via RaiseMsg is the thing that just failed to build its own
    ' lookup table, so it cannot depend on itself here.
    Err.Raise 5, "VLA-Messages", "message id '" & id & "' is already registered - ids are never reused (SD-9's discipline, at message-id scale); pick a different one"
End Sub

' The one chokepoint every migrated refusal calls. Named parameters as
' an alternating name/value ParamArray: kv(0)/kv(2)/... are names,
' kv(1)/kv(3)/... are the values substituted for each template's
' {name} slot. Err.Number/Err.Source come from the catalogue entry, not
' the call site - the id is now the single source of truth for what a
' refusal fundamentally is.
Public Sub RaiseMsg(ByVal id As String, ParamArray kv() As Variant)
    Dim cat As Collection
    Set cat = Catalogue()          ' outside the id-lookup's own error scope below, so a genuine AddMsg bug (e.g. a duplicate id) surfaces as itself, never masked as "unknown id"
    Dim rec As Variant
    On Error GoTo unknown
    rec = cat.Item(id)
    On Error GoTo 0
    ' VBA forbids forwarding a ParamArray identifier itself into another
    ' call's array parameter ("Invalid ParamArray use", compile-time,
    ' live-caught) - copy it into a plain Variant first, the standard
    ' workaround. SubstituteSlots/SlotValue below take a bare "kv As
    ' Variant" rather than "kv() As Variant" for the same reason: once
    ' copied out, it is no longer a ParamArray, but keeping every
    ' downstream signature untyped-array-free sidesteps the restriction
    ' entirely rather than tripping it one hop later.
    Dim kvArr As Variant
    kvArr = kv
    Err.Raise CLng(rec(0)), CStr(rec(1)), SubstituteSlots(CStr(rec(2)), kvArr)
    Exit Sub
unknown:
    Err.Raise 5, "VLA-Messages", "RaiseMsg: unknown message id '" & id & "'"
End Sub

' Same embedded-{slot}-scan shape as VLA_English.bas's own
' SpliceEmbeddedSlots, over a flat name/value ParamArray instead of
' that function's bn/bv bound-form Collections - a template here only
' ever substitutes a plain string, never a VLA form, so the heavier
' form-substitution machinery (FormSubstitute) would be a mismatch, not
' a reuse.
Private Function SubstituteSlots(ByVal tmpl As String, ByRef kv As Variant) As String
    Dim r As String, i As Long, openPos As Long, closePos As Long
    i = 1
    Do While i <= Len(tmpl)
        openPos = InStr(i, tmpl, "{")
        If openPos = 0 Then
            r = r & Mid$(tmpl, i)
            Exit Do
        End If
        closePos = InStr(openPos, tmpl, "}")
        If closePos = 0 Then
            r = r & Mid$(tmpl, i)
            Exit Do
        End If
        r = r & Mid$(tmpl, i, openPos - i)
        r = r & CStr(SlotValue(Mid$(tmpl, openPos + 1, closePos - openPos - 1), kv))
        i = closePos + 1
    Loop
    SubstituteSlots = r
End Function

Private Function SlotValue(ByVal slotName As String, ByRef kv As Variant) As Variant
    Dim i As Long
    For i = LBound(kv) To UBound(kv) Step 2
        If CStr(kv(i)) = slotName Then
            SlotValue = kv(i + 1)
            Exit Function
        End If
    Next i
    Err.Raise 5, "VLA-Messages", "RaiseMsg: no value supplied for slot '{" & slotName & "}'"
End Function
