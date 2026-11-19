<#
check_expand_golden.ps1 - the expand golden agrees with its fixture, every
record is EXPANDED with its text or REFUSED with a catalogue id, the
reference's lists of the expand-time primitives are the port's, and every
primitive is reached.

WHY: VLA.3 (the expand-time object model) added ten primitives to the
seventeen, in both implementations. The reference names each in five places
(src/VLA.bas: DefineMacro's reserved names, ExpandMacros' standalone Select,
Substitute's inline Select, EvalListopsPrim's dispatch, and since VLA.8
IsListopsName, where a value read computes a call handed in) and the port in two
(vla-lang/src/expand.rs: RESERVED and LISTOPS). The port's own test,
vla-lang/tests/expand_golden.rs, holds its expansion of scripts/expand.txt
to the reference's golden, scripts/expand_golden.txt (VlaWriteExpandGolden,
VLA_Tests.bas). That test proves agreement on what the fixture asks. It
cannot see the reference's own lists, so a primitive left out of one of
the reference's Selects would pass it; nor a primitive no case reaches, a
golden regenerated against a fixture that lost a case, or a refusal whose
id the catalogue does not hold. This check holds those, host-free, on
every push:

  1. the golden's headers are the fixture's, in order, and every fixture
     case has text;
  2. every golden case is one record: EXPANDED<TAB><count> and at least
     one line of text, or one REFUSED<TAB><id><TAB><text> line whose id is
     in the language's half of the catalogue (vla-lang/data/messages.vla);
  3. the reference's lists and the port's name the same primitives: the
     three recognition Selects (ExpandMacros', Substitute's and, since
     VLA.8, IsListopsName's) and EvalListopsPrim's Cases are the port's
     LISTOPS, and the reference's reserved words, one table since VLA.5
     (ExpandReservedDesc, which DefineMacro and a loop's names both ask),
     are the port's RESERVED (read from the sources, never retyped here);
  4. every primitive of LISTOPS heads a form in at least one fixture case;
     each of VLA.3's ten heads one in an EXPANDED case and in a REFUSED
     case; and each of VLA.3's eight ids is the id of a REFUSED record;
  5. the floors below: the cases, the expanded cases and the distinct ids;
  6. VLA.4's formula-expand, an operation of the walk and no LISTOPS
     primitive (Substitute leaves it alone, so it is in neither Select):
     the walk recognizes it in both implementations (ExpandMacros' test of
     headFold, expand_at's of head_fold), it heads a form in an EXPANDED
     case and in a REFUSED one, and each of VLA.4's eleven ids an expansion
     can raise is the id of a REFUSED record. The reference's formula
     writer reads no compile's state: none of its procedures calls
     ResolveHeadAlias, whose map only a compile sets, since formula-expand
     runs in the expander (the owner's pass of 2026-10-17 found
     (si ...) written SI(...) in a session no compile had run in);
  7. VLA.5's loop, each and each-index, a form of the walks and of the
     substitution and no LISTOPS primitive: both implementations take it at
     a list's head in their substitution (Substitute's Case, substitute's
     arm) and at a form's head in their walk (ExpandMacros' and expand_at's
     test of the head), both reserve it and neither lists it among the
     primitives, each form heads a form in an EXPANDED case and in a
     REFUSED one, VLA.5's primitive assoc-expand does too, and each of
     VLA.5's six ids is the id of a REFUSED record.
  8. VLA.6's define, a form of the walk alone, with a reference's two
     parts, sheet-expand and range-expand, and the formula syntax's last:
     both walks take define and the two at a form's head (ExpandMacros' and
     expand_at's test of the head), both reserve the three and list none
     among the primitives; a define binds from the next form of the source
     on only while every form of a source goes through one wrapper, so the
     reference's three entry points call ExpandSourceForm and nothing else
     calls ExpandMacros(f, 0), and the port's Expander::expand takes the
     defines (take_defines); both writers take last among their own heads
     (FsList's Case, formula_syntax's list); each of the four heads a form
     in an EXPANDED case and in a REFUSED one; and each of VLA.6's ten ids
     is the id of a REFUSED record.
  9. VLA.15's list?, the twenty-ninth primitive, which rules 3 and 4 hold
     among the others; a reference's first cell, first-cell-expand, a form
     of the walk alone, taken at a form's head by both walks, reserved in
     both and listed among neither's primitives; the formula syntax's three
     marks, absolute, absolute-row and absolute-column, among both writers'
     own heads (FsList's Case, formula_syntax's list); each of the five
     heading a form in an EXPANDED case and in a REFUSED one; and VLA.15's
     id the id of a REFUSED record.
 10. VLA.9's named pieces: both signature walks take a list as a declaration
     (DefineMacro's call of DeclarePiece, define_macro's of declare_piece) and
     both bindings take a default (BindPieces' and bind_pieces' substitution of
     one); each of VLA.9's eleven ids is the id of a REFUSED record; and the
     two cases of one call in two orders have one record, and that record the
     rows of a piece bound by its word and never by its place.
 11. VLA.10's place a verb writes: bind-expand and offset-expand, forms of the
     walk alone, taken at a form's head by both walks; those two and the word
     self reserved in both and listed among neither's primitives; each form
     heading a form in an EXPANDED case and in a REFUSED one; each of VLA.10's
     seven ids the id of a REFUSED record. The wall's one condition, by reading:
     the stack reset at each of the reference's three entry points, restored on
     its error path (ExpandBindCall's handler) and before the port's question
     mark (bind_expand); and the gate's premise, LISTOPS-CONFLUENCE's, that
     neither Substitute nor EvalListopsPrim nor the port's substitute names the
     stack. And two sibling verbs, in either order, one record of each row.
 12. VLA.8's functions and let-expand: let-expand and defexpand reserved in
     both and listed among neither's primitives, each heading a form in an
     EXPANDED case and in a REFUSED one; let-expand taken at a form's head by
     both walks; a function taken by every entry point as a macro is, the
     reference's three loops and the port's define; the working-out's depth
     restored on its error path in both (FunctionWorkOut's handler; every
     way out of function_work_out's loop answered through work_out_leave,
     which lowers it, and none by `?` or `return`), so that no refusal
     leaves a later working-out a level short; the two limits one number in
     both; and each of VLA.8's fifteen ids the id of a REFUSED record.
 13. VLA.11's definitions: a macro's second definition refused by name in
     both definers (DefineMacro, define_macro); parameter reserved in both
     and listed among neither's primitives, heading a form in an EXPANDED
     case and in a REFUSED one, both walks passing it by as a define (rule
     8's test of the head); the compile's parameter words reset at the
     reference's three entry points and a door's values held to them at the
     end of each (ParametersFinish); both golden writers giving a case's
     ";door " lines to the door (ExpandReportFor, the port's report_for), a
     case with one EXPANDED and one REFUSED; and each of its four ids the id
     of a REFUSED record.

-Control proves the check on scratch copies: the real files must pass, and
each of these must fail: a golden with one record's tag misspelled, one
with its last case dropped, one refusing with an id the catalogue lacks, a
reference with one name taken from Substitute's Select, a fixture with
every characters-expand renamed away, a reference whose walk no longer
recognizes formula-expand, a reference whose formula writer resolves an
alias through ResolveHeadAlias, a reference whose Substitute no longer
takes the loop at a list's head, a port whose RESERVED lost each, a
reference whose walk no longer passes a define by, a reference whose
VlaExpandText expands a form by ExpandMacros directly, a port whose writer
lost last, a reference whose walk no longer takes the first cell, a port
whose writer lost the three marks, a reference whose signature walk reads no
declaration, a port whose binding takes no default, a golden whose second
order of one call differs from its first, a reference whose walk no longer
takes bind-expand, a reference whose handler leaves the stack as it stood, a
reference whose Substitute reads the stack, a port whose binding never pops,
a golden whose second sibling writes the first's place, a reference whose
working-out leaves the depth as it stood on a refusal, a port whose
working-out never lowers it, a port whose loop answers a refusal of its fuel
by `?`, past the lowering, a reference whose DefineMacro lets a second
definition replace the first, a port whose define_macro does, and a reference
whose VlaCompileToForms no longer holds a door's values.

House style (tools/check_*.ps1): PowerShell 5.1, host-free, no Excel, no COM,
no network; exit 0 clean, exit 1 with every problem named.

Usage:  powershell -File tools\check_expand_golden.ps1
        powershell -File tools\check_expand_golden.ps1 -Control
#>
param(
    [switch]$Control
)
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot

# --- the floors, never lowered ------------------------------------------------
# 2026-10-10, VLA.3: 135 cases, 75 of them expanded, 19 distinct ids refused;
# the golden written by the port ahead of the reference's first run.
# 2026-10-17, VLA.4: 174 cases, 96 expanded, 30 distinct ids, formula-expand's
# 39 written by the port ahead of the reference's run, as VLA.3's were.
# 2026-10-24, VLA.5: 229 cases, 129 expanded, 36 distinct ids, the loop's and
# assoc-expand's 52 written by the port ahead of the reference's run.
# 2026-10-25, VLA.6: 294 cases, 155 expanded, 46 distinct ids, define's,
# last's, the reference parts' and the walk's returns' 65 written by the port
# ahead of the reference's run.
# 2026-10-29, VLA.15: 362 cases, 198 expanded, 47 distinct ids, list?'s, the
# three marks' and the first cell's 68 written by the port ahead of the
# reference's run.
# 2026-10-30, VLA.9: 471 cases, 245 expanded, 63 distinct ids, the named pieces'
# 109 written by the port ahead of the reference's run, and five ids of the
# macro system's a refused case had not reached before.
# 2026-10-30, VLA.10: 656 cases, 339 expanded, 71 distinct ids, the place a verb
# writes' 185 written by the port ahead of the reference's run.
# 2026-11-16, VLA.16 and VLA.17: 686 cases, 354 expanded, 72 distinct ids, an
# entry's names' 20 and funcall's 10 written by the port ahead of the reference's
# run, vla-each-entry-shape the one id more.
# 2026-11-16, VLA.17's IFERROR, from KERNEL.29's third save: 687 cases, 354
# expanded, 73 distinct ids, vla-formula-not-called-in-place the one id more.
# 2026-11-17, VLA.18: 701 cases, 363 expanded, 75 distinct ids, the raw
# string's 14 written by the port ahead of the reference's run,
# vla-raw-string-unclosed and vla-hash-word the two ids more.
# 2026-11-18, VLA.8: 761 cases, 393 expanded, 90 distinct ids, the functions'
# and let-expand's 60 written by the port ahead of the reference's run, its
# fifteen ids the ids more; one case renamed for its changed meaning,
# define-is-a-value-a-primitive-reads, where define-is-no-primitive-s-argument
# read true.
# 2026-11-18, VLA.8's second slice: 777 cases, 405 expanded, 91 distinct ids,
# the prelude's list library over functions, 16 written by the port ahead of
# the reference's run, vla-quote-if-test-is-list the one id more (quote-filter
# given a macro's test); the 761 before them unmoved by the prelude's rewrite.
# 2026-11-19, VLA.11's first slice: 802 cases, 414 expanded, 95 distinct ids,
# the definitions' 25 written by the port ahead of the reference's run, its four
# ids the ids more; the 777 before them unmoved, the define refusals' new slot
# filled as their texts were.
$caseFloor = 802
$expandedFloor = 414
$idFloor = 95

# VLA.3's ten primitives and eight ids, which rule 4 holds to the fixture.
$newPrimitives = @('-expand', '*expand', 'quotient-expand', 'mod-expand', '&expand',
    'characters-expand', 'address-expand', 'row-expand', 'column-expand', 'column-letters-expand')
$newIds = @('vla-expand-arity', 'vla-expand-arg-is-list', 'vla-concat-expand-arity',
    'vla-arith-expand-not-whole', 'vla-arith-expand-divide-by-zero',
    'vla-address-expand-row', 'vla-address-expand-column', 'vla-address-expand-not-address')

# VLA.4's operation of the walk and the eleven ids an expansion can raise
# (the reader's two, vla-formula-read-malformed and -unsupported, are raised
# by no expansion), which rule 6 holds to both implementations and the fixture.
$walkOperation = 'formula-expand'
$formulaIds = @('vla-formula-atom-unspelled', 'vla-formula-numeral-unspelled',
    'vla-formula-head-not-a-function', 'vla-formula-head-refused', 'vla-formula-arity',
    'vla-formula-unsupported', 'vla-formula-lambda-params', 'vla-formula-ref-invalid',
    'vla-formula-text-too-long', 'vla-formula-expand-arity', 'vla-formula-expand-nested')

# VLA.5's loop, a form of the walks and of the substitution, its primitive,
# and its seven ids, VLA.16's vla-each-entry-shape among them, which rule 7 holds to both implementations and the fixture.
$loopForms = @('each', 'each-index')
$loopPrimitives = @('assoc-expand')
$loopIds = @('vla-each-shape', 'vla-each-name-reserved', 'vla-each-name-bound',
    'vla-each-list-a-call', 'vla-expand-expected-list', 'vla-assoc-expand-entry',
    'vla-each-entry-shape')

# VLA.6's forms of the walk alone, the formula syntax's head, and the ten ids,
# which rule 8 holds to both implementations and the fixture.
$defineForms = @('define', 'sheet-expand', 'range-expand')
$syntaxHead = 'last'
$defineIds = @('vla-define-shape', 'vla-define-name-reference', 'vla-define-name-reserved',
    'vla-define-name-taken', 'vla-define-not-a-statement', 'vla-define-used-in-its-form',
    'vla-define-cycle', 'vla-reference-no-sheet', 'vla-reference-expected', 'vla-formula-last-twin')

# VLA.15's primitive, its form of the walk alone, the formula syntax's three
# marks, and its id, which rule 9 holds to both implementations and the fixture.
$addressPrimitive = 'list?'
$addressForm = 'first-cell-expand'
$markHeads = @('absolute', 'absolute-row', 'absolute-column')
$addressIds = @('vla-reference-whole-line')

# VLA.9's eleven ids, the definition's six and the call's five, and the two
# cases of one call written in two orders, which rule 10 holds to the fixture.
$pieceIds = @('vla-defmacro-piece-shape', 'vla-defmacro-param-after-piece',
    'vla-defmacro-piece-with-rest', 'vla-defmacro-param-twice',
    'vla-defmacro-piece-reserved', 'vla-defmacro-piece-default-unbound',
    'vla-macro-piece-unknown', 'vla-macro-piece-arity', 'vla-macro-piece-twice',
    'vla-macro-piece-missing', 'vla-macro-piece-for-param')
$twoOrders = @('=== a-put-by-pieces', '=== a-put-by-pieces-in-another-order')

# VLA.10's forms of the walk alone, the word they read, its seven ids, and the two
# cases of two sibling verbs in either order, which rule 11 holds.
$selfForms = @('bind-expand', 'offset-expand')
$selfWord = 'self'
$selfIds = @('vla-self-free', 'vla-self-in-a-primitive', 'vla-bind-expand-shape',
    'vla-bind-expand-word', 'vla-bind-expand-not-a-cell', 'vla-offset-expand-arity',
    'vla-offset-expand-size')
$twoSiblings = @('=== two-siblings', '=== two-siblings-in-the-other-order')
$siblingRows = @('(formula "A" "B2" "=A.last!B2")', '(formula "C" "D4" "=C.last!D4")')

# VLA.8's two forms and its fifteen ids, which rule 12 holds.
$functionForms = @('let-expand', 'defexpand')
$functionIds = @('vla-defexpand-shape', 'vla-defexpand-name-reference',
    'vla-defexpand-name-reserved', 'vla-defexpand-name-taken', 'vla-defexpand-piece',
    'vla-defexpand-rest', 'vla-defexpand-param-twice', 'vla-defexpand-body',
    'vla-defmacro-name-taken', 'vla-function-arity', 'vla-function-arity-min',
    'vla-function-fuel', 'vla-function-depth', 'vla-let-expand-shape',
    'vla-let-expand-value-forms')

# VLA.11's first slice, the definitions: its form and its four ids, which rule 13
# holds.
$parameterForm = 'parameter'
$definitionIds = @('vla-defmacro-twice', 'vla-parameter-given-twice',
    'vla-parameter-kind', 'vla-parameter-undeclared')

$fixturePath = Join-Path $repoRoot 'scripts/expand.txt'
$goldenPath  = Join-Path $repoRoot 'scripts/expand_golden.txt'
$vbaPath     = Join-Path $repoRoot 'src/VLA.bas'
$portPath    = Join-Path $repoRoot 'vla-lang/src/expand.rs'
$syntaxPath  = Join-Path $repoRoot 'vla-lang/src/formula_syntax.rs'
$catalogPath = Join-Path $repoRoot 'vla-lang/data/messages.vla'

function Get-ExpandLines([string]$path) {
    $text = [System.IO.File]::ReadAllText($path)
    $text = $text -replace "`r`n", "`n" -replace "`r", "`n"
    return ,@($text -split "`n")
}

function New-OrdinalSet {
    return ,(New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::Ordinal))
}

# The cases of a fixture or a golden: each header with the lines under it,
# the blank lines at a case's end dropped (the file's own last break).
function Read-ExpandCases([string]$path) {
    $cases = New-Object System.Collections.ArrayList
    $current = $null
    foreach ($ln in (Get-ExpandLines $path)) {
        if ($ln.StartsWith('=== ')) {
            if ($null -ne $current) { [void]$cases.Add($current) }
            $current = @{ Header = $ln; Lines = (New-Object System.Collections.ArrayList) }
        } elseif ($null -ne $current) {
            [void]$current.Lines.Add($ln)
        }
    }
    if ($null -ne $current) { [void]$cases.Add($current) }
    foreach ($c in $cases) {
        while ($c.Lines.Count -gt 0 -and $c.Lines[$c.Lines.Count - 1] -eq '') {
            $c.Lines.RemoveAt($c.Lines.Count - 1)
        }
    }
    return ,$cases
}

# The heads of the forms in a text: the token after each opening
# parenthesis, a string's own parentheses never one since a quote mark ends
# the token before it starts.
function Get-ExpandHeads([string]$text) {
    $set = New-OrdinalSet
    foreach ($m in [regex]::Matches($text, '\(([^\s()"]+)')) {
        [void]$set.Add($m.Groups[1].Value)
    }
    return ,$set
}

# The reference's lists, read from src/VLA.bas with continuations joined:
# the recognition Selects (each a Case opening on "car", "cdr"), the Cases
# of EvalListopsPrim, and the reserved words, the Cases of ExpandReservedDesc
# (VLA.5), the one table DefineMacro asks.
function Read-ReferenceLists([string]$path) {
    $logical = New-Object System.Collections.ArrayList
    $buffer = ''
    foreach ($ln in (Get-ExpandLines $path)) {
        if ($ln -match '\s_\s*$') {
            $buffer += ($ln -replace '\s_\s*$', ' ')
        } else {
            [void]$logical.Add($buffer + $ln)
            $buffer = ''
        }
    }
    $selects = New-Object System.Collections.ArrayList
    $prim = New-OrdinalSet
    $reserved = New-OrdinalSet
    $inPrim = $false
    $inReserved = $false
    foreach ($ln in $logical) {
        if ($ln -match '^\s*Case "car", "cdr", ') {
            $set = New-OrdinalSet
            foreach ($q in [regex]::Matches($ln, '"([^"]+)"')) { [void]$set.Add($q.Groups[1].Value) }
            [void]$selects.Add($set)
        }
        if ($ln -match '^\s*Private Function EvalListopsPrim\b') { $inPrim = $true; continue }
        if ($inPrim -and $ln -match '^\s*End Function') { $inPrim = $false }
        if ($inPrim -and $ln -match '^\s*Case\s+"') {
            foreach ($q in [regex]::Matches($ln, '"([^"]+)"')) { [void]$prim.Add($q.Groups[1].Value) }
        }
        if ($ln -match '^\s*Private Function ExpandReservedDesc\b') { $inReserved = $true; continue }
        if ($inReserved -and $ln -match '^\s*End Function') { $inReserved = $false }
        if ($inReserved -and $ln -match '^\s*Case\s+"') {
            foreach ($q in [regex]::Matches($ln, '"([^"]+)"')) { [void]$reserved.Add($q.Groups[1].Value) }
        }
    }
    return @{ Selects = $selects; Prim = $prim; Reserved = $reserved }
}

# The port's lists, read from vla-lang/src/expand.rs: LISTOPS, and the first
# word of each RESERVED pair.
function Read-PortLists([string]$path) {
    $text = [System.IO.File]::ReadAllText($path)
    $listops = New-OrdinalSet
    $reserved = New-OrdinalSet
    $m = [regex]::Match($text, '(?s)const LISTOPS: &\[&str\] = &\[(.*?)\];')
    if ($m.Success) {
        foreach ($q in [regex]::Matches($m.Groups[1].Value, '"([^"]+)"')) { [void]$listops.Add($q.Groups[1].Value) }
    }
    $r = [regex]::Match($text, '(?s)const RESERVED: &\[\(&str, &str\)\] = &\[(.*?)\];')
    if ($r.Success) {
        foreach ($q in [regex]::Matches($r.Groups[1].Value, '\(\s*"([^"]+)"\s*,')) { [void]$reserved.Add($q.Groups[1].Value) }
    }
    return @{ Listops = $listops; Reserved = $reserved }
}

# The ids the language's half of the catalogue holds.
function Read-CatalogIds([string]$path) {
    $ids = New-OrdinalSet
    foreach ($ln in (Get-ExpandLines $path)) {
        if ($ln -match '^\(message (\S+) ') { [void]$ids.Add($Matches[1]) }
    }
    return ,$ids
}

function Format-ExpandSetDiff([string]$what, $a, [string]$aName, $b, [string]$bName) {
    $out = New-Object System.Collections.ArrayList
    foreach ($x in $a) { if (-not $b.Contains($x)) { [void]$out.Add("$what '$x' is in $aName and not in $bName") } }
    foreach ($x in $b) { if (-not $a.Contains($x)) { [void]$out.Add("$what '$x' is in $bName and not in $aName") } }
    return ,$out
}

# Every problem of one fixture/golden/reference/port/writer/catalogue set,
# with the counts the floors read.
function Test-ExpandGolden([string]$fixture, [string]$golden, [string]$vba, [string]$port, [string]$catalog, [string]$syntax) {
    $problems = New-Object System.Collections.ArrayList
    $fcases = Read-ExpandCases $fixture
    $gcases = Read-ExpandCases $golden
    $ref = Read-ReferenceLists $vba
    $prt = Read-PortLists $port
    $catalogIds = Read-CatalogIds $catalog

    # 3. the lists
    if ($prt.Listops.Count -eq 0) { [void]$problems.Add("no LISTOPS array in $port") }
    if ($prt.Reserved.Count -eq 0) { [void]$problems.Add("no RESERVED array in $port") }
    if ($ref.Selects.Count -ne 3) {
        [void]$problems.Add(("{0} recognition Select(s) opening on `"car`", `"cdr`" in {1}, not the three (ExpandMacros, Substitute, IsListopsName)" -f $ref.Selects.Count, $vba))
    }
    $k = 0
    foreach ($s in $ref.Selects) {
        $k++
        foreach ($p in (Format-ExpandSetDiff 'primitive' $s "the reference's recognition Select $k" $prt.Listops "the port's LISTOPS")) { [void]$problems.Add($p) }
    }
    foreach ($p in (Format-ExpandSetDiff 'primitive' $ref.Prim "the reference's EvalListopsPrim" $prt.Listops "the port's LISTOPS")) { [void]$problems.Add($p) }
    foreach ($p in (Format-ExpandSetDiff 'reserved word' $ref.Reserved "the reference's ExpandReservedDesc" $prt.Reserved "the port's RESERVED")) { [void]$problems.Add($p) }
    foreach ($x in $prt.Listops) {
        if (-not $prt.Reserved.Contains($x)) { [void]$problems.Add("primitive '$x' is recognized by the port and not reserved") }
    }
    # VLA.5: the table is the reserved words only while DefineMacro asks it.
    $vbaSource = [System.IO.File]::ReadAllText($vba)
    if ($vbaSource -notmatch 'reservedDesc = ExpandReservedDesc\(VLA_Identity\.Fold\(mname\)\)') {
        [void]$problems.Add("the reference's DefineMacro does not ask ExpandReservedDesc, so its table is no macro's refusal")
    }

    # 1. the headers
    if ($fcases.Count -ne $gcases.Count) {
        [void]$problems.Add(("the fixture has {0} cases and the golden {1}" -f $fcases.Count, $gcases.Count))
    }
    $n = [Math]::Min($fcases.Count, $gcases.Count)
    for ($i = 0; $i -lt $n; $i++) {
        if ($fcases[$i].Header -cne $gcases[$i].Header) {
            [void]$problems.Add(("case {0}: the fixture says '{1}', the golden '{2}'" -f ($i + 1), $fcases[$i].Header, $gcases[$i].Header))
        }
    }
    foreach ($c in $fcases) {
        if ($c.Lines.Count -eq 0) { [void]$problems.Add(("'{0}': no text after its header" -f $c.Header)) }
    }

    # 2. the records, and what each case's text heads
    $expanded = 0
    $ids = New-OrdinalSet
    $headsAny = New-OrdinalSet
    $headsExpanded = New-OrdinalSet
    $headsRefused = New-OrdinalSet
    for ($i = 0; $i -lt $gcases.Count; $i++) {
        $g = $gcases[$i]
        $heads = New-OrdinalSet
        if ($i -lt $fcases.Count) { $heads = Get-ExpandHeads (($fcases[$i].Lines) -join "`n") }
        foreach ($h in $heads) { [void]$headsAny.Add($h) }
        if ($g.Lines.Count -eq 0) {
            [void]$problems.Add(("'{0}': no record" -f $g.Header))
            continue
        }
        $first = [string]$g.Lines[0]
        if ($first -cmatch '^EXPANDED\t\d+$') {
            if ($g.Lines.Count -lt 2) { [void]$problems.Add(("'{0}': EXPANDED with no text after it" -f $g.Header)) }
            $expanded++
            foreach ($h in $heads) { [void]$headsExpanded.Add($h) }
        } elseif ($first -cmatch '^REFUSED\t([a-z0-9-]+)\t(.+)$') {
            $id = $Matches[1]
            if ($g.Lines.Count -ne 1) { [void]$problems.Add(("'{0}': a refusal of {1} lines, want one" -f $g.Header, $g.Lines.Count)) }
            if (-not $catalogIds.Contains($id)) { [void]$problems.Add(("'{0}': refused with '{1}', which the language's catalogue does not hold" -f $g.Header, $id)) }
            [void]$ids.Add($id)
            foreach ($h in $heads) { [void]$headsRefused.Add($h) }
        } else {
            [void]$problems.Add(("'{0}': a record that is neither EXPANDED<TAB><count> nor REFUSED<TAB><id><TAB><text>: '{1}'" -f $g.Header, $first))
        }
    }

    # 4. coverage
    foreach ($x in $prt.Listops) {
        if (-not $headsAny.Contains($x)) { [void]$problems.Add("no case of the fixture heads a form with '$x'") }
    }
    foreach ($x in $newPrimitives) {
        if (-not $prt.Listops.Contains($x)) { [void]$problems.Add("'$x', one of VLA.3's ten, is not in the port's LISTOPS") }
        if (-not $headsExpanded.Contains($x)) { [void]$problems.Add("no EXPANDED case heads a form with '$x'") }
        if (-not $headsRefused.Contains($x)) { [void]$problems.Add("no REFUSED case heads a form with '$x'") }
    }
    foreach ($x in $newIds) {
        if (-not $ids.Contains($x)) { [void]$problems.Add("no REFUSED record of the golden is '$x'") }
    }
    foreach ($x in $loopPrimitives) {
        if (-not $prt.Listops.Contains($x)) { [void]$problems.Add("'$x', VLA.5's primitive, is not in the port's LISTOPS") }
        if (-not $headsExpanded.Contains($x)) { [void]$problems.Add("no EXPANDED case heads a form with '$x'") }
        if (-not $headsRefused.Contains($x)) { [void]$problems.Add("no REFUSED case heads a form with '$x'") }
    }

    # 6. VLA.4's operation of the walk
    $vbaText = [System.IO.File]::ReadAllText($vba)
    $portText = [System.IO.File]::ReadAllText($port)
    if ($vbaText -notmatch ('headFold = "' + [regex]::Escape($walkOperation) + '"')) {
        [void]$problems.Add("the reference's walk (ExpandMacros) does not recognize '$walkOperation'")
    }
    if ($portText -notmatch ('head_fold == "' + [regex]::Escape($walkOperation) + '"')) {
        [void]$problems.Add("the port's walk (expand_at) does not recognize '$walkOperation'")
    }
    if (-not $prt.Reserved.Contains($walkOperation)) { [void]$problems.Add("'$walkOperation' is not in the port's RESERVED") }
    if (-not $headsExpanded.Contains($walkOperation)) { [void]$problems.Add("no EXPANDED case heads a form with '$walkOperation'") }
    if (-not $headsRefused.Contains($walkOperation)) { [void]$problems.Add("no REFUSED case heads a form with '$walkOperation'") }
    foreach ($x in $formulaIds) {
        if (-not $ids.Contains($x)) { [void]$problems.Add("no REFUSED record of the golden is '$x'") }
    }
    # The reference's formula writer holds no compile's state: an alias
    # reaches its form by the head table (FsCanonicalHead), never by
    # ResolveHeadAlias, whose map only a compile sets.
    $writer = [regex]::Matches($vbaText, '(?ms)^(?:Private|Public) Function (Fs\w+|FormulaSyntaxWrite|FormulaHeadRefused)\(.*?^End Function')
    $writerNames = @($writer | ForEach-Object { $_.Groups[1].Value })
    if ($writerNames -notcontains 'FsList') { [void]$problems.Add("the reference's formula writer has no FsList for this rule to read") }
    foreach ($m in $writer) {
        if ($m.Value -match 'ResolveHeadAlias\(') {
            [void]$problems.Add("the reference's $($m.Groups[1].Value) resolves an alias through ResolveHeadAlias, a compile's state the expander does not set")
        }
    }

    # 7. VLA.5's loop, a form of the walks and of the substitution: taken at a
    # list's head by both implementations' substitution and at a form's head
    # by both walks, reserved in both, listed among neither's primitives, and
    # reached by the fixture; its ids each a refused record.
    if ($vbaText -notmatch 'Case "each", "each-index"') {
        [void]$problems.Add("the reference's Substitute does not take the loop at a list's head (Case `"each`", `"each-index`")")
    }
    if ($vbaText -notmatch 'If headFold = "each" Or headFold = "each-index" Then') {
        [void]$problems.Add("the reference's walk (ExpandMacros) does not take the loop at a form's head")
    }
    if ($portText -notmatch '"each" \| "each-index" if !in_quasi =>') {
        [void]$problems.Add("the port's substitute does not take the loop at a list's head")
    }
    if ($portText -notmatch 'if head_fold == "each" \|\| head_fold == "each-index" \{') {
        [void]$problems.Add("the port's walk (expand_at) does not take the loop at a form's head")
    }
    foreach ($x in $loopForms) {
        if (-not $prt.Reserved.Contains($x)) { [void]$problems.Add("the loop's form '$x' is not in the port's RESERVED") }
        if ($prt.Listops.Contains($x)) { [void]$problems.Add("the loop's form '$x' is in the port's LISTOPS, which takes data and splices nothing") }
        if (-not $headsExpanded.Contains($x)) { [void]$problems.Add("no EXPANDED case heads a form with '$x'") }
        if (-not $headsRefused.Contains($x)) { [void]$problems.Add("no REFUSED case heads a form with '$x'") }
    }
    foreach ($x in $loopIds) {
        if (-not $ids.Contains($x)) { [void]$problems.Add("no REFUSED record of the golden is '$x'") }
    }

    # 8. VLA.6's define and a reference's two parts, forms of the walk alone,
    # and the formula syntax's last: taken at a form's head by both walks,
    # reserved in both and listed among neither's primitives; every form of a
    # source through one wrapper, so that a define binds from the next form
    # on; last among both writers' own heads; reached by the fixture; its ids
    # each a refused record.
    $syntaxText = [System.IO.File]::ReadAllText($syntax)
    if ($vbaText -notmatch 'If headFold = "define" Or headFold = "parameter" Then') {
        [void]$problems.Add("the reference's walk (ExpandMacros) does not pass a define by at a form's head")
    }
    if ($vbaText -notmatch 'If headFold = "sheet-expand" Or headFold = "range-expand" Then') {
        [void]$problems.Add("the reference's walk (ExpandMacros) does not take a reference's two parts at a form's head")
    }
    if ($portText -notmatch 'if head_fold == "define" \|\| head_fold == "parameter" \{') {
        [void]$problems.Add("the port's walk (expand_at) does not pass a define by at a form's head")
    }
    if ($portText -notmatch 'if head_fold == "sheet-expand" \|\| head_fold == "range-expand" \{') {
        [void]$problems.Add("the port's walk (expand_at) does not take a reference's two parts at a form's head")
    }
    $wrapperCalls = [regex]::Matches($vbaText, 'ExpandSourceForm\(f\)').Count
    if ($wrapperCalls -ne 3) {
        [void]$problems.Add("the reference's entry points call ExpandSourceForm(f) $wrapperCalls time(s), not the three (VlaTranspile, VlaCompileToForms, VlaExpandText)")
    }
    $directCalls = [regex]::Matches($vbaText, 'ExpandMacros\(f, 0\)').Count
    if ($directCalls -ne 1) {
        [void]$problems.Add("the reference calls ExpandMacros(f, 0) $directCalls time(s), where ExpandSourceForm's one call is the only one, so a form of a source would expand with no define taken")
    }
    if ($portText -notmatch '(?s)pub fn expand\(&mut self, d: &Form, depth: u32\).*?self\.take_defines\(out\)') {
        [void]$problems.Add("the port's Expander::expand does not take a form's defines (take_defines)")
    }
    if ($vbaText -notmatch 'FsList = FsLastForm\(lst, headWritten\)') {
        [void]$problems.Add("the reference's formula writer (FsList) does not take last among its own heads")
    }
    if ($syntaxText -notmatch '"last" => last_form\(head, args\),') {
        [void]$problems.Add("the port's formula writer (formula_syntax's list) does not take last among its own heads")
    }
    foreach ($x in $defineForms) {
        if (-not $prt.Reserved.Contains($x)) { [void]$problems.Add("VLA.6's form '$x' is not in the port's RESERVED") }
        if ($prt.Listops.Contains($x)) { [void]$problems.Add("VLA.6's form '$x' is in the port's LISTOPS, which takes data and sees no define") }
    }
    foreach ($x in @($defineForms + $syntaxHead)) {
        if (-not $headsExpanded.Contains($x)) { [void]$problems.Add("no EXPANDED case heads a form with '$x'") }
        if (-not $headsRefused.Contains($x)) { [void]$problems.Add("no REFUSED case heads a form with '$x'") }
    }
    foreach ($x in $defineIds) {
        if (-not $ids.Contains($x)) { [void]$problems.Add("no REFUSED record of the golden is '$x'") }
    }

    # 9. VLA.15's list?, a primitive among the others (rules 3 and 4); a
    # reference's first cell, a form of the walk alone: taken at a form's head
    # by both walks, reserved in both and listed among neither's primitives;
    # the three marks among both writers' own heads; each of the five reached
    # by an expanded case and a refused one; its id a refused record.
    if (-not $prt.Listops.Contains($addressPrimitive)) {
        [void]$problems.Add("'$addressPrimitive', VLA.15's primitive, is not in the port's LISTOPS")
    }
    if ($vbaText -notmatch ('If headFold = "' + [regex]::Escape($addressForm) + '" Then')) {
        [void]$problems.Add("the reference's walk (ExpandMacros) does not take a reference's first cell at a form's head")
    }
    if ($portText -notmatch ('if head_fold == "' + [regex]::Escape($addressForm) + '" \{')) {
        [void]$problems.Add("the port's walk (expand_at) does not take a reference's first cell at a form's head")
    }
    if (-not $prt.Reserved.Contains($addressForm)) { [void]$problems.Add("VLA.15's form '$addressForm' is not in the port's RESERVED") }
    if ($prt.Listops.Contains($addressForm)) {
        [void]$problems.Add("VLA.15's form '$addressForm' is in the port's LISTOPS, which takes data and expands no argument")
    }
    if ($vbaText -notmatch 'Case "absolute", "absolute-row", "absolute-column"\s+FsList = FsMarkForm\(lst, h\)') {
        [void]$problems.Add("the reference's formula writer (FsList) does not take the three marks among its own heads")
    }
    if ($syntaxText -notmatch '"absolute" \| "absolute-row" \| "absolute-column" => mark_form\(head, &h, args\),') {
        [void]$problems.Add("the port's formula writer (formula_syntax's list) does not take the three marks among its own heads")
    }
    foreach ($x in @(@($addressPrimitive, $addressForm) + $markHeads)) {
        if (-not $headsExpanded.Contains($x)) { [void]$problems.Add("no EXPANDED case heads a form with '$x'") }
        if (-not $headsRefused.Contains($x)) { [void]$problems.Add("no REFUSED case heads a form with '$x'") }
    }
    foreach ($x in $addressIds) {
        if (-not $ids.Contains($x)) { [void]$problems.Add("no REFUSED record of the golden is '$x'") }
    }

    # 10. VLA.9's named pieces: a list in a signature a declaration and a default
    # substituted at the call, in both implementations; its ids each a refused
    # record; and one call in two orders, one record, since the pieces are bound
    # in the declaration's order and never the call's.
    if ($vbaText -notmatch 'DeclarePiece mname, sig\.Item\(j\), params, pieceWords, pieceDecls') {
        [void]$problems.Add("the reference's signature walk (DefineMacro) does not read a list as a piece's declaration")
    }
    if ($vbaText -notmatch 'bindings\.Add Substitute\(decl\.Item\(2\), bindings, "", noItems\), wordFold') {
        [void]$problems.Add("the reference's binding (BindPieces) does not substitute a default under the parameters bound before it")
    }
    if ($portText -notmatch 'pieces\.push\(declare_piece\(&mname, decl, &mut declared\)\?\);') {
        [void]$problems.Add("the port's signature walk (define_macro) does not read a list as a piece's declaration")
    }
    if ($portText -notmatch '\(None, Some\(default\)\) => substitute\(default, env, false\)\?,') {
        [void]$problems.Add("the port's binding (bind_pieces) does not substitute a default under the parameters bound before it")
    }
    foreach ($x in $pieceIds) {
        if (-not $ids.Contains($x)) { [void]$problems.Add("no REFUSED record of the golden is '$x'") }
    }
    $orders = @($gcases | Where-Object { $twoOrders -ccontains $_.Header })
    if ($orders.Count -ne 2) {
        [void]$problems.Add(("{0} of the two cases of one call in two orders in the golden" -f $orders.Count))
    } elseif ((($orders[0].Lines) -join "`n") -cne (($orders[1].Lines) -join "`n")) {
        [void]$problems.Add("one call written in two orders, $($twoOrders[0]) and $($twoOrders[1]), has two records, so a piece is bound by its place")
    }

    # 11. VLA.10's place a verb writes: two forms of the walk alone and the word they
    # read, reserved and no primitive; the stack empty at every entry and exit, read
    # from both implementations; Substitute reading none of it; its ids each a refused
    # record; and two sibling verbs, in either order, each row in each record.
    foreach ($x in $selfForms) {
        if ($vbaText -notmatch ('If headFold = "' + [regex]::Escape($x) + '" Then')) {
            [void]$problems.Add("the reference's walk (ExpandMacros) does not take '$x' at a form's head")
        }
        if ($portText -notmatch ('if head_fold == "' + [regex]::Escape($x) + '" \{')) {
            [void]$problems.Add("the port's walk (expand_at) does not take '$x' at a form's head")
        }
        if (-not $headsExpanded.Contains($x)) { [void]$problems.Add("no EXPANDED case heads a form with '$x'") }
        if (-not $headsRefused.Contains($x)) { [void]$problems.Add("no REFUSED case heads a form with '$x'") }
    }
    foreach ($x in @($selfForms + $selfWord)) {
        if (-not $prt.Reserved.Contains($x)) { [void]$problems.Add("VLA.10's word '$x' is not in the port's RESERVED") }
        if ($prt.Listops.Contains($x)) { [void]$problems.Add("VLA.10's word '$x' is in the port's LISTOPS, which takes data and reads no binding") }
    }
    $resets = [regex]::Matches($vbaText, '(?m)^\s*Set mSelfStack = Nothing\s+'' VLA\.10').Count
    if ($resets -ne 3) {
        [void]$problems.Add("the reference resets the stack at $resets entry point(s), not the three (VlaTranspile, VlaCompileToForms, VlaExpandText)")
    }
    $bindCall = [regex]::Match($vbaText, '(?ms)^Private Function ExpandBindCall\(.*?^End Function')
    if (-not $bindCall.Success) {
        [void]$problems.Add("the reference has no ExpandBindCall for this rule to read")
    } elseif ($bindCall.Value -notmatch '(?s)failed:.*?PopSelf.*?Err\.Raise errNum, errSrc, errDesc') {
        [void]$problems.Add("the reference's ExpandBindCall does not restore the stack on its error path, so a sibling could see another's self")
    }
    if ($portText -notmatch '(?s)let body = self\.bind_body\(&items\[2\.\.\], depth, walk\);\s+self\.self_stack\.pop\(\);\s+body') {
        [void]$problems.Add("the port's bind_expand does not pop the stack before its question mark")
    }
    foreach ($proc in @('Substitute', 'SubstituteValue', 'EvalListopsPrim')) {
        $body = [regex]::Match($vbaText, ('(?ms)^Private Function ' + $proc + '\(.*?^End Function'))
        if (-not $body.Success) {
            [void]$problems.Add("the reference has no $proc for this rule to read")
        } elseif ($body.Value -match 'mSelfStack|VlaSelfDepth') {
            [void]$problems.Add("the reference's $proc reads the stack, so a macro's result is no longer a function of its template and its arguments (LISTOPS-CONFLUENCE)")
        }
    }
    $portSubstitute = [regex]::Match($portText, '(?ms)^fn substitute\(.*?^\}')
    if (-not $portSubstitute.Success) {
        [void]$problems.Add("the port has no substitute for this rule to read")
    } elseif ($portSubstitute.Value -match 'self_stack') {
        [void]$problems.Add("the port's substitute reads the stack (LISTOPS-CONFLUENCE)")
    }
    foreach ($x in $selfIds) {
        if (-not $ids.Contains($x)) { [void]$problems.Add("no REFUSED record of the golden is '$x'") }
    }
    $siblings = @($gcases | Where-Object { $twoSiblings -ccontains $_.Header })
    if ($siblings.Count -ne 2) {
        [void]$problems.Add(("{0} of the two cases of two sibling verbs in the golden" -f $siblings.Count))
    } else {
        foreach ($c in $siblings) {
            $record = ($c.Lines) -join "`n"
            foreach ($row in $siblingRows) {
                if (-not $record.Contains($row)) {
                    [void]$problems.Add("$($c.Header) does not hold the row $row, so a sibling saw another's self")
                }
            }
        }
    }

    # 12. VLA.8's functions and let-expand: the two words reserved in both and no
    # primitive, each heading a form in an EXPANDED case and in a REFUSED one;
    # let-expand taken at a form's head by both walks; a function taken by every
    # entry point as a macro is; the working-out's depth restored on its error
    # path in both; the two limits one number in both; and its ids each a
    # refused record.
    foreach ($x in $functionForms) {
        if (-not $prt.Reserved.Contains($x)) { [void]$problems.Add("VLA.8's word '$x' is not in the port's RESERVED") }
        if ($prt.Listops.Contains($x)) { [void]$problems.Add("VLA.8's word '$x' is in the port's LISTOPS, which takes data") }
        if (-not $headsExpanded.Contains($x)) { [void]$problems.Add("no EXPANDED case heads a form with '$x'") }
        if (-not $headsRefused.Contains($x)) { [void]$problems.Add("no REFUSED case heads a form with '$x'") }
    }
    if ($vbaText -notmatch 'ElseIf headFold = "let-expand" Then') {
        [void]$problems.Add("the reference's walk (ExpandMacros) does not take 'let-expand' at a form's head")
    }
    if ($portText -notmatch 'if head_fold == "let-expand" \|\|') {
        [void]$problems.Add("the port's walk (expand_at) does not take 'let-expand' at a form's head")
    }
    $takes = [regex]::Matches($vbaText, 'ElseIf ListHeadIs\(f, "defexpand"\) Then').Count
    if ($takes -ne 3) {
        [void]$problems.Add("the reference takes a function at $takes entry point(s), not the three (VlaTranspile, VlaCompileToForms, VlaExpandText)")
    }
    if ($portText -notmatch 'if f\.head_is\("defexpand"\) \{') {
        [void]$problems.Add("the port's define does not take a function")
    }
    $workOut = [regex]::Match($vbaText, '(?ms)^Private Function FunctionWorkOut\(.*?^End Function')
    if (-not $workOut.Success) {
        [void]$problems.Add("the reference has no FunctionWorkOut for this rule to read")
    } elseif ($workOut.Value -notmatch '(?s)fail:.*?mFnDepth = mFnDepth - 1.*?Err\.Raise errNum, errSrc, errDesc') {
        [void]$problems.Add("the reference's FunctionWorkOut does not restore the depth on its error path, so a refusal leaves a later working-out a level short")
    }
    $portWorkOut = [regex]::Match($portText, '(?ms)^fn function_work_out\(.*?^\}')
    $portLeave = [regex]::Match($portText, '(?ms)^fn work_out_leave\(.*?^\}')
    if (-not $portWorkOut.Success -or -not $portLeave.Success) {
        [void]$problems.Add("the port has no function_work_out or no work_out_leave for this rule to read")
    } else {
        # The entry's `?` is the one way out before the depth is raised.
        $early = [regex]::Matches($portWorkOut.Value, '\)\?|\breturn\b').Count
        if ($portWorkOut.Value -notmatch '(?s)work_out_enter\(calls, &state\)\?;\s+let result = loop \{.*?\};\s+work_out_leave\(calls, result\)\s+\}$' -or $early -ne 1) {
            [void]$problems.Add("the port's function_work_out answers a way out of its loop other than through work_out_leave, so a refusal leaves a later working-out a level short")
        }
        if ($portLeave.Value -notmatch '(?s)calls\.set_depth\(calls\.depth\(\) - 1\);\s+result\s+\}$') {
            [void]$problems.Add("the port's work_out_leave does not restore the depth before it answers")
        }
    }
    foreach ($pair in @(@('VLA_FUNCTION_DEPTH_LIMIT', 'FUNCTION_DEPTH_LIMIT'), @('VLA_FUNCTION_FUEL', 'FUNCTION_FUEL'))) {
        $refLimit = [regex]::Match($vbaText, ('Private Const ' + $pair[0] + ' As Long = (\d+)'))
        $portLimit = [regex]::Match($portText, ('pub const ' + $pair[1] + ': u32 = (\d+);'))
        if (-not $refLimit.Success -or -not $portLimit.Success) {
            [void]$problems.Add("the limit $($pair[0]) or $($pair[1]) is not written where this rule reads it")
        } elseif ($refLimit.Groups[1].Value -ne $portLimit.Groups[1].Value) {
            [void]$problems.Add("the reference's $($pair[0]) is $($refLimit.Groups[1].Value) and the port's $($pair[1]) $($portLimit.Groups[1].Value)")
        }
    }
    foreach ($x in $functionIds) {
        if (-not $ids.Contains($x)) { [void]$problems.Add("no REFUSED record of the golden is '$x'") }
    }

    # 13. VLA.11's definitions: a macro defined once in a compile, refused by
    # name in both definers; parameter reserved in both and no primitive,
    # heading a form in an EXPANDED case and in a REFUSED one (rule 8 holds
    # both walks passing it by as a define); the compile's parameter words
    # reset at the reference's three entry points and the door's values held
    # to them at each; both golden writers giving a case's ";door " lines to
    # the door, a case with one EXPANDED and one REFUSED; and its four ids
    # each a refused record.
    if ($vbaText -notmatch 'If mMacros\.Exists\(VLA_Identity\.Fold\(mname\)\) Then VLA_Messages\.RaiseMsg "vla-defmacro-twice"') {
        [void]$problems.Add("the reference's DefineMacro does not refuse a macro's second definition")
    }
    if ($portText -notmatch '(?s)if let Some\(first\) = self\.rules\.get\(&mfold\) \{\s+return Err\(raise\(\s+"vla-defmacro-twice"') {
        [void]$problems.Add("the port's define_macro does not refuse a macro's second definition")
    }
    if (-not $prt.Reserved.Contains($parameterForm)) { [void]$problems.Add("VLA.11's word '$parameterForm' is not in the port's RESERVED") }
    if ($prt.Listops.Contains($parameterForm)) { [void]$problems.Add("VLA.11's word '$parameterForm' is in the port's LISTOPS, which takes data") }
    if (-not $headsExpanded.Contains($parameterForm)) { [void]$problems.Add("no EXPANDED case heads a form with '$parameterForm'") }
    if (-not $headsRefused.Contains($parameterForm)) { [void]$problems.Add("no REFUSED case heads a form with '$parameterForm'") }
    $paramResets = [regex]::Matches($vbaText, '(?m)^\s*Set mParameters = Nothing\s+'' VLA\.11').Count
    if ($paramResets -ne 3) {
        [void]$problems.Add("the reference resets its parameters at $paramResets entry point(s), not the three (VlaTranspile, VlaCompileToForms, VlaExpandText)")
    }
    $finishes = [regex]::Matches($vbaText, '(?m)^\s*(If toFixpoint Then )?ParametersFinish\b').Count
    if ($finishes -ne 3) {
        [void]$problems.Add("the reference holds a door's values to its parameters at $finishes entry point(s), not the three")
    }
    $testsText = [System.IO.File]::ReadAllText((Join-Path $repoRoot 'src/VLA_Tests.bas'))
    $reportFor = [regex]::Match($testsText, '(?ms)^Private Function ExpandReportFor\(.*?^End Function')
    if (-not $reportFor.Success -or $reportFor.Value -notmatch '";door "' -or $reportFor.Value -notmatch 'VLA\.VlaSetParameters door') {
        [void]$problems.Add("the reference's golden writer (ExpandReportFor) does not give a case's ;door lines to VlaSetParameters")
    }
    $portGoldenText = [System.IO.File]::ReadAllText((Join-Path $repoRoot 'vla-lang/tests/expand_golden.rs'))
    if ($portGoldenText -notmatch 'strip_prefix\(";door "\)' -or $portGoldenText -notmatch 'expand_text_given\(body, PRELUDE, &door\)') {
        [void]$problems.Add("the port's golden test does not give a case's ;door lines to expand_text_given")
    }
    $doorExpanded = $false
    $doorRefused = $false
    for ($i = 0; $i -lt [Math]::Min($fcases.Count, $gcases.Count); $i++) {
        if (@($fcases[$i].Lines | Where-Object { ([string]$_).StartsWith(';door ') }).Count -eq 0) { continue }
        if ($gcases[$i].Lines.Count -eq 0) { continue }
        $first = [string]$gcases[$i].Lines[0]
        if ($first.StartsWith("EXPANDED`t")) { $doorExpanded = $true }
        if ($first.StartsWith("REFUSED`t")) { $doorRefused = $true }
    }
    if (-not $doorExpanded) { [void]$problems.Add("no EXPANDED case gives a door's value by a ;door line") }
    if (-not $doorRefused) { [void]$problems.Add("no REFUSED case gives a door's value by a ;door line") }
    foreach ($x in $definitionIds) {
        if (-not $ids.Contains($x)) { [void]$problems.Add("no REFUSED record of the golden is '$x'") }
    }

    # 5. the floors
    if ($gcases.Count -lt $caseFloor) { [void]$problems.Add(("{0} cases, below the floor of {1}" -f $gcases.Count, $caseFloor)) }
    if ($expanded -lt $expandedFloor) { [void]$problems.Add(("{0} expanded cases, below the floor of {1}" -f $expanded, $expandedFloor)) }
    if ($ids.Count -lt $idFloor) { [void]$problems.Add(("{0} distinct ids refused, below the floor of {1}" -f $ids.Count, $idFloor)) }

    return @{
        Problems = $problems; Cases = $gcases.Count; Expanded = $expanded; Ids = $ids.Count
        Primitives = $prt.Listops.Count; Reserved = $prt.Reserved.Count
    }
}

if ($Control) {
    $tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('frazaro_expand_' + [System.IO.Path]::GetRandomFileName())
    New-Item -ItemType Directory -Path $tmp | Out-Null
    try {
        $utf8 = New-Object System.Text.UTF8Encoding($false)
        $real = Test-ExpandGolden $fixturePath $goldenPath $vbaPath $portPath $catalogPath $syntaxPath
        $lines = Get-ExpandLines $goldenPath
        $firstExpanded = -1
        $firstRefused = -1
        $lastHeader = -1
        for ($i = 0; $i -lt $lines.Count; $i++) {
            if ($firstExpanded -lt 0 -and $lines[$i].StartsWith("EXPANDED`t")) { $firstExpanded = $i }
            if ($firstRefused -lt 0 -and $lines[$i].StartsWith("REFUSED`t")) { $firstRefused = $i }
            if ($lines[$i].StartsWith('=== ')) { $lastHeader = $i }
        }
        if ($firstExpanded -lt 0 -or $firstRefused -lt 0 -or $lastHeader -lt 1) {
            Write-Output 'FAIL: the control found no EXPANDED record, no REFUSED record or no last header to mutate'
            exit 1
        }
        # (a) one record's tag misspelled
        $a = @($lines | ForEach-Object { $_ })
        $a[$firstExpanded] = 'EXPANDD' + $a[$firstExpanded].Substring(8)
        $pa = Join-Path $tmp 'a.txt'
        [System.IO.File]::WriteAllText($pa, ($a -join "`n"), $utf8)
        # (b) the last case dropped
        $b = @($lines[0..($lastHeader - 1)])
        $pb = Join-Path $tmp 'b.txt'
        [System.IO.File]::WriteAllText($pb, ($b -join "`n"), $utf8)
        # (c) a refusal with an id the catalogue lacks
        $c = @($lines | ForEach-Object { $_ })
        $c[$firstRefused] = ($c[$firstRefused] -creplace '^REFUSED\t[a-z0-9-]+\t', "REFUSED`tvla-no-such-id`t")
        if ($c[$firstRefused] -cnotmatch 'vla-no-such-id') { Write-Output 'FAIL: the control could not plant an unknown id'; exit 1 }
        $pc = Join-Path $tmp 'c.txt'
        [System.IO.File]::WriteAllText($pc, ($c -join "`n"), $utf8)
        # (d) the reference with row-expand taken from Substitute's Select, the second of the three
        $vbaText = [System.IO.File]::ReadAllText($vbaPath)
        $anchor = '"address-expand", "row-expand", "column-expand", "column-letters-expand"'
        $at = [regex]::Matches($vbaText, [regex]::Escape($anchor))
        if ($at.Count -ne 3) { Write-Output "FAIL: the control expected the Selects' last line three times in the reference, found $($at.Count)"; exit 1 }
        $last = $at[1].Index
        $dText = $vbaText.Substring(0, $last) + '"address-expand", "column-expand", "column-letters-expand"' + $vbaText.Substring($last + $anchor.Length)
        $pd = Join-Path $tmp 'd.bas'
        [System.IO.File]::WriteAllText($pd, $dText, $utf8)
        # (e) the fixture with every characters-expand renamed away
        $fixText = [System.IO.File]::ReadAllText($fixturePath)
        $renamed = [regex]::Matches($fixText, [regex]::Escape('(characters-expand')).Count
        if ($renamed -lt 1) { Write-Output 'FAIL: the control found no characters-expand in the fixture'; exit 1 }
        $pe = Join-Path $tmp 'e.txt'
        [System.IO.File]::WriteAllText($pe, $fixText.Replace('(characters-expand', '(chars-expand'), $utf8)
        # (f) the reference whose walk no longer recognizes formula-expand
        $walkAnchor = 'If headFold = "formula-expand" Then'
        if ([regex]::Matches($vbaText, [regex]::Escape($walkAnchor)).Count -ne 1) { Write-Output 'FAIL: the control expected the walk''s formula-expand test once in the reference'; exit 1 }
        $pf = Join-Path $tmp 'f.bas'
        [System.IO.File]::WriteAllText($pf, $vbaText.Replace($walkAnchor, 'If headFold = "formula-expandd" Then'), $utf8)
        # (g) the reference whose formula writer resolves an alias by a compile's state
        $aliasAnchor = 'h = FsCanonicalHead(VLA_Identity.Fold(headWritten))'
        if ([regex]::Matches($vbaText, [regex]::Escape($aliasAnchor)).Count -ne 1) { Write-Output 'FAIL: the control expected FsList''s FsCanonicalHead call once in the reference'; exit 1 }
        $pg = Join-Path $tmp 'g.bas'
        [System.IO.File]::WriteAllText($pg, $vbaText.Replace($aliasAnchor, 'h = ResolveHeadAlias(VLA_Identity.Fold(headWritten))'), $utf8)
        # (h) the reference whose Substitute no longer takes the loop at a list's head
        $loopAnchor = 'Case "each", "each-index"'
        if ([regex]::Matches($vbaText, [regex]::Escape($loopAnchor)).Count -ne 1) { Write-Output 'FAIL: the control expected Substitute''s loop Case once in the reference'; exit 1 }
        $ph = Join-Path $tmp 'h.bas'
        [System.IO.File]::WriteAllText($ph, $vbaText.Replace($loopAnchor, 'Case "eachh", "each-index"'), $utf8)
        # (i) the port whose RESERVED lost each
        $portSource = [System.IO.File]::ReadAllText($portPath)
        $reservedRow = '(?m)^    \("each", "expand-time loop form"\),\r?$'
        if ([regex]::Matches($portSource, $reservedRow).Count -ne 1) { Write-Output 'FAIL: the control expected the port''s RESERVED row for each once'; exit 1 }
        $pi = Join-Path $tmp 'i.rs'
        [System.IO.File]::WriteAllText($pi, [regex]::Replace($portSource, $reservedRow, ''), $utf8)
        # (j) the reference whose walk no longer passes a define by
        $defineAnchor = 'If headFold = "define" Or headFold = "parameter" Then'
        if ([regex]::Matches($vbaText, [regex]::Escape($defineAnchor)).Count -ne 1) { Write-Output 'FAIL: the control expected the walk''s define test once in the reference'; exit 1 }
        $pj = Join-Path $tmp 'j.bas'
        [System.IO.File]::WriteAllText($pj, $vbaText.Replace($defineAnchor, 'If headFold = "definee" Or headFold = "parameterr" Then'), $utf8)
        # (k) the reference whose VlaExpandText expands a form by ExpandMacros directly
        $wrapperAnchor = 'AssignVar ex, ExpandSourceForm(f)'
        if ([regex]::Matches($vbaText, [regex]::Escape($wrapperAnchor)).Count -ne 1) { Write-Output 'FAIL: the control expected VlaExpandText''s call of the wrapper once in the reference'; exit 1 }
        $pk = Join-Path $tmp 'k.bas'
        [System.IO.File]::WriteAllText($pk, $vbaText.Replace($wrapperAnchor, 'AssignVar ex, ExpandMacros(f, 0)'), $utf8)
        # (l) the port whose writer lost last
        $syntaxSource = [System.IO.File]::ReadAllText($syntaxPath)
        $lastArm = '"last" => last_form(head, args),'
        if ([regex]::Matches($syntaxSource, [regex]::Escape($lastArm)).Count -ne 1) { Write-Output 'FAIL: the control expected the port''s writer arm for last once'; exit 1 }
        $pl = Join-Path $tmp 'l.rs'
        [System.IO.File]::WriteAllText($pl, $syntaxSource.Replace($lastArm, ''), $utf8)
        # (m) the reference whose walk no longer takes a reference's first cell
        $firstCellAnchor = 'If headFold = "first-cell-expand" Then'
        if ([regex]::Matches($vbaText, [regex]::Escape($firstCellAnchor)).Count -ne 1) { Write-Output 'FAIL: the control expected the walk''s first-cell-expand test once in the reference'; exit 1 }
        $pm = Join-Path $tmp 'm.bas'
        [System.IO.File]::WriteAllText($pm, $vbaText.Replace($firstCellAnchor, 'If headFold = "first-cell-expandd" Then'), $utf8)
        # (n) the port whose writer lost the three marks
        $markArm = '"absolute" | "absolute-row" | "absolute-column" => mark_form(head, &h, args),'
        if ([regex]::Matches($syntaxSource, [regex]::Escape($markArm)).Count -ne 1) { Write-Output 'FAIL: the control expected the port''s writer arm for the three marks once'; exit 1 }
        $pn = Join-Path $tmp 'n.rs'
        [System.IO.File]::WriteAllText($pn, $syntaxSource.Replace($markArm, ''), $utf8)
        # (o) the reference whose signature walk reads no declaration
        $declareAnchor = 'DeclarePiece mname, sig.Item(j), params, pieceWords, pieceDecls'
        if ([regex]::Matches($vbaText, [regex]::Escape($declareAnchor)).Count -ne 1) { Write-Output 'FAIL: the control expected DefineMacro''s call of DeclarePiece once in the reference'; exit 1 }
        $po = Join-Path $tmp 'o.bas'
        [System.IO.File]::WriteAllText($po, $vbaText.Replace($declareAnchor, 'SymText sig.Item(j)'), $utf8)
        # (p) the port whose binding takes no default
        $defaultArm = '(None, Some(default)) => substitute(default, env, false)?,'
        if ([regex]::Matches($portSource, [regex]::Escape($defaultArm)).Count -ne 1) { Write-Output 'FAIL: the control expected the port''s arm for a default once'; exit 1 }
        $pp = Join-Path $tmp 'p.rs'
        [System.IO.File]::WriteAllText($pp, $portSource.Replace($defaultArm, '(None, Some(default)) => default.clone(),'), $utf8)
        # (q) a golden whose second order of one call differs from its first
        $q = @($lines | ForEach-Object { $_ })
        $orderAt = [array]::IndexOf($q, $twoOrders[1])
        if ($orderAt -lt 0) { Write-Output 'FAIL: the control found no second order of one call in the golden'; exit 1 }
        $q[$orderAt + 1] = "EXPANDED`t9"
        $pq = Join-Path $tmp 'q.txt'
        [System.IO.File]::WriteAllText($pq, ($q -join "`n"), $utf8)
        # (r) the reference whose walk no longer takes bind-expand
        $bindAnchor = 'If headFold = "bind-expand" Then'
        if ([regex]::Matches($vbaText, [regex]::Escape($bindAnchor)).Count -ne 1) { Write-Output 'FAIL: the control expected the walk''s bind-expand test once in the reference'; exit 1 }
        $pr = Join-Path $tmp 'r.bas'
        [System.IO.File]::WriteAllText($pr, $vbaText.Replace($bindAnchor, 'If headFold = "bind-expandd" Then'), $utf8)
        # (s) the reference whose handler leaves the stack as it stood
        $handlerAnchor = [regex]'(?s)(Private Function ExpandBindCall\(.*?failed:.*?)PopSelf'
        if (-not $handlerAnchor.IsMatch($vbaText)) { Write-Output 'FAIL: the control expected ExpandBindCall''s handler to pop the stack'; exit 1 }
        $ps = Join-Path $tmp 's.bas'
        [System.IO.File]::WriteAllText($ps, $handlerAnchor.Replace($vbaText, '$1', 1), $utf8)
        # (t) the reference whose Substitute reads the stack
        $substituteAnchor = 'Private Function Substitute(t As Variant, bindings As Collection, ByVal restName As String, restItems As Collection, Optional ByVal inQuasi As Boolean = False) As Variant'
        if ([regex]::Matches($vbaText, [regex]::Escape($substituteAnchor)).Count -ne 1) { Write-Output 'FAIL: the control expected Substitute''s signature once in the reference'; exit 1 }
        $pt = Join-Path $tmp 't.bas'
        [System.IO.File]::WriteAllText($pt, $vbaText.Replace($substituteAnchor, $substituteAnchor + "`r`n    If mSelfStack Is Nothing Then Exit Function"), $utf8)
        # (u) the port whose binding never pops
        $popAnchor = 'self.self_stack.pop();'
        if ([regex]::Matches($portSource, [regex]::Escape($popAnchor)).Count -ne 1) { Write-Output 'FAIL: the control expected the port''s pop once'; exit 1 }
        $pu = Join-Path $tmp 'u.rs'
        [System.IO.File]::WriteAllText($pu, $portSource.Replace($popAnchor, ''), $utf8)
        # (v) a golden whose second sibling writes the first's place
        $v = @($lines | ForEach-Object { $_ })
        $siblingAt = [array]::IndexOf($v, $twoSiblings[1])
        if ($siblingAt -lt 0) { Write-Output 'FAIL: the control found no second sibling in the golden'; exit 1 }
        $v[$siblingAt + 2] = $v[$siblingAt + 2].Replace('=C.last!D4', '=A.last!B2')
        if ($v[$siblingAt + 2] -match '=C\.last!D4') { Write-Output 'FAIL: the control could not plant the first''s place'; exit 1 }
        $pv = Join-Path $tmp 'v.txt'
        [System.IO.File]::WriteAllText($pv, ($v -join "`n"), $utf8)
        # (w) the reference whose working-out leaves the depth as it stood on a refusal
        $depthAnchor = [regex]'(?s)(Private Function FunctionWorkOut\(.*?fail:.*?)mFnDepth = mFnDepth - 1'
        if (-not $depthAnchor.IsMatch($vbaText)) { Write-Output 'FAIL: the control expected FunctionWorkOut''s handler to lower the depth'; exit 1 }
        $pw = Join-Path $tmp 'w.bas'
        [System.IO.File]::WriteAllText($pw, $depthAnchor.Replace($vbaText, '$1', 1), $utf8)
        # (x) the port whose working-out never lowers the depth
        $lowerAnchor = 'calls.set_depth(calls.depth() - 1);'
        if ([regex]::Matches($portSource, [regex]::Escape($lowerAnchor)).Count -ne 1) { Write-Output 'FAIL: the control expected the port''s lowering of the depth once'; exit 1 }
        $px = Join-Path $tmp 'x.rs'
        [System.IO.File]::WriteAllText($px, $portSource.Replace($lowerAnchor, ''), $utf8)
        # (y) the port whose loop answers a refusal of its fuel by `?`, past the lowering
        $fuelAnchor = [regex]'if let Err\(e\) = work_out_fuel\(calls, &state\) \{\s+break Err\(e\);\s+\}'
        if ($fuelAnchor.Matches($portSource).Count -ne 1) { Write-Output 'FAIL: the control expected the port''s loop to answer its fuel once'; exit 1 }
        $py = Join-Path $tmp 'y.rs'
        [System.IO.File]::WriteAllText($py, $fuelAnchor.Replace($portSource, 'work_out_fuel(calls, &state)?;', 1), $utf8)
        # (z) the reference whose DefineMacro lets a second definition replace the first
        $twiceAnchor = 'If mMacros.Exists(VLA_Identity.Fold(mname)) Then VLA_Messages.RaiseMsg "vla-defmacro-twice"'
        if ([regex]::Matches($vbaText, [regex]::Escape($twiceAnchor)).Count -ne 1) { Write-Output 'FAIL: the control expected DefineMacro''s refusal of a second definition once'; exit 1 }
        $pz = Join-Path $tmp 'z.bas'
        [System.IO.File]::WriteAllText($pz, $vbaText.Replace($twiceAnchor, 'If False Then VLA_Messages.RaiseMsg "vla-defmacro-twice"'), $utf8)
        # (aa) the port whose define_macro does
        $portTwiceAnchor = 'if let Some(first) = self.rules.get(&mfold) {'
        if ([regex]::Matches($portSource, [regex]::Escape($portTwiceAnchor)).Count -ne 1) { Write-Output 'FAIL: the control expected define_macro''s refusal of a second definition once'; exit 1 }
        $paa = Join-Path $tmp 'aa.rs'
        [System.IO.File]::WriteAllText($paa, $portSource.Replace($portTwiceAnchor, 'if let Some(first) = None::<&Rule> {'), $utf8)
        # (ab) the reference whose VlaCompileToForms no longer holds a door's values
        $finishAnchor = "    ParametersFinish   ' VLA.11"
        if ([regex]::Matches($vbaText, [regex]::Escape($finishAnchor)).Count -ne 1) { Write-Output 'FAIL: the control expected VlaCompileToForms'' ParametersFinish once'; exit 1 }
        $pab = Join-Path $tmp 'ab.bas'
        [System.IO.File]::WriteAllText($pab, $vbaText.Replace($finishAnchor, "    ' VLA.11"), $utf8)
        $ra = Test-ExpandGolden $fixturePath $pa $vbaPath $portPath $catalogPath $syntaxPath
        $rb = Test-ExpandGolden $fixturePath $pb $vbaPath $portPath $catalogPath $syntaxPath
        $rc = Test-ExpandGolden $fixturePath $pc $vbaPath $portPath $catalogPath $syntaxPath
        $rd = Test-ExpandGolden $fixturePath $goldenPath $pd $portPath $catalogPath $syntaxPath
        $re = Test-ExpandGolden $pe $goldenPath $vbaPath $portPath $catalogPath $syntaxPath
        $rf = Test-ExpandGolden $fixturePath $goldenPath $pf $portPath $catalogPath $syntaxPath
        $rg = Test-ExpandGolden $fixturePath $goldenPath $pg $portPath $catalogPath $syntaxPath
        $rh = Test-ExpandGolden $fixturePath $goldenPath $ph $portPath $catalogPath $syntaxPath
        $ri = Test-ExpandGolden $fixturePath $goldenPath $vbaPath $pi $catalogPath $syntaxPath
        $rj = Test-ExpandGolden $fixturePath $goldenPath $pj $portPath $catalogPath $syntaxPath
        $rk = Test-ExpandGolden $fixturePath $goldenPath $pk $portPath $catalogPath $syntaxPath
        $rl = Test-ExpandGolden $fixturePath $goldenPath $vbaPath $portPath $catalogPath $pl
        $rm = Test-ExpandGolden $fixturePath $goldenPath $pm $portPath $catalogPath $syntaxPath
        $rn = Test-ExpandGolden $fixturePath $goldenPath $vbaPath $portPath $catalogPath $pn
        $ro = Test-ExpandGolden $fixturePath $goldenPath $po $portPath $catalogPath $syntaxPath
        $rp = Test-ExpandGolden $fixturePath $goldenPath $vbaPath $pp $catalogPath $syntaxPath
        $rq = Test-ExpandGolden $fixturePath $pq $vbaPath $portPath $catalogPath $syntaxPath
        $rr = Test-ExpandGolden $fixturePath $goldenPath $pr $portPath $catalogPath $syntaxPath
        $rs = Test-ExpandGolden $fixturePath $goldenPath $ps $portPath $catalogPath $syntaxPath
        $rt = Test-ExpandGolden $fixturePath $goldenPath $pt $portPath $catalogPath $syntaxPath
        $ru = Test-ExpandGolden $fixturePath $goldenPath $vbaPath $pu $catalogPath $syntaxPath
        $rv = Test-ExpandGolden $fixturePath $pv $vbaPath $portPath $catalogPath $syntaxPath
        $rw = Test-ExpandGolden $fixturePath $goldenPath $pw $portPath $catalogPath $syntaxPath
        $rx = Test-ExpandGolden $fixturePath $goldenPath $vbaPath $px $catalogPath $syntaxPath
        $ry = Test-ExpandGolden $fixturePath $goldenPath $vbaPath $py $catalogPath $syntaxPath
        $rz = Test-ExpandGolden $fixturePath $goldenPath $pz $portPath $catalogPath $syntaxPath
        $raa = Test-ExpandGolden $fixturePath $goldenPath $vbaPath $paa $catalogPath $syntaxPath
        $rab = Test-ExpandGolden $fixturePath $goldenPath $pab $portPath $catalogPath $syntaxPath
        Write-Output ("control: the real set has {0} problem(s); the misspelt tag {1}, the dropped case {2}, the unknown id {3}, the reference's broken Select {4}, the renamed primitive {5}, the walk without formula-expand {6}, the writer reading a compile's aliases {7}, the substitution without the loop {8}, the port's RESERVED without each {9}, the walk passing no define by {10}, VlaExpandText past the wrapper {11}, the port's writer without last {12}, the walk without the first cell {13}, the port's writer without the marks {14}, the signature walk with no declaration {15}, the binding with no default {16}, one call's two orders apart {17}, the walk without bind-expand {18}, the handler leaving the stack {19}, Substitute reading the stack {20}, the port's binding never popping {21}, a sibling writing the first's place {22}, the working-out leaving the depth {23}, the port's working-out never lowering it {24}, the fuel leaving by ? {25}, the reference replacing a macro {26}, the port replacing one {27}, a compile holding no door's values {28}" -f $real.Problems.Count, $ra.Problems.Count, $rb.Problems.Count, $rc.Problems.Count, $rd.Problems.Count, $re.Problems.Count, $rf.Problems.Count, $rg.Problems.Count, $rh.Problems.Count, $ri.Problems.Count, $rj.Problems.Count, $rk.Problems.Count, $rl.Problems.Count, $rm.Problems.Count, $rn.Problems.Count, $ro.Problems.Count, $rp.Problems.Count, $rq.Problems.Count, $rr.Problems.Count, $rs.Problems.Count, $rt.Problems.Count, $ru.Problems.Count, $rv.Problems.Count, $rw.Problems.Count, $rx.Problems.Count, $ry.Problems.Count, $rz.Problems.Count, $raa.Problems.Count, $rab.Problems.Count)
        if ($real.Problems.Count -eq 0 -and $ra.Problems.Count -ge 1 -and $rb.Problems.Count -ge 1 -and $rc.Problems.Count -ge 1 -and $rd.Problems.Count -ge 1 -and $re.Problems.Count -ge 1 -and $rf.Problems.Count -ge 1 -and $rg.Problems.Count -ge 1 -and $rh.Problems.Count -ge 1 -and $ri.Problems.Count -ge 1 -and $rj.Problems.Count -ge 1 -and $rk.Problems.Count -ge 1 -and $rl.Problems.Count -ge 1 -and $rm.Problems.Count -ge 1 -and $rn.Problems.Count -ge 1 -and $ro.Problems.Count -ge 1 -and $rp.Problems.Count -ge 1 -and $rq.Problems.Count -ge 1 -and $rr.Problems.Count -ge 1 -and $rs.Problems.Count -ge 1 -and $rt.Problems.Count -ge 1 -and $ru.Problems.Count -ge 1 -and $rv.Problems.Count -ge 1 -and $rw.Problems.Count -ge 1 -and $rx.Problems.Count -ge 1 -and $ry.Problems.Count -ge 1 -and $rz.Problems.Count -ge 1 -and $raa.Problems.Count -ge 1 -and $rab.Problems.Count -ge 1) {
            Write-Output 'OK: the check passes the golden and fails each mutant'
            exit 0
        }
        Write-Output 'FAIL: the control did not behave'
        foreach ($p in $real.Problems) { Write-Output ('  real: ' + $p) }
        exit 1
    } finally {
        Remove-Item -Recurse -Force $tmp -ErrorAction SilentlyContinue
    }
}

$r = Test-ExpandGolden $fixturePath $goldenPath $vbaPath $portPath $catalogPath $syntaxPath
if ($r.Problems.Count -eq 0) {
    $above = ($r.Cases -gt $caseFloor) -or ($r.Expanded -gt $expandedFloor) -or ($r.Ids -gt $idFloor)
    $note = if ($above) { ' - above a floor: raise the floors in this file' } else { '' }
    Write-Output ("OK: {0} cases, {1} expanded, {2} distinct ids refused; {3} primitives and {4} reserved words, the reference's and the port's alike, every primitive reached (floors {5}, {6}, {7}){8}" -f $r.Cases, $r.Expanded, $r.Ids, $r.Primitives, $r.Reserved, $caseFloor, $expandedFloor, $idFloor, $note)
    exit 0
}
Write-Output ("FAIL: the expand golden has {0} problem(s):" -f $r.Problems.Count)
foreach ($p in $r.Problems) { Write-Output ('  ' + $p) }
exit 1
