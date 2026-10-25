<#
check_expand_golden.ps1 - the expand golden agrees with its fixture, every
record is EXPANDED with its text or REFUSED with a catalogue id, the
reference's lists of the expand-time primitives are the port's, and every
primitive is reached.

WHY: VLA.3 (the expand-time object model) added ten primitives to the
seventeen, in both implementations. The reference names each in four places
(src/VLA.bas: DefineMacro's reserved names, ExpandMacros' standalone Select,
Substitute's inline Select, EvalListopsPrim's dispatch) and the port in two
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
     two recognition Selects and EvalListopsPrim's Cases are the port's
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

-Control proves the check on scratch copies: the real files must pass, and
each of these must fail: a golden with one record's tag misspelled, one
with its last case dropped, one refusing with an id the catalogue lacks, a
reference with one name taken from Substitute's Select, a fixture with
every characters-expand renamed away, a reference whose walk no longer
recognizes formula-expand, a reference whose formula writer resolves an
alias through ResolveHeadAlias, a reference whose Substitute no longer
takes the loop at a list's head, and a port whose RESERVED lost each.

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
$caseFloor = 229
$expandedFloor = 129
$idFloor = 36

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
# and its six ids, which rule 7 holds to both implementations and the fixture.
$loopForms = @('each', 'each-index')
$loopPrimitives = @('assoc-expand')
$loopIds = @('vla-each-shape', 'vla-each-name-reserved', 'vla-each-name-bound',
    'vla-each-list-a-call', 'vla-expand-expected-list', 'vla-assoc-expand-entry')

$fixturePath = Join-Path $repoRoot 'scripts/expand.txt'
$goldenPath  = Join-Path $repoRoot 'scripts/expand_golden.txt'
$vbaPath     = Join-Path $repoRoot 'src/VLA.bas'
$portPath    = Join-Path $repoRoot 'vla-lang/src/expand.rs'
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

# Every problem of one fixture/golden/reference/port/catalogue set, with the
# counts the floors read.
function Test-ExpandGolden([string]$fixture, [string]$golden, [string]$vba, [string]$port, [string]$catalog) {
    $problems = New-Object System.Collections.ArrayList
    $fcases = Read-ExpandCases $fixture
    $gcases = Read-ExpandCases $golden
    $ref = Read-ReferenceLists $vba
    $prt = Read-PortLists $port
    $catalogIds = Read-CatalogIds $catalog

    # 3. the lists
    if ($prt.Listops.Count -eq 0) { [void]$problems.Add("no LISTOPS array in $port") }
    if ($prt.Reserved.Count -eq 0) { [void]$problems.Add("no RESERVED array in $port") }
    if ($ref.Selects.Count -ne 2) {
        [void]$problems.Add(("{0} recognition Select(s) opening on `"car`", `"cdr`" in {1}, not the two (ExpandMacros, Substitute)" -f $ref.Selects.Count, $vba))
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
        $real = Test-ExpandGolden $fixturePath $goldenPath $vbaPath $portPath $catalogPath
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
        # (d) the reference with row-expand taken from Substitute's Select, the second of the two
        $vbaText = [System.IO.File]::ReadAllText($vbaPath)
        $anchor = '"address-expand", "row-expand", "column-expand", "column-letters-expand"'
        $at = [regex]::Matches($vbaText, [regex]::Escape($anchor))
        if ($at.Count -ne 2) { Write-Output "FAIL: the control expected the Selects' last line twice in the reference, found $($at.Count)"; exit 1 }
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
        $ra = Test-ExpandGolden $fixturePath $pa $vbaPath $portPath $catalogPath
        $rb = Test-ExpandGolden $fixturePath $pb $vbaPath $portPath $catalogPath
        $rc = Test-ExpandGolden $fixturePath $pc $vbaPath $portPath $catalogPath
        $rd = Test-ExpandGolden $fixturePath $goldenPath $pd $portPath $catalogPath
        $re = Test-ExpandGolden $pe $goldenPath $vbaPath $portPath $catalogPath
        $rf = Test-ExpandGolden $fixturePath $goldenPath $pf $portPath $catalogPath
        $rg = Test-ExpandGolden $fixturePath $goldenPath $pg $portPath $catalogPath
        $rh = Test-ExpandGolden $fixturePath $goldenPath $ph $portPath $catalogPath
        $ri = Test-ExpandGolden $fixturePath $goldenPath $vbaPath $pi $catalogPath
        Write-Output ("control: the real set has {0} problem(s); the misspelt tag {1}, the dropped case {2}, the unknown id {3}, the reference's broken Select {4}, the renamed primitive {5}, the walk without formula-expand {6}, the writer reading a compile's aliases {7}, the substitution without the loop {8}, the port's RESERVED without each {9}" -f $real.Problems.Count, $ra.Problems.Count, $rb.Problems.Count, $rc.Problems.Count, $rd.Problems.Count, $re.Problems.Count, $rf.Problems.Count, $rg.Problems.Count, $rh.Problems.Count, $ri.Problems.Count)
        if ($real.Problems.Count -eq 0 -and $ra.Problems.Count -ge 1 -and $rb.Problems.Count -ge 1 -and $rc.Problems.Count -ge 1 -and $rd.Problems.Count -ge 1 -and $re.Problems.Count -ge 1 -and $rf.Problems.Count -ge 1 -and $rg.Problems.Count -ge 1 -and $rh.Problems.Count -ge 1 -and $ri.Problems.Count -ge 1) {
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

$r = Test-ExpandGolden $fixturePath $goldenPath $vbaPath $portPath $catalogPath
if ($r.Problems.Count -eq 0) {
    $above = ($r.Cases -gt $caseFloor) -or ($r.Expanded -gt $expandedFloor) -or ($r.Ids -gt $idFloor)
    $note = if ($above) { ' - above a floor: raise the floors in this file' } else { '' }
    Write-Output ("OK: {0} cases, {1} expanded, {2} distinct ids refused; {3} primitives and {4} reserved words, the reference's and the port's alike, every primitive reached (floors {5}, {6}, {7}){8}" -f $r.Cases, $r.Expanded, $r.Ids, $r.Primitives, $r.Reserved, $caseFloor, $expandedFloor, $idFloor, $note)
    exit 0
}
Write-Output ("FAIL: the expand golden has {0} problem(s):" -f $r.Problems.Count)
foreach ($p in $r.Problems) { Write-Output ('  ' + $p) }
exit 1
