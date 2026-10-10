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
     LISTOPS, and DefineMacro's reserved words are the port's RESERVED
     (read from the sources, never retyped here);
  4. every primitive of LISTOPS heads a form in at least one fixture case;
     each of VLA.3's ten heads one in an EXPANDED case and in a REFUSED
     case; and each of VLA.3's eight ids is the id of a REFUSED record;
  5. the floors below: the cases, the expanded cases and the distinct ids.

-Control proves the check on scratch copies: the real files must pass, and
each of these must fail: a golden with one record's tag misspelled, one
with its last case dropped, one refusing with an id the catalogue lacks, a
reference with one name taken from Substitute's Select, and a fixture with
every characters-expand renamed away.

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
$caseFloor = 135
$expandedFloor = 75
$idFloor = 19

# VLA.3's ten primitives and eight ids, which rule 4 holds to the fixture.
$newPrimitives = @('-expand', '*expand', 'quotient-expand', 'mod-expand', '&expand',
    'characters-expand', 'address-expand', 'row-expand', 'column-expand', 'column-letters-expand')
$newIds = @('vla-expand-arity', 'vla-expand-arg-is-list', 'vla-concat-expand-arity',
    'vla-arith-expand-not-whole', 'vla-arith-expand-divide-by-zero',
    'vla-address-expand-row', 'vla-address-expand-column', 'vla-address-expand-not-address')

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
# of EvalListopsPrim, and DefineMacro's reserved words.
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
        foreach ($q in [regex]::Matches($ln, 'RaiseMsg "vla-defmacro-reserved-name", "word", "([^"]+)"')) {
            [void]$reserved.Add($q.Groups[1].Value)
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
    foreach ($p in (Format-ExpandSetDiff 'reserved word' $ref.Reserved "the reference's DefineMacro" $prt.Reserved "the port's RESERVED")) { [void]$problems.Add($p) }
    foreach ($x in $prt.Listops) {
        if (-not $prt.Reserved.Contains($x)) { [void]$problems.Add("primitive '$x' is recognized by the port and not reserved") }
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
        $ra = Test-ExpandGolden $fixturePath $pa $vbaPath $portPath $catalogPath
        $rb = Test-ExpandGolden $fixturePath $pb $vbaPath $portPath $catalogPath
        $rc = Test-ExpandGolden $fixturePath $pc $vbaPath $portPath $catalogPath
        $rd = Test-ExpandGolden $fixturePath $goldenPath $pd $portPath $catalogPath
        $re = Test-ExpandGolden $pe $goldenPath $vbaPath $portPath $catalogPath
        Write-Output ("control: the real set has {0} problem(s); the misspelt tag {1}, the dropped case {2}, the unknown id {3}, the reference's broken Select {4}, the renamed primitive {5}" -f $real.Problems.Count, $ra.Problems.Count, $rb.Problems.Count, $rc.Problems.Count, $rd.Problems.Count, $re.Problems.Count)
        if ($real.Problems.Count -eq 0 -and $ra.Problems.Count -ge 1 -and $rb.Problems.Count -ge 1 -and $rc.Problems.Count -ge 1 -and $rd.Problems.Count -ge 1 -and $re.Problems.Count -ge 1) {
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
