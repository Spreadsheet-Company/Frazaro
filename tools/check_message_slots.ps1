<#
=== every refusal supplies exactly the slots its own template names ===

WHY THIS EXISTS, and like its sibling check_module_heads.ps1 it is a
fresh scar. OPTIMIZE.1's follow-up appended a `{defined}` slot to two
message templates and updated the two raise sites in the procedure it was
reading. There were FOUR sites for those two ids. The other two - one in
the keyed-atom desugaring path, one a defensive post-fixpoint guard -
still raised without the new slot, so instead of a refusal the user got

    RaiseMsg: no value supplied for slot '{defined}'

which is VLA_Messages doing exactly the right thing loudly, and it was
caught by a test that happened to exercise that path. Nothing mechanical
was watching, and all 24 checks passed.

THE REASON A STATIC CHECK IS WORTH IT HERE, rather than leaning on the
suite: there are 519 catalogue entries and 713 raise sites, and no test
exercises most of them. A refusal path with a missing slot ships broken
and announces itself only when a user finally reaches it - at which point
the message they get is about Frazaro's internals instead of about their
own spreadsheet. That is the worst moment for a refusal to fail.

WHAT IS CHECKED, over both catalogues (VLA_Messages.bas's, and
VLA_Runtime.bas's own mini-catalogue below the inject boundary):

  1. Every `RaiseMsg`/`RaiseRuntimeMsg` site names an id the matching
     catalogue actually holds. A site naming an unknown id would raise
     "unknown message id" instead of its refusal.
  2. Every site supplies EVERY slot its template names. This is the
     defect above.
  3. No site supplies a slot the template does NOT name. Those are
     silently ignored at runtime, so a mistyped slot name ("{cells}" for
     "{cell}") would leave the real slot unfilled AND look deliberate at
     the call site.
  4. Reported, not failed: catalogue entries no site raises. Some are
     deliberate - a reserved id, or one raised from a path this scan
     cannot see - so this is information, not a verdict.

Slot names and ids must be string LITERALS at the call site, which every
one of them is today; a site whose id or slot name were computed would be
reported as unparseable rather than silently skipped.

House shape, per the standing convention for this project's static
scans: PowerShell, host-independent, hardcoded and reviewable baseline,
never wired into VlaSelfTest.

Usage:  powershell -File tools\check_message_slots.ps1 [-List]
Exit 0 clean, exit 1 with every failure listed.
#>
param([switch]$List)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$srcDir = Join-Path $root 'src'
$failures = New-Object System.Collections.Generic.List[string]

# --- VBA source, with line continuations joined -----------------------
function Get-LogicalLines([string]$path) {
    $lines = [System.IO.File]::ReadAllText($path) -split "`r`n|`n"
    $out = New-Object System.Collections.Generic.List[object]
    $buf = ''
    $startNo = 0
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $ln = $lines[$i]
        if ($buf -eq '') { $startNo = $i + 1 }
        if ($ln -match '\s_\s*$') {
            $buf = $buf + ($ln -replace '\s_\s*$', ' ')
            continue
        }
        $out.Add([pscustomobject]@{ No = $startNo; Text = $buf + $ln })
        $buf = ''
    }
    if ($buf -ne '') { $out.Add([pscustomobject]@{ No = $startNo; Text = $buf }) }
    # PowerShell unrolls a collection on return, and a one-element result
    # comes back as a bare scalar whose [0] is its first CHARACTER rather
    # than its first element. The leading comma stops the unroll, and
    # every helper below needs the same guard.
    return ,$out.ToArray()
}

# Every VBA string literal in a line, "" un-doubled, in order.
function Get-StringLiterals([string]$text) {
    $out = New-Object System.Collections.Generic.List[string]
    $i = 0
    while ($i -lt $text.Length) {
        if ($text[$i] -ne '"') { $i++; continue }
        $j = $i + 1
        $lit = ''
        while ($j -lt $text.Length) {
            if ($text[$j] -eq '"') {
                if ($j + 1 -lt $text.Length -and $text[$j + 1] -eq '"') { $lit += '"'; $j += 2; continue }
                break
            }
            $lit += $text[$j]
            $j++
        }
        $out.Add($lit)
        $i = $j + 1
    }
    return ,$out.ToArray()
}

# The arguments after a RaiseMsg id, split on TOP-LEVEL commas only, so
# a value like DefinedNamesSentence(relations) stays one argument.
function Split-TopLevelArgs([string]$text) {
    $out = New-Object System.Collections.Generic.List[string]
    $depth = 0
    $cur = ''
    $i = 0
    while ($i -lt $text.Length) {
        $ch = $text[$i]
        if ($ch -eq '"') {
            $cur += $ch
            $i++
            while ($i -lt $text.Length) {
                $cur += $text[$i]
                if ($text[$i] -eq '"') {
                    if ($i + 1 -lt $text.Length -and $text[$i + 1] -eq '"') { $cur += $text[$i + 1]; $i += 2; continue }
                    $i++
                    break
                }
                $i++
            }
            continue
        }
        if ($ch -eq '(') { $depth++ }
        elseif ($ch -eq ')') { $depth-- }
        if ($ch -eq ',' -and $depth -le 0) { $out.Add($cur.Trim()); $cur = ''; $i++; continue }
        $cur += $ch
        $i++
    }
    if ($cur.Trim() -ne '') { $out.Add($cur.Trim()) }
    return ,$out.ToArray()
}

function Get-TemplateSlots([string]$template) {
    $out = New-Object System.Collections.Generic.HashSet[string]
    foreach ($mt in [regex]::Matches($template, '\{([A-Za-z][A-Za-z0-9_]*)\}')) {
        [void]$out.Add($mt.Groups[1].Value)
    }
    return ,$out
}

# --- the two catalogues ----------------------------------------------
$catalogues = @{}
$builders = @(
    @{ File = 'VLA_Messages.bas'; Add = 'AddMsg';        Raiser = 'RaiseMsg' },
    @{ File = 'VLA_Runtime.bas';  Add = 'RuntimeAddMsg'; Raiser = 'RaiseRuntimeMsg' }
)
foreach ($b in $builders) {
    $path = Join-Path $srcDir $b.File
    if (-not (Test-Path -LiteralPath $path)) { Write-Error "Missing $path" }
    $entries = @{}
    foreach ($row in (Get-LogicalLines $path)) {
        $mt = [regex]::Match($row.Text, ('^\s*' + [regex]::Escape($b.Add) + '\s+m\s*,\s*'))
        if (-not $mt.Success) { continue }
        $lits = Get-StringLiterals $row.Text.Substring($mt.Length)
        if ($lits.Count -lt 2) {
            $failures.Add("$($b.File):$($row.No): could not read an id and a template from this $($b.Add) line")
            continue
        }
        $id = $lits[0]
        # The template is EVERY literal from the third onward, joined:
        # literal 0 is the id, literal 1 is the Err.Source tag, and a
        # template is often built by concatenation -
        #   "{loc}: test FAILED" & vbCrLf & "  sentence: {sentence}" & ...
        # - so reading only the last literal saw only its final slot and
        # reported every earlier one as "supplied but not named". That was
        # this script's own first bug, and a clean tree reporting 14
        # problems is what found it.
        $template = ''
        for ($li = 2; $li -lt $lits.Count; $li++) { $template += $lits[$li] }
        if ($entries.ContainsKey($id)) {
            $failures.Add("$($b.File):$($row.No): duplicate catalogue id '$id'")
            continue
        }
        $entries[$id] = [pscustomobject]@{ Slots = (Get-TemplateSlots $template); Raised = 0; Line = $row.No }
    }
    $catalogues[$b.Raiser] = @{ Name = $b.File; Entries = $entries }
    Write-Output ("  {0,-20} {1} catalogue entries" -f $b.File, $entries.Count)
}

# --- every raise site -------------------------------------------------
$siteCount = 0
$sourceFiles = @(Get-ChildItem -LiteralPath $srcDir -File | Where-Object { $_.Extension -in @('.bas', '.cls', '.frm') } | Sort-Object Name)
foreach ($f in $sourceFiles) {
    foreach ($row in (Get-LogicalLines $f.FullName)) {
        $text = $row.Text
        if ($text -match "^\s*'") { continue }
        foreach ($raiser in @('RaiseRuntimeMsg', 'RaiseMsg')) {
            $mt = [regex]::Match($text, ('(?<![A-Za-z0-9_])' + [regex]::Escape($raiser) + '\s+(?=")'))
            if (-not $mt.Success) { continue }
            $rest = $text.Substring($mt.Index + $mt.Length)
            # NOT $args: that is a PowerShell automatic variable.
            $argList = Split-TopLevelArgs $rest
            if ($argList.Count -lt 1) { continue }
            $idLits = Get-StringLiterals $argList[0]
            if ($idLits.Count -lt 1) {
                $failures.Add("$($f.Name):$($row.No): $raiser's id is not a string literal, so it cannot be checked")
                continue
            }
            $siteCount++
            $id = $idLits[0]
            $cat = $catalogues[$raiser]
            if (-not $cat.Entries.ContainsKey($id)) {
                $failures.Add("$($f.Name):$($row.No): $raiser names '$id', which $($cat.Name) does not hold - this would raise ""unknown message id"" instead of a refusal")
                continue
            }
            $entry = $cat.Entries[$id]
            $entry.Raised = $entry.Raised + 1
            $supplied = New-Object System.Collections.Generic.HashSet[string]
            for ($a = 1; $a -lt $argList.Count; $a += 2) {
                $nameLits = Get-StringLiterals $argList[$a]
                if ($nameLits.Count -lt 1) {
                    $failures.Add("$($f.Name):$($row.No): $raiser '$id' has a slot NAME that is not a string literal ($($argList[$a]))")
                    continue
                }
                [void]$supplied.Add($nameLits[0])
            }
            foreach ($want in $entry.Slots) {
                if (-not $supplied.Contains($want)) {
                    $failures.Add("$($f.Name):$($row.No): $raiser '$id' does not supply '{$want}', which its template names - this raises ""no value supplied for slot"" instead of the refusal")
                }
            }
            foreach ($got in $supplied) {
                if (-not $entry.Slots.Contains($got)) {
                    $failures.Add("$($f.Name):$($row.No): $raiser '$id' supplies '$got', which its template does not name - silently ignored at runtime, so a real slot may be going unfilled")
                }
            }
            if ($List) { Write-Output ("  site  {0}:{1}  {2}  [{3}]" -f $f.Name, $row.No, $id, (@($supplied) -join ' ')) }
            break
        }
    }
}

Write-Output ("  {0,-20} {1} raise sites parsed" -f 'src\*', $siteCount)

$never = New-Object System.Collections.Generic.List[string]
foreach ($raiser in $catalogues.Keys) {
    foreach ($id in $catalogues[$raiser].Entries.Keys) {
        if ($catalogues[$raiser].Entries[$id].Raised -eq 0) { $never.Add($id) }
    }
}
if ($never.Count -gt 0) {
    Write-Output ''
    Write-Output ("  note  {0} catalogue id(s) no site raises - a reserved id, or one raised from a path this scan cannot see" -f $never.Count)
    if ($List) { foreach ($id in ($never | Sort-Object)) { Write-Output "          $id" } }
}

Write-Output ''
if ($failures.Count -eq 0) {
    Write-Output '=== CHECK: clean - every refusal site names a real id and supplies exactly its template''s slots ==='
    exit 0
} else {
    foreach ($f in $failures) { Write-Output "FAIL: $f" }
    Write-Output ''
    Write-Output "=== CHECK: FAILED - $($failures.Count) problem(s) ==="
    Write-Output "A refusal that cannot render its own template is worse than no refusal: the user is shown Frazaro's internals instead of what is wrong with their spreadsheet. Add the missing slot value at the call site, or drop the slot from the template if nothing can supply it."
    exit 1
}
