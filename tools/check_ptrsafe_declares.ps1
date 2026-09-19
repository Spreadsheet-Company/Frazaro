<#
check_ptrsafe_declares.ps1 - EN.8's mechanical check: 64-bit declaration
discipline, for every Declare in the repository.

WHY. A Declare - a call into a Windows DLL - is the one VBA statement whose
correctness depends on which Office compiles it, and each way of getting it
wrong fails somewhere its author's own Excel cannot see:

  * 64-bit Office - the default install of Office 2019, 2021 and Microsoft
    365, and every Mac Office since 2016 - refuses to compile a Declare
    without PtrSafe ("The code in this project must be updated for use on
    64-bit systems"), and that refusal stops the whole module, not one call.
  * Office 2007 and older (VBA6, and Mac Office 2011) knows neither PtrSafe
    nor LongPtr nor LongLong, so the corrected line fails there instead.
    That is why a Declare meant for both lives in the two branches of one
    #If VBA7 Then ... #Else ... #End If.
  * No modern Office compiles that #Else branch, so no compile, self-test or
    live pass ever reads it. It can drift from its twin for years, unseen.
  * The quiet one: a handle or pointer typed Long compiles everywhere, works
    on 32-bit Office, and on 64-bit Office loses the top half of the value -
    on some runs, depending on the address Windows hands back.

A compile or a live pass answers for the one Office it ran on. This scan
answers for every host in $hosts below, from the text, on every push (CI)
and before every release (tools/release.ps1 runs every tools/check_*.ps1).
Today the entire bitness-sensitive surface is frmCLI.frm's four user32
calls, declared once per branch (docs/SUBSTRATE.md's census); this script
holds any Declare that arrives later to the same discipline, and pins that
Frazaro's generators write none.

WHAT IS CHECKED. Every Declare statement (continuation lines joined) is
placed in its #If/#ElseIf/#Else nesting, and each enclosing condition is
evaluated for every host in $hosts, giving the set of hosts that compile
that line. Then, each rule a tag in the output:

  [STRUCTURE] #If/#ElseIf/#Else/#End If out of order or never closed - the
              nesting every rule below reads cannot be trusted.
  [CONDITION] a Declare under a condition this script cannot evaluate: a
              name other than VBA7, VBA6, Win64, Win32, Win16 and Mac, or an
              operator other than Not, And, Or, Xor, = and <>. A constant set
              in the project's properties is invisible to a text scan, so the
              script refuses to guess which hosts compile the line.
  [PARSE]     a Declare line this script cannot read.
  [SPLIT]     a Declare compiled both by a 64-bit Office and by a pre-VBA7
              Office - one outside any #If, or in the #Else of #If Win64,
              which 64-bit Mac Office compiles as well. No single line is
              right on both.
  [PTRSAFE]   a Declare 64-bit Office compiles, without PtrSafe.
  [PRE-VBA7]  a Declare pre-VBA7 Office compiles, with PtrSafe, LongPtr or
              LongLong.
  [LONGLONG]  LongLong in a Declare that 32-bit Office 2010+ compiles.
  [UNTYPED]   a parameter or return with no type - an implicit Variant.
  [WIDTH]     a Long, Integer, Byte or Boolean parameter or return, in a
              Declare 64-bit Office compiles, that $reviewedWidths does not
              list. See WIDTHS.
  [DUPLICATE] two Declares of one name that one host compiles together
              ("Ambiguous name detected").
  [COVERAGE]  a Declare name some host has no Declare for, so a call to it
              cannot compile there - unless $partialHosts lists it.
  [PARITY]    two Declares of one name that differ in anything but a Long
              that became LongPtr or LongLong: the branch nobody compiles has
              drifted from the branch everybody does.
  [GENERATED] a Declare in text Frazaro generates code from or as: a string
              literal in VBA source, VLA text (.vla - the phrasebooks, the
              prelude, translated programs), or generated VBA (.vba). Neither the emitter nor any phrasebook
              writes one today - the only "declare" outside comments is a
              word in VLA_SentenceEngine's reserved-word list - and a
              template that did would put a DLL call, reviewed by none of the
              rules above, into every workbook built from it.

WIDTHS - a reviewed list, not a naming rule. Whether a slot holds a pointer
is a fact about the Windows API, not about the VBA text. A rule keyed on
names (hwnd, h*, lp*) fails open: ByVal window As Long passes it, and so does
every handle-returning function, because a return has no name. So in a
Declare that 64-bit Office compiles, every slot narrower than a pointer is a
reviewed fact - listed in $reviewedWidths with the API's own type as the
reason - or it fails. LongPtr, String, Any, the floating types and user types
need no entry.

SCANNED SET: every .bas, .cls, .frm, .vba and .vla file under the repository
root except .git, tracked or not - shipped modules, dev-only modules, tools/
diagnostics and generated goldens alike. Wider than check_no_network.ps1's
shipped set on purpose: a module that will not compile on 64-bit Office stops
the dev workbook, the self-tests and VlaBuildAddin on the owner's own Excel,
not only a user's. Each site is labelled shipped, dev-only or tool from
VLA_Build.bas's mods array, so a failure says who it would reach.

HOSTS: $hosts below. The Windows rows follow Microsoft's documented compiler
constants (Win32 is True in 64-bit Office too; VBA6 is True on every VBA7
host). What Mac Office reports for Win32 and Win64 is not assumed: each Mac
row appears once per possible value, so a condition that leans on either is
judged both ways.

WHAT IS NOT CHECKED
  * That a reviewed 32-bit slot really is one. That is the review
    $reviewedWidths records; each reason is its evidence.
  * Pointer members inside a user type passed ByRef - the layout is not
    visible from the Declare.
  * Which DLLs may be called at all: check_no_network.ps1 (SD-13) holds the
    reviewed list of Declares in shipped code.
  * A Declare assembled from pieces ("Decl" & "are") in generated code, a
    phrasebook string that spans lines, or a `raw` rule in a phrasebook a
    user loads - SEC.2-consent-gated opaque VBA (THREAT_MODEL.md 1.4).
  * DefLng-style defaults: an untyped slot is reported [UNTYPED] even where a
    Def statement would type it.

BASELINES: $reviewedWidths and $partialHosts, hand-maintained below in the
reviewable shape of check_no_network.ps1's lists. Adding an entry is a
reviewed act: give the reason in the entry, and say why in the commit.

NOT wired into VlaSelfTest, same reasoning F.12/F.14/AS.8/SEC.13 gave for the
other tools/check_*.ps1 scans.

Usage:  powershell -File tools\check_ptrsafe_declares.ps1
Exit code: 0 if every rule holds; 1 otherwise.
#>

$ErrorActionPreference = 'Stop'

$repoRoot = [IO.Path]::GetFullPath((Split-Path -Parent $PSScriptRoot))
$srcDir   = Join-Path $repoRoot 'src'

# --- Shipped set, for labels: VLA_Build.bas's own mods array ---
$buildFile = Join-Path $srcDir 'VLA_Build.bas'
$buildText = [IO.File]::ReadAllText($buildFile)
$modsMatch = [regex]::Match($buildText, 'mods\s*=\s*Array\(([^)]*)\)')
if (-not $modsMatch.Success) {
    Write-Error "No 'mods = Array(...)' line found in $buildFile - VLA_Build.bas's shipped-module list shape may have changed; update this script's `$modsMatch pattern."
}
$shipped = @{}
foreach ($mm in [regex]::Matches($modsMatch.Groups[1].Value, '"([^"]*)"')) { $shipped[$mm.Groups[1].Value] = $true }
if ($shipped.Count -eq 0) {
    Write-Error "Parsed zero module names out of the mods array in $buildFile - regex mismatch, not an empty ship list."
}

# --- Hosts: every Office a Declare line can meet ---
# Is64 decides whether PtrSafe is required; VBA7 decides whether PtrSafe,
# LongPtr and LongLong exist at all.
$hosts = [System.Collections.Generic.List[object]]::new()
function Add-HostRow([string]$name, [bool]$vba7, [bool]$win64, [bool]$win32, [bool]$mac, [bool]$is64) {
    $hosts.Add([pscustomobject]@{ Name = $name; VBA7 = $vba7; Win64 = $win64; Win32 = $win32; Mac = $mac; Is64 = $is64 })
}
Add-HostRow 'Windows Office 2007 or older (VBA6)' $false $false $true $false $false
Add-HostRow 'Windows Office 2010+ 32-bit'         $true  $false $true $false $false
Add-HostRow 'Windows Office 2010+ 64-bit'         $true  $true  $true $false $true
foreach ($w32 in @($false, $true)) {
    Add-HostRow "Mac Office 2011 (VBA6), Win32=$w32" $false $false $w32 $true $false
    foreach ($w64 in @($false, $true)) {
        Add-HostRow "Mac Office 2016+ (64-bit), Win64=$w64 Win32=$w32" $true $w64 $w32 $true $true
    }
}

$allMask  = (1 -shl $hosts.Count) - 1
$mask64   = 0
$maskOld  = 0
$mask32v7 = 0
for ($k = 0; $k -lt $hosts.Count; $k++) {
    $bit = 1 -shl $k
    if ($hosts[$k].Is64) { $mask64 = $mask64 -bor $bit }
    if (-not $hosts[$k].VBA7) { $maskOld = $maskOld -bor $bit }
    elseif (-not $hosts[$k].Is64) { $mask32v7 = $mask32v7 -bor $bit }
}
$maskVba7 = $allMask -band (-bnot $maskOld)

# --- Baseline 1: reviewed 32-bit slots in Declares 64-bit Office compiles ---
# "<Module>::<Declare>::<parameter, or return>" = the Windows API's own type,
# the evidence that the slot is not a pointer. 2026-09-14, EN.8: frmCLI's
# #If VBA7 branch. (Its #Else branch needs no entry: only pre-VBA7 Office
# compiles it, and there Long is as wide as a pointer.)
$reviewedWidths = [ordered]@{
    'frmCLI::GetWindowLong::nIndex'    = 'int nIndex (GWL_STYLE)'
    'frmCLI::GetWindowLong::return'    = 'LONG - whole for GWL_STYLE, the only index frmCLI reads; a GWLP_ pointer index would need GetWindowLongPtrA, As LongPtr'
    'frmCLI::SetWindowLong::nIndex'    = 'int nIndex (GWL_STYLE)'
    'frmCLI::SetWindowLong::dwNewLong' = 'LONG - the new style bits; the same GWL_STYLE-only caveat'
    'frmCLI::SetWindowLong::return'    = 'LONG - the previous style bits; the same caveat'
    'frmCLI::SetWindowPos::x'          = 'int X'
    'frmCLI::SetWindowPos::y'          = 'int Y'
    'frmCLI::SetWindowPos::cx'         = 'int cx'
    'frmCLI::SetWindowPos::cy'         = 'int cy'
    'frmCLI::SetWindowPos::wFlags'     = 'UINT uFlags'
    'frmCLI::SetWindowPos::return'     = 'BOOL - a 32-bit int'
}

# --- Baseline 2: Declares deliberately absent on some hosts ---
# "<Module>::<Declare>" = how every call to it is guarded the same way.
# Empty: frmCLI declares all four calls for every host.
$partialHosts = [ordered]@{
}

# VBA names fold case; so do these lookups (PowerShell hashtables do).
$widthKeys = @{}
foreach ($wk in $reviewedWidths.Keys) { $widthKeys[$wk] = $true }
$partialKeys = @{}
foreach ($pk0 in $partialHosts.Keys) { $partialKeys[$pk0] = $true }

$reIf      = [regex]::new('^\s*#\s*If\s+(?<c>.+?)\s+Then\s*$', 'IgnoreCase')
$reElseIf  = [regex]::new('^\s*#\s*ElseIf\s+(?<c>.+?)\s+Then\s*$', 'IgnoreCase')
$reElse    = [regex]::new('^\s*#\s*Else\s*$', 'IgnoreCase')
$reEndIf   = [regex]::new('^\s*#\s*End\s*If\s*$', 'IgnoreCase')
$reConst   = [regex]::new('^\s*#\s*Const\b', 'IgnoreCase')
$reToken   = [regex]::new('\G\s*(&H[0-9A-F]+|\d+|[A-Z_][A-Z0-9_]*|<>|=|\(|\))', 'IgnoreCase')
$reDeclareAnywhere = [regex]::new('(?:^|:)\s*(?:(?:Public|Private)\s+)?Declare\s', 'IgnoreCase')
$reDeclareFull = [regex]::new('^\s*(?:(?:Public|Private)\s+)?Declare\s+(?<ptrsafe>PtrSafe\s+)?(?<kind>Function|Sub)\s+(?<name>[A-Z][A-Z0-9_]*)(?<ntc>[%&^!#@$])?\s+Lib\s+"(?<lib>[^"]*)"(?:\s+Alias\s+"(?<alias>[^"]*)")?\s*(?:\((?<params>[^()]*(?:\(\s*\)[^()]*)*)\))?(?:\s+As\s+(?<ret>[A-Z][A-Z0-9_]*)(?<retarr>\s*\(\s*\))?)?\s*$', 'IgnoreCase')
$reParam   = [regex]::new('^\s*(?<opt>Optional\s+)?(?:(?<mode>ByVal|ByRef)\s+)?(?<pa>ParamArray\s+)?(?<name>[A-Z][A-Z0-9_]*)(?<tc>[%&^!#@$])?\s*(?<arr>\(\s*\))?(?:\s+As\s+(?<type>[A-Z][A-Z0-9_]*))?(?:\s*=\s*.+)?\s*$', 'IgnoreCase')
$reGenerated = [regex]::new('\bDeclare\s+(?:PtrSafe\s+)?(?:Function|Sub)\s+[A-Z]', 'IgnoreCase')

$typeChars = @{ '%' = 'Integer'; '&' = 'Long'; '^' = 'LongLong'; '!' = 'Single'; '#' = 'Double'; '@' = 'Currency'; '$' = 'String' }

# --- #If conditions: tokens, then one evaluation per host ---
function Get-Tokens([string]$expr) {
    $list = [System.Collections.Generic.List[string]]::new()
    $text = $expr.TrimEnd()
    $at = 0
    while ($at -lt $text.Length) {
        $m = $reToken.Match($text, $at)
        if (-not $m.Success) { return $null }
        $list.Add($m.Groups[1].Value)
        $at = $m.Index + $m.Length
    }
    return ,$list
}

# VBA precedence: comparison, then Not, And, Or, Xor. True is -1, as in VBA.
function Read-Xor {
    $v = Read-Or
    while ($script:pos -lt $script:tok.Count -and $script:tok[$script:pos] -eq 'Xor') { $script:pos++; $v = $v -bxor (Read-Or) }
    return $v
}
function Read-Or {
    $v = Read-And
    while ($script:pos -lt $script:tok.Count -and $script:tok[$script:pos] -eq 'Or') { $script:pos++; $v = $v -bor (Read-And) }
    return $v
}
function Read-And {
    $v = Read-Not
    while ($script:pos -lt $script:tok.Count -and $script:tok[$script:pos] -eq 'And') { $script:pos++; $v = $v -band (Read-Not) }
    return $v
}
function Read-Not {
    if ($script:pos -lt $script:tok.Count -and $script:tok[$script:pos] -eq 'Not') { $script:pos++; return (-1 - (Read-Not)) }
    return (Read-Cmp)
}
function Read-Cmp {
    $v = Read-Prim
    while ($script:pos -lt $script:tok.Count -and ($script:tok[$script:pos] -eq '=' -or $script:tok[$script:pos] -eq '<>')) {
        $op = $script:tok[$script:pos]
        $script:pos++
        $r = Read-Prim
        $same = ($v -eq $r)
        if ($op -eq '<>') { $same = -not $same }
        if ($same) { $v = -1 } else { $v = 0 }
    }
    return $v
}
function Read-Prim {
    if ($script:pos -ge $script:tok.Count) { throw 'the condition ends early' }
    $t = $script:tok[$script:pos]
    $script:pos++
    if ($t -eq '(') {
        $v = Read-Xor
        if ($script:pos -ge $script:tok.Count -or $script:tok[$script:pos] -ne ')') { throw 'a ( is never closed' }
        $script:pos++
        return $v
    }
    if ([regex]::IsMatch($t, '^\d+$')) { return [int]$t }
    if ([regex]::IsMatch($t, '^&H[0-9A-F]+$', 'IgnoreCase')) { return [Convert]::ToInt32($t.Substring(2), 16) }
    $hr = $script:hostRow
    switch ($t) {
        'True'  { return -1 }
        'False' { return 0 }
        'VBA6'  { return -1 }
        'Win16' { return 0 }
        'VBA7'  { if ($hr.VBA7)  { return -1 } else { return 0 } }
        'Win64' { if ($hr.Win64) { return -1 } else { return 0 } }
        'Win32' { if ($hr.Win32) { return -1 } else { return 0 } }
        'Mac'   { if ($hr.Mac)   { return -1 } else { return 0 } }
    }
    throw "'$t' is not a constant this script models"
}

function Get-ConditionMask([string]$expr) {
    $t = Get-Tokens $expr
    if ($null -eq $t -or $t.Count -eq 0) { return @{ Mask = $null; Why = "cannot read the condition '$expr'" } }
    $mask = 0
    for ($k = 0; $k -lt $hosts.Count; $k++) {
        $script:tok = $t
        $script:pos = 0
        $script:hostRow = $hosts[$k]
        try {
            $v = Read-Xor
            if ($script:pos -ne $t.Count) { throw "'$($t[$script:pos])' is not expected there" }
        } catch {
            return @{ Mask = $null; Why = "in '$expr', $($_.Exception.Message)" }
        }
        if ($v -ne 0) { $mask = $mask -bor (1 -shl $k) }
    }
    return @{ Mask = $mask; Why = '' }
}

function New-Frame([int]$line) {
    return [pscustomobject]@{
        Line   = $line
        Conds  = [System.Collections.Generic.List[object]]::new()
        Whys   = [System.Collections.Generic.List[string]]::new()
        InElse = $false
    }
}

function Add-Cond($frame, [string]$expr, [int]$line) {
    $r = Get-ConditionMask $expr
    $frame.Conds.Add($r.Mask)
    $frame.Whys.Add("line ${line}: $($r.Why)")
}

# The hosts that compile a line inside $stack: each frame admits the hosts
# where its current branch's condition holds and every earlier one did not.
function Get-StackMask($stack) {
    $mask = $allMask
    foreach ($fr in $stack) {
        $excluded = 0
        $last = $fr.Conds.Count - 1
        $upto = $last - 1
        if ($fr.InElse) { $upto = $last }
        for ($k = 0; $k -le $upto; $k++) {
            if ($null -eq $fr.Conds[$k]) { return @{ Mask = $null; Why = $fr.Whys[$k] } }
            $excluded = $excluded -bor $fr.Conds[$k]
        }
        if ($fr.InElse) {
            $branch = $allMask -band (-bnot $excluded)
        } else {
            if ($null -eq $fr.Conds[$last]) { return @{ Mask = $null; Why = $fr.Whys[$last] } }
            $branch = $fr.Conds[$last] -band (-bnot $excluded)
        }
        $mask = $mask -band $branch
    }
    return @{ Mask = $mask; Why = '' }
}

function Get-MaskText([int]$mask) {
    if ($mask -eq $allMask)  { return 'every host' }
    if ($mask -eq $maskVba7) { return 'VBA7 hosts' }
    if ($mask -eq $maskOld)  { return 'pre-VBA7 hosts' }
    if ($mask -eq 0)         { return 'no host' }
    $names = [System.Collections.Generic.List[string]]::new()
    for ($k = 0; $k -lt $hosts.Count; $k++) {
        if (($mask -band (1 -shl $k)) -ne 0) { $names.Add($hosts[$k].Name) }
    }
    return ($names -join '; ')
}

# --- VBA text: a line's code (literals intact), the same with literal
# contents blanked, and the literals themselves; comments dropped ---
function Split-VbaCode([string]$line) {
    $lits = [System.Collections.Generic.List[string]]::new()
    $bare = [System.Text.StringBuilder]::new()
    $n = $line.Length
    $i = 0
    $end = $n
    while ($i -lt $n) {
        $c = $line[$i]
        if ($c -eq [char]'"') {
            $j = $i + 1
            $lit = [System.Text.StringBuilder]::new()
            while ($j -lt $n) {
                if ($line[$j] -eq [char]'"') {
                    if (($j + 1) -lt $n -and $line[$j + 1] -eq [char]'"') { [void]$lit.Append('"'); $j += 2; continue }
                    break
                }
                [void]$lit.Append($line[$j])
                $j++
            }
            $lits.Add($lit.ToString())
            [void]$bare.Append('""')
            $i = $j + 1
            continue
        }
        if ($c -eq [char]"'") { $end = $i; break }
        [void]$bare.Append($c)
        $i++
    }
    $code = $line.Substring(0, [Math]::Min($end, $n))
    $bareText = $bare.ToString()
    if ([regex]::IsMatch($bareText, '^\s*Rem(\s|$)', 'IgnoreCase')) { $code = ''; $bareText = ''; $lits.Clear() }
    return @{ Code = $code; Bare = $bareText; Literals = $lits }
}

# A phrasebook line without its ; comment. Strings are "..." with \" escapes.
function Get-VlaCodeText([string]$line) {
    $n = $line.Length
    $i = 0
    $inStr = $false
    while ($i -lt $n) {
        $c = $line[$i]
        if ($inStr) {
            if ($c -eq [char]'\') { $i += 2; continue }
            if ($c -eq [char]'"') { $inStr = $false }
        } elseif ($c -eq [char]'"') {
            $inStr = $true
        } elseif ($c -eq [char]';') {
            return $line.Substring(0, $i)
        }
        $i++
    }
    return $line
}

function Get-SlotType([string]$asType, [string]$typeChar) {
    if ($typeChar.Length -gt 0) { return $typeChars[$typeChar] }
    return $asType
}

function Get-WidthClass([string]$t) {
    switch ($t) {
        ''         { return 'untyped' }
        'LongPtr'  { return 'ptr' }
        'LongLong' { return 'longlong' }
        'Long'     { return 'fixed' }
        'Integer'  { return 'fixed' }
        'Byte'     { return 'fixed' }
        'Boolean'  { return 'fixed' }
    }
    return 'other'
}

function Read-DeclareSite([string]$code, [string]$rel, [int]$line, [string]$module, [string]$label) {
    $dm = $reDeclareFull.Match($code)
    if (-not $dm.Success) { return $null }
    $params = [System.Collections.Generic.List[object]]::new()
    $ptext = $dm.Groups['params'].Value
    if ($ptext.Trim().Length -gt 0) {
        foreach ($piece in $ptext.Split(',')) {
            $pm = $reParam.Match($piece)
            if (-not $pm.Success) { return $null }
            $mode = 'ByRef'
            if ($pm.Groups['mode'].Success) { $mode = $pm.Groups['mode'].Value }
            $ptype = Get-SlotType $pm.Groups['type'].Value $pm.Groups['tc'].Value
            $params.Add([pscustomobject]@{
                Name     = $pm.Groups['name'].Value
                Mode     = $mode
                Optional = $pm.Groups['opt'].Success
                ParamArr = $pm.Groups['pa'].Success
                IsArray  = $pm.Groups['arr'].Success
                Type     = $ptype
            })
        }
    }
    $kind = $dm.Groups['kind'].Value
    $ret = ''
    if ($kind -eq 'Function') { $ret = Get-SlotType $dm.Groups['ret'].Value $dm.Groups['ntc'].Value }
    return [pscustomobject]@{
        File     = $rel
        Line     = $line
        Module   = $module
        Label    = $label
        Name     = $dm.Groups['name'].Value
        Kind     = $kind
        PtrSafe  = $dm.Groups['ptrsafe'].Success
        Lib      = $dm.Groups['lib'].Value
        Alias    = $dm.Groups['alias'].Value
        Params   = $params
        Ret      = $ret
        RetArray = $dm.Groups['retarr'].Success
        Mask     = $null
        MaskWhy  = ''
    }
}

function Get-LibKey([string]$lib) {
    $l = $lib.Trim().ToLowerInvariant()
    if ($l.EndsWith('.dll')) { $l = $l.Substring(0, $l.Length - 4) }
    return $l
}

function Get-ParityType([string]$t) {
    $c = Get-WidthClass $t
    if ($c -eq 'ptr' -or $c -eq 'longlong') { return 'long' }
    return $t.ToLowerInvariant()
}

# '' when two Declares of one name agree up to Long <-> LongPtr/LongLong.
function Get-ParityDiff($x, $y) {
    if ($x.Kind -ne $y.Kind) { return "one is a $($x.Kind), the other a $($y.Kind)" }
    if ((Get-LibKey $x.Lib) -ne (Get-LibKey $y.Lib)) { return "Lib `"$($x.Lib)`" against `"$($y.Lib)`"" }
    if ($x.Alias -cne $y.Alias) { return "Alias `"$($x.Alias)`" against `"$($y.Alias)`"" }
    if ($x.Params.Count -ne $y.Params.Count) { return "$($x.Params.Count) parameter(s) against $($y.Params.Count)" }
    for ($k = 0; $k -lt $x.Params.Count; $k++) {
        $p = $x.Params[$k]
        $q = $y.Params[$k]
        if ($p.Name -ne $q.Name) { return "parameter $($k + 1) is '$($p.Name)' against '$($q.Name)'" }
        if ($p.Mode -ne $q.Mode) { return "'$($p.Name)' is $($p.Mode) against $($q.Mode)" }
        if ($p.Optional -ne $q.Optional -or $p.ParamArr -ne $q.ParamArr -or $p.IsArray -ne $q.IsArray) { return "'$($p.Name)' differs in Optional, ParamArray or ()" }
        if ((Get-ParityType $p.Type) -ne (Get-ParityType $q.Type)) { return "'$($p.Name)' is As $($p.Type) against As $($q.Type)" }
    }
    if ((Get-ParityType $x.Ret) -ne (Get-ParityType $y.Ret) -or $x.RetArray -ne $y.RetArray) { return "the return is As $($x.Ret) against As $($y.Ret)" }
    return ''
}

function Get-RelPath($f) {
    return ($f.FullName.Substring($repoRoot.Length).TrimStart('\', '/') -replace '\\', '/')
}

function Get-Label($f, [string]$rel) {
    $ext = $f.Extension.ToLowerInvariant()
    if ($ext -eq '.vla') { return 'vla' }
    if ($ext -eq '.vba') { return 'generated' }
    if ($rel -like 'src/*') {
        if ($shipped.ContainsKey([IO.Path]::GetFileNameWithoutExtension($f.Name))) { return 'shipped' }
        return 'dev-only'
    }
    if ($rel -like 'tools/*') { return 'tool' }
    return 'other'
}

# --- The scanned set: every VBA-bearing file under the root but .git ---
$exts = @('.bas', '.cls', '.frm', '.vba', '.vla')
$found = [System.Collections.Generic.List[object]]::new()
foreach ($top in @(Get-ChildItem -LiteralPath $repoRoot -Force)) {
    if ($top.PSIsContainer) {
        if ($top.Name -eq '.git') { continue }
        foreach ($f in @(Get-ChildItem -LiteralPath $top.FullName -Recurse -File -Force)) {
            if ($exts -contains $f.Extension.ToLowerInvariant()) { $found.Add($f) }
        }
    } elseif ($exts -contains $top.Extension.ToLowerInvariant()) {
        $found.Add($top)
    }
}
$files = @($found | Sort-Object { Get-RelPath $_ })

$failed = [System.Collections.Generic.List[string]]::new()
$sites  = [System.Collections.Generic.List[object]]::new()
$counts = [ordered]@{ 'shipped' = 0; 'dev-only' = 0; 'tool' = 0; 'generated' = 0; 'vla' = 0; 'other' = 0 }

foreach ($f in $files) {
    $rel   = Get-RelPath $f
    $label = Get-Label $f $rel
    $counts[$label] = $counts[$label] + 1
    $ext   = $f.Extension.ToLowerInvariant()
    $lines = [IO.File]::ReadAllText($f.FullName) -split "`r?`n"

    if ($ext -eq '.vla') {
        for ($i = 0; $i -lt $lines.Count; $i++) {
            if ($lines[$i].IndexOf('declare', [StringComparison]::OrdinalIgnoreCase) -lt 0) { continue }
            if ($reGenerated.IsMatch((Get-VlaCodeText $lines[$i]))) {
                $failed.Add("[GENERATED] ${rel}:$($i + 1) ($label) - VLA text (a phrasebook, the prelude or a translated program) spells a Declare; the VBA built from a line like it would carry a DLL call none of these rules has reviewed.")
            }
        }
        continue
    }

    $module = [IO.Path]::GetFileNameWithoutExtension($f.Name)
    $stack = [System.Collections.Generic.List[object]]::new()
    $i = 0
    while ($i -lt $lines.Count) {
        $startLine = $i + 1
        $logical = $lines[$i]
        # A line ending in " _" continues onto the next - comments included, as in VBA.
        while (($i + 1) -lt $lines.Count) {
            $t = $logical.TrimEnd()
            if ($t.Length -lt 2 -or $t[$t.Length - 1] -ne [char]'_' -or -not [char]::IsWhiteSpace($t[$t.Length - 2])) { break }
            $logical = $t.Substring(0, $t.Length - 1) + $lines[$i + 1]
            $i++
        }

        if ($logical.TrimStart().StartsWith('#')) {
            $d = (Split-VbaCode $logical).Code
            if ($reIf.IsMatch($d)) {
                $fr = New-Frame $startLine
                Add-Cond $fr $reIf.Match($d).Groups['c'].Value $startLine
                $stack.Add($fr)
            } elseif ($reElseIf.IsMatch($d)) {
                if ($stack.Count -eq 0) {
                    $failed.Add("[STRUCTURE] ${rel}:$startLine ($label) - #ElseIf with no #If open.")
                } elseif ($stack[$stack.Count - 1].InElse) {
                    $failed.Add("[STRUCTURE] ${rel}:$startLine ($label) - #ElseIf after the #Else of the #If on line $($stack[$stack.Count - 1].Line).")
                } else {
                    Add-Cond $stack[$stack.Count - 1] $reElseIf.Match($d).Groups['c'].Value $startLine
                }
            } elseif ($reElse.IsMatch($d)) {
                if ($stack.Count -eq 0) {
                    $failed.Add("[STRUCTURE] ${rel}:$startLine ($label) - #Else with no #If open.")
                } elseif ($stack[$stack.Count - 1].InElse) {
                    $failed.Add("[STRUCTURE] ${rel}:$startLine ($label) - a second #Else for the #If on line $($stack[$stack.Count - 1].Line).")
                } else {
                    $stack[$stack.Count - 1].InElse = $true
                }
            } elseif ($reEndIf.IsMatch($d)) {
                if ($stack.Count -eq 0) {
                    $failed.Add("[STRUCTURE] ${rel}:$startLine ($label) - #End If with no #If open.")
                } else {
                    $stack.RemoveAt($stack.Count - 1)
                }
            } elseif (-not $reConst.IsMatch($d)) {
                $failed.Add("[STRUCTURE] ${rel}:$startLine ($label) - a directive this script does not know: $($logical.Trim())")
            }
        } elseif ($logical.IndexOf('declare', [StringComparison]::OrdinalIgnoreCase) -ge 0) {
            $sp = Split-VbaCode $logical
            foreach ($lit in $sp.Literals) {
                if ($reGenerated.IsMatch($lit)) {
                    $failed.Add("[GENERATED] ${rel}:$startLine ($label) - a string literal spells a Declare; code generated from it would carry a DLL call none of these rules has reviewed.")
                }
            }
            if ($reDeclareAnywhere.IsMatch($sp.Bare)) {
                if ($ext -eq '.vba') {
                    $failed.Add("[GENERATED] ${rel}:$startLine ($label) - generated VBA declares a DLL call.")
                } else {
                    $site = Read-DeclareSite $sp.Code $rel $startLine $module $label
                    if ($null -eq $site) {
                        $failed.Add("[PARSE] ${rel}:$startLine ($label) - a Declare this script cannot read (one per line, Lib and Alias as string literals, each parameter a name with an optional As type): $($logical.Trim())")
                    } else {
                        $sm = Get-StackMask $stack
                        $site.Mask = $sm.Mask
                        $site.MaskWhy = $sm.Why
                        $sites.Add($site)
                    }
                }
            }
        }
        $i++
    }
    foreach ($fr in $stack) {
        $failed.Add("[STRUCTURE] ${rel}:$($fr.Line) ($label) - #If never closed by #End If.")
    }
}

# --- Per-line rules ---
$seenWidths = @{}
foreach ($s in $sites) {
    $where = "$($s.File):$($s.Line) $($s.Module)::$($s.Name) ($($s.Label))"
    if ($null -eq $s.Mask) {
        $failed.Add("[CONDITION] $where - cannot tell which Office compiles it: $($s.MaskWhy). Only VBA7, VBA6, Win64, Win32, Win16 and Mac, with Not, And, Or, Xor, = and <>, are modelled.")
        continue
    }
    $has64  = ($s.Mask -band $mask64) -ne 0
    $hasOld = ($s.Mask -band $maskOld) -ne 0
    $slots = [System.Collections.Generic.List[object]]::new()
    foreach ($p in $s.Params) { $slots.Add([pscustomobject]@{ Slot = $p.Name; Type = $p.Type }) }
    if ($s.Kind -eq 'Function') { $slots.Add([pscustomobject]@{ Slot = 'return'; Type = $s.Ret }) }
    $usesPtr = $false
    $usesLL  = $false
    foreach ($sl in $slots) {
        $wc = Get-WidthClass $sl.Type
        if ($wc -eq 'ptr') { $usesPtr = $true }
        if ($wc -eq 'longlong') { $usesLL = $true }
    }

    if ($has64 -and $hasOld) {
        $failed.Add("[SPLIT] $where - compiled both by 64-bit Office ($(Get-MaskText ($s.Mask -band $mask64))) and by pre-VBA7 Office ($(Get-MaskText ($s.Mask -band $maskOld))); no one line compiles on both. Declare it under #If VBA7 Then (PtrSafe, LongPtr for handles and pointers), #Else (no PtrSafe, Long), #End If.")
    } else {
        if ($has64 -and -not $s.PtrSafe) {
            $failed.Add("[PTRSAFE] $where - 64-bit Office compiles this line ($(Get-MaskText ($s.Mask -band $mask64))) and it has no PtrSafe, so the module will not compile there.")
        }
        if ($hasOld -and ($s.PtrSafe -or $usesPtr -or $usesLL)) {
            $failed.Add("[PRE-VBA7] $where - pre-VBA7 Office compiles this line ($(Get-MaskText ($s.Mask -band $maskOld))) and it uses PtrSafe, LongPtr or LongLong, which that VBA does not have.")
        }
    }
    if ($usesLL -and ($s.Mask -band $mask32v7) -ne 0) {
        $failed.Add("[LONGLONG] $where - LongLong exists only in 64-bit Office, and 32-bit Office 2010+ compiles this line too. Use LongPtr, or move the line under #If Win64.")
    }
    foreach ($sl in $slots) {
        $wc = Get-WidthClass $sl.Type
        if ($wc -eq 'untyped') {
            $failed.Add("[UNTYPED] $where - '$($sl.Slot)' has no type, so it is passed as a Variant; give it the API's type.")
        } elseif ($wc -eq 'fixed' -and $has64) {
            $key = "$($s.Module)::$($s.Name)::$($sl.Slot)"
            $seenWidths[$key] = $true
            if (-not $widthKeys.ContainsKey($key)) {
                $failed.Add("[WIDTH] $where - '$($sl.Slot)' is As $($sl.Type) in a line 64-bit Office compiles. A handle or pointer must be LongPtr there. If the API's own type is 32-bit (int, UINT, LONG, DWORD, BOOL), list '$key' in `$reviewedWidths with that type as the reason.")
            }
        }
    }
}

# --- Per-name rules: one module's Declares of one name, taken together ---
$groups = [ordered]@{}
foreach ($s in $sites) {
    if ($null -eq $s.Mask) { continue }
    $gk = "$($s.File)::$($s.Name)".ToLowerInvariant()
    if (-not $groups.Contains($gk)) { $groups[$gk] = [System.Collections.Generic.List[object]]::new() }
    $groups[$gk].Add($s)
}
$usedPartial = @{}
foreach ($gk in @($groups.Keys)) {
    $g = $groups[$gk]
    $first = $g[0]
    $where = "$($first.File) $($first.Module)::$($first.Name) ($($first.Label))"
    $union = 0
    for ($a = 0; $a -lt $g.Count; $a++) {
        $union = $union -bor $g[$a].Mask
        for ($b = $a + 1; $b -lt $g.Count; $b++) {
            $both = $g[$a].Mask -band $g[$b].Mask
            if ($both -ne 0) {
                $failed.Add("[DUPLICATE] $where - lines $($g[$a].Line) and $($g[$b].Line) are both compiled by $(Get-MaskText $both): 'Ambiguous name detected'.")
            }
            $diff = Get-ParityDiff $g[$a] $g[$b]
            if ($diff.Length -gt 0) {
                $failed.Add("[PARITY] $where - lines $($g[$a].Line) and $($g[$b].Line) differ beyond a Long that became LongPtr: $diff.")
            }
        }
    }
    if ($union -ne $allMask) {
        $pk = "$($first.Module)::$($first.Name)"
        if ($partialKeys.ContainsKey($pk)) {
            $usedPartial[$pk] = $true
        } else {
            $failed.Add("[COVERAGE] $where - no Declare of it for $(Get-MaskText ($allMask -band (-bnot $union))), so a call to it cannot compile there. Declare it in the other branch too; or, if every call to it is guarded the same way, list '$pk' in `$partialHosts with the guard as the reason.")
        }
    }
}

# --- Report ---
Write-Output '=== 64-BIT DECLARATION DISCIPLINE (EN.8) ==='
Write-Output ("Scanned {0} file(s): {1} shipped, {2} dev-only, {3} tool, {4} generated VBA, {5} VLA text, {6} other - each Declare judged for {7} Office hosts" -f `
    $files.Count, $counts['shipped'], $counts['dev-only'], $counts['tool'], $counts['generated'], $counts['vla'], $counts['other'], $hosts.Count)
Write-Output ''
foreach ($s in $sites) {
    $hostText = 'hosts unknown'
    if ($null -ne $s.Mask) { $hostText = Get-MaskText $s.Mask }
    $safeText = 'plain'
    if ($s.PtrSafe) { $safeText = 'PtrSafe' }
    Write-Output ("  {0,-30} {1,-20} {2,-9} {3}, {4}" -f "$($s.Module)::$($s.Name)", "$($s.File):$($s.Line)", $s.Label, $hostText, $safeText)
}
if ($sites.Count -eq 0) { Write-Output '  (no Declare statements)' }

# Stale baseline entries: reported, not failed - the list then names a slot
# that is gone, the safe direction to be wrong in.
$stale = [System.Collections.Generic.List[string]]::new()
foreach ($wk in $reviewedWidths.Keys) { if (-not $seenWidths.ContainsKey($wk)) { $stale.Add("`$reviewedWidths: $wk") } }
foreach ($pk0 in $partialHosts.Keys) { if (-not $usedPartial.ContainsKey($pk0)) { $stale.Add("`$partialHosts: $pk0") } }
if ($stale.Count -gt 0) {
    Write-Output ''
    Write-Output '=== BASELINE ENTRIES WITH NO LIVE SLOT (removed or renamed - drop them above) ==='
    $stale | ForEach-Object { Write-Output "  $_" }
}

Write-Output ''
if ($failed.Count -eq 0) {
    $nameCount = $groups.Count
    Write-Output "=== CHECK: clean - $($sites.Count) Declare line(s), $nameCount name(s): PtrSafe wherever 64-bit Office compiles one and plain wherever pre-VBA7 Office does, each name declared for every host with its branches in step, $($seenWidths.Count) reviewed 32-bit slot(s); no generated text spells a Declare ==="
} else {
    Write-Output "=== CHECK: $($failed.Count) problem(s) ==="
    $failed | ForEach-Object { Write-Output "  $_" }
    Write-Output 'A Declare is right only if it is right on every Office that compiles it, and any one Excel compiles one branch at most. Fix each line as its problem says. Where a 32-bit slot or a one-branch Declare is genuinely correct, list it in the baseline above with the API''s own type, or the guard, as the reason - and say why in the commit message.'
}
exit ([Math]::Min($failed.Count, 1))
