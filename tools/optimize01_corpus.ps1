# optimize01_corpus.ps1 - OPTIMIZE.0.1's answer keys, re-derived by enumeration.
#
# scripts/pareto_logic.txt section 17 holds the OPTIMIZE corpus: twenty
# manager-sized questions, each with a tiny fixture and an ANSWER KEY written
# by hand before the engine exists - whether any world satisfies every rule,
# how many do, and the optimal cost in the objective's order with how many
# worlds reach it. This script re-derives every key a second way, by walking
# every candidate world the entry's choice allows and checking each rule on
# it, and exits non-zero if any key disagrees - with the enumeration, or with
# the key line printed in the corpus.
#
# What "candidates" means: the product of every choice's own options (a pair
# of people for each shift, a table for each guest), before any rule. The
# search walks all of them; the toy alone prunes "never two in a row" as it
# goes (sound: a pruned branch has already broken a rule), since 10,000,000
# leaves is too many for PowerShell. Two entries are too large to enumerate
# and are PROVED instead - their proofs' numbers are recomputed here from the
# same closed-form fixture definitions tools/optimize0_lp.ps1 uses - and the
# reference roster's own key is clingo's (OPTIMIZE.0 Part D), of which only
# the hand floor is checked here.
#
# The encodings here are deliberately NOT OPTIMIZE's and not clingo's: people
# are bits, choices are integers, every rule is a line of PowerShell over one
# complete world. A key that agrees here, by hand, and (after the owner runs
# tools/optimize01_lp.ps1's exports) in clingo, agrees three ways.
#
# Host-independent: PowerShell 5.1, no Excel, no network, writes nothing.
#
#   powershell -NoProfile -ExecutionPolicy Bypass -File tools\optimize01_corpus.ps1
#   ... -Only optimize-seat-wedding   one entry
#   ... -Control                      plant a wrong key in every entry and
#                                     require every one to be caught (exit 0
#                                     only if all are), so a green run is
#                                     known to be able to go red

param(
    [string]$Only = '',
    [switch]$Control
)

Set-StrictMode -Version 2
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$corpusPath = Join-Path $repoRoot 'scripts\pareto_logic.txt'
$dot = [string][char]0x00B7   # the corpus's own separator; built, so this file stays ASCII

# --- helpers (Verb-O1Noun: short names are aliases in PowerShell) -----------

function Format-O1Count([double]$x) {
    [string]::Format([Globalization.CultureInfo]::InvariantCulture, '{0:N0}', $x)
}

function Get-O1Bits([int]$mask) {
    $count = 0
    while ($mask -ne 0) { $count += ($mask -band 1); $mask = $mask -shr 1 }
    $count
}

# Every k-subset of n people, as bitmasks in increasing order.
function Get-O1Subsets([int]$n, [int]$k) {
    $out = New-Object 'System.Collections.Generic.List[int]'
    for ($m = 0; $m -lt (1 -shl $n); $m++) { if ((Get-O1Bits $m) -eq $k) { $out.Add($m) } }
    , $out.ToArray()
}

# A list of domains, built without PowerShell unrolling a one-element array.
function New-O1Domains([int]$count, [int[]]$values) {
    $list = New-Object 'System.Collections.Generic.List[object]'
    for ($i = 0; $i -lt $count; $i++) { $list.Add([int[]]$values) }
    , $list
}

function Compare-O1Cost([int[]]$a, [object]$b) {
    if ($null -eq $b) { return -1 }
    $bb = [int[]]$b
    for ($i = 0; $i -lt $a.Count; $i++) {
        if ($a[$i] -lt $bb[$i]) { return -1 }
        if ($a[$i] -gt $bb[$i]) { return 1 }
    }
    0
}

# Walks every candidate world: one value per variable from its domain,
# depth first and iteratively. Prune (optional) sees a partial world up to
# $depth and may only return $false for one that has already broken a rule.
function Invoke-O1Search($domains, [scriptblock]$prune, [scriptblock]$valid, [scriptblock]$cost) {
    $doms = @($domains.ToArray())
    $n = $doms.Count
    $pos = New-Object int[] $n
    $v = New-Object int[] $n
    $res = @{ Worlds = 0; Best = $null; BestCount = 0; Candidates = [double]1 }
    foreach ($d in $doms) { $res.Candidates *= ([int[]]$d).Count }
    $depth = 0
    $pos[0] = -1
    while ($depth -ge 0) {
        $pos[$depth]++
        $dom = [int[]]$doms[$depth]
        if ($pos[$depth] -ge $dom.Count) { $depth--; continue }
        $v[$depth] = $dom[$pos[$depth]]
        if ($null -ne $prune) { if (-not (& $prune $v $depth)) { continue } }
        if ($depth -eq $n - 1) {
            if (& $valid $v) {
                $res.Worlds++
                if ($null -ne $cost) {
                    $c = [int[]]@(& $cost $v)
                    $cmp = Compare-O1Cost $c $res.Best
                    if ($cmp -lt 0) { $res.Best = $c; $res.BestCount = 1 }
                    elseif ($cmp -eq 0) { $res.BestCount++ }
                }
            }
            continue
        }
        $depth++
        $pos[$depth] = -1
    }
    $res
}

# How many of a world's values equal $who (a person, a table, a slot).
function Measure-O1Count([int[]]$v, [int[]]$idx, [int]$who) {
    $count = 0
    foreach ($i in $idx) { if ($v[$i] -eq $who) { $count++ } }
    $count
}

# The key line the corpus must print for an entry, built from its expected answer.
function Get-O1KeyLine($e) {
    $cand = (Format-O1Count $e.Candidates) + $(if ($e.Candidates -eq 1) { ' candidate' } else { ' candidates' })
    if (-not $e.Sat) { return "key      no world $dot 0 valid worlds of $cand" }
    $w = Format-O1Count $e.Worlds
    $worldWord = if ($e.Worlds -eq 1) { 'valid world' } else { 'valid worlds' }
    if ($null -eq $e.Opt) { return "key      satisfiable $dot $w $worldWord of $cand $dot no objective" }
    $o = ($e.Opt | ForEach-Object { Format-O1Count $_ }) -join ', '
    $reach = if ($e.OptCount -eq 1) { '1 world reaches it' } else { "$(Format-O1Count $e.OptCount) worlds reach it" }
    "key      satisfiable $dot $w $worldWord of $cand $dot optimum ($o), $reach"
}

# One entry's lines in section 17: from its id line to the next blank line.
function Get-O1CorpusBlock([string[]]$lines, [string]$id) {
    $start = -1
    for ($i = 0; $i -lt $lines.Count; $i++) { if ($lines[$i] -match ('^' + [regex]::Escape($id) + '\s')) { $start = $i; break } }
    if ($start -lt 0) { return $null }
    $sb = New-Object System.Text.StringBuilder
    for ($i = $start; $i -lt $lines.Count; $i++) {
        if ($i -gt $start -and $lines[$i].Trim() -eq '') { break }
        [void]$sb.AppendLine($lines[$i].TrimEnd())
    }
    $sb.ToString()
}

# --- the closed-form roster fixture (tools/optimize0_lp.ps1's definitions) ---

function Get-O1ShiftDay([int]$shiftNo) { [int][math]::Floor(($shiftNo - 1) / 3) + 1 }
function Get-O1ShiftSlot([int]$shiftNo) { ($shiftNo - 1) % 3 }
function Test-O1OnLeave([int]$person, [int]$dayNo) { (($person + $dayNo) % 7) -eq 0 }
function Get-O1TightNeed([int]$people) { $t = [int][math]::Floor($people / 5); if ($t -lt 1) { 1 } else { $t } }

# Kept assignments that fall on the person's own leave day: each forces at
# least one drop and one add, so twice this is a floor under "changes".
function Measure-O1KeptOnLeave([int]$people, [int]$weeks) {
    $cursor = 0
    $clash = 0
    for ($s = 1; $s -le 21 * $weeks; $s++) {
        $need = Get-O1TightNeed $people
        for ($j = 0; $j -lt $need; $j++) {
            $p = (($cursor + $j) % $people) + 1
            if (Test-O1OnLeave $p (Get-O1ShiftDay $s)) { $clash++ }
        }
        $cursor += $need
    }
    $clash
}

# --- the entries ---------------------------------------------------------------
# Each: Id; Domains, Prune, Valid, Cost for the search; Maximize when the
# objective is a maximum (Cost returns the negation); and the hand-derived
# key - Sat, Worlds, Opt (in the objective's order, $null for none), OptCount.

$entries = New-Object 'System.Collections.Generic.List[object]'

# Segment 1 - compliance assignments
# optimize-sod-close. People Ann 0, Bob 1 (senior), Cy 2, Di 3; tasks T1-T3;
# variables: T1 preparer, T1 reviewer, T2 preparer, ... (index 2t, 2t+1).
$entries.Add(@{
    Id = 'optimize-sod-close'
    Domains = (New-O1Domains 6 @(0, 1, 2, 3))
    Prune = $null
    Valid = { param([int[]]$v)
        for ($t = 0; $t -lt 3; $t++) {
            if ($v[2 * $t] -eq $v[2 * $t + 1]) { return $false }   # nobody reviews their own work
            if ($v[2 * $t + 1] -gt 1) { return $false }             # a reviewer is senior
        }
        for ($p = 0; $p -lt 4; $p++) { if ((Measure-O1Count $v @(0, 1, 2, 3, 4, 5) $p) -gt 2) { return $false } }
        $true }
    Cost = { param([int[]]$v) $c = 0; foreach ($i in 0, 2, 4) { if ($v[$i] -le 1) { $c++ } }; $c }
    Sat = $true; Worlds = 84; Opt = @(0); OptCount = 36
})

# optimize-sod-check: no choice. The one world is this month's assignment
# (T1 Ann/Bob, T2 Cy/Cy, T3 Bob/Ann); the key is that it breaks two rules.
$entries.Add(@{
    Id = 'optimize-sod-check'
    Domains = $null   # handled below: one world, rules listed rather than counted
    Fixed = @(0, 1, 2, 2, 1, 0)
    Sat = $false; Worlds = 0; Opt = $null; OptCount = 0
    Violations = @('T2: Cy reviews their own work', 'T2: the reviewer, Cy, is not senior')
})

# optimize-sod-short: four tasks, three people (Ann 0, Bob 1, Cy 2), any
# reviewer; eight roles against room for six.
$entries.Add(@{
    Id = 'optimize-sod-short'
    Domains = (New-O1Domains 8 @(0, 1, 2))
    Prune = $null
    Valid = { param([int[]]$v)
        for ($t = 0; $t -lt 4; $t++) { if ($v[2 * $t] -eq $v[2 * $t + 1]) { return $false } }
        for ($p = 0; $p -lt 3; $p++) { if ((Measure-O1Count $v @(0, 1, 2, 3, 4, 5, 6, 7) $p) -gt 2) { return $false } }
        $true }
    Cost = $null
    Sat = $false; Worlds = 0; Opt = $null; OptCount = 0
})

# optimize-audit-independence. Staff Ann 0, Bob 1, Cy 2, Di 3, Ed 4 (bits);
# engagements Acme, Beta, Cafe each take a pair.
$pairs5 = Get-O1Subsets 5 2
$entries.Add(@{
    Id = 'optimize-audit-independence'
    Domains = (New-O1Domains 3 $pairs5)
    Prune = $null
    Valid = { param([int[]]$v)
        if (($v[0] -band ((1 -shl 0) -bor (1 -shl 2))) -ne 0) { return $false }   # Ann, Cy hold Acme
        if (($v[1] -band (1 -shl 1)) -ne 0) { return $false }                        # Bob is related to Beta
        if (($v[2] -band (1 -shl 3)) -ne 0) { return $false }                        # Di holds Cafe
        for ($p = 0; $p -lt 5; $p++) {
            $n = 0; foreach ($m in $v) { if (($m -band (1 -shl $p)) -ne 0) { $n++ } }
            if ($n -gt 2) { return $false } }
        $true }
    Cost = { param([int[]]$v) $c = 0; foreach ($m in $v) { if (($m -band (1 -shl 4)) -ne 0) { $c++ } }; $c }
    Sat = $true; Worlds = 90; Opt = @(0); OptCount = 9
})

# optimize-duty-rotation. Variable 3m + p is the duty of person p (Ann 0,
# Bob 1, Cy 2) in month m (M1-M3); duties Cash 0, Payables 1, Payroll 2.
# Last month was Ann Cash, Bob Payables, Cy Payroll.
$entries.Add(@{
    Id = 'optimize-duty-rotation'
    Domains = (New-O1Domains 9 @(0, 1, 2))
    Prune = $null
    Valid = { param([int[]]$v)
        for ($m = 0; $m -lt 3; $m++) {   # every duty has exactly one person each month
            if ($v[3 * $m] -eq $v[3 * $m + 1] -or $v[3 * $m] -eq $v[3 * $m + 2] -or $v[3 * $m + 1] -eq $v[3 * $m + 2]) { return $false } }
        for ($p = 0; $p -lt 3; $p++) {
            if ($v[$p] -eq $p) { return $false }   # not last month's duty in M1
            if ($v[$p] -eq $v[3 + $p] -or $v[$p] -eq $v[6 + $p] -or $v[3 + $p] -eq $v[6 + $p]) { return $false }   # not twice in the quarter
        }
        $true }
    Cost = $null
    Sat = $true; Worlds = 4; Opt = $null; OptCount = 0
})

# Segment 2 - shift rosters
# optimize-toy: 5 people, 7 shifts, a pair each, never two in a row.
$pairs5b = Get-O1Subsets 5 2
$entries.Add(@{
    Id = 'optimize-toy'
    Domains = (New-O1Domains 7 $pairs5b)
    Prune = { param([int[]]$v, [int]$depth) ($depth -eq 0) -or (($v[$depth] -band $v[$depth - 1]) -eq 0) }
    Valid = { param([int[]]$v) for ($s = 1; $s -lt 7; $s++) { if (($v[$s] -band $v[$s - 1]) -ne 0) { return $false } }; $true }
    Cost = $null
    Sat = $true; Worlds = 7290; Opt = $null; OptCount = 0
})

# optimize-roster-apart. Brianna 0, Tyler 1, Uma 2, Vic 3; Fri, Sat, Sun.
$pairs4 = Get-O1Subsets 4 2
$entries.Add(@{
    Id = 'optimize-roster-apart'
    Domains = (New-O1Domains 3 $pairs4)
    Prune = $null
    Valid = { param([int[]]$v)
        foreach ($m in $v) { if ($m -eq 3) { return $false } }   # 3 = Brianna and Tyler together
        for ($p = 0; $p -lt 4; $p++) {
            $n = 0; foreach ($m in $v) { if (($m -band (1 -shl $p)) -ne 0) { $n++ } }
            if ($n -gt 2) { return $false } }
        $true }
    Cost = { param([int[]]$v) if (($v[1] -band 1) -ne 0) { 1 } else { 0 } }   # Brianna on Saturday
    Sat = $true; Worlds = 60; Opt = @(0); OptCount = 34
})

# optimize-roster-kept. Ann 0, Bob 1, Cy 2; Mon-Thu, one each; kept A B A B.
$entries.Add(@{
    Id = 'optimize-roster-kept'
    Domains = (New-O1Domains 4 @(0, 1, 2))
    Prune = $null
    Valid = { param([int[]]$v)
        if ($v[3] -eq 1) { return $false }   # Bob is on leave Thursday
        for ($d = 1; $d -lt 4; $d++) { if ($v[$d] -eq $v[$d - 1]) { return $false } }
        for ($p = 0; $p -lt 3; $p++) { if ((Measure-O1Count $v @(0, 1, 2, 3) $p) -gt 2) { return $false } }
        $true }
    Cost = { param([int[]]$v) $kept = @(0, 1, 0, 1); $c = 0; for ($d = 0; $d -lt 4; $d++) { if ($v[$d] -ne $kept[$d]) { $c += 2 } }; $c }
    Sat = $true; Worlds = 16; Opt = @(2); OptCount = 1
})

# Segment 3 - event seating
# optimize-seat-dinner. Ann 0, Bob 1, Cy 2, Di 3, Ed 4, Fay 5; tables 0, 1.
$entries.Add(@{
    Id = 'optimize-seat-dinner'
    Domains = (New-O1Domains 6 @(0, 1))
    Prune = $null
    Valid = { param([int[]]$v)
        if ((Measure-O1Count $v @(0, 1, 2, 3, 4, 5) 0) -ne 3) { return $false }   # three at each table
        if ($v[0] -ne 0) { return $false }          # Ann, the host, at the head table
        if ($v[0] -ne $v[1]) { return $false }      # Ann and Bob together
        if ($v[2] -eq $v[3]) { return $false }      # Cy and Di apart
        $true }
    Cost = { param([int[]]$v) $c = 0; if ($v[4] -ne $v[5]) { $c++ }; if ($v[2] -ne $v[4]) { $c++ }; $c }
    Sat = $true; Worlds = 2; Opt = @(0); OptCount = 1
})

# optimize-seat-wedding. Ann 0, Bob 1, Cy 2, Di 3, Ed 4, Fay 5, Gus 6, Hal 7,
# Ivy 8; tables 0 (the head table), 1, 2, three seats each.
$entries.Add(@{
    Id = 'optimize-seat-wedding'
    Domains = (New-O1Domains 9 @(0, 1, 2))
    Prune = { param([int[]]$v, [int]$depth)
        $idx = 0..$depth
        for ($t = 0; $t -lt 3; $t++) { if ((Measure-O1Count $v $idx $t) -gt 3) { return $false } }
        $true }
    Valid = { param([int[]]$v)
        for ($t = 0; $t -lt 3; $t++) { if ((Measure-O1Count $v @(0, 1, 2, 3, 4, 5, 6, 7, 8) $t) -ne 3) { return $false } }
        if ($v[0] -ne $v[1]) { return $false }   # Ann and Bob together
        if ($v[2] -ne $v[3]) { return $false }   # Cy and Di together
        if ($v[6] -ne 0) { return $false }       # Gus at the head table
        if ($v[4] -eq $v[5]) { return $false }   # Ed and Fay apart
        if ($v[0] -eq $v[4]) { return $false }   # Ann and Ed apart
        $true }
    Cost = { param([int[]]$v) $c = 0; if ($v[6] -ne $v[4]) { $c += 2 }; if ($v[7] -ne $v[8]) { $c += 1 }; $c }
    Sat = $true; Worlds = 20; Opt = @(1); OptCount = 8
})

# optimize-seat-overflow: seven guests, two tables of at most three.
$entries.Add(@{
    Id = 'optimize-seat-overflow'
    Domains = (New-O1Domains 7 @(0, 1))
    Prune = $null
    Valid = { param([int[]]$v)
        for ($t = 0; $t -lt 2; $t++) { if ((Measure-O1Count $v @(0, 1, 2, 3, 4, 5, 6) $t) -gt 3) { return $false } }
        $true }
    Cost = $null
    Sat = $false; Worlds = 0; Opt = $null; OptCount = 0
})

# Segment 4 - configuration and quoting
# optimize-config-laptop. laptop x 0 (900), y 1 (700); dock d1 0 (150),
# d2 1 (120); monitor m4 0 (300), m5 1 (250).
$entries.Add(@{
    Id = 'optimize-config-laptop'
    Domains = (New-O1Domains 3 @(0, 1))
    Prune = $null
    Valid = { param([int[]]$v)
        $fitsLD = @('0-0', '1-0', '1-1')   # x-d1, y-d1, y-d2
        $fitsDM = @('0-0', '0-1', '1-1')   # d1-m4, d1-m5, d2-m5
        if ($fitsLD -notcontains "$($v[0])-$($v[1])") { return $false }
        if ($fitsDM -notcontains "$($v[1])-$($v[2])") { return $false }
        $price = @(900, 700)[$v[0]] + @(150, 120)[$v[1]] + @(300, 250)[$v[2]]
        $price -le 1200 }
    Cost = { param([int[]]$v) @(900, 700)[$v[0]] + @(150, 120)[$v[1]] + @(300, 250)[$v[2]] }
    Sat = $true; Worlds = 3; Opt = @(1070); OptCount = 1
})

# optimize-quote-bike. frame alloy 0 / carbon 1; wheels standard 0 / carbon 1;
# child seat, lights, rack each 0 or 1. Price and margin per item below.
$bikePrice = { param([int[]]$v) @(800, 1500)[$v[0]] + @(300, 700)[$v[1]] + 150 * $v[2] + 80 * $v[3] + 100 * $v[4] }
$bikeMargin = { param([int[]]$v) @(200, 500)[$v[0]] + @(60, 250)[$v[1]] + 40 * $v[2] + 30 * $v[3] + 35 * $v[4] }
$entries.Add(@{
    Id = 'optimize-quote-bike'
    Domains = (New-O1Domains 5 @(0, 1))
    Prune = $null
    Valid = { param([int[]]$v)
        if ($v[1] -eq 1 -and $v[0] -ne 1) { return $false }   # carbon wheels need the carbon frame
        if ($v[2] -eq 1 -and $v[0] -eq 1) { return $false }   # the child seat is not rated for carbon
        if ($v[2] -ne 1) { return $false }                    # the customer wants the child seat
        (& $bikePrice $v) -le 1350 }
    Cost = { param([int[]]$v) - (& $bikeMargin $v) }
    Maximize = $true
    Sat = $true; Worlds = 3; Opt = @(335); OptCount = 1
})

# optimize-quote-conflict: the same catalogue; the customer wants the carbon
# frame and the child seat, and there is no price limit.
$entries.Add(@{
    Id = 'optimize-quote-conflict'
    Domains = (New-O1Domains 5 @(0, 1))
    Prune = $null
    Valid = { param([int[]]$v)
        if ($v[1] -eq 1 -and $v[0] -ne 1) { return $false }
        if ($v[2] -eq 1 -and $v[0] -eq 1) { return $false }
        if ($v[2] -ne 1) { return $false }                    # wants the child seat
        if ($v[0] -ne 1) { return $false }                    # wants the carbon frame
        $true }
    Cost = $null
    Sat = $false; Worlds = 0; Opt = $null; OptCount = 0
})

# Network design - the sentence recursion through choices needs (SD-7).
# Offices HQ 0, North 1, South 2, West 3; candidate links, each on or off:
# 0 HQ-North 3, 1 HQ-South 4, 2 North-South 2, 3 North-West 5, 4 South-West 6.
$linkEnds = @(@(0, 1), @(0, 2), @(1, 2), @(1, 3), @(2, 3))
$linkCost = @(3, 4, 2, 5, 6)
$entries.Add(@{
    Id = 'optimize-network-connect'
    Domains = (New-O1Domains 5 @(0, 1))
    Prune = $null
    Valid = { param([int[]]$v)
        # Reachability from HQ over the chosen links, the recursion itself:
        # grown to a fixpoint rather than assumed.
        $seen = New-Object bool[] 4
        $seen[0] = $true
        $grew = $true
        while ($grew) {
            $grew = $false
            for ($i = 0; $i -lt 5; $i++) {
                if ($v[$i] -ne 1) { continue }
                $a = $linkEnds[$i][0]; $b = $linkEnds[$i][1]
                if ($seen[$a] -and -not $seen[$b]) { $seen[$b] = $true; $grew = $true }
                if ($seen[$b] -and -not $seen[$a]) { $seen[$a] = $true; $grew = $true }
            }
        }
        for ($o = 0; $o -lt 4; $o++) { if (-not $seen[$o]) { return $false } }
        $true }
    Cost = { param([int[]]$v) $c = 0; for ($i = 0; $i -lt 5; $i++) { if ($v[$i] -eq 1) { $c += $linkCost[$i] } }; $c }
    Sat = $true; Worlds = 14; Opt = @(10); OptCount = 1
})

# Segment 5 - timetabling
# optimize-exam-slots. Stats 0, ML 1, Law 2, Art 3; slots Mon AM 0, Mon PM 1,
# Tue AM 2. Zoe sits Stats+ML, Yan ML+Law, Xi Law+Art.
$entries.Add(@{
    Id = 'optimize-exam-slots'
    Domains = (New-O1Domains 4 @(0, 1, 2))
    Prune = $null
    Valid = { param([int[]]$v)
        foreach ($pr in @(@(0, 1), @(1, 2), @(2, 3))) { if ($v[$pr[0]] -eq $v[$pr[1]]) { return $false } }
        for ($s = 0; $s -lt 3; $s++) { if ((Measure-O1Count $v @(0, 1, 2, 3) $s) -gt 2) { return $false } }
        $true }
    Cost = { param([int[]]$v) $c = 0
        foreach ($pr in @(@(0, 1), @(1, 2), @(2, 3))) { if (($v[$pr[0]] + $v[$pr[1]]) -eq 1 -and $v[$pr[0]] -le 1 -and $v[$pr[1]] -le 1) { $c++ } }
        $c }
    Sat = $true; Worlds = 24; Opt = @(0); OptCount = 8
})

# optimize-class-timetable. Variables: 7A Math, 7A English, 7A Science, 7B
# Math, 7B English, 7B Science; periods 0-2. One teacher per subject.
$entries.Add(@{
    Id = 'optimize-class-timetable'
    Domains = (New-O1Domains 6 @(0, 1, 2))
    Prune = $null
    Valid = { param([int[]]$v)
        foreach ($c in 0, 3) { if ($v[$c] -eq $v[$c + 1] -or $v[$c] -eq $v[$c + 2] -or $v[$c + 1] -eq $v[$c + 2]) { return $false } }
        for ($s = 0; $s -lt 3; $s++) { if ($v[$s] -eq $v[$s + 3]) { return $false } }   # one teacher, one room at a time
        $true }
    Cost = { param([int[]]$v) $c = 0; if ($v[0] -eq 0) { $c++ }; if ($v[3] -eq 0) { $c++ }; $c }
    Sat = $true; Worlds = 12; Opt = @(0); OptCount = 4
})

# optimize-exam-rooms: five exams, two slots, two rooms a slot.
$entries.Add(@{
    Id = 'optimize-exam-rooms'
    Domains = (New-O1Domains 5 @(0, 1))
    Prune = $null
    Valid = { param([int[]]$v)
        for ($s = 0; $s -lt 2; $s++) { if ((Measure-O1Count $v @(0, 1, 2, 3, 4) $s) -gt 2) { return $false } }
        $true }
    Cost = $null
    Sat = $false; Worlds = 0; Opt = $null; OptCount = 0
})

# --- run -------------------------------------------------------------------------

$corpusLines = [IO.File]::ReadAllLines($corpusPath, [Text.Encoding]::UTF8)
$mismatches = 0
$caught = 0
$checked = 0

function Write-O1Result([string]$id, [string]$what, [bool]$agrees) {
    $script:checked++
    if ($agrees) { Write-Output "O01|$id|$what|agrees" }
    else { $script:mismatches++; Write-Output "O01|$id|$what|MISMATCH" }
}

foreach ($e in $entries) {
    if ($Only -and $e.Id -ne $Only) { continue }
    $key = @{ Sat = $e.Sat; Worlds = $e.Worlds; Opt = $e.Opt; OptCount = $e.OptCount }
    if ($Control) {
        # A planted wrong key: one more world than stated. Every entry must catch it.
        $key.Worlds = $key.Worlds + 1
        $key.Sat = $true
    }

    if ($e.Id -eq 'optimize-sod-check') {
        $v = [int[]]$e.Fixed
        $names = @('Ann', 'Bob', 'Cy', 'Di')
        $found = New-Object 'System.Collections.Generic.List[string]'
        for ($t = 0; $t -lt 3; $t++) {
            if ($v[2 * $t] -eq $v[2 * $t + 1]) { $found.Add("T$($t + 1): $($names[$v[2 * $t]]) reviews their own work") }
            if ($v[2 * $t + 1] -gt 1) { $found.Add("T$($t + 1): the reviewer, $($names[$v[2 * $t + 1]]), is not senior") }
        }
        for ($p = 0; $p -lt 4; $p++) { if ((Measure-O1Count $v @(0, 1, 2, 3, 4, 5) $p) -gt 2) { $found.Add("$($names[$p]) holds more than two roles") } }
        $worlds = if ($found.Count -eq 0) { 1 } else { 0 }
        $agree = ($worlds -eq $key.Worlds) -and (($found -join '; ') -eq ($e.Violations -join '; '))
        Write-O1Result $e.Id ("one world, breaks {0} rule(s): {1}" -f $found.Count, ($found -join '; ')) $agree
        if (-not $agree) { $caught++ }
        $e.Candidates = 1
    } else {
        $t0 = [DateTime]::UtcNow
        $res = Invoke-O1Search $e.Domains $e.Prune $e.Valid $e.Cost
        $secs = ([DateTime]::UtcNow - $t0).TotalSeconds
        $e.Candidates = $res.Candidates
        $opt = $null
        if ($null -ne $res.Best) {
            $opt = [int[]]$res.Best
            if ($e.ContainsKey('Maximize') -and $e.Maximize) { $opt = [int[]]@($opt | ForEach-Object { - $_ }) }
        }
        $sat = ($res.Worlds -gt 0)
        $agree = ($sat -eq $key.Sat) -and ($res.Worlds -eq $key.Worlds)
        if ($sat -and $null -ne $key.Opt) {
            $agree = $agree -and ($null -ne $opt) -and ((@($opt) -join ',') -eq (@($key.Opt) -join ',')) -and ($res.BestCount -eq $key.OptCount)
        }
        $optText = if ($null -eq $opt) { 'no optimum' } else { "optimum ($((@($opt)) -join ', ')), $($res.BestCount) reach it" }
        Write-O1Result $e.Id ("{0} valid of {1} candidates, {2} ({3:N1} s)" -f $res.Worlds, (Format-O1Count $res.Candidates), $optText, $secs) $agree
        if (-not $agree) { $caught++ }
    }

    if (-not $Control) {
        $block = Get-O1CorpusBlock $corpusLines $e.Id
        $line = Get-O1KeyLine @{ Sat = $e.Sat; Worlds = $e.Worlds; Opt = $e.Opt; OptCount = $e.OptCount; Candidates = $e.Candidates }
        if ($null -eq $block) { Write-O1Result $e.Id 'corpus: no entry with this id' $false }
        else { Write-O1Result $e.Id "corpus prints: $($line.Substring(9))" ($block.Contains($line)) }
    }
}

if (-not $Control -and (-not $Only -or $Only -in @('optimize-roster', 'optimize-roster-senior', 'optimize-roster-loose'))) {
    # optimize-roster-senior: 5 people x 1 week, tight (need 1). Two proofs.
    $seniors = @(1..5 | Where-Object { ($_ % 4) -eq 0 })
    $nights = @(1..21 | Where-Object { (Get-O1ShiftSlot $_) -eq 2 })
    $p4Leave = @(1..7 | Where-Object { Test-O1OnLeave 4 $_ })
    $uncovered = @($nights | Where-Object { $d = Get-O1ShiftDay $_; @($seniors | Where-Object { -not (Test-O1OnLeave $_ $d) }).Count -eq 0 })
    $ok = ($seniors.Count -eq 1 -and $seniors[0] -eq 4 -and $nights.Count -eq 7 -and $p4Leave.Count -eq 1 -and $p4Leave[0] -eq 3 -and $uncovered.Count -eq 1)
    Write-O1Result 'optimize-roster-senior' ("seniors {0}; {1} nights; P4 on leave day {2}; nights with no senior free: {3}" -f ($seniors -join ','), $nights.Count, ($p4Leave -join ','), ($uncovered -join ',')) $ok
    $block = Get-O1CorpusBlock $corpusLines 'optimize-roster-senior'
    $ok2 = ($null -ne $block) -and $block.Contains('P4') -and $block.Contains('day 3') -and $block.Contains('7 nights') -and $block.Contains('at most 5')
    Write-O1Result 'optimize-roster-senior' 'corpus names P4, day 3, 7 nights and at most 5' $ok2

    # optimize-roster-loose: 10 people x 1 week, needs Early 2, Late 3, Night 4.
    $slots = 0; for ($s = 1; $s -le 21; $s++) { $slots += 2 + (Get-O1ShiftSlot $s) }
    $capacity = 10 * 5
    Write-O1Result 'optimize-roster-loose' ("{0} slots a week against a capacity of {1}" -f $slots, $capacity) ($slots -eq 63 -and $capacity -eq 50 -and $slots -gt $capacity)
    $block = Get-O1CorpusBlock $corpusLines 'optimize-roster-loose'
    Write-O1Result 'optimize-roster-loose' 'corpus names 63 and 50' (($null -ne $block) -and $block.Contains('63') -and $block.Contains('50'))

    # optimize-roster: the reference roster's hand floor on changes.
    $clash50 = Measure-O1KeptOnLeave 50 4
    $clash10 = Measure-O1KeptOnLeave 10 1
    Write-O1Result 'optimize-roster' ("kept assignments on leave: {0} at 50 x 4 (floor {1}), {2} at 10 x 1 (floor {3})" -f $clash50, (2 * $clash50), $clash10, (2 * $clash10)) ($clash50 -eq 118)
    $block = Get-O1CorpusBlock $corpusLines 'optimize-roster'
    Write-O1Result 'optimize-roster' 'corpus names the floor 236 and the proven 24' (($null -ne $block) -and $block.Contains('236') -and $block.Contains('24'))
}

if ($Control) {
    $expectedCaught = @($entries | Where-Object { -not $Only -or $_.Id -eq $Only }).Count
    Write-Output ("O01|control|planted a wrong key in {0} entries, caught {1}" -f $expectedCaught, $caught)
    if ($caught -ne $expectedCaught) { Write-Output 'O01|control|FAILED: a planted wrong key went unnoticed'; exit 1 }
    Write-Output 'O01|control|every planted key was caught'
    exit 0
}

Write-Output ("O01|done|{0} checks, {1} mismatches" -f $checked, $mismatches)
if ($mismatches -gt 0) { exit 1 }
exit 0
