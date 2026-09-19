# optimize0_expected.ps1 - OPTIMIZE.0's expected answers, derived independently.
#
# archive/VLA_DiagO0.bas checks every DATALOG answer against closed forms
# (S * C(P,3), 18 * P * W, ...). This script derives the same numbers a second
# way: by ENUMERATING the rows from the fixture's definitions (who is on leave,
# which shifts are nights, which pairs are consecutive) and counting them. It
# then compares each count with the closed form the harness uses and says
# MISMATCH if they differ. Where a rung is too large to enumerate whole (over
# EnumCap rows), it enumerates ONE group and multiplies by the number of
# groups, and says so: every group of those shapes is identical by definition.
#
# It also counts the toy (5 people, 7 shifts, exactly 2 per shift, never two in
# a row) by exhaustive search, summarises the reference fixtures exactly as
# O0Fixture prints them, and with -PaperModel reprints Entry 1's table in
# docs/OPTIMIZATION.md from its formulas.
#
# Host-independent: PowerShell 5.1, no Excel, no network, writes nothing.
#
#   powershell -NoProfile -ExecutionPolicy Bypass -File tools\optimize0_expected.ps1
#   ... -Step 8          one step's rungs only
#   ... -PaperModel      the paper model table as well

param(
    [int]$Step = 0,
    [switch]$PaperModel,
    [int]$EnumCap = 400000
)

Set-StrictMode -Version 2
$ErrorActionPreference = 'Stop'

# --- the fixture's definitions (the same closed forms as VLA_DiagO0.bas) -----

function Get-ShiftDay([int]$shiftNo) { [int][math]::Floor(($shiftNo - 1) / 3) + 1 }
function Get-ShiftSlot([int]$shiftNo) { ($shiftNo - 1) % 3 }
function Get-ShiftWeek([int]$shiftNo) { [int][math]::Floor(((Get-ShiftDay $shiftNo) - 1) / 7) + 1 }
function Test-OnLeave([int]$person, [int]$dayNo) { (($person + $dayNo) % 7) -eq 0 }
function Test-Senior([int]$person) { ($person % 4) -eq 0 }
function Get-Contract([int]$person) { if (($person % 2) -eq 1) { 4 } else { 5 } }
function Get-TightNeed([int]$people) { $t = [int][math]::Floor($people / 5); if ($t -lt 1) { 1 } else { $t } }
function Get-Need([int]$shiftNo, [int]$people, [bool]$loose) {
    if ($loose) { 2 + (Get-ShiftSlot $shiftNo) } else { Get-TightNeed $people }
}

function Get-Comb([int]$n, [int]$r) {
    if ($r -lt 0 -or $r -gt $n) { return [double]0 }
    $v = [double]1
    for ($i = 1; $i -le $r; $i++) { $v = $v * ($n - $r + $i) / $i }
    [math]::Round($v)
}

# The kept schedule, built by walking the round-robin rather than by the
# harness's KeptStart formula: a cursor that advances by each shift's need.
function Get-KeptRows([int]$people, [int]$weeks, [bool]$loose) {
    $rows = New-Object System.Collections.Generic.List[object]
    $cursor = 0
    for ($is = 1; $is -le 21 * $weeks; $is++) {
        $need = Get-Need $is $people $loose
        for ($j = 0; $j -lt $need; $j++) {
            # Every element parenthesised: in PowerShell `,` binds tighter than
            # `+`, so `$a, ($x) + 1, $b` concatenates two arrays into FOUR elements.
            $rows.Add(@($is, ((($cursor + $j) % $people) + 1), (Get-ShiftWeek $is)))
        }
        $cursor += $need
    }
    , $rows
}

# Counts increasing tuples of `size` drawn from $members, by walking every one.
function Measure-Increasing([int[]]$members, [int]$size) {
    $n = $members.Count
    if ($size -gt $n) { return 0 }
    $idx = New-Object int[] $size
    for ($i = 0; $i -lt $size; $i++) { $idx[$i] = $i }
    $count = 0
    while ($true) {
        $count++
        $pos = $size - 1
        while ($pos -ge 0 -and $idx[$pos] -eq $n - $size + $pos) { $pos-- }
        if ($pos -lt 0) { break }
        $idx[$pos]++
        for ($q = $pos + 1; $q -lt $size; $q++) { $idx[$q] = $idx[$q - 1] + 1 }
    }
    $count
}

# Counts ordered triples of distinct members, by walking every one.
function Measure-OrderedTriples([int[]]$members) {
    $count = 0
    foreach ($x1 in $members) { foreach ($x2 in $members) { if ($x2 -eq $x1) { continue }
        foreach ($x3 in $members) { if ($x3 -ne $x1 -and $x3 -ne $x2) { $count++ } } } }
    $count
}

function Get-Members([int]$people, [scriptblock]$keep) {
    $list = New-Object System.Collections.Generic.List[int]
    for ($ip = 1; $ip -le $people; $ip++) { if (& $keep $ip) { $list.Add($ip) } }
    , $list.ToArray()
}

# --- each step's closed form (what the harness checks against) ---------------

function Get-ClosedForm([int]$stepNo, [int]$people, [int]$weeks, [int]$kk) {
    $nS = 21 * $weeks
    switch ($stepNo) {
        1 { return [double]$people * $nS }
        2 { return 18.0 * $people * $weeks }
        3 { return 14.0 * $weeks * $people + 7.0 * $weeks * [math]::Floor($people / 4) }
        4 { return [double]$nS }
        5 { return [double]$people * $weeks }
        { $_ -eq 6 -or $_ -eq 7 } { return [double]$people * ($nS - 1) }
        8 { return $nS * (Get-Comb $people ($kk + 1)) }
        9 { return [double]$nS * $people * ($people - 1) * ($people - 2) }
        10 { return [double]$people * $weeks * (Get-Comb 21 ($kk + 1)) }
        11 { return $null }
    }
}

# --- each step's enumeration -------------------------------------------------

function Get-Enumerated([int]$stepNo, [int]$people, [int]$weeks, [int]$kk) {
    $nS = 21 * $weeks
    $how = 'enumerated'
    $count = [double]0
    $detail = ''
    switch ($stepNo) {
        1 {
            for ($is = 1; $is -le $nS; $is++) { for ($ip = 1; $ip -le $people; $ip++) { $count++ } }
            $detail = "every shift $people"
        }
        2 {
            $perShift = @{}
            for ($is = 1; $is -le $nS; $is++) {
                $n = 0
                for ($ip = 1; $ip -le $people; $ip++) { if (-not (Test-OnLeave $ip (Get-ShiftDay $is))) { $n++ } }
                $count += $n; $perShift[$n] = 1
            }
            $detail = 'rows per shift: ' + (($perShift.Keys | Sort-Object) -join ' or ')
        }
        3 {
            for ($is = 1; $is -le $nS; $is++) { for ($ip = 1; $ip -le $people; $ip++) {
                if ((Get-ShiftSlot $is) -ne 2 -or (Test-Senior $ip)) { $count++ } } }
            $detail = "nights $([math]::Floor($people / 4)), other shifts $people"
        }
        4 {
            $vals = @{}
            for ($is = 1; $is -le $nS; $is++) {
                $n = 0
                for ($ip = 1; $ip -le $people; $ip++) { if (-not (Test-OnLeave $ip (Get-ShiftDay $is))) { $n++ } }
                $count++; $vals[$n] = 1
            }
            $detail = 'N per shift: ' + (($vals.Keys | Sort-Object) -join ' or ')
        }
        5 {
            $vals = @{}
            for ($ip = 1; $ip -le $people; $ip++) { for ($iw = 1; $iw -le $weeks; $iw++) {
                $n = 0
                for ($is = 1; $is -le $nS; $is++) {
                    if ((Get-ShiftWeek $is) -eq $iw -and -not (Test-OnLeave $ip (Get-ShiftDay $is))) { $n++ } }
                $count++; $vals[$n] = 1 } }
            $detail = 'N per person-week: ' + (($vals.Keys | Sort-Object) -join ' or ')
        }
        6 {
            for ($ip = 1; $ip -le $people; $ip++) { for ($is = 1; $is -lt $nS; $is++) { $count++ } }
            $detail = "per person $($nS - 1)"
        }
        7 {
            if ([double]$people * $nS * $nS -le 4 * $EnumCap) {
                for ($ip = 1; $ip -le $people; $ip++) { for ($ia = 1; $ia -le $nS; $ia++) { for ($ib = 1; $ib -le $nS; $ib++) {
                    if ($ib - $ia -eq 1) { $count++ } } } }
            } else {
                $one = 0
                for ($ia = 1; $ia -le $nS; $ia++) { for ($ib = 1; $ib -le $nS; $ib++) { if ($ib - $ia -eq 1) { $one++ } } }
                $count = [double]$one * $people; $how = "one person enumerated x $people"
            }
            $detail = "all pairs walked: $([double]$people * $nS * $nS)"
        }
        8 {
            $members = Get-Members $people { param($x) $true }
            $closed = Get-Closedform 8 $people $weeks $kk
            if ($closed -le $EnumCap) {
                for ($is = 1; $is -le $nS; $is++) { $count += Measure-Increasing $members ($kk + 1) }
            } elseif ((Get-Comb $people ($kk + 1)) -le $EnumCap) {
                $count = [double](Measure-Increasing $members ($kk + 1)) * $nS; $how = "one shift enumerated x $nS"
            } else { return $null }
            $detail = "per shift $($count / $nS)"
        }
        9 {
            $members = Get-Members $people { param($x) $true }
            $one = Measure-OrderedTriples $members
            if ([double]$one * $nS -le $EnumCap) {
                for ($is = 1; $is -le $nS; $is++) { $count += Measure-OrderedTriples $members }
            } else { $count = [double]$one * $nS; $how = "one shift enumerated x $nS" }
            $detail = "per shift $one"
        }
        10 {
            $groups = $people * $weeks
            $one = 0
            if ((Get-Comb 21 ($kk + 1)) -le $EnumCap) {
                $week1 = New-Object System.Collections.Generic.List[int]
                for ($is = 1; $is -le $nS; $is++) { if ((Get-ShiftWeek $is) -eq 1) { $week1.Add($is) } }
                $one = Measure-Increasing $week1.ToArray() ($kk + 1)
            } else { return $null }
            if ([double]$one * $groups -le $EnumCap) {
                for ($g = 1; $g -le $groups; $g++) { $count += $one }
                $how = "each person-week enumerated alike ($one)"
            } else { $count = [double]$one * $groups; $how = "one person-week enumerated x $groups" }
            $detail = "per person-week $one"
        }
        11 {
            $kept = Get-KeptRows $people $weeks $false
            $worked = @{}
            foreach ($row in $kept) { $key = "$($row[1])|$($row[2])"; if ($worked.ContainsKey($key)) { $worked[$key]++ } else { $worked[$key] = 1 } }
            $vals = @{}
            for ($ip = 1; $ip -le $people; $ip++) { for ($iw = 1; $iw -le $weeks; $iw++) {
                $key = "$ip|$iw"; $n = 0; if ($worked.ContainsKey($key)) { $n = $worked[$key] }
                $over = $n - (Get-Contract $ip)
                if ($over -gt 0) { $count++; $vals[$over] = 1 } } }
            $detail = 'overtime values: ' + (($vals.Keys | Sort-Object) -join ' or ')
        }
    }
    [pscustomobject]@{ Count = $count; How = $how; Detail = $detail }
}

# --- the rungs, the same lists as VLA_DiagO0.bas's RungList ------------------

function Get-Rungs([int]$stepNo) {
    switch ($stepNo) {
        { $_ -in 1, 2, 3, 4, 5, 6, 7, 11 } { return '5,1 10,1 5,4 20,1 10,4 50,1 20,4 50,4' -split ' ' }
        8 { return '5,1 7,1 10,1 5,4 14,1 7,4 20,1 10,4 14,4 20,4 50,1 50,4 50,4,10' -split ' ' }
        9 { return '5,1 7,1 10,1 5,4 14,1 7,4 20,1 10,4 14,4 20,4 50,1 50,4' -split ' ' }
        10 { return '1,1,1 5,1,1 1,1,2 10,1,1 20,1,1 5,1,2 1,1,3 10,1,2 20,1,2 5,1,3 10,1,3 1,1,5 20,1,3 5,1,5 50,4,5' -split ' ' }
    }
}

$mismatches = 0
$steps = if ($Step -gt 0) { @($Step) } else { 1..11 }
Write-Output 'O0E|step|people|weeks|k|expected rows|closed form|agrees|how|detail'
foreach ($st in $steps) {
    foreach ($rung in (Get-Rungs $st)) {
        $parts = $rung -split ','
        $people = [int]$parts[0]; $weeks = [int]$parts[1]; $kk = 2
        if ($parts.Count -ge 3) { $kk = [int]$parts[2] }
        $kText = if ($st -eq 8 -or $st -eq 10) { "$kk" } else { '' }
        $closed = Get-ClosedForm $st $people $weeks $kk
        $big = ($null -ne $closed -and $closed -gt 50 * $EnumCap)
        if ($big) {
            Write-Output ("O0E|{0}|{1}|{2}|{3}||{4:N0}|model only|too large to enumerate|" -f $st, $people, $weeks, $kText, $closed)
            continue
        }
        $e = Get-Enumerated $st $people $weeks $kk
        if ($null -eq $e) {
            Write-Output ("O0E|{0}|{1}|{2}|{3}||{4:N0}|model only|too large to enumerate|" -f $st, $people, $weeks, $kText, $closed)
            continue
        }
        if ($null -eq $closed) { $agree = 'no closed form (harness counts the kept rows)'; $closedText = '' }
        elseif ($closed -eq $e.Count) { $agree = 'yes'; $closedText = '{0:N0}' -f $closed }
        else { $agree = 'MISMATCH'; $closedText = '{0:N0}' -f $closed; $mismatches++ }
        Write-Output ("O0E|{0}|{1}|{2}|{3}|{4:N0}|{5}|{6}|{7}|{8}" -f $st, $people, $weeks, $kText, $e.Count, $closedText, $agree, $e.How, $e.Detail)
    }
}

# --- the toy: 5 people, 7 shifts, exactly 2 per shift, never two in a row ----

$pairs = New-Object System.Collections.Generic.List[object]
for ($x1 = 1; $x1 -le 5; $x1++) { for ($x2 = $x1 + 1; $x2 -le 5; $x2++) { $pairs.Add(@($x1, $x2)) } }
$toyWorlds = 0
$stack = New-Object System.Collections.Generic.Stack[object]
foreach ($pr in $pairs) { $stack.Push(@(1, $pr)) }
while ($stack.Count -gt 0) {
    $top = $stack.Pop(); $depth = $top[0]; $last = $top[1]
    if ($depth -eq 7) { $toyWorlds++; continue }
    foreach ($pr in $pairs) {
        if ($pr[0] -ne $last[0] -and $pr[0] -ne $last[1] -and $pr[1] -ne $last[0] -and $pr[1] -ne $last[1]) {
            $stack.Push(@(($depth + 1), $pr))
        }
    }
}
$toyHand = 10 * [math]::Pow(3, 6)
$toyAgree = if ($toyWorlds -eq $toyHand) { 'yes' } else { $mismatches++; 'MISMATCH' }
Write-Output ("O0E|toy|valid worlds {0:N0}|by hand 10 x 3^6 = {1:N0}|agrees {2}|of {3:N0} candidates" -f $toyWorlds, $toyHand, $toyAgree, [math]::Pow(10, 7))

# --- the reference fixtures, as O0Fixture prints them ------------------------

Write-Output 'O0EF|fixture|need|weekly demand|tightness|seniors|leave rows|kept rows|on leave|two in a row|over five|nights no senior|over contract|a schedule exists'
foreach ($loose in @($false, $true)) {
    foreach ($people in 5, 10, 20, 50) {
        foreach ($weeks in 1, 4) {
            $nS = 21 * $weeks
            $demand = 0; for ($is = 1; $is -le $nS; $is++) { $demand += Get-Need $is $people $loose }
            $leaveRows = 0
            for ($ip = 1; $ip -le $people; $ip++) { for ($id = 1; $id -le 7 * $weeks; $id++) { if (Test-OnLeave $ip $id) { $leaveRows++ } } }
            $kept = Get-KeptRows $people $weeks $loose
            $on = @{}; $wk = @{}; $clash = 0
            foreach ($row in $kept) {
                $on["$($row[0])|$($row[1])"] = 1
                $key = "$($row[1])|$($row[2])"; if ($wk.ContainsKey($key)) { $wk[$key]++ } else { $wk[$key] = 1 }
                if (Test-OnLeave $row[1] (Get-ShiftDay $row[0])) { $clash++ }
            }
            $inRow = 0; $noSenior = 0
            for ($is = 1; $is -le $nS; $is++) {
                if ($is -lt $nS) { for ($ip = 1; $ip -le $people; $ip++) { if ($on.ContainsKey("$is|$ip") -and $on.ContainsKey("$($is + 1)|$ip")) { $inRow++ } } }
                if ((Get-ShiftSlot $is) -eq 2) {
                    $has = $false
                    for ($ip = 1; $ip -le $people; $ip++) { if ($on.ContainsKey("$is|$ip") -and (Test-Senior $ip)) { $has = $true } }
                    if (-not $has) { $noSenior++ }
                }
            }
            $overFive = 0; $overContract = 0
            for ($ip = 1; $ip -le $people; $ip++) { for ($iw = 1; $iw -le $weeks; $iw++) {
                $n = 0; if ($wk.ContainsKey("$ip|$iw")) { $n = $wk["$ip|$iw"] }
                if ($n -gt 5) { $overFive++ }
                if ($n -gt (Get-Contract $ip)) { $overContract++ } } }
            $seniors = [math]::Floor($people / 4)
            $exists = if ($demand / $weeks -gt 5 * $people) { 'NO, provably (weekly demand over 5 per person)' }
                      elseif (5 * $seniors -lt 7) { 'NO, provably (too few seniors for 7 nights)' }
                      else { 'not decided by hand' }
            $needText = if ($loose) { 'Early 2, Late 3, Night 4' } else { "$(Get-TightNeed $people) on every shift" }
            $name = "$people" + 'x' + "$weeks" + $(if ($loose) { 'L' } else { '' })
            Write-Output ("O0EF|{0}|{1}|{2}|{3:0.00}|{4}|{5}|{6}|{7}|{8}|{9}|{10}|{11}|{12}" -f $name, $needText, ($demand / $weeks), ($demand / $weeks / (5.0 * $people)), $seniors, $leaveRows, $kept.Count, $clash, $inRow, $overFive, $noSenior, $overContract, $exists)
        }
    }
}

# --- the paper model (docs/OPTIMIZATION.md, Entry 1) -------------------------

if ($PaperModel) {
    Write-Output 'O0P|people|weeks|pool|pool minus leave|amo2 shift combos|amo2 native|amo5 week combos|amo5 native|row pairs|all-pairs join rows|one-a-day combos|one-a-day native|alo2 clauses|overtime combos (c=4)|worlds log10'
    foreach ($people in 10, 20, 50) {
        foreach ($weeks in 1, 4) {
            $nS = 21 * $weeks
            Write-Output ("O0P|{0}|{1}|{2:N0}|{3:N0}|{4:N0}|{5}|{6:N0}|{7}|{8:N0}|{9:N0}|{10:N0}|{11:N0}|{12:N0}|{13:N0}|{14:N1}" -f `
                $people, $weeks, ($people * $nS), (18 * $people * $weeks), ($nS * (Get-Comb $people 3)), $nS, `
                ($people * $weeks * (Get-Comb 21 6)), ($people * $weeks), ($people * ($nS - 1)), ([double]$people * $nS * $nS), `
                ($people * 7 * $weeks * 3), ($people * 7 * $weeks), ($nS * $people), ($people * $weeks * (Get-Comb 21 5)), `
                ($nS * [math]::Log10((Get-Comb $people 2))))
        }
    }
}

if ($mismatches -gt 0) {
    Write-Output "=== $mismatches MISMATCH(ES) - the harness's closed forms disagree with the enumeration ==="
    exit 1
}
Write-Output '=== every enumerated count agrees with the closed form the harness checks against ==='
exit 0
