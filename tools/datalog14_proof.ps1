# SPDX-License-Identifier: 0BSD
#
# DATALOG.14 - the specialisation's decision procedure and its answers,
# proven before a line of it is imported into Excel.
#
# WHAT THIS PROVES, AND WHAT IT DOES NOT. The filter fix (an atom plan built
# once, and a relation handed back whole when the atom constrains nothing)
# is a refactor with no decision in it. The SPECIALISATION - pushing the
# question's own bound argument into a recursive predicate - is a rewrite
# that can change an answer, and its whole safety rests on five conditions.
# So this script transliterates those five conditions and a small Datalog
# evaluator, and asks two things of every corpus program:
#
#   1. Does the decision procedure reach the verdict the entry claims -
#      specialise at this position with this constant, or decline?
#   2. Where it specialises, does the program still answer IDENTICALLY?
#
# THE CONTROL COMES FIRST. Every corpus program states its answer, derived
# by hand from the program and its data. The unspecialised evaluator must
# reproduce that answer before the specialised one is compared against it.
# A harness that is wrong about the semantics would otherwise bless a
# rewrite by agreeing with itself - which is exactly how a green control
# over a construct it does not contain gets you nothing.
#
# It calls nothing in Frazaro and touches no workbook. Host-independent.

[CmdletBinding()]
param(
    [ValidateSet('none', 'no-query-guard', 'same-constant', 'variable-pins', 'passthrough', 'value-use')]
    [string]$Mutant = 'none',
    [switch]$Mutations
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Which condition, if any, is switched off for this run. -Mutations drives
# every one of them in turn and requires each to break something: a condition
# no mutant can break is a condition this corpus does not actually test.
$script:Mut = $Mutant

# body item kinds, VLA_Datalog.bas' own BI_* constants
$BI_POS = 1
$BI_NOT = 2
$BI_CMP = 3
$BI_COUNT = 4

# ---------------------------------------------------------------- atoms

function New-Atom {
    # NOT $Args - a PowerShell automatic variable binds nothing there.
    param([string]$Pred, [string[]]$ArgTexts)
    $recs = [System.Collections.Generic.List[object]]::new()
    foreach ($a in $ArgTexts) {
        if ($a.StartsWith('"')) { [void]$recs.Add(@{ Var = $false; Text = $a.Substring(1) }) }
        else { [void]$recs.Add(@{ Var = $true; Text = $a }) }
    }
    @{ Pred = $Pred; Args = $recs }
}

function New-Item2 {
    param([int]$Kind, $Atom, [string]$ResultVar = '')
    @{ Kind = $Kind; Atom = $Atom; ResultVar = $ResultVar }
}

function New-Rule {
    param($Head, $Body)
    @{ Head = $Head; Body = $Body }
}

function Test-SameVar {
    param([string]$A, [string]$B)
    return ($A.ToUpperInvariant() -eq $B.ToUpperInvariant())
}

function Copy-Atom {
    param($Atom)
    $recs = [System.Collections.Generic.List[object]]::new()
    foreach ($a in $Atom.Args) { [void]$recs.Add(@{ Var = $a.Var; Text = $a.Text }) }
    @{ Pred = $Atom.Pred; Args = $recs }
}

function Copy-Program {
    param($P)
    $rules = [System.Collections.Generic.List[object]]::new()
    foreach ($r in $P.Rules) {
        $body = [System.Collections.Generic.List[object]]::new()
        foreach ($b in $r.Body) {
            [void]$body.Add(@{ Kind = $b.Kind; Atom = (Copy-Atom $b.Atom); ResultVar = $b.ResultVar })
        }
        [void]$rules.Add(@{ Head = (Copy-Atom $r.Head); Body = $body })
    }
    $qa = $null
    if ($null -ne $P.QueryAtom) { $qa = Copy-Atom $P.QueryAtom }
    @{ Rules = $rules; Facts = $P.Facts; Edges = $P.Edges; Query = $P.Query; QueryAtom = $qa }
}

# ------------------------------------------------- the decision procedure
#
# A line-for-line transliteration of PushBoundArguments / SpecialisePredicate
# / SpecialiseAtPosition in VLA_Datalog.bas. Returns the verdicts it reached,
# as "pred@position=constant", and rewrites $P in place.

function Test-BodyItemMentionsVar {
    param($Item, [string]$V)
    if ($Item.ResultVar -and (Test-SameVar $Item.ResultVar $V)) { return $true }
    foreach ($a in $Item.Atom.Args) {
        if ($a.Var -and (Test-SameVar $a.Text $V)) { return $true }
    }
    return $false
}

function Test-OccurrencePins {
    param($Atom, [int]$K, [ref]$Wanted, [ref]$HaveOne)
    if ($Atom.Args.Count -lt $K) { return $false }
    $a = $Atom.Args[$K - 1]
    if ($a.Var) {
        if ($script:Mut -eq 'variable-pins') { return $true }
        return $false
    }
    if ($HaveOne.Value) {
        if ($Wanted.Value -cne $a.Text) {
            if ($script:Mut -ne 'same-constant') { return $false }
            return $true
        }
    } else {
        $Wanted.Value = $a.Text
        $HaveOne.Value = $true
    }
    return $true
}

function Invoke-SpecialiseAtPosition {
    param($P, [string]$Pred, [int]$K, [string]$C)

    # Condition 5, over EVERY Pred-headed rule, before any is rewritten.
    foreach ($r in $P.Rules) {
        if ($r.Head.Pred -cne $Pred) { continue }
        if ($r.Head.Args.Count -lt $K) { return $false }
        $headArg = $r.Head.Args[$K - 1]
        if (-not $headArg.Var) { return $false }
        $v = $headArg.Text
        foreach ($b in $r.Body) {
            if ($b.Kind -ne $script:BI_POS) {
                if ((Test-BodyItemMentionsVar -Item $b -V $v) -and $script:Mut -ne 'value-use') { return $false }
                if ($b.Atom.Pred -ceq $Pred) { return $false }
            } elseif ($b.Atom.Pred -ceq $Pred) {
                if ($b.Atom.Args.Count -lt $K) { return $false }
                $ba = $b.Atom.Args[$K - 1]
                if (-not $ba.Var) { return $false }
                if ((-not (Test-SameVar $ba.Text $v)) -and $script:Mut -ne 'passthrough') { return $false }
            }
        }
    }

    foreach ($r in $P.Rules) {
        if ($r.Head.Pred -cne $Pred) { continue }
        $v = $r.Head.Args[$K - 1].Text
        foreach ($a in $r.Head.Args) {
            if ($a.Var -and (Test-SameVar $a.Text $v)) { $a.Var = $false; $a.Text = $C }
        }
        foreach ($b in $r.Body) {
            if ($b.Kind -ne $script:BI_POS) { continue }
            foreach ($a in $b.Atom.Args) {
                if ($a.Var -and (Test-SameVar $a.Text $v)) { $a.Var = $false; $a.Text = $C }
            }
        }
    }
    return $true
}

function Invoke-SpecialisePredicate {
    param($P, [string]$Pred, [ref]$Verdicts)

    if ($Pred -ceq $P.Query -and $script:Mut -ne 'no-query-guard') { return }   # condition 1
    if ($P.Facts.ContainsKey($Pred)) { return }                # condition 2

    $selfRec = $false
    $arity = 0
    foreach ($r in $P.Rules) {
        if ($r.Head.Pred -cne $Pred) { continue }
        $arity = $r.Head.Args.Count
        foreach ($b in $r.Body) {
            if ($b.Kind -eq $script:BI_POS -and $b.Atom.Pred -ceq $Pred) { $selfRec = $true }
        }
    }
    if (-not $selfRec -or $arity -eq 0) { return }             # condition 3

    for ($k = 1; $k -le $arity; $k++) {                        # condition 4
        $wanted = ''
        $haveOne = $false
        $ok = $true
        if ($null -ne $P.QueryAtom -and $P.QueryAtom.Pred -ceq $Pred) {
            if (-not (Test-OccurrencePins -Atom $P.QueryAtom -K $k -Wanted ([ref]$wanted) -HaveOne ([ref]$haveOne))) { $ok = $false }
        }
        if ($ok) {
            foreach ($r in $P.Rules) {
                if ($r.Head.Pred -ceq $Pred) { continue }
                foreach ($b in $r.Body) {
                    if ($b.Atom.Pred -ceq $Pred) {
                        if (-not (Test-OccurrencePins -Atom $b.Atom -K $k -Wanted ([ref]$wanted) -HaveOne ([ref]$haveOne))) { $ok = $false }
                    }
                }
            }
        }
        if ($ok -and $haveOne) {
            if (Invoke-SpecialiseAtPosition -P $P -Pred $Pred -K $k -C $wanted) {
                [void]$Verdicts.Value.Add("$Pred@$k=$wanted")
                return
            }
        }
    }
}

function Invoke-PushBoundArguments {
    param($P)
    $verdicts = [System.Collections.Generic.List[string]]::new()
    $heads = [System.Collections.Generic.List[string]]::new()
    foreach ($r in $P.Rules) {
        if (-not $heads.Contains($r.Head.Pred)) { [void]$heads.Add($r.Head.Pred) }
    }
    foreach ($h in $heads) { Invoke-SpecialisePredicate -P $P -Pred $h -Verdicts ([ref]$verdicts) }
    return , $verdicts
}

# ------------------------------------------------------------- evaluator
#
# Enough of VLA_Datalog's fixpoint to decide answers: positive atoms with
# constants and repeated variables, `not` as an anti-join, `count` as a
# grouped aggregate over a fully-computed relation, and `=` comparison.
# Naive rather than semi-naive - it computes the same least fixed point, and
# what is being proven here is the ANSWER, not the cost.

function Get-RelKey { param([string[]]$T) return [string]::Join([char]31, $T) }

function Invoke-Body {
    param($Rule, $Relations)

    $rows = [System.Collections.Generic.List[object]]::new()
    [void]$rows.Add(@{})                    # one empty binding

    foreach ($b in $Rule.Body) {
        $next = [System.Collections.Generic.List[object]]::new()

        if ($b.Kind -eq $script:BI_CMP) {
            foreach ($env in $rows) {
                $l = if ($b.Atom.Args[0].Var) { $env[$b.Atom.Args[0].Text.ToUpperInvariant()] } else { $b.Atom.Args[0].Text }
                $r = if ($b.Atom.Args[1].Var) { $env[$b.Atom.Args[1].Text.ToUpperInvariant()] } else { $b.Atom.Args[1].Text }
                if ($l -ceq $r) { [void]$next.Add($env) }
            }
            $rows = $next
            continue
        }

        if ($b.Kind -eq $script:BI_NOT) {
            $rel = if ($Relations.ContainsKey($b.Atom.Pred)) { $Relations[$b.Atom.Pred] } else { @{} }
            foreach ($env in $rows) {
                $probe = [string[]]::new($b.Atom.Args.Count)
                for ($i = 0; $i -lt $b.Atom.Args.Count; $i++) {
                    $a = $b.Atom.Args[$i]
                    $probe[$i] = if ($a.Var) { $env[$a.Text.ToUpperInvariant()] } else { $a.Text }
                }
                if (-not $rel.ContainsKey((Get-RelKey $probe))) { [void]$next.Add($env) }
            }
            $rows = $next
            continue
        }

        if ($b.Kind -eq $script:BI_COUNT) {
            $rel = if ($Relations.ContainsKey($b.Atom.Pred)) { $Relations[$b.Atom.Pred] } else { @{} }
            foreach ($env in $rows) {
                $n = 0
                foreach ($key in $rel.Keys) {
                    $t = $rel[$key]
                    $consistent = $true
                    for ($i = 0; $i -lt $b.Atom.Args.Count; $i++) {
                        $a = $b.Atom.Args[$i]
                        if (-not $a.Var) {
                            if ($a.Text -cne $t[$i]) { $consistent = $false; break }
                        } elseif ($env.ContainsKey($a.Text.ToUpperInvariant())) {
                            if ($env[$a.Text.ToUpperInvariant()] -cne $t[$i]) { $consistent = $false; break }
                        }
                    }
                    if ($consistent) { $n++ }
                }
                $e2 = @{}
                foreach ($kk in $env.Keys) { $e2[$kk] = $env[$kk] }
                $e2[$b.ResultVar.ToUpperInvariant()] = "$n"
                [void]$next.Add($e2)
            }
            $rows = $next
            continue
        }

        # BI_POS
        $rel = if ($Relations.ContainsKey($b.Atom.Pred)) { $Relations[$b.Atom.Pred] } else { @{} }
        foreach ($env in $rows) {
            foreach ($key in $rel.Keys) {
                $t = $rel[$key]
                $e2 = @{}
                foreach ($kk in $env.Keys) { $e2[$kk] = $env[$kk] }
                $fit = $true
                for ($i = 0; $i -lt $b.Atom.Args.Count; $i++) {
                    $a = $b.Atom.Args[$i]
                    if (-not $a.Var) {
                        if ($a.Text -cne $t[$i]) { $fit = $false; break }
                    } else {
                        $nm = $a.Text.ToUpperInvariant()
                        if ($e2.ContainsKey($nm)) {
                            if ($e2[$nm] -cne $t[$i]) { $fit = $false; break }
                        } else {
                            $e2[$nm] = $t[$i]
                        }
                    }
                }
                if ($fit) { [void]$next.Add($e2) }
            }
        }
        $rows = $next
    }

    $out = [System.Collections.Generic.List[object]]::new()
    foreach ($env in $rows) {
        $head = [string[]]::new($Rule.Head.Args.Count)
        for ($i = 0; $i -lt $Rule.Head.Args.Count; $i++) {
            $a = $Rule.Head.Args[$i]
            $head[$i] = if ($a.Var) { $env[$a.Text.ToUpperInvariant()] } else { $a.Text }
        }
        [void]$out.Add($head)
    }
    return , $out
}

# Strata, so `not` and `count` read a finished relation. A body item that is
# not BI_POS forces its source predicate into a strictly earlier stratum.
function Get-Strata {
    param($P)
    $strata = @{}
    foreach ($r in $P.Rules) { $strata[$r.Head.Pred] = 0 }
    for ($pass = 0; $pass -le $P.Rules.Count + 1; $pass++) {
        foreach ($r in $P.Rules) {
            foreach ($b in $r.Body) {
                if ($b.Kind -eq $script:BI_CMP) { continue }
                $src = $b.Atom.Pred
                $srcS = if ($strata.ContainsKey($src)) { $strata[$src] } else { 0 }
                $need = if ($b.Kind -eq $script:BI_POS) { $srcS } else { $srcS + 1 }
                if ($strata[$r.Head.Pred] -lt $need) { $strata[$r.Head.Pred] = $need }
            }
        }
    }
    return $strata
}

function Invoke-Program {
    param($P)

    $relations = @{}
    foreach ($nm in $P.Edges.Keys) {
        $rel = @{}
        foreach ($t in $P.Edges[$nm]) { $rel[(Get-RelKey $t)] = $t }
        $relations[$nm] = $rel
    }
    foreach ($nm in $P.Facts.Keys) {
        if (-not $relations.ContainsKey($nm)) { $relations[$nm] = @{} }
        foreach ($t in $P.Facts[$nm]) { $relations[$nm][(Get-RelKey $t)] = $t }
    }
    foreach ($r in $P.Rules) {
        if (-not $relations.ContainsKey($r.Head.Pred)) { $relations[$r.Head.Pred] = @{} }
    }

    $strata = Get-Strata -P $P
    $levels = ($strata.Values | Sort-Object -Unique)
    foreach ($lvl in $levels) {
        $active = @($P.Rules | Where-Object { $strata[$_.Head.Pred] -eq $lvl })
        $changed = $true
        $guard = 0
        while ($changed) {
            $guard++
            if ($guard -gt 5000) { throw 'runaway fixpoint' }
            $changed = $false
            foreach ($r in $active) {
                $derived = Invoke-Body -Rule $r -Relations $relations
                foreach ($t in $derived) {
                    $k = Get-RelKey $t
                    if (-not $relations[$r.Head.Pred].ContainsKey($k)) {
                        $relations[$r.Head.Pred][$k] = $t
                        $changed = $true
                    }
                }
            }
        }
    }

    $ans = @()
    if ($relations.ContainsKey($P.Query)) {
        $ans = @($relations[$P.Query].Keys | Sort-Object)
    }
    return , $ans
}

# ---------------------------------------------------------------- corpus
#
# Every program states the answer its own data implies, worked out by hand.
# `Verdicts` is what the decision procedure must reach, in order.

function New-Chain {
    param([int]$N)
    $rows = [System.Collections.Generic.List[object]]::new()
    for ($i = 1; $i -le $N; $i++) { [void]$rows.Add([string[]]@("E$($i + 1)", "E$i")) }
    return , $rows
}

function New-ClosureRules {
    param([string]$Any = 'any', [string]$Edge = 'edge', [switch]$LeftRecursive)
    $r = [System.Collections.Generic.List[object]]::new()
    [void]$r.Add((New-Rule (New-Atom $Any @('X', 'Y')) @((New-Item2 $script:BI_POS (New-Atom $Edge @('X', 'Y'))))))
    if ($LeftRecursive) {
        [void]$r.Add((New-Rule (New-Atom $Any @('X', 'Y')) @(
            (New-Item2 $script:BI_POS (New-Atom $Any @('X', 'Z'))),
            (New-Item2 $script:BI_POS (New-Atom $Edge @('Z', 'Y'))))))
    } else {
        [void]$r.Add((New-Rule (New-Atom $Any @('X', 'Y')) @(
            (New-Item2 $script:BI_POS (New-Atom $Edge @('X', 'Z'))),
            (New-Item2 $script:BI_POS (New-Atom $Any @('Z', 'Y'))))))
    }
    return , $r
}

function New-Case {
    param([string]$Name, $Rules, $Edges, [string]$Query, $QueryAtom, [string[]]$Verdicts,
          [string[]]$Answer, $Facts, [switch]$VerdictOnly)
    if ($null -eq $Facts) { $Facts = @{} }
    @{
        Name = $Name
        Program = @{ Rules = $Rules; Facts = $Facts; Edges = $Edges; Query = $Query; QueryAtom = $QueryAtom }
        Verdicts = $Verdicts
        Answer = $Answer
        VerdictOnly = [bool]$VerdictOnly
    }
}

function Get-Corpus {
    $cases = [System.Collections.Generic.List[object]]::new()

    # A four-link chain: E5->E4->E3->E2->E1. Everyone reaches E1.
    $chain = New-Chain -N 4
    $edges = @{ edge = $chain }

    # 1. The ladder's own shape: who reports-to E1 directly or not.
    #    Answer: E2, E3, E4, E5 - every employee reaches E1.
    $r = New-ClosureRules
    [void]$r.Add((New-Rule (New-Atom 'ask' @('W')) @((New-Item2 $script:BI_POS (New-Atom 'any' @('W', '"E1'))))))
    [void]$cases.Add((New-Case -Name 'closure, bound second argument' -Rules $r -Edges $edges `
        -Query 'ask' -QueryAtom $null -Verdicts @('any@2=E1') -Answer @('E2', 'E3', 'E4', 'E5')))

    # 2. The other direction. Right-recursive, so the binding does NOT pass
    #    through position 1 - the recursive atom holds Z there, not X.
    #    Answer: what E5 reaches = E4, E3, E2, E1.
    $r = New-ClosureRules
    [void]$r.Add((New-Rule (New-Atom 'ask' @('W')) @((New-Item2 $script:BI_POS (New-Atom 'any' @('"E5', 'W'))))))
    [void]$cases.Add((New-Case -Name 'closure, bound FIRST argument, right-recursive' -Rules $r -Edges $edges `
        -Query 'ask' -QueryAtom $null -Verdicts @() -Answer @('E1', 'E2', 'E3', 'E4')))

    # 3. The mirror image: LEFT-recursive, bound first argument. Now the
    #    binding does pass straight through, so it specialises.
    $r = New-ClosureRules -LeftRecursive
    [void]$r.Add((New-Rule (New-Atom 'ask' @('W')) @((New-Item2 $script:BI_POS (New-Atom 'any' @('"E5', 'W'))))))
    [void]$cases.Add((New-Case -Name 'closure, bound first argument, LEFT-recursive' -Rules $r -Edges $edges `
        -Query 'ask' -QueryAtom $null -Verdicts @('any@1=E5') -Answer @('E1', 'E2', 'E3', 'E4')))

    # 4. Two readers, same constant: still one narrowing that suits both.
    $r = New-ClosureRules
    [void]$r.Add((New-Rule (New-Atom 'ask' @('W')) @((New-Item2 $script:BI_POS (New-Atom 'any' @('W', '"E1'))))))
    [void]$r.Add((New-Rule (New-Atom 'ask' @('W')) @(
        (New-Item2 $script:BI_POS (New-Atom 'any' @('W', '"E1'))),
        (New-Item2 $script:BI_POS (New-Atom 'edge' @('W', '"E1'))))))
    [void]$cases.Add((New-Case -Name 'two readers, one constant' -Rules $r -Edges $edges `
        -Query 'ask' -QueryAtom $null -Verdicts @('any@2=E1') -Answer @('E2', 'E3', 'E4', 'E5')))

    # 5. Two readers, DIFFERENT constants: no single narrowing serves both.
    #    TWO SEPARATE chains, not one - down a single chain the reachers of
    #    E2 are a subset of the reachers of E1, so narrowing to the wrong one
    #    of the two would still answer correctly and the case would prove
    #    nothing. A3->A2->A1 and B3->B2->B1 make the two answers disjoint.
    $twoChains = [System.Collections.Generic.List[object]]::new()
    [void]$twoChains.Add([string[]]@('A2', 'A1'))
    [void]$twoChains.Add([string[]]@('A3', 'A2'))
    [void]$twoChains.Add([string[]]@('B2', 'B1'))
    [void]$twoChains.Add([string[]]@('B3', 'B2'))
    $edges2 = @{ edge = $twoChains }
    $r = New-ClosureRules
    [void]$r.Add((New-Rule (New-Atom 'ask' @('W')) @((New-Item2 $script:BI_POS (New-Atom 'any' @('W', '"A1'))))))
    [void]$r.Add((New-Rule (New-Atom 'ask' @('W')) @((New-Item2 $script:BI_POS (New-Atom 'any' @('W', '"B1'))))))
    [void]$cases.Add((New-Case -Name 'two readers, two constants' -Rules $r -Edges $edges2 `
        -Query 'ask' -QueryAtom $null -Verdicts @() -Answer @('A2', 'A3', 'B2', 'B3')))

    # 6. A reader that projects both columns: nothing is bound, nothing is
    #    narrowed. Answer: all ten pairs of a four-link chain.
    $r = New-ClosureRules
    [void]$r.Add((New-Rule (New-Atom 'ask' @('A', 'B')) @((New-Item2 $script:BI_POS (New-Atom 'any' @('A', 'B'))))))
    [void]$cases.Add((New-Case -Name 'reader projects both columns' -Rules $r -Edges $edges `
        -Query 'ask' -QueryAtom $null -Verdicts @() `
        -Answer @("E2$([char]31)E1", "E3$([char]31)E1", "E3$([char]31)E2", "E4$([char]31)E1",
                  "E4$([char]31)E2", "E4$([char]31)E3", "E5$([char]31)E1", "E5$([char]31)E2",
                  "E5$([char]31)E3", "E5$([char]31)E4")))

    # 7. The closure IS the query: its whole relation is the answer.
    $r = New-ClosureRules
    [void]$cases.Add((New-Case -Name 'the closure is the query predicate' -Rules $r -Edges $edges `
        -Query 'any' -QueryAtom $null -Verdicts @() `
        -Answer @("E2$([char]31)E1", "E3$([char]31)E1", "E3$([char]31)E2", "E4$([char]31)E1",
                  "E4$([char]31)E2", "E4$([char]31)E3", "E5$([char]31)E1", "E5$([char]31)E2",
                  "E5$([char]31)E3", "E5$([char]31)E4")))

    # 8. A fact already populates the closure predicate.
    $r = New-ClosureRules
    [void]$r.Add((New-Rule (New-Atom 'ask' @('W')) @((New-Item2 $script:BI_POS (New-Atom 'any' @('W', '"E1'))))))
    # A List, not @(...): `@( [string[]]@('Z9','E1') )` does NOT wrap the
    # array, it flattens to the two strings, and the evaluator then read
    # 'Z9' as a one-column tuple and indexed into the STRING for column two.
    # No error, a silently wrong relation, and the only thing that caught it
    # was this case's hand-derived answer.
    $factRows = [System.Collections.Generic.List[object]]::new()
    [void]$factRows.Add([string[]]@('Z9', 'E1'))
    $facts = @{ any = $factRows }
    [void]$cases.Add((New-Case -Name 'a fact already defines the closure' -Rules $r -Edges $edges `
        -Query 'ask' -QueryAtom $null -Verdicts @() -Facts $facts `
        -Answer @('E2', 'E3', 'E4', 'E5', 'Z9')))

    # 9. The reader is a NEGATION. It still pins position 2 to E1, and every
    #    row it keeps is one the narrowed relation would have refused too.
    #    Answer: employees that do NOT reach E1 - none, so `ask` is empty.
    $r = New-ClosureRules
    [void]$r.Add((New-Rule (New-Atom 'ask' @('W')) @(
        (New-Item2 $script:BI_POS (New-Atom 'edge' @('W', 'Q'))),
        (New-Item2 $script:BI_NOT (New-Atom 'any' @('W', '"E1'))))))
    [void]$cases.Add((New-Case -Name 'the reader is a negation' -Rules $r -Edges $edges `
        -Query 'ask' -QueryAtom $null -Verdicts @('any@2=E1') -Answer @()))

    # 10. The reader is a COUNT. Four employees reach E1, so ask holds "4".
    $r = New-ClosureRules
    [void]$r.Add((New-Rule (New-Atom 'ask' @('N')) @(
        (New-Item2 $script:BI_COUNT (New-Atom 'any' @('VlaCounted', '"E1')) 'N'))))
    [void]$cases.Add((New-Case -Name 'the reader is a count' -Rules $r -Edges $edges `
        -Query 'ask' -QueryAtom $null -Verdicts @('any@2=E1') -Answer @('4')))

    # 11. The head variable is also a comparison operand: substituting the
    #     constant's TEXT where a value used to be read is declined.
    $r = [System.Collections.Generic.List[object]]::new()
    [void]$r.Add((New-Rule (New-Atom 'any' @('X', 'Y')) @((New-Item2 $script:BI_POS (New-Atom 'edge' @('X', 'Y'))))))
    # The comparison must GATE THE ANSWER, or the case proves nothing: with
    # (Y = "E2") the recursive step contributes only pairs ending at E2, which
    # the E1 reader never looks at, so narrowing wrongly would still answer
    # right. With (Y = "E1") the recursion is exactly what carries E3, E4 and
    # E5 into the answer. Substituting E1 for Y in the head and the positive
    # atoms leaves Y UNBOUND in the comparison - in the real engine that is a
    # loud VlaDictGet miss, here simply a row that never survives; either way
    # the answer collapses to E2, which is the unsoundness this declines.
    [void]$r.Add((New-Rule (New-Atom 'any' @('X', 'Y')) @(
        (New-Item2 $script:BI_POS (New-Atom 'edge' @('X', 'Z'))),
        (New-Item2 $script:BI_POS (New-Atom 'any' @('Z', 'Y'))),
        (New-Item2 $script:BI_CMP (New-Atom '=' @('Y', '"E1'))))))
    [void]$r.Add((New-Rule (New-Atom 'ask' @('W')) @((New-Item2 $script:BI_POS (New-Atom 'any' @('W', '"E1'))))))
    [void]$cases.Add((New-Case -Name 'the bound variable is a comparison operand' -Rules $r -Edges $edges `
        -Query 'ask' -QueryAtom $null -Verdicts @() -Answer @('E2', 'E3', 'E4', 'E5')))

    # 12. Not recursive at all: nothing to materialise early.
    $r = [System.Collections.Generic.List[object]]::new()
    [void]$r.Add((New-Rule (New-Atom 'any' @('X', 'Y')) @((New-Item2 $script:BI_POS (New-Atom 'edge' @('X', 'Y'))))))
    [void]$r.Add((New-Rule (New-Atom 'ask' @('W')) @((New-Item2 $script:BI_POS (New-Atom 'any' @('W', '"E1'))))))
    [void]$cases.Add((New-Case -Name 'the predicate is not recursive' -Rules $r -Edges $edges `
        -Query 'ask' -QueryAtom $null -Verdicts @() -Answer @('E2')))

    # 13. A repeated variable inside the closure's own atom must keep
    #     filtering after the rewrite. edge holds no self-loop, so nothing
    #     reaches the head and ask is empty - the point is that the answer
    #     is the SAME with and without the rewrite.
    $r = New-ClosureRules
    [void]$r.Add((New-Rule (New-Atom 'ask' @('W')) @(
        (New-Item2 $script:BI_POS (New-Atom 'any' @('W', '"E1'))),
        (New-Item2 $script:BI_POS (New-Atom 'edge' @('W', 'W'))))))
    [void]$cases.Add((New-Case -Name 'a repeated variable still filters' -Rules $r -Edges $edges `
        -Query 'ask' -QueryAtom $null -Verdicts @('any@2=E1') -Answer @()))

    # 14. The closure is the query AND another rule reads it with a constant.
    #     Without condition 1 the query's own relation would be narrowed to
    #     the four pairs ending at E1 - the cell would show six rows fewer.
    $r = New-ClosureRules
    [void]$r.Add((New-Rule (New-Atom 'side' @('W')) @((New-Item2 $script:BI_POS (New-Atom 'any' @('W', '"E1'))))))
    [void]$cases.Add((New-Case -Name 'the query predicate, with a pinned reader' -Rules $r -Edges $edges `
        -Query 'any' -QueryAtom $null -Verdicts @() `
        -Answer @("E2$([char]31)E1", "E3$([char]31)E1", "E3$([char]31)E2", "E4$([char]31)E1",
                  "E4$([char]31)E2", "E4$([char]31)E3", "E5$([char]31)E1", "E5$([char]31)E2",
                  "E5$([char]31)E3", "E5$([char]31)E4")))

    # 15. One reader pins position 2, another leaves it free. The free one is
    #     what makes the narrowing visible, so the rewrite must decline.
    $r = New-ClosureRules
    [void]$r.Add((New-Rule (New-Atom 'ask' @('A', 'B')) @((New-Item2 $script:BI_POS (New-Atom 'any' @('A', 'B'))))))
    [void]$r.Add((New-Rule (New-Atom 'flag' @('W')) @((New-Item2 $script:BI_POS (New-Atom 'any' @('W', '"E1'))))))
    [void]$cases.Add((New-Case -Name 'one reader pins, one reader does not' -Rules $r -Edges $edges `
        -Query 'ask' -QueryAtom $null -Verdicts @() `
        -Answer @("E2$([char]31)E1", "E3$([char]31)E1", "E3$([char]31)E2", "E4$([char]31)E1",
                  "E4$([char]31)E2", "E4$([char]31)E3", "E5$([char]31)E1", "E5$([char]31)E2",
                  "E5$([char]31)E3", "E5$([char]31)E4")))

    return , $cases
}

# ------------------------------------------------------------------ run

function Invoke-Corpus {
    $cases = Get-Corpus
    $bad = [System.Collections.Generic.List[string]]::new()
    $ctl = 0
    $rows = [System.Collections.Generic.List[object]]::new()
    foreach ($c in $cases) {
        $plain = Copy-Program $c.Program
        $savedMut = $script:Mut
        $script:Mut = 'none'
        $before = Invoke-Program -P $plain
        $script:Mut = $savedMut

        $want = @($c.Answer)
        $controlOk = (@(Compare-Object -ReferenceObject $want -DifferenceObject @($before) -SyncWindow 999).Count -eq 0)
        if (-not $controlOk) {
            $ctl++
            [void]$bad.Add("$($c.Name): CONTROL - hand-derived answer [$($want -join ', ')] but the evaluator said [$(@($before) -join ', ')]")
        }

        $rw = Copy-Program $c.Program
        $verdicts = Invoke-PushBoundArguments -P $rw
        $gotV = @($verdicts) -join ','
        $wantV = @($c.Verdicts) -join ','
        $verdictOk = ($gotV -ceq $wantV)
        if (-not $verdictOk) { [void]$bad.Add("$($c.Name): verdict - expected [$wantV] but reached [$gotV]") }

        $moved = $false
        try {
            $after = Invoke-Program -P $rw
            $moved = -not (@(Compare-Object -ReferenceObject @($before) -DifferenceObject @($after) -SyncWindow 999).Count -eq 0)
            if ($moved) { [void]$bad.Add("$($c.Name): ANSWER MOVED - [$(@($before) -join ', ')] became [$(@($after) -join ', ')]") }
        } catch {
            $moved = $true
            [void]$bad.Add("$($c.Name): ANSWER MOVED - the rewritten program could not be evaluated: $($_.Exception.Message)")
        }
        [void]$rows.Add(@{ Name = $c.Name; ControlOk = $controlOk; Verdict = $gotV; VerdictOk = $verdictOk; Moved = $moved })
    }
    return , @{ Rows = $rows; Bad = $bad; ControlFails = $ctl; Count = $cases.Count }
}

if ($Mutations) {
    Write-Output 'DATALOG.14 - mutation control over the specialisation''s conditions'
    Write-Output ''
    $muts = @(
        @{ Name = 'no-query-guard'; What = '1: the closure may be the query predicate' },
        @{ Name = 'same-constant';  What = '4: two readers may pin different constants' },
        @{ Name = 'variable-pins';  What = '4: a reader may leave the position free' },
        @{ Name = 'passthrough';    What = '5: the recursive atom may name another variable' },
        @{ Name = 'value-use';      What = '5: the bound variable may be read as a value' }
    )
    Write-Output 'mutant           condition switched off                                 caught by'
    Write-Output '---------------  ----------------------------------------------------  ---------'
    $survivors = 0
    foreach ($m in $muts) {
        $script:Mut = $m.Name
        $res = Invoke-Corpus
        $moved = @($res.Rows | Where-Object { $_.Moved })
        $wrongV = @($res.Rows | Where-Object { -not $_.VerdictOk })
        $caught = if ($moved.Count -gt 0) { "$($moved.Count) answer(s) moved" }
                  elseif ($wrongV.Count -gt 0) { "$($wrongV.Count) verdict(s) only" }
                  else { 'NOTHING' }
        if ($moved.Count -eq 0) { $survivors++ }
        '{0,-15}  {1,-52}  {2}' -f $m.Name, $m.What, $caught | Write-Output
    }
    $script:Mut = 'none'
    Write-Output ''
    if ($survivors -eq 0) {
        Write-Output 'MUTATIONS GREEN: every load-bearing condition has a mutant that moves an answer.'
        Write-Output ''
        Write-Output 'Conditions 2 (no fact defines it) and 3 (directly self-recursive) have no mutant here,'
        Write-Output 'and that is recorded rather than hidden: neither is load-bearing. Dropping either one'
        Write-Output 'moves no answer in this corpus - they decline cases that would be safe - and they are'
        Write-Output 'kept because each removes a class of reasoning rather than a class of bugs.'
        exit 0
    } else {
        Write-Output "MUTATIONS RED: $survivors mutant(s) moved no answer. The corpus does not test what it claims to."
        exit 1
    }
}

$corpus = Get-Corpus
$fails = [System.Collections.Generic.List[string]]::new()
$controlFails = 0

Write-Output 'DATALOG.14 - the specialisation, proven before import'
Write-Output ''
Write-Output 'case                                          control  verdict            rewritten  identical'
Write-Output '--------------------------------------------  -------  -----------------  ---------  ---------'

foreach ($c in $corpus) {
    $plain = Copy-Program $c.Program
    $before = Invoke-Program -P $plain

    $want = @($c.Answer)
    $controlOk = (@(Compare-Object -ReferenceObject $want -DifferenceObject @($before) -SyncWindow 999).Count -eq 0)
    if (-not $controlOk) {
        $controlFails++
        [void]$fails.Add("$($c.Name): CONTROL - hand-derived answer [$($want -join ', ')] but the evaluator said [$(@($before) -join ', ')]")
    }

    $rw = Copy-Program $c.Program
    $verdicts = Invoke-PushBoundArguments -P $rw
    $gotV = @($verdicts) -join ','
    $wantV = @($c.Verdicts) -join ','
    $verdictOk = ($gotV -ceq $wantV)
    if (-not $verdictOk) {
        [void]$fails.Add("$($c.Name): verdict - expected [$wantV] but reached [$gotV]")
    }

    $after = Invoke-Program -P $rw
    $sameOk = (@(Compare-Object -ReferenceObject @($before) -DifferenceObject @($after) -SyncWindow 999).Count -eq 0)
    if (-not $sameOk) {
        [void]$fails.Add("$($c.Name): ANSWER MOVED - [$(@($before) -join ', ')] became [$(@($after) -join ', ')]")
    }

    '{0,-44}  {1,-7}  {2,-17}  {3,-9}  {4,-9}' -f `
        $c.Name,
        $(if ($controlOk) { 'ok' } else { 'FAIL' }),
        $(if ($verdictOk) { $(if ($gotV) { $gotV } else { 'declined' }) } else { "FAIL:$gotV" }),
        $(if ($gotV) { 'yes' } else { 'no' }),
        $(if ($sameOk) { 'yes' } else { 'NO' }) | Write-Output
}

Write-Output ''
if ($fails.Count -eq 0) {
    Write-Output "PROOF GREEN: $($corpus.Count) programs. Every hand-derived answer reproduced, every verdict as claimed,"
    Write-Output '  and every specialised program answers exactly what the unspecialised one answered.'
    exit 0
} else {
    Write-Output "PROOF RED: $($fails.Count) failure(s)."
    foreach ($f in $fails) { Write-Output "  - $f" }
    if ($controlFails -gt 0) {
        Write-Output ''
        Write-Output '  A CONTROL failure means the harness is wrong about the semantics. Fix it before'
        Write-Output '  reading anything else here: a harness that agrees with itself proves nothing.'
    }
    exit 1
}
