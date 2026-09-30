<#
=== OPTIMIZE.1 - the parity claim is EVERY program, and this is what
    makes "every" true ===

WHY THIS EXISTS. OPTIMIZE.1's whole claim is that it is DATALOG's own
engine wearing the OPTIMIZE name. The owner approved one proof for it
(2026-09-18): every DATALOG test program in VLA_Tests_Query.bas run
through OPTIMIZE and required to answer identically, so that "DATALOG
wearing the name" is a measured claim and not a description.

A VBA test cannot read its own module's source without VBIDE, and this
tranche's tests are pure. So the parity suite holds a TABLE of programs
(DatalogParityPrograms, in VLA_Tests_Query.bas), and a table is exactly
the shape that silently stops being complete: somebody adds a DATALOG
test six months from now, the parity suite still passes, and "every
program" has quietly become "every program as of OPTIMIZE.1". That is
the same defect tools/check_devrig_mods_parity.ps1 was written for, in
its own words: "two independent things that must agree, with nothing
mechanical holding them together, is the actual defect; the repeated
comments are a workaround for it, not a fix."

WHAT IS CHECKED

  1. Every DATALOG program literal in VLA_Tests_Query.bas - from every
     VLA_Datalog.DatalogRun(...) and VLA_Datalog.DATALOG(...) call site,
     with VBA line continuations joined and "" un-doubled - appears in
     the parity table. A missing one FAILS.

  2. The parity table holds no program that no call site has. Reported,
     not failed: a DATALOG test deleted later leaves a stale entry that
     is still a perfectly good parity case, so this is information
     rather than a defect.

WHAT THE PARITY PIN ITSELF COMPARES, for a reader of this file: each
program is run through both engines with no table arguments, and the two
must agree completely - the same queried relation with the same tuples
in the same order, the same header names, the same headless flag, the
same Boolean answer; or both must refuse, with the same error number and
the same words. A program whose tables were supplied by its original
test refuses in both engines, identically, because its predicates are
undefined - which is still parity, and is why this check needs no
per-program fixture. (METAPROOF.4: a proof in the corpus that carries its
tables in a (tables ...) clause is asked of both engines WITH them, since
the corpus holds the fixture beside the program; only the table's
programs run bare.)

  -Emit prints the table as VBA source, for pasting into
  DatalogParityPrograms when the list grows. It is a convenience for a
  person, not a build step: nothing generates src/.

METAPROOF.1 - THE PROOF CORPUS, AND A CEILING. DATALOG tests can be proofs
now: forms in scripts/proofs/datalog.vla, each program beside the answer it
must give, run by TestDatalogProofs. Both parity loops read that file's
programs directly (DatalogProofPrograms), so a proof needs no table line -
the table covers the VBA call sites and nothing else, and a test that moves
to the corpus takes its line out of the table (192 -> 177 at METAPROOF.1).
The direction is held by one more rule:

  3. The DATALOG call sites in VLA_Tests_Query.bas - every
     VLA_Datalog.DatalogRun/DATALOG call carrying a literal program - may
     not number more than the held ceiling below. A new DATALOG test
     belongs in the corpus. One that truly needs VBA - a live Table, a
     program built in a loop, an engine internal - raises the ceiling, with
     its reason, in the same commit. Fewer than the ceiling is reported,
     not failed, as room to lower it (check_raise_ratchet.ps1's shape).

House shape, per the standing convention for this project's static
scans: PowerShell, host-independent, hardcoded and reviewable baseline,
never wired into VlaSelfTest.

Usage:  powershell -File tools\check_optimize_parity.ps1 [-Emit]
Exit 0 clean, exit 1 with every failure listed.
#>
param([switch]$Emit)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$failures = New-Object System.Collections.Generic.List[string]

# --- the held ceiling on DATALOG call sites (rule 3, header) ------------
# 2026-09-27, METAPROOF.1: 188 - the 204 at 0.7.0 less the sixteen tests
# that moved to scripts/proofs/datalog.vla (eleven of
# TestDatalogBoundArgument, five of TestDatalog's refusals).
# 2026-09-28, METAPROOF.3: 109 - seventy-nine more call sites moved, as
# seventy-eight proofs (TestDatalog's (edge X X) pin was already one).
# 2026-09-30, METAPROOF.4: 84 - twenty-five call sites that passed a table
# argument moved, as proofs that carry their tables in a (tables ...) clause.
$callSiteCeiling = 84
$testsPath = Join-Path $root 'src\VLA_Tests_Query.bas'
if (-not (Test-Path -LiteralPath $testsPath)) {
    Write-Error "Missing $testsPath"
}
$raw = [System.IO.File]::ReadAllText($testsPath)
$lines = $raw -split "`r`n"

# --- VBA line continuations joined into one logical line --------------
$logical = New-Object System.Collections.Generic.List[string]
$buf = ''
foreach ($ln in $lines) {
    if ($ln -match '\s_\s*$') {
        $buf = $buf + ($ln -replace '\s_\s*$', '')
    } else {
        $logical.Add($buf + $ln)
        $buf = ''
    }
}
if ($buf.Length -gt 0) { $logical.Add($buf) }

# --- the concatenated string literals of a call's FIRST argument ------
# Stops at the first top-level comma or close paren, so a bases-dict or
# table argument after the program is not swallowed. "" inside a VBA
# literal is one quote character and is un-doubled here, which is what
# makes the extracted text comparable to the table's own.
function Get-LeadingLiterals([string]$rest) {
    $parts = New-Object System.Collections.Generic.List[string]
    $i = 0
    $depth = 0
    while ($i -lt $rest.Length) {
        $ch = $rest[$i]
        if ($ch -eq '"') {
            $j = $i + 1
            $lit = ''
            while ($j -lt $rest.Length) {
                if ($rest[$j] -eq '"') {
                    if ($j + 1 -lt $rest.Length -and $rest[$j + 1] -eq '"') { $lit += '"'; $j += 2; continue }
                    break
                }
                $lit += $rest[$j]
                $j++
            }
            $parts.Add($lit)
            $i = $j + 1
            continue
        }
        if ($ch -eq '(') { $depth++ }
        elseif ($ch -eq ')') { if ($depth -eq 0) { break }; $depth-- }
        elseif ($ch -eq ',' -and $depth -eq 0) { break }
        $i++
    }
    return ($parts -join '')
}

# --- 1. every program a DATALOG test names ----------------------------
$callPrograms = New-Object System.Collections.Generic.List[string]
foreach ($ln in $logical) {
    foreach ($mt in [regex]::Matches($ln, 'VLA_Datalog\.(DatalogRun|DATALOG)\s*\(?')) {
        $prog = Get-LeadingLiterals $ln.Substring($mt.Index + $mt.Length)
        if ($prog.Length -gt 0) { $callPrograms.Add($prog) }
    }
}
$wanted = @($callPrograms | Sort-Object -Unique)

# --- 2. the parity table, as the VBA source holds it ------------------
$tableFn = [regex]::Match($raw, '(?ms)^Private Function DatalogParityPrograms\(\).*?^End Function')
$have = @()
if (-not $tableFn.Success) {
    if (-not $Emit) {
        $failures.Add('VLA_Tests_Query.bas no longer defines DatalogParityPrograms - this check cannot verify a table it cannot find. If it was renamed, rename it here too.')
    }
} else {
    $tableLines = $tableFn.Value -split "`r`n"
    $acc = New-Object System.Collections.Generic.List[string]
    $tbuf = ''
    foreach ($ln in $tableLines) {
        if ($ln -match '\s_\s*$') { $tbuf = $tbuf + ($ln -replace '\s_\s*$', ''); continue }
        $acc.Add($tbuf + $ln)
        $tbuf = ''
    }
    $entries = New-Object System.Collections.Generic.List[string]
    foreach ($ln in $acc) {
        $mt = [regex]::Match($ln, '^\s*p\.Add\s+')
        if (-not $mt.Success) { continue }
        $prog = Get-LeadingLiterals $ln.Substring($mt.Length)
        if ($prog.Length -gt 0) { $entries.Add($prog) }
    }
    $have = @($entries | Sort-Object -Unique)
}

if ($Emit) {
    Write-Output "    ' Generated by tools\check_optimize_parity.ps1 -Emit, then reviewed."
    Write-Output ("    ' " + $wanted.Count + " distinct DATALOG programs, every one this module names.")
    foreach ($p in $wanted) {
        Write-Output ('    p.Add "' + $p.Replace('"', '""') + '"')
    }
    exit 0
}

Write-Output '=== OPTIMIZE.1 PARITY COVERAGE (VLA_Tests_Query.bas) ==='
Write-Output ("  DATALOG programs named by a call site : {0}" -f $wanted.Count)
Write-Output ("  programs in the parity table          : {0}" -f $have.Count)

$missing = @($wanted | Where-Object { $have -notcontains $_ })
$extra = @($have | Where-Object { $wanted -notcontains $_ })

Write-Output ''
if ($missing.Count -eq 0) {
    Write-Output '  ok    every DATALOG program in the module is in the parity table'
} else {
    foreach ($mp in $missing) {
        $show = $mp
        if ($show.Length -gt 96) { $show = $show.Substring(0, 96) + '...' }
        Write-Output ("  FAIL  not in the parity table: {0}" -f $show)
    }
    $failures.Add("$($missing.Count) DATALOG program(s) in VLA_Tests_Query.bas are absent from DatalogParityPrograms, so the parity pin no longer covers every program. Run this script with -Emit and paste the table it prints.")
}

if ($extra.Count -gt 0) {
    Write-Output ''
    Write-Output ("  note  {0} table entr(ies) match no current call site - a DATALOG test was changed or removed. Harmless: each is still a parity case." -f $extra.Count)
}

# --- 3. the ceiling on DATALOG call sites (METAPROOF.1) ------------------
Write-Output ''
Write-Output ("  DATALOG call sites in the module    : {0} (ceiling {1})" -f $callPrograms.Count, $callSiteCeiling)
if ($callPrograms.Count -gt $callSiteCeiling) {
    $failures.Add("VLA_Tests_Query.bas now holds $($callPrograms.Count) DATALOG call sites, above the held ceiling of $callSiteCeiling. A new DATALOG test belongs in scripts/proofs/datalog.vla; if this one truly needs VBA (a live Table, a program built in a loop, an engine internal), raise the ceiling in this script with the reason.")
} elseif ($callPrograms.Count -lt $callSiteCeiling) {
    Write-Output ("  note  {0} below the ceiling - lower it to {1} in this script, so the room cannot be spent on a new VBA test unnoticed." -f ($callSiteCeiling - $callPrograms.Count), $callPrograms.Count)
} else {
    Write-Output '  ok    at the ceiling'
}

Write-Output ''
if ($failures.Count -eq 0) {
    Write-Output '=== CHECK: clean - the parity pin covers every DATALOG program in the module ==='
    exit 0
} else {
    foreach ($f in $failures) { Write-Output "FAIL: $f" }
    Write-Output ''
    Write-Output "=== CHECK: FAILED - $($failures.Count) problem(s) ==="
    exit 1
}
