<#
check_test_assertion_safety.ps1 - PROLOG.20's mechanical check.

WHY THIS EXISTS. VBA's `And` does not short-circuit. Every operand of a
combined expression is evaluated however the earlier ones answered, so an
assertion written as one expression can RAISE at the exact moment it was
about to report a failure:

    ok = (UBound(result, 1) = 3 And CStr(result(2, 1)) = "a" _
                                And CStr(result(3, 1)) = "b")

While that test PASSES it is harmless. The moment it genuinely FAILS -
the query returns one solution instead of two, so UBound is 2 -
`result(3, 1)` is evaluated anyway and raises "Subscript out of range".
The run dies where it was about to say what broke, and every remaining
test in the module never runs. An assertion that cannot survive its own
failure is not a test.

THIS DEFECT IS INVISIBLE IN A GREEN RUN, BY CONSTRUCTION. That is why it
needs a static check rather than a test: nothing about a passing suite
can reveal it, and it only ever announces itself on the day something
else has already gone wrong - the worst possible day to lose the report.
PROLOG.8's own ResultCol1Is header records it happening live.

THE RULE, one sentence: IF AN OPERAND GUARDS A VARIABLE, NO LATER
OPERAND MAY TOUCH THAT VARIABLE EXCEPT THROUGH ANOTHER GUARD OR A
GUARDED HELPER.

A guard is IsArray(V), VarType(V), UBound(V, ..), LBound(V, ..) or
V.Count - an operand asking what shape V has. The moment a statement asks
that question, it has admitted V might be the other shape; every later
mention of V then runs on the shape it just admitted to. The observed
failures, all four from this repo:

    IsArray(one) And one(1, 1) = 42            Type mismatch on a scalar
    VarType(r) = vbBoolean And r = True        Type mismatch on an array
    VarType(a) = vbString And CStr(a) = ""     Type mismatch on an array
    UBound(r, 1) = 3 And CStr(r(3, 1)) = "b"   Subscript out of range
    paths.Count = 2 And paths.Item(2) = "x"    Subscript out of range

Stating it as one rule rather than as a list of shapes is deliberate: the
first version of this check paired each guard with the specific risk it
had been seen with, and missed two real instances - a CStr() of a
VarType-guarded variable - because CStr was not on the risk list. A rule
about what may be TOUCHED has no such list to be incomplete.

THE FIX IN EVERY CASE is a guarded helper that returns rather than
raises - ResultRowCount, ResultColCount, ResultCellIs, ResultCol1Is,
ResultBoolIs in VLA_Tests_Query.bas - each of which answers False or -1
for a shape it cannot read. Those helpers were introduced to fix exactly
this and then never migrated onto; this check is what makes the migration
stick.

A RATCHET, NOT A CLIFF, and for the same reason check_raise_ratchet.ps1
is one: the migration lands module by module, and a check that can only
say "all clean or all broken" would have to stay red across the whole of
it, which means nobody can tell a new defect from the backlog. Each
module holds a CEILING. Exceeding it fails. Coming in under it also fails
- with instructions to lower the ceiling - so the count can only ever go
down.

WHAT THIS SCAN DOES NOT SEE, stated so its count is not read as a
ceiling on the defect itself. It only fires where the statement guards
the variable somewhere. An assertion that indexes a container with NO
guard in the statement at all is invisible to it:

    okAll = (nRowsAll = 4) And (CStr(arrAll(r0, c0)) = "Name")

nRowsAll was computed on an earlier line, so the bound really is checked
- just not somewhere this scan can connect to the index.

PROLOG.20 measured EIGHT of them and deliberately did NOT fix them. They
are named here by VARIABLE, not by line number, because a line number in
a comment goes stale on the next edit and this file has already shifted
by 43 lines once:

  VLA_Tests_Query.bas   arrNarrow, arrAll, arrJoin, arrGroup, arrOrder,
                        arrCompound, arrChain   (the SQL spill tests)
  VLA_Tests_Host.bas    arr                     (the slab read)

Not fixed because telling a genuine one from a safe loop index over an
Array() literal needs dataflow, not a regex - and fixing what this check
cannot hold is precisely how the backlog it exists to drain was created
in the first place. Filed as its own follow-up.

The module list and ceilings below are the reviewable baseline -
deliberately NAMES and NUMBERS, not a glob: a new test module must be
added on purpose, and a rename breaks the run loudly rather than
silently scanning nothing.

House style (tools/check_*.ps1): PowerShell 5.1, host-independent, a
hard-coded reviewable baseline, no live Excel needed. Never wired into
VlaSelfTest - it reads source text, not runtime behaviour.

Usage:  powershell -File tools\check_test_assertion_safety.ps1
        powershell -File tools\check_test_assertion_safety.ps1 -List
Exit code: 0 clean, 1 if any module is off its held ceiling either way.
#>
param([switch]$List)

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot

# ---- baseline: the test modules and the count each is allowed ----------
# The defect is not PROLOG's - it is a property of VBA's And - so all four
# test modules are governed, not just the one where it was found.
#
# PROLOG.20 STAGE 1 took VLA_Tests_Query.bas from 146 to 0. The other
# three hold at what they measure today, because their fixes need helpers
# of their own in each module and land as stage 2 - and because a check
# stuck red across a staged migration cannot tell a NEW defect from the
# backlog, which is the whole point of ratcheting rather than gating.
# Every number here is a debt that may only shrink: the check fails just
# as loudly if a module comes in UNDER its ceiling without the ceiling
# being lowered in the same diff.
$ceilings = [ordered]@{
    'src\VLA_Tests.bas'         = 4
    'src\VLA_Tests_Grammar.bas' = 1
    'src\VLA_Tests_Host.bas'    = 3
    'src\VLA_Tests_Query.bas'   = 0
}

# ---- baseline: accepted mentions, with the reason each is safe ---------
# A guarded helper answers False or -1 for a shape it cannot read, so
# mentioning a guarded variable inside one is exactly the fix this check
# asks for and must not itself be reported.
$guardedHelpers = @(
    'ResultRowCount', 'ResultColCount', 'ResultCellIs', 'ResultCol1Is',
    'ResultBoolIs', 'ResultDescribe', 'JoinColumn',
    'ResultTextIs', 'ResultTextStartsWith', 'CollItemIs'
)

# An operand that asks what shape V is. Touching V after one of these has
# admitted V might be the other shape.
#
# GUARDS COME IN TWO STRENGTHS, and conflating them is a hole this check
# had for one revision. `Exists` asks whether V is an array AT ALL, so
# after it NOTHING may touch V - not even UBound, because
# `IsArray(a) And UBound(a) < LBound(a)` raises on the very scalar
# IsArray just answered False about. `Extent` asks how BIG V is, which
# would itself have raised in that operand had V not been an array, so a
# later UBound/LBound is genuinely safe - `UBound(a,1) = 3 And
# UBound(a,2) = 2` is fine - while indexing V is not. Same for a
# Collection: after `.Count`, another `.Count` is safe and `.Item` is not.
#
# THE GUARDED HELPERS ARE THEMSELVES GUARDS, and that closes a hole the
# migration itself would otherwise open. Rewriting
#
#   UBound(r,1) = 2 And CStr(r(2,1)) = "a",  "got: " & r(2, 1)
#
# into ResultRowCount/ResultCellIs makes the ASSERTION safe and leaves the
# DETAIL argument raw-indexing - and a Report's detail is evaluated on
# every call, pass or fail, so the run still dies on exactly the failing
# case. Without this entry the statement would stop being reported the
# moment it was half-fixed, which is worse than never having flagged it.
# Calling a helper on V is itself an admission that V's shape is
# uncertain, so it guards like IsArray does.
$guardKinds = @(
    @{ Kind = 'Exists'; Form = '\bIsArray\s*\(\s*{0}\s*\)' },
    @{ Kind = 'Exists'; Form = '\bVarType\s*\(\s*{0}\s*\)' },
    @{ Kind = 'Exists'; Form = '\bResult(?:RowCount|ColCount|CellIs|Col1Is|BoolIs|TextIs|TextStartsWith)\s*\(\s*{0}\s*[,)]' },
    @{ Kind = 'Extent'; Form = '\b(?:UBound|LBound)\s*\(\s*{0}\s*[,)]' },
    @{ Kind = 'Count';  Form = '\b{0}\s*\.\s*Count\b' }
)
# What a LATER operand may still do to V, keyed on the guard it followed.
$allowedAfter = @{
    'Exists' = @()
    'Extent' = @('\b(?:UBound|LBound)\s*\(\s*{0}\s*[,)]')
    'Count'  = @('\b{0}\s*\.\s*Count\b')
}

$failures = New-Object System.Collections.Generic.List[string]
$findings = New-Object System.Collections.ArrayList

# Strip a trailing comment, respecting string literals. VBA escapes a
# quote by doubling it, which this handles naturally: the second quote of
# a doubled pair simply toggles back in.
function Strip-Comment {
    param([string]$line)
    $inStr = $false
    for ($i = 0; $i -lt $line.Length; $i++) {
        $c = $line[$i]
        if ($c -eq '"') { $inStr = -not $inStr; continue }
        if ($c -eq "'" -and -not $inStr) { return $line.Substring(0, $i) }
    }
    return $line
}

# Blank out the CONTENTS of every string literal, keeping the quotes and
# the length. Without this a Report message like "...(atom? X)..." would
# be scanned for code, and an expected literal containing "result(" would
# be read as an array index.
function Blank-Strings {
    param([string]$s)
    $sb = New-Object System.Text.StringBuilder
    $inStr = $false
    foreach ($c in $s.ToCharArray()) {
        if ($c -eq '"') { $inStr = -not $inStr; [void]$sb.Append($c); continue }
        if ($inStr) { [void]$sb.Append(' ') } else { [void]$sb.Append($c) }
    }
    return $sb.ToString()
}

foreach ($rel in $ceilings.Keys) {
    $path = Join-Path $repoRoot $rel
    if (-not (Test-Path -LiteralPath $path)) {
        Write-Output "FAIL: missing test module $rel (renamed or removed - update this baseline deliberately)"
        exit 1
    }
    $raw = Get-Content -LiteralPath $path

    # ---- join continuations into the statements VBA actually executes
    $stmts = New-Object System.Collections.ArrayList
    $acc = ''; $accLine = 0
    for ($i = 0; $i -lt $raw.Count; $i++) {
        $s = Strip-Comment $raw[$i]
        if ($acc -eq '') { $accLine = $i + 1 }
        if ($s -match '\s_\s*$') { $acc += ($s -replace '\s_\s*$', ' '); continue }
        $acc += $s
        if ($acc.Trim().Length -gt 0) { [void]$stmts.Add([pscustomobject]@{ Line = $accLine; Text = $acc.Trim() }) }
        $acc = ''
    }
    if ($acc.Trim().Length -gt 0) { [void]$stmts.Add([pscustomobject]@{ Line = $accLine; Text = $acc.Trim() }) }

    foreach ($st in $stmts) {
        $code = Blank-Strings $st.Text
        if ($code -notmatch '\bAnd\b') { continue }

        # Every variable this statement asks the shape of.
        $vars = @([regex]::Matches($code, '\b(?:IsArray|UBound|LBound|VarType)\s*\(\s*([A-Za-z_]\w*)') |
                  ForEach-Object { $_.Groups[1].Value }) +
                @([regex]::Matches($code, '\bResult(?:RowCount|ColCount|CellIs|Col1Is|BoolIs|TextIs|TextStartsWith)\s*\(\s*([A-Za-z_]\w*)') |
                  ForEach-Object { $_.Groups[1].Value }) +
                @([regex]::Matches($code, '\b([A-Za-z_]\w*)\s*\.\s*Count\b') |
                  ForEach-Object { $_.Groups[1].Value })
        $vars = @($vars | Sort-Object -Unique)
        if ($vars.Count -eq 0) { continue }

        # Operands in source order. Splitting on the bare keyword is sound
        # for assertion expressions: no VBA identifier here contains "And"
        # as a whole word.
        $ops = [regex]::Split($code, '\bAnd\b')

        foreach ($v in $vars) {
            $esc = [regex]::Escape($v)

            # Where the statement first asks V's shape, and how strongly.
            # The WEAKEST guard wins when several appear in one operand:
            # asking IsArray at all admits V might be a scalar, and that
            # admission is not undone by also asking UBound.
            $guardAt = -1; $guardKind = ''
            for ($k = 0; $k -lt $ops.Count -and $guardAt -lt 0; $k++) {
                foreach ($g in $guardKinds) {
                    if ($ops[$k] -match [string]::Format($g.Form, $esc)) {
                        $guardAt = $k
                        if ($g.Kind -eq 'Exists') { $guardKind = 'Exists'; break }
                        if ($guardKind -eq '') { $guardKind = $g.Kind }
                    }
                }
            }
            if ($guardAt -lt 0) { continue }

            for ($m = $guardAt + 1; $m -lt $ops.Count; $m++) {
                $op = $ops[$m]
                if ($op -notmatch ('\b' + $esc + '\b')) { continue }

                # Blank out every mention this guard still permits...
                $onlySafe = $op
                foreach ($a in $allowedAfter[$guardKind]) {
                    $onlySafe = [regex]::Replace($onlySafe, [string]::Format($a, $esc), ' ')
                }
                # ...and every mention inside a guarded helper, which
                # answers False or -1 rather than raising whatever V is.
                foreach ($h in $guardedHelpers) {
                    $onlySafe = [regex]::Replace($onlySafe, '\b' + [regex]::Escape($h) + '\s*\([^)]*\)', ' ')
                }
                if ($onlySafe -match ('\b' + $esc + '\b')) {
                    [void]$findings.Add([pscustomobject]@{
                        Module = $rel; Line = $st.Line; Var = $v; Text = $st.Text })
                    break
                }
            }
        }
    }
}

# One statement can offend for more than one variable; report it once, so
# the count is a count of ASSERTIONS rather than of pattern hits.
$unique = @($findings | Group-Object { "$($_.Module)|$($_.Line)" } | ForEach-Object { $_.Group[0] })

Write-Output '=== PROLOG.20 - no assertion may raise on the case it exists to catch ==='
Write-Output ''
Write-Output '--- module, held ceiling, actual ---'
foreach ($rel in $ceilings.Keys) {
    $n   = @($unique | Where-Object { $_.Module -eq $rel }).Count
    $cap = $ceilings[$rel]
    if ($n -gt $cap) {
        Write-Output ("  {0,-30} ceiling {1,3}   actual {2,3}   OVER" -f $rel, $cap, $n)
        $failures.Add("$rel has $n unsafe assertion(s) but its held ceiling is $cap - migrate them onto the guarded helpers")
    } elseif ($n -lt $cap) {
        Write-Output ("  {0,-30} ceiling {1,3}   actual {2,3}   LOWER THE CEILING" -f $rel, $cap, $n)
        $failures.Add("$rel is at $n, below its held ceiling of $cap - lower the ceiling in this script to $n so the gain cannot be silently given back")
    } else {
        Write-Output ("  {0,-30} ceiling {1,3}   actual {2,3}   ok" -f $rel, $cap, $n)
    }
}

if ($List -and $unique.Count -gt 0) {
    Write-Output ''
    Write-Output '--- every unsafe assertion ---'
    foreach ($f in ($unique | Sort-Object Module, Line)) {
        $t = $f.Text
        if ($t.Length -gt 150) { $t = $t.Substring(0, 150) + '...' }
        Write-Output ("  {0}:{1}  [{2}]" -f $f.Module, $f.Line, $f.Var)
        Write-Output ("      {0}" -f $t)
    }
}

Write-Output ''
Write-Output ("  TOTAL unsafe assertions: {0}" -f $unique.Count)
Write-Output ''
if ($failures.Count -eq 0) {
    Write-Output '=== CHECK: clean - every combined assertion survives its own failure ==='
} else {
    Write-Output "=== CHECK: $($failures.Count) module(s) off their held ceiling ==="
    $failures | ForEach-Object { Write-Output "  $_" }
    Write-Output ''
    Write-Output 'Re-run with -List to see each assertion.'
    Write-Output ''
    Write-Output 'VBA evaluates every operand of an And however the earlier ones answered, so once a statement asks what shape a variable is, every later mention of it runs on the shape it just admitted to. Move the assertion onto the guarded helpers - ResultRowCount, ResultColCount, ResultCellIs, ResultCol1Is, ResultBoolIs - which answer False or -1 for a shape they cannot read instead of raising. Do not widen this check to accommodate an assertion; the assertion is what is wrong.'
}
exit ([Math]::Min($failures.Count, 1))
