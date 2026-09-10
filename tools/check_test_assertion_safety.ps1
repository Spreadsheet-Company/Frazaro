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

RULES E AND F - THE RAW SUBSCRIPT, added by PROLOG.20's own follow-up.
Rules A-D fire only where the statement guards the variable itself, so an
assertion whose bound was computed on an EARLIER line was invisible to
them:

    okAll = (nRowsAll = 4) And (CStr(arrAll(r0, c0)) = "Name")

  E  an And-joined statement may not raw-subscript a bare local. That is
     the assertion class, and it is exactly the shape above.

  F  PROCEDURE-SCOPED, and it is what makes E's fix stick: once a
     procedure passes V to a guarded helper it has ADMITTED V's shape is
     uncertain, so no statement in that procedure may raw-subscript V.
     Without F, fixing the assertion and leaving the neighbouring
     `detailAll = ... & CStr(arrAll(r0, c0))` still kills the run - a
     Report's detail is evaluated on every call - and that statement has
     no `And` in it, so E cannot reach it.

F was the answer to what PROLOG.20 filed as needing "dataflow, not a
regex". It needs neither: it is one syntactic question asked over one
procedure. What it will not do is judge a raw subscript in a procedure
that never uses a helper at all - a blanket rule there was measured at
99 statements, most of them legitimate (loop bodies indexed by their own
loop bounds; write-backs like `arr(2, 2) = "CHANGED"` that are the test's
ACTION, not its assertion), and a rule that flags 90 false positives gets
widened until it stops noticing.

An assignment TARGET is dropped for that reason, and the one genuine
false positive - a loop index over a local Array() literal - is exempted
BY NAME in $rawIndexExempt with its reason, rather than by loosening the
rule.

RULE G - AN ASSERTION THAT CANNOT PASS, which is the other half of the
same idea. Rules A-F catch a test that cannot survive its own FAILURE;
G catches one that has already failed permanently, because it expects a
rendering the writer can no longer emit. Both are tests that are wrong
about THEMSELVES rather than about the code, which is why they share a
script.

PROLOG.13 is the incident: it changed how a findall bag renders,
re-pointed the pins it REMEMBERED, and two stale ones reached the owner's
live run. $retiredRenderings holds each retired spelling with what
retired it. That list only ever grows - a rendering that stopped being
emitted is a fact about the past - so an item changing a rendering adds
its line BEFORE changing the writer, and the check then names exactly the
pins to re-point instead of leaving it to memory.

G COUNTS WHAT IT EXAMINED and fails if that collapses, because this rule
shipped broken once: its extraction ran past the expected value and
captured `got: ` out of the DETAIL argument, so it inspected the wrong
string and reported clean. Only the mutation that failed to turn it red
revealed it. A rule that looks at nothing passes by not looking.

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
# PROLOG.20 STAGE 1 took VLA_Tests_Query.bas from 146 to 0 and held the
# other three at 4/1/3; STAGE 2 took those to 0 as well. The whole
# baseline is zero now, and the ratchet is what it is FOR from here on:
# every number may only shrink, and the check fails just as loudly if a
# module comes in UNDER its ceiling without the ceiling being lowered in
# the same diff, so a gain cannot be silently given back.
#
# The staging mattered while it lasted: a check that can only say
# all-clean-or-all-broken sits red across a whole migration, and nobody
# can then tell a NEW defect from the backlog. Ratcheting rather than
# gating is what let stage 1 land green with stage 2's work still
# outstanding and visible.
$ceilings = [ordered]@{
    'src\VLA_Tests.bas'         = 0
    'src\VLA_Tests_Grammar.bas' = 0
    'src\VLA_Tests_Host.bas'    = 0
    'src\VLA_Tests_Query.bas'   = 0
}

# ---- baseline: RETIRED RENDERINGS - rule G ------------------------------
# An expected literal spelling a rendering the writer can no longer emit is
# a test that CANNOT PASS, whatever the code does. PROLOG.13 hit this for
# real: it changed how a findall bag renders, re-pointed the pins it
# remembered, and TWO stale ones survived to the owner's live run.
#
# Each entry is a regex over the expected literal, plus what retired it.
# The list only ever grows: a rendering that stops being emitted is a fact
# about the past, so no entry here can go wrong later. An item that
# changes a rendering adds its own line BEFORE changing the writer, and
# the check then lists exactly the pins to re-point - which is what
# PROLOG.13 had to do by memory.
$retiredRenderings = @(
    @{ Pattern = '^\(\s*cons\b.*\bnil\s*\)*\s*$'
       Why     = 'a PROPER cons chain; PROLOG.21 contracts those to (list ...) at write time, so no result can equal this' }
)

# ---- baseline: raw subscripts accepted as safe, with the reason --------
# Rule E flags a raw subscript of a bare local inside an And-joined
# assertion. One instance is genuinely safe and is exempted BY NAME with
# its reason, rather than by loosening the rule until it stops noticing.
# Keyed "<module>|<procedure>|<variable>".
$rawIndexExempt = @{
    'src\VLA_Tests_Host.bas|TestRuntimeModule|names' =
        'indexed only by `For i = LBound(names) To UBound(names)` over a local Array() literal, so the subscript is in range by the loop bounds themselves'
}

# ---- baseline: accepted mentions, with the reason each is safe ---------
# A guarded helper answers False or -1 for a shape it cannot read, so
# mentioning a guarded variable inside one is exactly the fix this check
# asks for and must not itself be reported.
$guardedHelpers = @(
    'ResultRowCount', 'ResultColCount', 'ResultCellIs', 'ResultCol1Is',
    'ResultBoolIs', 'ResultDescribe', 'JoinColumn',
    'ResultTextIs', 'ResultTextStartsWith', 'CollItemIs',
    # PROLOG.20 stage 2: the same shapes for the other three modules -
    # 1-D and 2-D arrays, and a Collection.
    'Arr1DIsEmpty', 'Arr1DItemIs', 'Arr1DText', 'Arr2DItemIs', 'Arr2DText',
    # PROLOG.20 follow-up: the numeric and display-text cell readers.
    'ResultCellNumIs', 'ResultCellText'
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
    @{ Kind = 'Exists'; Form = '\b(?:Result(?:RowCount|ColCount|CellIs|CellNumIs|CellText|Col1Is|BoolIs|TextIs|TextStartsWith)|Arr1D(?:IsEmpty|ItemIs|Text)|Arr2D(?:ItemIs|Text)|CollItemIs)\s*\(\s*{0}\s*[,)]' },
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
                @([regex]::Matches($code, '\b(?:Result(?:RowCount|ColCount|CellIs|CellNumIs|CellText|Col1Is|BoolIs|TextIs|TextStartsWith)|Arr1D(?:IsEmpty|ItemIs|Text)|Arr2D(?:ItemIs|Text)|CollItemIs)\s*\(\s*([A-Za-z_]\w*)') |
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

# ======================================================================
#  Rules E and F - the RAW SUBSCRIPT, which rules A-D cannot see.
#
#  Rules A-D only fire where the statement itself guards the variable. An
#  assertion whose bound was computed on an EARLIER line is invisible to
#  them:
#
#      okAll = (nRowsAll = 4) And (CStr(arrAll(r0, c0)) = "Name")
#
#  E: an And-joined statement may not raw-subscript a bare local. That is
#     the assertion class, and it is exactly the shape above.
#
#  F: PROCEDURE-SCOPED, and it is what makes E's fix stick. Once a
#     procedure passes V to a guarded helper it has ADMITTED V's shape is
#     uncertain, so no statement in that procedure may raw-subscript V.
#     Without F, fixing the assertion and leaving the neighbouring
#     `detailAll = ... & CStr(arrAll(r0, c0))` still kills the run - a
#     detail is evaluated on every call - and that statement has no `And`
#     in it, so E cannot reach it. F needs no dataflow: it is one
#     syntactic question asked over one procedure.
#
#  WHY NOT SIMPLY "no raw subscript anywhere in a test module": measured,
#  and it is 99 statements, most of them legitimate - loop bodies indexed
#  by their own loop bounds, and write-backs like `arr(2, 2) = "CHANGED"`
#  that are the test's ACTION rather than its assertion. A rule that
#  flags 90 false positives gets widened until it stops noticing, which
#  is the failure mode this file exists to avoid.
# ======================================================================
$vbaKeywords = @'
and or not then else elseif if end select case for each next do loop while wend with
to step in new call set let get is mod xor eqv imp byval byref optional as dim const
redim erase exit function sub property public private friend static on error resume goto
'@ -split '\s+' | Where-Object { $_ }
$vbaBuiltins = @'
instr instrrev left right mid ltrim rtrim trim len lenb replace split join filter strcomp string space
chr chrw asc ascw cstr clng cint cdbl csng cbool cbyte cdate cvar ccur cdec val str format formatnumber
abs int fix sgn sqr exp log round ubound lbound isarray isdate isempty iserror ismissing isnull isnumeric
isobject typename vartype array iif choose switch hex oct rgb dateadd datediff datepart dateserial
timeserial now date time rnd msgbox dir filelen shell environ callbyname createobject getobject nz
'@ -split '\s+' | Where-Object { $_ }
$notAnIndex = @{}
foreach ($w in ($vbaKeywords + $vbaBuiltins)) { $notAnIndex[$w] = $true }
foreach ($sf in (Get-ChildItem -Path (Join-Path $repoRoot 'src') -Filter '*.bas')) {
    foreach ($l in (Get-Content -LiteralPath $sf.FullName)) {
        if ($l -match '^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+([A-Za-z_]\w*)') {
            $notAnIndex[$Matches[1].ToLower()] = $true
        }
    }
}
$helperAlt = ($guardedHelpers | ForEach-Object { [regex]::Escape($_) }) -join '|'

foreach ($rel in $ceilings.Keys) {
    $raw = Get-Content -LiteralPath (Join-Path $repoRoot $rel)

    # statements again, this time carrying the enclosing procedure name
    $stmts = New-Object System.Collections.ArrayList
    $acc = ''; $accLine = 0; $proc = '(module)'
    for ($i = 0; $i -lt $raw.Count; $i++) {
        if ($raw[$i] -match '^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+([A-Za-z_]\w*)') { $proc = $Matches[1] }
        $s = Strip-Comment $raw[$i]
        if ($acc -eq '') { $accLine = $i + 1 }
        if ($s -match '\s_\s*$') { $acc += ($s -replace '\s_\s*$', ' '); continue }
        $acc += $s
        if ($acc.Trim().Length -gt 0) { [void]$stmts.Add([pscustomobject]@{ Line = $accLine; Text = $acc.Trim(); Proc = $proc }) }
        $acc = ''
    }

    # pass 1: which locals does each procedure hand to a guarded helper?
    $admitted = @{}
    foreach ($st in $stmts) {
        $code = Blank-Strings $st.Text
        foreach ($m in [regex]::Matches($code, '\b(?:' + $helperAlt + ')\s*\(\s*([A-Za-z_]\w*)')) {
            $admitted["$($st.Proc)|$($m.Groups[1].Value)"] = $true
        }
    }

    # pass 2: raw subscripts
    foreach ($st in $stmts) {
        $code = Blank-Strings $st.Text
        if ($code -match '^\s*(Dim|Const|ReDim|Declare|Erase)\b') { continue }
        # An assignment TARGET is the test's action, not its assertion -
        # `arr(2, 2) = "CHANGED"` writes a cell back. Dropped so the rule
        # judges what a test READS.
        $body = [regex]::Replace($code, '^\s*[A-Za-z_]\w*\s*\([^)]*\)\s*=', ' ')
        foreach ($m in [regex]::Matches($body, '(?<![.\w$])([A-Za-z_]\w*)\s*\(')) {
            $v = $m.Groups[1].Value
            if ($notAnIndex.ContainsKey($v.ToLower())) { continue }
            $key = "$rel|$($st.Proc)|$v"
            if ($rawIndexExempt.ContainsKey($key)) { continue }
            $isE = ($code -match '\bAnd\b')
            $isF = $admitted.ContainsKey("$($st.Proc)|$v")
            if (-not ($isE -or $isF)) { continue }
            [void]$findings.Add([pscustomobject]@{
                Module = $rel; Line = $st.Line; Var = $v; Text = $st.Text })
        }
    }
}

# ======================================================================
#  Rule G - AN ASSERTION THAT CANNOT PASS.
#
#  Rules A-F are about an assertion that cannot survive its own FAILURE.
#  This one is the other way round: an assertion that expects a value the
#  engine can no longer produce has already failed, permanently, and no
#  amount of correct code will satisfy it. Same family - a test that is
#  wrong about ITSELF rather than about the code - which is why it lives
#  in this script rather than an eighteenth one.
#
#  PROLOG.13's own live pass is the incident: it changed a rendering,
#  re-pointed the pins it REMEMBERED, and two stale ones reached the
#  owner. Searching for the literals you expect rather than enumerating
#  the class is the same error as a hand count.
# ======================================================================
#  THE EXTRACTION IS LAZY, NOT GREEDY, and that was a live bug in this
#  very rule. Written `\([^,]+,(?:[^,]+,){0,2}\s*"(...)"` the optional
#  argument group ran PAST the expected value and captured `got: ` out of
#  the DETAIL argument instead - so the rule examined the wrong string and
#  reported clean. Caught only because the mutation that was supposed to
#  turn it red did not. A lazy `*?` stops at the first string literal,
#  which is the expected value in every one of these helpers.
$staleLiterals = New-Object System.Collections.ArrayList
$literalsSeen  = 0
$expectedArg = '(?:ResultCol1Is|ResultCellIs|ResultCellNumIs|ResultCellText|ResultTextIs|ResultTextStartsWith|Arr1DItemIs|Arr2DItemIs|CollItemIs)\s*\(\s*[A-Za-z_]\w*\s*,(?:\s*[^,"]+,){0,8}?\s*"((?:[^"]|"")*)"'
foreach ($rel in $ceilings.Keys) {
    $raw = Get-Content -LiteralPath (Join-Path $repoRoot $rel)
    for ($i = 0; $i -lt $raw.Count; $i++) {
        if ($raw[$i] -match "^\s*'") { continue }
        foreach ($m in [regex]::Matches($raw[$i], $expectedArg)) {
            $lit = $m.Groups[1].Value
            $literalsSeen++
            foreach ($r in $retiredRenderings) {
                if ($lit -match $r.Pattern) {
                    [void]$staleLiterals.Add([pscustomobject]@{
                        Module = $rel; Line = $i + 1; Lit = $lit; Why = $r.Why })
                }
            }
        }
    }
}
# A rule that examined nothing would report "clean" forever. The suite has
# hundreds of these assertions, so zero means the extraction broke - which
# is exactly what happened once already.
if ($literalsSeen -lt 50) {
    $failures.Add("rule G examined only $literalsSeen expected literal(s) - the extraction has broken, and a rule that looks at nothing passes by not looking")
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
Write-Output '--- rule G: no assertion may expect a rendering the writer cannot emit ---'
foreach ($r in $retiredRenderings) {
    Write-Output ("  retired: {0}" -f $r.Why)
}
Write-Output ("  examined: {0} expected literal(s)" -f $literalsSeen)
if ($staleLiterals.Count -eq 0) {
    Write-Output '  none - every expected literal is a rendering the engine can still produce'
} else {
    foreach ($s in $staleLiterals) {
        Write-Output ("  {0}:{1}  STALE  {2}" -f $s.Module, $s.Line, $s.Lit)
        $failures.Add("$($s.Module):$($s.Line) expects '$($s.Lit)' - $($s.Why)")
    }
}
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
