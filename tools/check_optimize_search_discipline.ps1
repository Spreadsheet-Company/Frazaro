<#
check_optimize_search_discipline.ps1 - OPTIMIZE.0.C's data discipline, held
over OPTIMIZE's search: Long arrays and numbers, and nothing else.

WHY: OPTIMIZE.0 measured one CreateObject per search node at 160 seconds a
million nodes before any search happened, and DATALOG.13 measured one
TypeName at 85% of a DATALOG row. A search visits 10^5 to 10^7 nodes. The
regression this guards against is SILENT: an object, a dictionary lookup or
a Variant put back into the search loop gives exactly the same answers, so
every suite stays green while the search gets a hundred times slower, and
only a live timing at real size would notice. OPTIMIZE.3's entry names this
scan as the item's own pin.

WHAT IS CHECKED, over src/VLA_OptimizeSearch.bas - the WHOLE module, since
all of it is the search and nothing else lives there:
  - every procedure named in the baseline below still exists in it, and it
    holds no other: a renamed or moved procedure must not disarm the pin,
    and a new one is added here deliberately, where a reviewer sees it;
  - no non-comment code names CreateObject, GetObject, TypeName, TypeOf,
    Dictionary, Collection, Variant, Object, New, VlaDict..., CallByName,
    IIf, Array(, ParamArray, IsMissing or a Def... statement (string
    literals and comments are stripped first, so a word in either is
    fine);
  - no name is IMPLICITLY a Variant: every Dim, ReDim, Static, Private,
    Public and Const declaration, every parameter, and every Function's
    own return carries an explicit As. `Dim a, b As Long` makes `a` a
    Variant, and is the easiest way to break this module without writing
    the word.

WHAT IS NOT CHECKED: VLA_Optimize.bas, which grounds the problem this
module searches and may use what it likes - grounding runs once per
program, not once per node. Nor what the code does with a Long.

BASELINE: the procedure list below, hand-maintained, in
check_vladict_guard.ps1's reviewable shape.

MUTATION CONTROL, 2026-09-24 (OPTIMIZE.3 slice 2): -Control runs every
planted defect below against a scratch copy of the module and requires each
to be caught, and the unmutated copy to be clean. A new check whose first
run over a known-good tree is not clean is reporting its own defect - and
this one's was: it flagged a ReDim of an array PARAMETER as a ReDim of an
undeclared name, because it collected declared names from Dim lines only.
Parameters count now, and a planted ReDim of a truly undeclared name is
one of the mutants.

NOT wired into VlaSelfTest, the same reasoning the other tools/check_*.ps1
scans give: a step a human runs, not a per-run gate.

Usage:  powershell -File tools\check_optimize_search_discipline.ps1 [-SrcDir <folder>] [-Control]
Exit code: 0 when the module holds the discipline (and, with -Control, when
every planted defect is caught); 1 otherwise.
#>

param(
    [string]$SrcDir = '',
    [switch]$Control
)

$ErrorActionPreference = 'Stop'

if ($SrcDir -eq '') {
    $repoRoot = Split-Path -Parent $PSScriptRoot
    $SrcDir   = Join-Path $repoRoot 'src'
}

$moduleName = 'VLA_OptimizeSearch.bas'

# --- The procedures, hand-maintained. ---
# 2026-09-24, OPTIMIZE.3 slice 2: the problem builder, the search, and
# every helper the search loop calls.
$procedures = @(
    'OptProblemInit', 'OptProblemAddClause', 'OptProblemAddCounter',
    'OptSearchRun', 'ProblemIsWellFormed', 'LoadProblem', 'AssignAtom',
    'RootSeed', 'Propagate', 'ClauseForce', 'CounterForceHigh',
    'CounterForceLow', 'Backtrack', 'UndoTo', 'ExplainRootConflict',
    'CollectTags'
)

$forbidden = @(
    @{ Pattern = '\bCreateObject\b'; Word = 'CreateObject' },
    @{ Pattern = '\bGetObject\b';    Word = 'GetObject' },
    @{ Pattern = '\bTypeName\b';     Word = 'TypeName' },
    @{ Pattern = '\bTypeOf\b';       Word = 'TypeOf' },
    @{ Pattern = '\bDictionary\b';   Word = 'Dictionary' },
    @{ Pattern = '\bCollection\b';   Word = 'Collection' },
    @{ Pattern = '\bVariant\b';      Word = 'Variant' },
    @{ Pattern = '\bObject\b';       Word = 'Object' },
    @{ Pattern = '\bNew\b';          Word = 'New' },
    @{ Pattern = '\bVlaDict\w*';     Word = 'VlaDict' },
    @{ Pattern = '\bCallByName\b';   Word = 'CallByName' },
    @{ Pattern = '\bIIf\b';          Word = 'IIf' },
    @{ Pattern = '\bArray\s*\(';     Word = 'Array(' },
    @{ Pattern = '\bParamArray\b';   Word = 'ParamArray' },
    @{ Pattern = '\bIsMissing\b';    Word = 'IsMissing' },
    @{ Pattern = '^\s*Def(Var|Obj|Int|Lng|Sng|Dbl|Str|Bool|Byte|Cur|Date|Dec|LngLng|LngPtr)\b'; Word = 'a Def... statement' }
)

# One logical line per entry: continuations joined, string literals
# blanked, comments dropped. Keeps the physical number of its first line.
function Get-LogicalLines([string[]]$lines) {
    $out = New-Object System.Collections.Generic.List[object]
    $buf = ''
    $first = 0
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $raw = $lines[$i]
        # Blank string literals ("" inside one is an escaped quote), then
        # cut the first apostrophe outside them.
        $sb = New-Object System.Text.StringBuilder
        $inStr = $false
        $cut = $false
        for ($k = 0; $k -lt $raw.Length -and -not $cut; $k++) {
            $ch = $raw[$k]
            if ($inStr) {
                if ($ch -eq '"') {
                    if ($k + 1 -lt $raw.Length -and $raw[$k + 1] -eq '"') { $k++; [void]$sb.Append('  '); continue }
                    $inStr = $false
                    [void]$sb.Append('"')
                } else {
                    [void]$sb.Append(' ')
                }
            } elseif ($ch -eq '"') {
                $inStr = $true
                [void]$sb.Append('"')
            } elseif ($ch -eq "'") {
                $cut = $true
            } else {
                [void]$sb.Append($ch)
            }
        }
        $code = $sb.ToString()
        if ($code -match '^\s*Rem\b') { $code = '' }
        if ($buf -eq '') { $first = $i + 1 }
        if ($code -match '\s_\s*$') {
            $buf += ($code -replace '\s_\s*$', ' ')
            continue
        }
        $buf += $code
        $out.Add([pscustomobject]@{ No = $first; Text = $buf })
        $buf = ''
    }
    if ($buf -ne '') { $out.Add([pscustomobject]@{ No = $first; Text = $buf }) }
    return ,$out
}

# Splits on commas at parenthesis depth zero.
function Split-TopLevel([string]$s) {
    $parts = New-Object System.Collections.Generic.List[string]
    $depth = 0
    $cur = New-Object System.Text.StringBuilder
    foreach ($ch in $s.ToCharArray()) {
        if ($ch -eq '(') { $depth++ }
        elseif ($ch -eq ')') { $depth-- }
        if ($ch -eq ',' -and $depth -eq 0) {
            $parts.Add($cur.ToString())
            $cur = New-Object System.Text.StringBuilder
            continue
        }
        [void]$cur.Append($ch)
    }
    $parts.Add($cur.ToString())
    return ,$parts
}

# Every problem in one module's text, as "line N: why".
function Get-Problems([string[]]$lines) {
    $problems = New-Object System.Collections.Generic.List[string]
    $logical = Get-LogicalLines $lines
    # Every name something declares - a Dim, a Static, a module-level line
    # or a PARAMETER. A ReDim of a name in here resizes it; a ReDim of any
    # other plain name declares a new Variant array. (This check's first
    # run over the known-good module flagged a ReDim of an array
    # parameter, which is its own defect: parameters count.)
    $declared = New-Object System.Collections.Generic.HashSet[string]([System.StringComparer]::OrdinalIgnoreCase)

    # The procedures: each in the baseline present, and no other.
    $found = New-Object System.Collections.Generic.List[string]
    $procPattern = '^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?(Sub|Function|Property\s+(?:Get|Let|Set))\s+(\w+)\s*\((.*)\)\s*(.*)$'
    foreach ($l in $logical) {
        if ($l.Text -match $procPattern) {
            $kind = $Matches[1]
            $name = $Matches[2]
            $params = $Matches[3]
            $tail = $Matches[4]
            $found.Add($name)
            if ($kind -eq 'Function' -and -not ($tail -match '^\s*As\s+\w+')) {
                $problems.Add("line $($l.No): Function $name has no As for what it returns - that is a Variant")
            }
            if ($params.Trim().Length -gt 0) {
                foreach ($p in (Split-TopLevel $params)) {
                    if ($p.Trim() -match '^(?:Optional\s+)?(?:ByVal\s+|ByRef\s+)?(\w+)') { [void]$declared.Add($Matches[1]) }
                    if (-not ($p -match '\bAs\s+\w+')) {
                        $problems.Add("line $($l.No): parameter '$($p.Trim())' of $name has no As - that is a Variant")
                    }
                }
            }
        }
    }
    foreach ($p in $procedures) {
        if (-not $found.Contains($p)) {
            $problems.Add("procedure $p not found - renamed, moved or removed? Update this check's list deliberately.")
        }
    }
    foreach ($f in $found) {
        if ($procedures -notcontains $f) {
            $problems.Add("procedure $f is not in this check's list - add it deliberately, where a reviewer sees it")
        }
    }

    # Declarations: every name carries an explicit As. A ReDim resizes a
    # name declared elsewhere and takes no As of its own - but a ReDim of a
    # name declared NOWHERE declares a new Variant array, so each ReDim'd
    # plain name must be one a Dim, Static, module-level line or parameter
    # declares.
    $redims = New-Object System.Collections.Generic.List[object]
    foreach ($l in $logical) {
        $t = $l.Text
        foreach ($fb in $forbidden) {
            if ($t -match $fb.Pattern) {
                $problems.Add("line $($l.No): uses $($fb.Word) - $($t.Trim())")
            }
        }
        $decl = $null
        if ($t -match '^\s*ReDim(?:\s+Preserve)?\s+(.*)$') {
            foreach ($d in (Split-TopLevel $Matches[1])) {
                $redims.Add([pscustomobject]@{ No = $l.No; Target = $d.Trim() })
            }
            continue
        }
        if ($t -match '^\s*(?:Dim|Static)\s+(.*)$') {
            $decl = $Matches[1]
        } elseif ($t -match '^\s*(?:Public|Private|Global)\s+(?!Sub\b|Function\b|Property\b|Type\b|Enum\b|Declare\b|Const\b)(.*)$') {
            $decl = $Matches[1]
        } elseif ($t -match '^\s*(?:Public\s+|Private\s+)?Const\s+(.*)$') {
            $decl = $Matches[1]
        }
        if ($null -ne $decl) {
            foreach ($d in (Split-TopLevel $decl)) {
                $dt = $d.Trim()
                if ($dt.Length -eq 0) { continue }
                if ($dt -match '^(?:WithEvents\s+)?(\w+)') { [void]$declared.Add($Matches[1]) }
                if (-not ($dt -match '\bAs\s+\w+')) {
                    $problems.Add("line $($l.No): '$dt' is declared with no As - that is a Variant")
                }
            }
        }
    }
    foreach ($r in $redims) {
        if ($r.Target -match '^([\w.]+)') {
            $nm = $Matches[1]
            if ($nm.Contains('.')) { continue }
            if (-not $declared.Contains($nm)) {
                $problems.Add("line $($r.No): ReDim of '$nm', which nothing declares - that declares a Variant array")
            }
        }
    }

    # Members of a Type block carry an As too.
    $inType = $false
    foreach ($l in $logical) {
        $t = $l.Text
        if ($t -match '^\s*(?:Public\s+|Private\s+)?Type\s+\w+') { $inType = $true; continue }
        if ($t -match '^\s*End\s+Type\b') { $inType = $false; continue }
        if ($inType -and $t.Trim().Length -gt 0 -and -not ($t -match '\bAs\s+\w+')) {
            $problems.Add("line $($l.No): Type member '$($t.Trim())' has no As - that is a Variant")
        }
    }
    return ,$problems
}

function Test-Module([string]$path) {
    $lines = [System.IO.File]::ReadAllLines($path)
    return Get-Problems $lines
}

$path = Join-Path $SrcDir $moduleName
if (-not (Test-Path -LiteralPath $path)) {
    Write-Output "=== OPTIMIZE SEARCH DISCIPLINE ==="
    Write-Output "  FAIL  $moduleName not found at $path"
    exit 1
}

Write-Output '=== OPTIMIZE SEARCH DISCIPLINE (OPTIMIZE.0.C, pinned by OPTIMIZE.3) ==='
Write-Output ''
$problems = Test-Module $path
if ($problems.Count -gt 0) {
    Write-Output "  FAIL  $moduleName"
    foreach ($p in $problems) { Write-Output "        $p" }
} else {
    Write-Output "  ok    $moduleName - $($procedures.Count) procedures, Long arrays and numbers only, every name typed"
}

if ($Control) {
    Write-Output ''
    Write-Output '--- mutation control: each planted defect must be caught ---'
    $orig = [System.IO.File]::ReadAllLines($path)
    $text = [System.IO.File]::ReadAllText($path)
    $mutants = @(
        @{ Name = 'a Collection in Propagate';  Find = '    Dim a As Long, v As Long, i As Long, c As Long, g As Long' + "`r`n" + '    Do While mQHead <= mTLen'; Replace = '    Dim a As Long, v As Long, i As Long, c As Long, g As Long' + "`r`n" + '    Dim junk As Collection' + "`r`n" + '    Do While mQHead <= mTLen' },
        @{ Name = 'CreateObject in ClauseForce'; Find = '    Dim nOpen As Long, openLit As Long'; Replace = '    Dim nOpen As Long, openLit As Long' + "`r`n" + '    If nOpen < 0 Then Debug.Print CreateObject("Scripting.Dictionary") Is Nothing' },
        @{ Name = 'an implicit Variant (Dim a, b As Long)'; Find = '    Dim j As Long, a As Long, nTrue As Long'; Replace = '    Dim j, a As Long, nTrue As Long' },
        @{ Name = 'an untyped parameter'; Find = 'Private Sub AssignAtom(ByVal a As Long, ByVal v As Long'; Replace = 'Private Sub AssignAtom(ByVal a, ByVal v As Long' },
        @{ Name = 'a Function with no return type'; Find = 'Private Function ClauseForce(ByVal c As Long) As Boolean'; Replace = 'Private Function ClauseForce(ByVal c As Long)' },
        @{ Name = 'a procedure renamed out of the list'; Find = 'Private Sub Backtrack()'; Replace = 'Private Sub Backtrack2()' },
        @{ Name = 'IIf in RootSeed'; Find = '            lit = mLits(mCStart(c))'; Replace = '            lit = IIf(True, mLits(mCStart(c)), 0)' },
        @{ Name = 'an untyped Type member'; Find = '    rootConflict As Boolean'; Replace = '    rootConflict' },
        @{ Name = 'a ReDim that declares a Variant array'; Find = '    Dim g As Long, j As Long' + "`r`n" + '    For g = 1 To mNG'; Replace = '    Dim g As Long, j As Long' + "`r`n" + '    ReDim junk(1 To 2)' + "`r`n" + '    For g = 1 To mNG' }
    )
    $scratch = Join-Path ([System.IO.Path]::GetTempPath()) ('vla_optsearch_control_' + [System.Guid]::NewGuid().ToString('N'))
    [void](New-Item -ItemType Directory -Path $scratch)
    $controlFailed = New-Object System.Collections.Generic.List[string]
    $clean = Get-Problems $orig
    if ($clean.Count -gt 0) { $controlFailed.Add('the unmutated module is not clean') }
    $n = 0
    foreach ($m in $mutants) {
        $n++
        $idx = $text.IndexOf($m.Find)
        if ($idx -lt 0) {
            $controlFailed.Add("mutant '$($m.Name)': its anchor is gone from the module, so it cannot be planted - update the mutant")
            Write-Output "  ??    $($m.Name) (anchor missing)"
            continue
        }
        if ($text.IndexOf($m.Find, $idx + 1) -ge 0) {
            $controlFailed.Add("mutant '$($m.Name)': its anchor appears twice - make it unique")
            Write-Output "  ??    $($m.Name) (anchor not unique)"
            continue
        }
        $mutated = $text.Substring(0, $idx) + $m.Replace + $text.Substring($idx + $m.Find.Length)
        $mPath = Join-Path $scratch ("m$n.bas")
        [System.IO.File]::WriteAllText($mPath, $mutated)
        $got = Test-Module $mPath
        if ($got.Count -gt 0) {
            Write-Output "  ok    caught: $($m.Name)"
        } else {
            Write-Output "  MISS  not caught: $($m.Name)"
            $controlFailed.Add("mutant '$($m.Name)' was not caught")
        }
    }
    if ($controlFailed.Count -gt 0) {
        Write-Output ''
        Write-Output '=== CONTROL: FAILED ==='
        foreach ($f in $controlFailed) { Write-Output "  $f" }
        exit 1
    }
    Write-Output "  ok    the unmutated module is clean"
}

Write-Output ''
if ($problems.Count -gt 0) {
    Write-Output '=== CHECK: FAILED ==='
    exit 1
}
if ($Control) {
    Write-Output "=== CHECK: clean, and all $($mutants.Count) planted defects caught ==="
} else {
    Write-Output '=== CHECK: clean ==='
}
exit 0
