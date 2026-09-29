<#
proofs_lib.ps1 - what tools\check_proofs.ps1 and tools\proofs_lp.ps1 share.
Dot-sourced by both and never run on its own: it defines functions and does
nothing else.

THE READER (METAPROOF.1, moved here unchanged from check_proofs.ps1, so the
file that checks the corpus and the file that exports it read it one way).
VLA.bas's Tokenize, followed rule for rule: a ; comment runs to the end of
its line, "..." takes \" and \\ as escapes, and every other run of
characters up to a paren, a space, a ; or a " is one atom.

THE TRANSLATION (METAPROOF.2). A DATALOG proof is very nearly an ASP
program, so clingo - the Potassco answer-set solver, written by other
people from other sources - can say whether the answer a proof expects is
the answer its program means. A proof's own words are carried over; where
DATALOG and clingo could read the same words differently, the proof is
declined, by name, and never translated by a guess.

  (fact (p a b))           p(a,b).
  (rule (h X) (b X) ...)   h(X) :- b(X), ... .
  (not (p X))              not p(X). DATALOG binds a negated atom's
                           variables first and takes only a stratified
                           program, whose one answer set is its least
                           model.
  (count N (p X Y))        N = #count { X,Y : p(X,Y) }. For each binding
                           of the variables bound EARLIER in the rule,
                           DATALOG counts the rows that match, and no row
                           is a real 0 (ApplyAggregate); clingo counts the
                           distinct tuples, the variables bound outside
                           held fixed. The two agree when every variable
                           DATALOG counts appears nowhere else in the
                           rule; clingo would read one that does as bound,
                           so such a count is declined.
  (= A B)                  A = B. DATALOG binds both sides first, and the
                           values below never set a word beside a number.
  (query p)                the rows of p, compared as a SET: their order
                           and the header row are DATALOG's own.
  (headless)               nothing - it shapes the spill, not the rows.

  NAMES. A predicate is folded to lower case, as VLA_Identity.Fold folds
  it, and must then be a clingo name: a letter, then letters, digits and _.
  A variable is a bare word starting A-Z, as DATALOG reads one. DATALOG
  matches variable names without case (VlaDictNew) and clingo with case, so
  a rule that spells one variable two ways is declined.
  VALUES. Lowercase words (a-z, then a-z, 0-9 and _) and whole numbers of up
  to nine digits: both lineages read those alike, and neither equates a word
  with a number. Anything else is declined - a quoted "4", for one, is the
  number 4 to DATALOG and a string to clingo.
  THE vla_ PREFIX is the judge's own: a predicate or a value using it is
  declined.

  DECLINED, with the reason printed: a (refuses ...) proof (a refusal is
  DATALOG's own policy, and no other lineage raises its message ids), an
  (answer ...) proof, a query written as a fact, and sum, let, textjoin, the
  text tests, the comparisons other than =, and keyed atoms. No proof in the
  corpus uses any of them yet, so a translation of them would have no
  witness; each comes in with the first proof that needs it.

THE JUDGE, written after every program. vla_want/N holds the rows the proof
expects; vla_extra/N, the rows clingo derives that the proof does not list;
vla_missing/N, the reverse. The file shows vla_verdict(agrees) or
vla_verdict(differs), and beside a difference the rows that make it. A
CONTROL is one proof with its expectation made wrong on purpose - a row
left out, a row added, no row at all - and shows vla_verdict(caught), or
vla_verdict(missed) if the judge failed to notice: a judge never shown to
catch a wrong answer proves nothing.
#>

# ===========================================================================
#  THE READER
# ===========================================================================

function Get-ProofTokens([string]$src, [string]$label) {
    $toks = New-Object System.Collections.Generic.List[object]
    $n = $src.Length
    $i = 0
    $lineNo = 1
    while ($i -lt $n) {
        $ch = $src[$i]
        if ($ch -eq "`n") { $lineNo++; $i++; continue }
        if ($ch -eq ' ' -or $ch -eq "`t" -or $ch -eq "`r") { $i++; continue }
        if ($ch -eq '(' -or $ch -eq ')') {
            $toks.Add([pscustomobject]@{ Kind = [string]$ch; Text = [string]$ch; Line = $lineNo })
            $i++
            continue
        }
        if ($ch -eq ';') {
            while ($i -lt $n -and $src[$i] -ne "`r" -and $src[$i] -ne "`n") { $i++ }
            continue
        }
        if ($ch -eq '"') {
            $startLine = $lineNo
            $i++
            $sb = New-Object System.Text.StringBuilder
            $closed = $false
            while ($i -lt $n) {
                $c2 = $src[$i]
                if ($c2 -eq '\') {
                    $d2 = ''
                    if ($i + 1 -lt $n) { $d2 = [string]$src[$i + 1] }
                    if ($d2 -eq '"') { [void]$sb.Append('"'); $i += 2 }
                    elseif ($d2 -eq '\') { [void]$sb.Append('\'); $i += 2 }
                    else { [void]$sb.Append($c2); $i++ }
                } elseif ($c2 -eq '"') {
                    $i++
                    $closed = $true
                    break
                } else {
                    if ($c2 -eq "`n") { $lineNo++ }
                    [void]$sb.Append($c2)
                    $i++
                }
            }
            if (-not $closed) { throw "${label}: a string opened on line $startLine is never closed" }
            $toks.Add([pscustomobject]@{ Kind = 'str'; Text = $sb.ToString(); Line = $startLine })
            continue
        }
        $start = $i
        while ($i -lt $n) {
            $c3 = $src[$i]
            if ($c3 -eq '(' -or $c3 -eq ')' -or $c3 -eq ' ' -or $c3 -eq "`t" -or $c3 -eq "`r" -or $c3 -eq "`n" -or $c3 -eq ';' -or $c3 -eq '"') { break }
            $i++
        }
        $toks.Add([pscustomobject]@{ Kind = 'atom'; Text = $src.Substring($start, $i - $start); Line = $lineNo })
    }
    return ,$toks
}

# Tokens into forms: a node is a list (Items), an atom or a string.
function ConvertTo-ProofForms($toks, [string]$label) {
    $forms = New-Object System.Collections.Generic.List[object]
    $script:ptPos = 0
    while ($script:ptPos -lt $toks.Count) {
        $forms.Add((Read-ProofForm $toks $label))
    }
    return ,$forms
}

function Read-ProofForm($toks, [string]$label) {
    $t = $toks[$script:ptPos]
    if ($t.Kind -eq ')') { throw "${label}: a ')' on line $($t.Line) closes nothing" }
    if ($t.Kind -eq '(') {
        $node = [pscustomobject]@{ Kind = 'list'; Text = ''; Line = $t.Line; Items = (New-Object System.Collections.Generic.List[object]) }
        $script:ptPos++
        while ($true) {
            if ($script:ptPos -ge $toks.Count) { throw "${label}: the '(' on line $($t.Line) is never closed" }
            if ($toks[$script:ptPos].Kind -eq ')') { $script:ptPos++; break }
            $node.Items.Add((Read-ProofForm $toks $label))
        }
        return $node
    }
    $script:ptPos++
    return [pscustomobject]@{ Kind = $t.Kind; Text = $t.Text; Line = $t.Line; Items = $null }
}

function Get-ProofHeadWord($node) {
    if ($node.Kind -ne 'list') { return '' }
    if ($node.Items.Count -lt 1) { return '' }
    if ($node.Items[0].Kind -ne 'atom') { return '' }
    return $node.Items[0].Text.ToLowerInvariant()
}

# ===========================================================================
#  THE TRANSLATION (METAPROOF.2)
# ===========================================================================

# A proof the translation declines is thrown with this prefix and caught by
# ConvertTo-ClingoExport, which records the reason. Anything else thrown in
# here is a bug in this file, and is let through.
function Stop-ClingoProof([string]$reason) {
    throw ('clingo-declines: ' + $reason)
}

# The words a rule body may open with that the translation declines: the
# rest of DATALOG's wrappers and operators (VLA_Datalog.bas, ParseProgram)
# once not, count and = are taken out.
function Get-ClingoDeclinedWords {
    return @('sum', 'let', 'textjoin', 'text-starts-with', 'text-ends-with', 'text-contains', '<', '>', '<=', '>=', '<>')
}

function Get-ProofSlug([string]$name) {
    $s = ([regex]::Replace($name.ToLowerInvariant(), '[^a-z0-9]+', '-')).Trim('-')
    if ($s -eq '') { throw "the proof name ""$name"" has no letter or digit to make a file name from" }
    return $s
}

# A value both lineages read alike (VALUES, in the header).
function ConvertTo-ClingoValue([string]$t, [string]$what) {
    if ([regex]::IsMatch($t, '^(0|-?[1-9][0-9]{0,8})$')) { return $t }
    if ([regex]::IsMatch($t, '^[a-z][a-z0-9_]*$')) {
        if ($t.StartsWith('vla_')) { Stop-ClingoProof "the $what $t - the vla_ prefix is the judge's own" }
        return $t
    }
    Stop-ClingoProof "the $what $t - only lowercase words and whole numbers are translated"
}

# One argument of an atom in the program: a variable, or a value. $vars is
# the rule's variables, first spelling by name without case; $null in a fact.
function ConvertTo-ClingoTerm($node, $vars) {
    if ($node.Kind -eq 'list') { Stop-ClingoProof 'a list in an argument place (a keyed atom, or a compound term)' }
    if ($node.Kind -eq 'str') { Stop-ClingoProof ('the quoted string "' + $node.Text + '" - a quoted "4" is the number 4 to DATALOG and a string to clingo') }
    $t = [string]$node.Text
    if ([regex]::IsMatch($t, '^[A-Z]')) {
        if (-not [regex]::IsMatch($t, '^[A-Z][A-Za-z0-9_]*$')) { Stop-ClingoProof "the variable $t - clingo spells a variable as a capital, then letters, digits and _" }
        if ($null -eq $vars) { Stop-ClingoProof "the variable $t in a fact" }
        if ($vars.ContainsKey($t)) {
            if ($vars[$t] -cne $t) { Stop-ClingoProof ('the variable written both ' + $vars[$t] + " and $t - DATALOG reads one variable, clingo two") }
        } else {
            $vars.Add($t, $t)
        }
        return $t
    }
    return (ConvertTo-ClingoValue $t 'value')
}

# One cell of an expected row. A bare word is text here, whatever its first
# letter, and true and false are TRUE and FALSE (datalog.vla's notation).
function ConvertTo-ClingoCell($node) {
    if ($node.Kind -eq 'list') { Stop-ClingoProof 'a list where a cell belongs' }
    if ($node.Kind -eq 'str') { Stop-ClingoProof ('the quoted cell "' + $node.Text + '" - a quoted "4" is the number 4 to DATALOG and a string to clingo') }
    $t = [string]$node.Text
    $low = $t.ToLowerInvariant()
    if ($low -ceq 'true' -or $low -ceq 'false') { Stop-ClingoProof "the cell $t - a TRUE or FALSE cell is not translated" }
    if ([regex]::IsMatch($t, '^[A-Z]')) { Stop-ClingoProof "the cell $t - a value with a capital is not translated" }
    return (ConvertTo-ClingoValue $t 'cell')
}

function ConvertTo-ClingoPredicate([string]$t) {
    if (-not [regex]::IsMatch($t, '^[A-Za-z][A-Za-z0-9_]*$')) { Stop-ClingoProof "the predicate $t - clingo spells a predicate as a letter, then letters, digits and _" }
    $f = $t.ToLowerInvariant()
    if ($f.StartsWith('vla_')) { Stop-ClingoProof "the predicate $t - the vla_ prefix is the judge's own" }
    if ($f -ceq 'not') { Stop-ClingoProof 'a predicate named not, a word clingo keeps for itself' }
    return $f
}

# (pred arg ...) -> Pred, Arity, Text - pred(arg,...) - and VarNames, its
# distinct variables in the order they first appear.
function ConvertTo-ClingoAtom($node, $vars) {
    if ($node.Kind -ne 'list' -or $node.Items.Count -lt 2) { Stop-ClingoProof 'an atom that is not (predicate argument ...)' }
    $p = $node.Items[0]
    if ($p.Kind -ne 'atom') { Stop-ClingoProof 'a predicate written as a string or a list' }
    $pred = ConvertTo-ClingoPredicate $p.Text
    $terms = New-Object System.Collections.Generic.List[string]
    $varNames = New-Object System.Collections.Generic.List[string]
    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    for ($i = 1; $i -lt $node.Items.Count; $i++) {
        $t = ConvertTo-ClingoTerm $node.Items[$i] $vars
        $terms.Add($t)
        if ([regex]::IsMatch($t, '^[A-Z]')) {
            if ($seen.Add($t)) { $varNames.Add($t) }
        }
    }
    return [pscustomobject]@{ Pred = $pred; Arity = $terms.Count; Text = ($pred + '(' + ($terms -join ',') + ')'); VarNames = $varNames }
}

# The atom inside (not ...) or (count ...): a plain atom, never a wrapper.
function ConvertTo-ClingoInnerAtom($node, $vars) {
    $w = Get-ProofHeadWord $node
    if ($w -ceq 'not' -or $w -ceq 'count' -or $w -ceq '=' -or (Get-ClingoDeclinedWords) -ccontains $w) {
        Stop-ClingoProof "($w ...) inside a (not ...) or a (count ...)"
    }
    return (ConvertTo-ClingoAtom $node $vars)
}

function ConvertTo-ClingoRule($node) {
    if ($node.Items.Count -lt 3) { Stop-ClingoProof 'a rule with no body' }
    $vars = New-Object 'System.Collections.Generic.Dictionary[string,string]' ([StringComparer]::OrdinalIgnoreCase)
    $head = ConvertTo-ClingoAtom $node.Items[1] $vars
    # Every variable used outside a (count ...): in the head, a plain atom, a
    # (not ...) or a comparison, or a count's own result. A counted variable
    # found here would be read by clingo as bound, not counted.
    $outside = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($v in $head.VarNames) { [void]$outside.Add($v) }
    $items = New-Object System.Collections.Generic.List[object]
    for ($bi = 2; $bi -lt $node.Items.Count; $bi++) {
        $b = $node.Items[$bi]
        if ($b.Kind -ne 'list') { Stop-ClingoProof 'a bare word in a rule body' }
        $w = Get-ProofHeadWord $b
        if ($w -ceq 'not') {
            if ($b.Items.Count -ne 2) { Stop-ClingoProof 'a (not ...) around other than one atom' }
            $a = ConvertTo-ClingoInnerAtom $b.Items[1] $vars
            foreach ($v in $a.VarNames) { [void]$outside.Add($v) }
            $items.Add([pscustomobject]@{ Kind = 'not'; Atom = $a; Text = ('not ' + $a.Text); Result = '' })
        } elseif ($w -ceq 'count') {
            if ($b.Items.Count -ne 3) { Stop-ClingoProof 'a (count ...) that is not (count Var (atom))' }
            $rv = $b.Items[1]
            if ($rv.Kind -ne 'atom' -or -not [regex]::IsMatch([string]$rv.Text, '^[A-Z]')) { Stop-ClingoProof 'a (count ...) whose result is not a variable' }
            $a = ConvertTo-ClingoInnerAtom $b.Items[2] $vars
            $result = ConvertTo-ClingoTerm $rv $vars
            [void]$outside.Add($result)
            $items.Add([pscustomobject]@{ Kind = 'count'; Atom = $a; Text = ''; Result = $result })
        } elseif ($w -ceq '=') {
            if ($b.Items.Count -ne 3) { Stop-ClingoProof 'a (= ...) with other than two operands' }
            $l = ConvertTo-ClingoTerm $b.Items[1] $vars
            $r = ConvertTo-ClingoTerm $b.Items[2] $vars
            foreach ($t in @($l, $r)) {
                if ([regex]::IsMatch($t, '^[A-Z]')) { [void]$outside.Add($t) }
            }
            $items.Add([pscustomobject]@{ Kind = 'eq'; Atom = $null; Text = ($l + ' = ' + $r); Result = '' })
        } elseif ((Get-ClingoDeclinedWords) -ccontains $w) {
            Stop-ClingoProof "($w ...) - not translated yet"
        } else {
            $a = ConvertTo-ClingoAtom $b $vars
            foreach ($v in $a.VarNames) { [void]$outside.Add($v) }
            $items.Add([pscustomobject]@{ Kind = 'pos'; Atom = $a; Text = $a.Text; Result = '' })
        }
    }
    # Written order, the order DATALOG binds in: a count's variable is a
    # group key when an earlier item bound it, and counted when none did.
    $bound = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    $parts = New-Object System.Collections.Generic.List[string]
    foreach ($it in $items) {
        if ($it.Kind -ceq 'count') {
            if ($it.Atom.VarNames.Count -eq 0) { Stop-ClingoProof 'a (count ...) over an atom with no variable - not translated yet' }
            foreach ($v in $it.Atom.VarNames) {
                if (-not $bound.Contains($v) -and $outside.Contains($v)) {
                    Stop-ClingoProof "the counted variable $v is used again outside its (count ...) - clingo would read it as bound there, not counted"
                }
            }
            $parts.Add($it.Result + ' = #count { ' + ($it.Atom.VarNames -join ',') + ' : ' + $it.Atom.Text + ' }')
            [void]$bound.Add($it.Result)
        } else {
            $parts.Add($it.Text)
            if ($it.Kind -ceq 'pos') {
                foreach ($v in $it.Atom.VarNames) { [void]$bound.Add($v) }
            }
        }
    }
    return [pscustomobject]@{ Head = $head; Text = ($head.Text + ' :- ' + ($parts -join ', ') + '.') }
}

# One proof, as a clingo program and the rows it expects - or $null for a
# refusal proof, which no other lineage can witness.
function ConvertTo-ClingoProof($form) {
    $programNode = $null
    $expectNode = $null
    for ($ci = 2; $ci -lt $form.Items.Count; $ci++) {
        $cw = Get-ProofHeadWord $form.Items[$ci]
        if ($cw -ceq 'program') { $programNode = $form.Items[$ci] }
        elseif (@('rows', 'rows-in-any-order', 'answer', 'refuses') -ccontains $cw) { $expectNode = $form.Items[$ci] }
    }
    if ($null -eq $programNode -or $null -eq $expectNode) { Stop-ClingoProof 'not a well-formed proof - tools\check_proofs.ps1 says why' }
    $expectWord = Get-ProofHeadWord $expectNode
    if ($expectWord -ceq 'refuses') { return $null }
    if ($expectWord -ceq 'answer') { Stop-ClingoProof '(answer ...) - not translated yet' }

    $program = New-Object System.Collections.Generic.List[string]
    $arities = @{}
    $query = ''
    for ($fi = 1; $fi -lt $programNode.Items.Count; $fi++) {
        $f = $programNode.Items[$fi]
        $w = Get-ProofHeadWord $f
        $defined = $null
        if ($w -ceq 'fact') {
            if ($f.Items.Count -ne 2) { Stop-ClingoProof 'a (fact ...) that is not (fact (atom))' }
            $defined = ConvertTo-ClingoAtom $f.Items[1] $null
            $program.Add($defined.Text + '.')
        } elseif ($w -ceq 'rule') {
            $r = ConvertTo-ClingoRule $f
            $defined = $r.Head
            $program.Add($r.Text)
        } elseif ($w -ceq 'query') {
            if ($f.Items.Count -ne 2) { Stop-ClingoProof 'a (query ...) that is not (query name)' }
            $q = $f.Items[1]
            if ($q.Kind -eq 'list') { Stop-ClingoProof 'a query written as a fact, answered TRUE or FALSE - not translated yet' }
            if ($q.Kind -ne 'atom') { Stop-ClingoProof 'a query naming its predicate as a quoted string' }
            $query = ConvertTo-ClingoPredicate $q.Text
        } elseif ($w -ceq 'headless') {
            # The spill's header row: presentation, which the judge does not compare.
        } else {
            Stop-ClingoProof "($w ...) - not a DATALOG form this translation knows"
        }
        if ($null -ne $defined) {
            if ($arities.ContainsKey($defined.Pred) -and $arities[$defined.Pred] -ne $defined.Arity) {
                Stop-ClingoProof "the predicate $($defined.Pred) defined at two arities"
            }
            $arities[$defined.Pred] = $defined.Arity
        }
    }
    if ($query -eq '') { Stop-ClingoProof 'a program with no (query ...)' }
    if (-not $arities.ContainsKey($query)) { Stop-ClingoProof "the query names $query, which no fact or rule here defines" }
    $arity = [int]$arities[$query]

    $want = New-Object System.Collections.Generic.List[string]
    $rowsSeen = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::Ordinal)
    for ($ri = 2; $ri -lt $expectNode.Items.Count; $ri++) {
        $row = $expectNode.Items[$ri]
        if ($row.Kind -ne 'list') { Stop-ClingoProof 'a row that is not a list' }
        if ($row.Items.Count -ne $arity) { Stop-ClingoProof "a row $($row.Items.Count) wide, for $query of arity $arity" }
        $cells = New-Object System.Collections.Generic.List[string]
        foreach ($c in $row.Items) { $cells.Add((ConvertTo-ClingoCell $c)) }
        $rowText = $cells -join ','
        if (-not $rowsSeen.Add($rowText)) { Stop-ClingoProof "the row ($rowText) listed twice - clingo answers a set, so it cannot witness a repeat" }
        $want.Add($rowText)
    }
    return [pscustomobject]@{ Name = [string]$form.Items[1].Text; Program = $program; Query = $query; Arity = $arity; Want = $want }
}

# The whole .lp file: a header, the program, the rows expected, the judge.
# LF line ends and ASCII only, so the text compares the same on any machine.
function Get-ClingoFileText([string]$fileName, $about, $proof, $want, [bool]$isControl) {
    $verdict = 'agrees'
    if ($isControl) { $verdict = 'caught' }
    $xs = (1..$proof.Arity | ForEach-Object { 'X' + $_ }) -join ','
    $q = $proof.Query
    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add("% $fileName")
    foreach ($line in $about) { $lines.Add('% ' + [regex]::Replace([string]$line, '[^\x20-\x7E]', '?')) }
    $lines.Add('% Generated by tools/proofs_lp.ps1; do not edit by hand. tools/check_proofs.ps1')
    $lines.Add('% fails when this file no longer matches its proof.')
    $lines.Add("% Run:  clingo $fileName")
    $lines.Add("% Expect:  vla_verdict($verdict)")
    $lines.Add('')
    $lines.Add('% --- the program ---')
    foreach ($pl in $proof.Program) { $lines.Add($pl) }
    $lines.Add('')
    if ($want.Count -gt 0) {
        $lines.Add('% --- the rows the proof expects, in any order ---')
        foreach ($w in $want) { $lines.Add("vla_want($w).") }
    } else {
        $lines.Add('% --- the rows the proof expects: none ---')
    }
    $lines.Add('')
    if ($want.Count -gt 0) {
        $lines.Add('% --- the judge: a row clingo derives that the proof does not list is extra,')
        $lines.Add('% a row the proof lists that clingo does not derive is missing ---')
    } else {
        $lines.Add('% --- the judge: the proof expects no row, so any row clingo derives is extra ---')
    }
    if ($want.Count -gt 0) {
        $lines.Add("vla_extra($xs) :- $q($xs), not vla_want($xs).")
        $lines.Add("vla_missing($xs) :- vla_want($xs), not $q($xs).")
        $lines.Add("vla_differs :- vla_extra($xs).")
        $lines.Add("vla_differs :- vla_missing($xs).")
    } else {
        $lines.Add("vla_extra($xs) :- $q($xs).")
        $lines.Add("vla_differs :- vla_extra($xs).")
    }
    if ($isControl) {
        $lines.Add('vla_verdict(caught) :- vla_differs.')
        $lines.Add('vla_verdict(missed) :- not vla_differs.')
    } else {
        $lines.Add('vla_verdict(agrees) :- not vla_differs.')
        $lines.Add('vla_verdict(differs) :- vla_differs.')
    }
    $lines.Add('#show vla_verdict/1.')
    $lines.Add("#show vla_extra/$($proof.Arity).")
    if ($want.Count -gt 0) { $lines.Add("#show vla_missing/$($proof.Arity).") }
    return (($lines -join "`n") + "`n")
}

# Every proof in one file's forms -> Files (Name, Text), in the corpus's
# order, then its three controls; Exported, the proofs written; Refusals,
# the refusal proofs passed over; Skipped (Proof, Reason), the rest.
function ConvertTo-ClingoExport($forms, [string]$fileBase, [string]$label) {
    $files = New-Object System.Collections.Generic.List[object]
    $skipped = New-Object System.Collections.Generic.List[object]
    $names = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    $refusals = 0
    $exported = 0
    $base = $null
    foreach ($form in $forms) {
        # A form without a head and a quoted name is check_proofs.ps1's to report.
        if ($form.Kind -ne 'list') { continue }
        if ($form.Items.Count -lt 2) { continue }
        if ($form.Items[1].Kind -ne 'str') { continue }
        $head = Get-ProofHeadWord $form
        $name = [string]$form.Items[1].Text
        if ($head -cne 'test-datalog') {
            $skipped.Add([pscustomobject]@{ Proof = $name; Reason = "($head ...) - clingo is an oracle for DATALOG only" })
            continue
        }
        $proof = $null
        try {
            $proof = ConvertTo-ClingoProof $form
        } catch {
            $msg = $_.Exception.Message
            if (-not $msg.StartsWith('clingo-declines: ')) { throw }
            $skipped.Add([pscustomobject]@{ Proof = $name; Reason = $msg.Substring(17) })
            continue
        }
        if ($null -eq $proof) { $refusals++; continue }
        $fileName = 'proof-' + $fileBase + '-' + (Get-ProofSlug $name) + '.lp'
        if (-not $names.Add($fileName)) { throw "${label}: two proofs make the one file name $fileName - rename one" }
        $about = @(
            ('The proof "' + $name + '"'),
            ('in ' + $label + ', as a clingo program. Only the SET of rows is compared:'),
            'their order and the header row are DATALOG''s own.'
        )
        $files.Add([pscustomobject]@{ Name = $fileName; Text = (Get-ClingoFileText $fileName $about $proof $proof.Want $false) })
        $exported++
        if ($null -eq $base -and $proof.Want.Count -ge 2) { $base = $proof }
    }
    if ($exported -gt 0) {
        # The controls stand on the first exported proof expecting two rows or more.
        if ($null -eq $base) { throw "${label}: no exported proof expects two rows or more, so the controls have nothing to stand on" }
        $rowsLeft = New-Object System.Collections.Generic.List[string]
        for ($k = 0; $k -lt $base.Want.Count - 1; $k++) { $rowsLeft.Add($base.Want[$k]) }
        $lastRow = $base.Want[$base.Want.Count - 1]
        $rowsMore = New-Object System.Collections.Generic.List[string]
        foreach ($w in $base.Want) { $rowsMore.Add($w) }
        $absent = (1..$base.Arity | ForEach-Object { 'vla_not_an_answer' }) -join ','
        $rowsMore.Add($absent)
        $rowsNone = New-Object System.Collections.Generic.List[string]
        $controls = @(
            @{ Slug = 'a-row-missing'; Rows = $rowsLeft; What = ('its last expected row, (' + $lastRow + '), left out') },
            @{ Slug = 'a-row-too-many'; Rows = $rowsMore; What = ('a row no program here can derive, (' + $absent + '), added') },
            @{ Slug = 'nothing-expected'; Rows = $rowsNone; What = 'no row expected at all' }
        )
        foreach ($ct in $controls) {
            $fileName = 'proof-' + $fileBase + '-control-' + $ct.Slug + '.lp'
            if (-not $names.Add($fileName)) { throw "${label}: a proof's file name, $fileName, is a control's - rename the proof" }
            $about = @(
                ('A CONTROL, which the judge must catch: the proof "' + $base.Name + '"'),
                ('in ' + $label + ', with ' + $ct.What + ' - wrong on purpose.'),
                'A judge never shown to catch a wrong answer proves nothing.'
            )
            $files.Add([pscustomobject]@{ Name = $fileName; Text = (Get-ClingoFileText $fileName $about $base $ct.Rows $true) })
        }
    }
    return [pscustomobject]@{ Files = $files; Exported = $exported; Refusals = $refusals; Skipped = $skipped }
}
