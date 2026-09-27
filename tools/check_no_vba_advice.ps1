<#
check_no_vba_advice.ps1 - the owner's rule of 2026-09-26, made mechanical.

=== no message a Frazaro user can see tells them to use VBA ===

WHY THIS EXISTS. SOP.6's live pass loaded a whole SOP with no tag, and the
red row said: "No loaded sentence starts with 'expense' - click the 'Known
Sentences' button ... (or, in VBA, print the list with ?EnglishListPhrases in
the Immediate window)." The owner: "99% of Frazaro users will neither know nor
care about VBA, so we should remove VBA recommendations from refusal messages
entirely. VBA users can read the advanced documentation" - which is
docs/IMMEDIATE.md. The sweep that followed found seven more of the kind, and
one pointed at VlaTryValue, a dev-rig command no installed Frazaro even has.
Nothing mechanical was watching; this is what watches now.

WHAT IS CHECKED, case-sensitively, so that generated code such as
"(vlaCaption)" or the word "immediately" is not mistaken for advice:

  1. Every message template in both catalogues - VLA_Messages.bas's AddMsg
     lines and VLA_Runtime.bas's RuntimeAddMsg lines.
  2. Every string literal on a code line (not a comment) of every SHIPPED
     module, read from VLA_Build.bas's own mods array exactly as
     check_raise_ratchet.ps1 reads it, so inline text - a VlaShowError, a
     message assembled with msg & "..." - is held to the same rule. Dev-only
     modules (VLA_DevRig, the test modules) are not shipped and not scanned.

against eight shapes advice takes: the Immediate window, the VBA editor
(Alt+F11), "in VBA", an Immediate-window print (?Name), running a procedure
("with VlaSomething"), a procedure in brackets ("(EnglishResetGrammar)"),
editing a module ("in module VLA_IDE"), and asking whether a procedure ran.

A match that is genuinely not advice goes in $allowed below, with the reason,
where a reviewer can see it. Do not widen a pattern to let a message through:
the message is what is wrong. The fix is a ribbon button by its real name,
or plain words, or nothing - a refusal may simply say what happened.

House style (tools/check_*.ps1): PowerShell 5.1, host-independent, a
hardcoded and reviewable baseline, never wired into VlaSelfTest. Exit 0
clean, exit 1 with every failure listed.

Usage:  powershell -File tools\check_no_vba_advice.ps1
#>
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
$srcDir   = Join-Path $repoRoot 'src'

$advice = @(
    @{ Name = 'the Immediate window';           Rx = 'Immediate window|Immediate pane' },
    @{ Name = 'the VBA editor';                 Rx = 'VBA editor|Visual Basic Editor|Alt\+F11|\bVBE\b' },
    @{ Name = '"in VBA"';                       Rx = '\bin VBA\b' },
    @{ Name = 'an Immediate-window print';      Rx = '\?[A-Z][A-Za-z0-9_]+' },
    @{ Name = 'running a VBA procedure';        Rx = '\b(run|call|with|use|using|try)\s+(Vla|English|Ide)[A-Z]\w*' },
    @{ Name = 'a VBA procedure in brackets';    Rx = '\((Vla|English|Ide)[A-Z]\w*\)' },
    @{ Name = 'editing a module';               Rx = '\bin module\s+[A-Z]|\bmodule VLA_' },
    @{ Name = 'whether a VBA procedure ran';    Rx = '\bwas (Vla|English|Ide)[A-Z]\w* run\b' }
)

# Reviewed exceptions: "<file>|<exact text that matched>" = "why it is not advice".
# Empty on purpose - every match found when this check was written was advice.
$allowed = @{}

# --- the shipped set: VLA_Build.bas's own mods array ---
$buildText = Get-Content -LiteralPath (Join-Path $srcDir 'VLA_Build.bas') -Raw
$modsMatch = [regex]::Match($buildText, 'mods\s*=\s*Array\(([^)]*)\)')
if (-not $modsMatch.Success) { Write-Error "No 'mods = Array(...)' line in VLA_Build.bas - its shipped-module list shape changed; update this script." }
$modNames = @([regex]::Matches($modsMatch.Groups[1].Value, '"([^"]*)"') | ForEach-Object { $_.Groups[1].Value })
if ($modNames.Count -eq 0) { Write-Error 'Parsed zero module names out of the mods array - a regex mismatch, not an empty ship list.' }

function Get-ModulePath([string]$name) {
    foreach ($ext in '.bas', '.cls', '.frm') {
        $p = Join-Path $srcDir "$name$ext"
        if (Test-Path -LiteralPath $p) { return $p }
    }
    Write-Error "Shipped module '$name' has no .bas/.cls/.frm file in src\ - the mods array is stale."
}

# Physical lines joined across VBA's " _" continuations, each with the line
# number it started on; comment lines dropped.
function Get-CodeStatements([string]$path) {
    $out = New-Object System.Collections.Generic.List[object]
    $acc = ''; $startNo = 0; $no = 0
    foreach ($line in [System.IO.File]::ReadAllLines($path)) {
        $no++
        if ($acc.Length -eq 0) {
            if ($line -match "^\s*'") { continue }
            $startNo = $no
        }
        if ($line -match ' _$') { $acc += $line.Substring(0, $line.Length - 2) + ' '; continue }
        $acc += $line
        $out.Add([pscustomobject]@{ No = $startNo; Text = $acc })
        $acc = ''
    }
    return , $out
}

function Get-Literals([string]$text) {
    return @([regex]::Matches($text, '"((?:[^"]|"")*)"') | ForEach-Object { $_.Groups[1].Value.Replace('""', '"') })
}

$failures = New-Object System.Collections.Generic.List[string]
$scanned = 0
function Test-Text([string]$file, [int]$lineNo, [string]$text) {
    foreach ($a in $advice) {
        foreach ($m in [regex]::Matches($text, $a.Rx)) {
            $key = "$file|$($m.Value)"
            if ($allowed.ContainsKey($key)) { continue }
            $shown = if ($text.Length -gt 150) { $text.Substring(0, 150) + '...' } else { $text }
            $script:failures.Add("  ${file}:$lineNo  [$($a.Name): '$($m.Value)']  $shown")
        }
    }
}

# 1. both catalogues' templates
foreach ($cat in @(@{ File = 'VLA_Messages.bas'; Rx = '^\s*AddMsg\s+m,' }, @{ File = 'VLA_Runtime.bas'; Rx = '^\s*RuntimeAddMsg\s+m,' })) {
    $n = 0
    foreach ($st in (Get-CodeStatements (Join-Path $srcDir $cat.File))) {
        if ($st.Text -notmatch $cat.Rx) { continue }
        $lits = Get-Literals $st.Text
        if ($lits.Count -eq 0) { continue }
        Test-Text $cat.File $st.No $lits[$lits.Count - 1]
        $n++
    }
    Write-Output ("  {0,-20} {1} message templates" -f $cat.File, $n)
    $scanned += $n
}

# 2. every string literal in every shipped module
$litCount = 0
foreach ($name in $modNames) {
    $path = Get-ModulePath $name
    $file = Split-Path -Leaf $path
    foreach ($st in (Get-CodeStatements $path)) {
        if ($st.Text -match '^\s*(AddMsg|RuntimeAddMsg)\s+m,') { continue }   # already read as a template
        foreach ($lit in (Get-Literals $st.Text)) { Test-Text $file $st.No $lit; $litCount++ }
    }
}
Write-Output ("  {0} shipped modules, {1} string literals" -f $modNames.Count, $litCount)

if ($failures.Count -gt 0) {
    Write-Output ''
    Write-Output "=== CHECK: $($failures.Count) place(s) tell a Frazaro user to use VBA ==="
    $failures | ForEach-Object { Write-Output $_ }
    Write-Output ''
    Write-Output 'Say what happened in plain words, or name the ribbon button that helps (by its real label). The commands themselves belong in docs/IMMEDIATE.md.'
    exit 1
}
Write-Output "=== CHECK: clean - no message a user can see recommends VBA ($scanned templates, $litCount literals) ==="
exit 0
