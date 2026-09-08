<#
check_sec8_provenance_gate.ps1 - SEC.8's mechanical pin.

SEC.8: a Frazaro program lives in CELLS, not in a VBA project, so Office's
2022 block on internet-sourced macros has never applied to it. A plain .xlsx
mailed from outside runs its program with the ADD-IN's privilege after the
ordinary "Enable Editing" click. `VLA_Provenance` closes that by reading the
host workbook's Mark-of-the-Web and refusing external-effect members. This
script pins the two properties that make it real, neither of which any test
in this project can reach.

WHY A STATIC SCAN AND NOT A TEST: the pure suite does no COM and cannot
fabricate a Zone.Identifier alternate data stream, so it cannot prove that a
real marked workbook is refused. (It does cover the whole DECISION - parse,
path classification, policy, and the guard's own refusal through a memo seam;
see VLA_Tests.bas's TestSec8Provenance.) A host test that mails itself a
marked workbook and clicks through Protected View is not something this
project does. That leaves the CALL SITES with no automated coverage at all:
a future edit could drop a guard, or add a ninth external-effect member, and
nothing would notice. Same reasoning check_word_automation_security.ps1 gave
for SEC.13, and the same shape.

WHAT IS CHECKED, property 1 - every known external-effect site is guarded.
A guard is a `VlaProvenanceGuard...` call on a non-comment line EARLIER in
the same procedure, with NO intervening `Case ` line. Both halves matter:
  * same procedure - a guard in a caller can be bypassed by a second caller;
  * no intervening `Case` - `DynamicCall` and `DynamicNamedCall` each hold
    several gated members in ONE procedure, so a single guard high in the
    procedure must NOT be allowed to vouch for every Case below it. This is
    the rule that makes the check site-by-site rather than procedure-wide,
    and it is also why `VlaSendMail` (which has no Case at all, and guards at
    the top of the Sub) passes under the same rule with no special case.

WHAT IS CHECKED, property 2 - the capture still happens. `VLA_IDE`'s
`CaptureHost` must call `VlaProvenanceCapture`. The guards read a memo bound
to the workbook that CARRIED the program; if capture stops happening, every
guard silently falls through to its uncaptured branch and the gate is gone
with no other symptom. This is the single most load-bearing line in the item
and it is one line, in a procedure that exists for an unrelated reason (D1's
ambient-state capture), so it is exactly the kind of line a later refactor
deletes without noticing.

WHAT IS NOT CHECKED: that the Mark-of-the-Web is actually READ correctly (an
NTFS/COM property - live test), that Excel does not strip the mark when the
person clicks "Enable Editing" (live test, and the finding that decides
whether this item works at all), and whether the guarded LIST is complete -
a member with external effect that nobody ever added here is invisible to a
scan that works from a baseline. That last gap is the honest one: the
baseline below is a human judgement about which members reach outside the
workbook, re-derived by reading DynamicCall/DynamicNamedCall, not something
this script can discover.

SCANNED SET: the SHIPPED modules, read from `VLA_Build.bas`'s own `mods`
array the way check_raise_ratchet.ps1 does - deliberately NOT all of src/.
`VLA_Tests_Host.bas` legitimately calls `.Close SaveChanges:=False` on its
own scratch workbooks as direct VBA in the harness, which is not interpreter
dispatch and must not be gated; scanning only what ships draws that line in
the right place.

BASELINE: hand-maintained below, one entry per known site, in the same
reviewable shape as check_raise_ratchet.ps1's `$ceilings` and
check_word_automation_security.ps1's own. A NEW site is not a failure by
itself - it fails only if unguarded - but it is reported loudly, because
adding a member that reaches outside the workbook should always be a
deliberate, reviewed act.

NOT wired into VlaSelfTest, same reasoning F.12/F.14/AS.8 and SEC.13 gave for
the other tools/check_*.ps1 scans: this is a version-close step a human runs,
not a per-run gate.

Usage:  pwsh -File tools/check_sec8_provenance_gate.ps1
Exit code: 0 if every known external-effect site is guarded AND CaptureHost
still captures; 1 otherwise.
#>

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$srcDir   = Join-Path $repoRoot 'src'

# --- Shipped set: VLA_Build.bas's own mods array, not a hand-copied list ---
$buildFile = Join-Path $srcDir 'VLA_Build.bas'
$buildText = [IO.File]::ReadAllText($buildFile)
$modsMatch = [regex]::Match($buildText, 'mods\s*=\s*Array\(([^)]*)\)')
if (-not $modsMatch.Success) {
    Write-Error "No 'mods = Array(...)' line found in $buildFile - VLA_Build.bas's shipped-module list shape may have changed; update this script's `$modsMatch pattern."
}
$modNames = [regex]::Matches($modsMatch.Groups[1].Value, '"([^"]*)"') |
            ForEach-Object { $_.Groups[1].Value }
if ($modNames.Count -eq 0) {
    Write-Error "Parsed zero module names out of the mods array in $buildFile - regex mismatch, not an empty ship list."
}
if ($modNames -notcontains 'VLA_Provenance') {
    Write-Error "VLA_Provenance is not in VLA_Build.bas's mods array - an add-in built from this tree would ship WITHOUT the provenance gate, and every guard below would be a call into a module that is not there. Add it to the mods array."
}

# --- Known external-effect sites, hand-maintained. "<Module>::<Procedure>::<label>" ---
# 2026-09-08, SEC.8: nine sites. Three kinds, kept distinct in the roadmap
# entry rather than blurred: EGRESS (content leaves this machine - Open,
# SaveAs, SaveCopyAs, PrintOut, ExportAsFixedFormat, the Outlook draft) and
# DESTRUCTIVE-TO-LOCAL-DATA (Close-with-save, Protect, Unprotect - no
# attacker-chosen path, but a password-protect an internet-sourced program
# applies locks a person out of their own sheet).
$baseline = @(
    'VLA_Interpreter::DynamicCall::Workbooks.Open'
    'VLA_Interpreter::DynamicCall::Workbook.SaveAs'
    'VLA_Interpreter::DynamicCall::Workbook.SaveCopyAs'
    'VLA_Interpreter::DynamicNamedCall::Worksheet.Protect'
    'VLA_Interpreter::DynamicNamedCall::Worksheet.Unprotect'
    'VLA_Interpreter::DynamicNamedCall::Workbook.Close'
    'VLA_Interpreter::DynamicNamedCall::Worksheet.PrintOut'
    'VLA_Interpreter::DynamicNamedCall::Worksheet.ExportAsFixedFormat'
    'VLA_Runtime::VlaSendMail::Outlook.CreateItem'
)

# --- Explicitly NOT gated, hand-maintained, one entry per site plus its reason ---
# SEC.8 gates what a PROGRAM can reach through interpreter dispatch. A site
# that only a person can reach from the menu is a different thing, and gating
# it would refuse the person their own command. Each exemption is written
# down here rather than left to fall outside a pattern by luck, so that
# "this one is fine" stays a claim someone made and can re-check.
$exempt = @{
    'VLA_IDE::EnglishIdeExportSkeleton::Workbook.Close' =
        'Export standalone copy: a menu command the PERSON invokes, not anything a cell-carried program can dispatch. It closes the copy IT just created, and already sits behind VlaHasVbProjectTrust() - a strictly stronger gate than a Mark-of-the-Web check. Found by this script on its first run, which is the argument for the script.'
}

# Each pattern is tied to the CALL's own shape (its argument list), not to a
# bare member name: a bare '\.Open' would also match ADODB's st.Open in
# VLA_Loader and VLA_Provenance, which are not external effect at all.
$effectPatterns = [ordered]@{
    'Workbooks.Open'                 = '\.Open\s+ArgAt\('
    'Workbook.SaveAs'                = '\.SaveAs\s+ArgAt\('
    'Workbook.SaveCopyAs'            = '\.SaveCopyAs\s+ArgAt\('
    'Worksheet.Protect'              = '\.Protect\s+Password:='
    'Worksheet.Unprotect'            = '\.Unprotect\s+Password:='
    'Workbook.Close'                 = '\.Close\s+SaveChanges:='
    'Worksheet.PrintOut'             = '\.PrintOut\s+Preview:='
    'Worksheet.ExportAsFixedFormat'  = '\.ExportAsFixedFormat\s+Type:='
    'Outlook.CreateItem'             = '\.CreateItem\(0\)'
}

$procPattern  = '^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+([A-Za-z_][A-Za-z0-9_]*)'
$casePattern  = '^\s*Case\b'
$guardPattern = 'VlaProvenanceGuard'

function Test-IsComment([string]$line) {
    $t = $line.TrimStart()
    return ($t.Length -eq 0 -or $t.StartsWith("'"))
}

$failed     = New-Object System.Collections.Generic.List[string]
$found      = New-Object System.Collections.Generic.List[string]
$exemptSeen = New-Object System.Collections.Generic.List[string]

Write-Output '=== SEC.8 PROVENANCE GATE PIN ==='
Write-Output "Shipped modules scanned: $($modNames.Count) (from VLA_Build.bas's mods array)"
Write-Output ''

foreach ($name in $modNames) {
    $file = Get-ChildItem -LiteralPath $srcDir -File |
            Where-Object { [IO.Path]::GetFileNameWithoutExtension($_.Name) -eq $name -and
                           $_.Extension -in @('.bas', '.cls', '.frm') } |
            Select-Object -First 1
    if (-not $file) {
        Write-Error "Shipped module '$name' (from VLA_Build.bas's mods array) has no .bas/.cls/.frm file in $srcDir - export it, or the mods array is stale."
    }

    $lines = [IO.File]::ReadAllText($file.FullName) -split "`r?`n"

    $currentProc  = '(module level)'
    $procStartIdx = 0

    for ($i = 0; $i -lt $lines.Count; $i++) {
        $line = $lines[$i]
        if (Test-IsComment $line) { continue }

        $m = [regex]::Match($line, $procPattern)
        if ($m.Success) {
            $currentProc  = $m.Groups[1].Value
            $procStartIdx = $i
            continue
        }

        foreach ($label in $effectPatterns.Keys) {
            if ($line -notmatch $effectPatterns[$label]) { continue }

            $site = "$name::$currentProc::$label"

            if ($exempt.ContainsKey($site)) {
                Write-Output ("  {0,-58} line {1,5}  exempt     (see `$exempt)" -f $site, ($i + 1))
                $exemptSeen.Add($site)
                continue
            }

            $found.Add($site)

            # Walk BACKWARDS from the effect line to the start of the
            # procedure. A guard counts only if it is reached before any
            # `Case` line - that is what keeps one guard from vouching for
            # every sibling Case in the same Select.
            $guarded = $false
            for ($j = $i - 1; $j -ge $procStartIdx; $j--) {
                if (Test-IsComment $lines[$j]) { continue }
                if ($lines[$j] -match $guardPattern) { $guarded = $true; break }
                if ($lines[$j] -match $casePattern)  { break }
            }

            $known = if ($baseline -contains $site) { 'known' } else { 'NEW SITE' }
            if ($guarded) {
                Write-Output ("  {0,-58} line {1,5}  guarded    ({2})" -f $site, ($i + 1), $known)
            } else {
                Write-Output ("  {0,-58} line {1,5}  UNGUARDED  ({2})  FAIL" -f $site, ($i + 1), $known)
                $failed.Add("$site (src/$($file.Name):$($i + 1))")
            }
        }
    }
}

if ($found.Count -eq 0) {
    Write-Output '  (no external-effect site found in any shipped module - see this script`s $effectPatterns)'
}

# --- Property 2: the capture that every guard depends on ---
Write-Output ''
Write-Output '--- CaptureHost still captures ---'
$ideFile = Join-Path $srcDir 'VLA_IDE.bas'
$ideLines = [IO.File]::ReadAllText($ideFile) -split "`r?`n"
$inCapture = $false
$captureFound = $false
for ($i = 0; $i -lt $ideLines.Count; $i++) {
    $line = $ideLines[$i]
    $m = [regex]::Match($line, $procPattern)
    if ($m.Success) { $inCapture = ($m.Groups[1].Value -eq 'CaptureHost'); continue }
    if (-not $inCapture) { continue }
    if (Test-IsComment $line) { continue }
    if ($line -match 'VlaProvenanceCapture') {
        Write-Output ("  VLA_IDE::CaptureHost                 line {0,5}  captures" -f ($i + 1))
        $captureFound = $true
        break
    }
}
if (-not $captureFound) {
    Write-Output '  VLA_IDE::CaptureHost                            NO CAPTURE  FAIL'
    $failed.Add('VLA_IDE::CaptureHost does not call VlaProvenanceCapture')
}

$staleExempt = $exempt.Keys | Where-Object { $exemptSeen -notcontains $_ }
if ($staleExempt) {
    Write-Output ''
    Write-Output '=== EXEMPTIONS WITH NO LIVE SITE (the code moved - drop them from `$exempt above and re-read the reason) ==='
    $staleExempt | ForEach-Object { Write-Output "  $_" }
}

$missing = $baseline | Where-Object { $found -notcontains $_ }
if ($missing) {
    Write-Output ''
    Write-Output '=== BASELINE ENTRIES WITH NO LIVE SITE (removed or renamed - drop them from `$baseline above) ==='
    $missing | ForEach-Object { Write-Output "  $_" }
}

$newSites = $found | Where-Object { $baseline -notcontains $_ } | Select-Object -Unique
if ($newSites) {
    Write-Output ''
    Write-Output '=== NEW EXTERNAL-EFFECT SITE(S) NOT IN THE BASELINE ==='
    $newSites | ForEach-Object { Write-Output "  $_" }
    Write-Output 'Guarded or not, adding a member that reaches outside the workbook is a reviewable act: add it to `$baseline above and say why in the commit message.'
}

Write-Output ''
if ($failed.Count -eq 0) {
    Write-Output "=== CHECK: clean - all $($found.Count) external-effect site(s) guarded, and CaptureHost still captures ==="
} else {
    Write-Output "=== CHECK: $($failed.Count) problem(s) ==="
    $failed | ForEach-Object { Write-Output "  $_" }
    Write-Output 'Call "VLA_Provenance.VlaProvenanceGuardCaptured ""<what the program tried to do>""" on a line inside the same Case, before the effect. Without it a workbook mailed in from outside runs that effect with the add-in''s own privilege, which is exactly what Office''s 2022 macro block removed and what SEC.8 restores. See SEC.8 in docs/BETA_ROADMAP1.md.'
}
exit ([Math]::Min($failed.Count, 1))
