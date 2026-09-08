<#
check_word_automation_security.ps1 - SEC.13's mechanical pin.

SEC.13: Word's `Application.AutomationSecurity` default under automation is
`msoAutomationSecurityLow`, so a `Documents.Open` with no guard runs an
opened document's `AutoOpen`/`Document_Open` macro silently. Frazaro reads
Word documents on a live, one-click path (*Import Program File...* on both
the menu and the ribbon -> `EnglishIdeImport` -> `ImportFromPath` -> 
`ReadWordFile`), so this is not a hypothetical surface.

WHY A STATIC SCAN AND NOT A TEST: the pure suite does no COM and cannot open
Word at all, and a host test that launches Office is exactly what this
project does not do casually (feedback: the owner runs the live passes).
That leaves this property with no automated coverage whatsoever - a future
edit to `ReadWordFile`, or a SECOND Word-opening site, would drop the guard
with nothing to notice. `SOP.1` is a filed roadmap item that adds more Word
intake, so this is the first of a family of call sites, not one line forever;
that is what tips a one-line fix over into deserving its own pin.

WHAT IS CHECKED: every non-comment source line containing `Documents.Open`
must sit in a procedure that ALSO sets `AutomationSecurity = 3`
(msoAutomationSecurityForceDisable) on a non-comment line EARLIER in that
same procedure. Same-procedure and earlier-line are both deliberate: a guard
set in a caller can be bypassed by a second caller, and a guard set after
the Open is no guard at all.

WHAT IS NOT CHECKED: that the prior value is restored afterwards (a
correctness/politeness property, not a security one - a missed restore
leaves the user's Word MORE restrictive, never less), and Excel's own
`Workbooks.Open`, which is a real but separate surface belonging to SEC.8's
mark-of-the-web scope, not this item's.

SCANNED SET: every source file under src/, NOT just the shipped modules
`check_raise_ratchet.ps1` reads from `VLA_Build.bas`. A deliberate
difference: an unguarded `Documents.Open` is worth catching in a dev-only or
not-yet-shipped module too, before it ever reaches the build list.

BASELINE: hand-maintained below, one entry per known site, in the same
reviewable shape as check_raise_ratchet.ps1's `$ceilings` and
check_id_registry.ps1's `$retired`. A new site is not a failure by itself -
it fails only if unguarded - but it is reported loudly so that adding one
stays a deliberate, reviewed act.

NOT wired into VlaSelfTest, same reasoning F.12/F.14/AS.8 gave for the other
tools/check_*.ps1 scans: this is a version-close step a human runs, not a
per-run gate.

Usage:  pwsh -File tools/check_word_automation_security.ps1
Exit code: 0 if every Documents.Open site is guarded; 1 if any is not.
#>

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$srcDir   = Join-Path $repoRoot 'src'

# --- Known sites, hand-maintained. "<Module>::<Procedure>". ---
# 2026-09-08, SEC.13: one site. ReadWordFile is the only place Frazaro has
# ever opened a Word document; the guard landed with this baseline.
$baseline = @(
    'VLA_IDE::ReadWordFile'
)

$procPattern  = '^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+([A-Za-z_][A-Za-z0-9_]*)'
$openPattern  = 'Documents\s*\.\s*Open'
$guardPattern = 'AutomationSecurity\s*=\s*3\b'

function Test-IsComment([string]$line) {
    $t = $line.TrimStart()
    return ($t.Length -eq 0 -or $t.StartsWith("'"))
}

$files = Get-ChildItem -LiteralPath $srcDir -File |
         Where-Object { $_.Extension -in @('.bas', '.cls', '.frm') } |
         Sort-Object Name

$failed = New-Object System.Collections.Generic.List[string]
$found  = New-Object System.Collections.Generic.List[string]

Write-Output '=== WORD AutomationSecurity PIN (SEC.13) ==='
Write-Output "Source files scanned: $($files.Count) (every .bas/.cls/.frm under src/)"
Write-Output ''

foreach ($f in $files) {
    $module = [IO.Path]::GetFileNameWithoutExtension($f.Name)
    $lines  = [IO.File]::ReadAllText($f.FullName) -split "`r?`n"

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

        if ($line -notmatch $openPattern) { continue }

        $site = "$module::$currentProc"
        $found.Add($site)

        # Guard must appear on a non-comment line earlier in THIS procedure.
        $guarded = $false
        for ($j = $procStartIdx; $j -lt $i; $j++) {
            if (Test-IsComment $lines[$j]) { continue }
            if ($lines[$j] -match $guardPattern) { $guarded = $true; break }
        }

        $known = if ($baseline -contains $site) { 'known' } else { 'NEW SITE' }
        if ($guarded) {
            Write-Output ("  {0,-34} line {1,5}  guarded    ({2})" -f $site, ($i + 1), $known)
        } else {
            Write-Output ("  {0,-34} line {1,5}  UNGUARDED  ({2})  FAIL" -f $site, ($i + 1), $known)
            $failed.Add("$site (src/$($f.Name):$($i + 1))")
        }
    }
}

if ($found.Count -eq 0) {
    Write-Output '  (no Documents.Open site found anywhere under src/)'
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
    Write-Output '=== NEW Documents.Open SITE(S) NOT IN THE BASELINE ==='
    $newSites | ForEach-Object { Write-Output "  $_" }
    Write-Output 'Guarded or not, adding a Word-opening site is a reviewable act: add it to `$baseline above and say why in the commit message.'
}

Write-Output ''
if ($failed.Count -eq 0) {
    Write-Output "=== CHECK: clean - all $($found.Count) Documents.Open site(s) set AutomationSecurity = 3 first ==="
} else {
    Write-Output "=== CHECK: $($failed.Count) unguarded Documents.Open site(s) ==="
    $failed | ForEach-Object { Write-Output "  $_" }
    Write-Output 'Set "wordApp.AutomationSecurity = 3" (msoAutomationSecurityForceDisable - the literal, since Word here is late-bound and the Office Object Library is not a checked reference) on a line before the Open, inside the same procedure. Without it Word''s automation default is msoAutomationSecurityLow and an opened .docm''s AutoOpen/Document_Open macro runs silently. See SEC.13 in docs/BETA_ROADMAP1.md.'
}
exit ([Math]::Min($failed.Count, 1))
