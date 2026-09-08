<#
check_devrig_mods_parity.ps1 - the two module lists must agree.

This project keeps TWO independent arrays of module names, and nothing
mechanical has ever held them together:

  * `VLA_Build.bas`'s  `mods` - what a built .xlam SHIPS. The builder starts
    from `Workbooks.Add` (a fresh, empty workbook) and imports exactly these,
    so a module absent here does not exist in the add-in at all.
  * `VLA_DevRig.bas`'s `mods` - what `VlaDevReload` refreshes into the dev
    workbook. Deliberately a superset: it also covers dev-only modules
    (`VLA_Tests*`, the builder itself) that never ship.

Both files carry a comment block warning about this, and the warnings are
themselves the evidence that comments are not enough. The recorded
recurrences, every one of them the same root cause - a module named by code
that ships, in a list that does not carry it:

    F5.0   VlaFrame            LX3.0  (build list)
    IN1.0  (build list)        IN.6   VLA_Interpreter
    IN.7   VLA_Events/Sink     CLI.0  frmCLI
    LX2.0  VLA_Messages        SEC.8  VLA_Provenance

Eight. The failure mode is always the same and always late: the module sits
on disk unimported, and the FIRST live `Debug > Compile` of a call site that
references it fails with "Variable not defined" (or, for a type, "User-defined
type not defined"). It is never caught by inspection, and never by any test -
the pure suite does not build an add-in and does not reload the dev rig.

WHAT IS CHECKED

  1. Every module in the BUILD list is also in the DEVRIG list. This is the
     SEC.8 shape: it ships, so a developer editing it must be able to reload
     it, and the dev workbook must compile the same code the add-in will.

  2. Every module in the DEVRIG list is also in the BUILD list, EXCEPT the
     hand-maintained dev-only set below. This is the F5.0/IN.6 shape, and the
     more dangerous direction: the dev workbook compiles fine (the rig loaded
     it) while a freshly built add-in does not, so the gap is invisible until
     someone builds and compiles.

  3. Every name in either list has a real .bas/.cls/.frm under src/.

  4. Reported, not failed: source files under src/ that appear in NEITHER
     list. `VLA_DevRig` itself is the legitimate case (it cannot reload
     itself while running, so it is absent from its own array by design).

WHAT IS NOT CHECKED: that the shipped set is SUFFICIENT - i.e. that no
shipped module references something outside both lists entirely. That needs a
reference graph, not a set comparison. What this catches is the narrower and
historically actual failure: a module that exists, is known to one list, and
was forgotten by the other.

NOT wired into VlaSelfTest, same reasoning F.12/F.14/AS.8/SEC.13 gave for the
other tools/check_*.ps1 scans: a version-close step a human runs.

Usage:  pwsh -File tools/check_devrig_mods_parity.ps1
Exit code: 0 if the two lists agree (modulo the dev-only set); 1 otherwise.
#>

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$srcDir   = Join-Path $repoRoot 'src'

# --- Dev-only, hand-maintained: in the DEVRIG list and rightly NOT shipped ---
# One entry per module, each with a reason, in the same reviewable shape as
# check_raise_ratchet.ps1's $ceilings. Adding to this list is how you assert
# "this genuinely must not ship" - it is the one place that claim is written
# down, so it should be a deliberate edit, never a way to quiet a failure.
$devOnly = @{
    'VLA_Build'        = 'the builder itself - it constructs the add-in and has no business inside it'
    'VLA_Tests'        = 'the pure self-test suite; VlaSelfTest is a dev-workbook activity'
    'VLA_Tests_Grammar'= 'test suite'
    'VLA_Tests_Host'   = 'test suite (the host half, which drives real Excel)'
    'VLA_Tests_Query'  = 'test suite'
}

function Get-ModsArray([string]$fileName) {
    $path = Join-Path $srcDir $fileName
    if (-not (Test-Path $path)) { Write-Error "$fileName not found in $srcDir" }
    $text  = [IO.File]::ReadAllText($path)
    $match = [regex]::Match($text, 'mods\s*=\s*Array\(([^)]*)\)')
    if (-not $match.Success) {
        Write-Error "No 'mods = Array(...)' found in $fileName - its shape may have changed; update this script's pattern."
    }
    $names = [regex]::Matches($match.Groups[1].Value, '"([^"]*)"') |
             ForEach-Object { $_.Groups[1].Value }
    if ($names.Count -eq 0) { Write-Error "Parsed zero names out of $fileName's mods array - regex mismatch, not an empty list." }
    return $names
}

$shipped = Get-ModsArray 'VLA_Build.bas'
$devrig  = Get-ModsArray 'VLA_DevRig.bas'

$failed = New-Object System.Collections.Generic.List[string]

Write-Output '=== MODULE LIST PARITY (VLA_Build.bas vs VLA_DevRig.bas) ==='
Write-Output ("Shipped (VLA_Build):  {0} modules" -f $shipped.Count)
Write-Output ("Dev rig (VLA_DevRig): {0} modules" -f $devrig.Count)
Write-Output ''

# --- 1. shipped but not reloadable ---
Write-Output '--- In the BUILD list but missing from the DEVRIG list ---'
$missingFromDev = $shipped | Where-Object { $devrig -notcontains $_ }
if ($missingFromDev) {
    foreach ($m in $missingFromDev) {
        Write-Output ("  {0,-22} SHIPS, but VlaDevReload will never refresh it  FAIL" -f $m)
        $failed.Add("$m is in VLA_Build.bas's mods but not VLA_DevRig.bas's - add it to VLA_DevRig.bas")
    }
} else {
    Write-Output '  (none)'
}

# --- 2. reloadable but not shipped ---
Write-Output ''
Write-Output '--- In the DEVRIG list but missing from the BUILD list ---'
$missingFromShip = $devrig | Where-Object { $shipped -notcontains $_ }
if ($missingFromShip) {
    foreach ($m in $missingFromShip) {
        if ($devOnly.ContainsKey($m)) {
            Write-Output ("  {0,-22} dev-only: {1}" -f $m, $devOnly[$m])
        } else {
            Write-Output ("  {0,-22} the dev workbook compiles it, a BUILT ADD-IN WOULD NOT  FAIL" -f $m)
            $failed.Add("$m is in VLA_DevRig.bas's mods but not VLA_Build.bas's - either add it to VLA_Build.bas (if shipped code names it) or record it in this script's `$devOnly with a reason")
        }
    }
} else {
    Write-Output '  (none)'
}

# --- 3. every listed name has a file ---
Write-Output ''
Write-Output '--- Listed names with no source file ---'
$exts = @('.bas', '.cls', '.frm')
$anyMissingFile = $false
foreach ($m in (@($shipped) + @($devrig) | Select-Object -Unique)) {
    $hit = $exts | ForEach-Object { Join-Path $srcDir ($m + $_) } | Where-Object { Test-Path $_ }
    if (-not $hit) {
        Write-Output ("  {0,-22} named in a mods array, no .bas/.cls/.frm in src/  FAIL" -f $m)
        $failed.Add("$m is named in a mods array but has no source file in src/")
        $anyMissingFile = $true
    }
}
if (-not $anyMissingFile) { Write-Output '  (none)' }

# --- 4. informational: source files in neither list ---
Write-Output ''
Write-Output '--- Source files in NEITHER list (informational) ---'
$known = @($shipped) + @($devrig) | Select-Object -Unique
$orphans = Get-ChildItem -LiteralPath $srcDir -File |
           Where-Object { $_.Extension -in $exts } |
           ForEach-Object { [IO.Path]::GetFileNameWithoutExtension($_.Name) } |
           Where-Object { $known -notcontains $_ } |
           Sort-Object -Unique
if ($orphans) {
    foreach ($o in $orphans) {
        $note = if ($o -eq 'VLA_DevRig') { 'expected - a module cannot reload itself while running' } else { 'neither shipped nor reloaded; is it still used?' }
        Write-Output ("  {0,-22} {1}" -f $o, $note)
    }
} else {
    Write-Output '  (none)'
}

$staleDevOnly = $devOnly.Keys | Where-Object { $devrig -notcontains $_ }
if ($staleDevOnly) {
    Write-Output ''
    Write-Output '=== $devOnly ENTRIES NO LONGER IN THE DEVRIG LIST (drop them and re-read the reason) ==='
    $staleDevOnly | ForEach-Object { Write-Output "  $_" }
}

Write-Output ''
if ($failed.Count -eq 0) {
    Write-Output '=== CHECK: clean - the two module lists agree (modulo the recorded dev-only set) ==='
} else {
    Write-Output "=== CHECK: $($failed.Count) parity problem(s) ==="
    $failed | ForEach-Object { Write-Output "  $_" }
    Write-Output 'A module named by code that ships must be in BOTH arrays. Missing from VLA_Build.bas, a freshly built add-in fails Debug > Compile; missing from VLA_DevRig.bas, the dev workbook does. See the comment blocks in both files for the eight prior recurrences.'
}
exit ([Math]::Min($failed.Count, 1))
