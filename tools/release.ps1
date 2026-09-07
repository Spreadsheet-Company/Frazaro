<#
release.ps1 - the "Push" step of Scope -> Implement -> Test -> Build -> Push.

One command cuts a release: verifies the tree is in a releasable state, runs
every static ratchet, tags, pushes main and the tag, and publishes a GitHub
release carrying the two edition add-ins with notes taken from
docs/RELEASES.md. Cadence (owner decision, 2026-09-07): a 0.5.N patch release
at the end of each working day, a 0.N.0 minor release at the end of each
week, security and safety fixes front-loaded into the patches. SD-14's
MAJOR.MINOR.PATCH numbering is unchanged; this only fixes the tempo.

What it verifies before touching the remote, in order:
  1. Run from the repository root, on main, with a clean tree, in sync.
  2. VLA_RELEASE_VERSION in src/VLA.bas equals -Version (DI.3a: bumped by
     hand, at release, and nowhere else).
  3. Tag v<Version> does not already exist, locally or on origin.
  4. docs/RELEASES.md has a "## <Version>" section - the release notes are
     written before the release, never after.
  5. Every tools/check_*.ps1 exits 0.
  6. Frazaro_English.xlam and Frazaro_Espanol.xlam exist beside this
     repository's root and are newer than every file under src/ and
     scripts/ - i.e. VlaBuildAddin ran after the last source edit.
  7. -Locked was passed: the owner's attestation that both .xlam VBA
     projects were locked in the VBE (DEPLOY.md, "Building the add-in",
     step 4). VBA offers no API to check this, so the switch IS the check.

Usage:
  powershell -File tools\release.ps1 -Version 0.5.1 -Locked
  powershell -File tools\release.ps1 -Version 0.5.1 -Locked -DryRun
#>
param(
    [Parameter(Mandatory = $true)][string]$Version,
    [switch]$Locked,
    [switch]$DryRun,
    [string]$Repo = 'Spreadsheet-Company/Frazaro'
)
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
Set-Location $root
$fail = New-Object System.Collections.Generic.List[string]

if ($Version -notmatch '^\d+\.\d+\.\d+$') { $fail.Add("version '$Version' is not MAJOR.MINOR.PATCH (SD-14)") }
$tag = "v$Version"

# 1. repository state
$branch = (git rev-parse --abbrev-ref HEAD).Trim()
if ($branch -ne 'main') { $fail.Add("on branch '$branch', not main") }
$dirty = git status --porcelain
if ($dirty) { $fail.Add("working tree not clean:`n" + ($dirty -join "`n")) }
git fetch -q origin
$ahead = (git rev-list --count origin/main..HEAD).Trim()
$behind = (git rev-list --count HEAD..origin/main).Trim()
if ($behind -ne '0') { $fail.Add("main is $behind commit(s) behind origin/main - pull first") }

# 2. version constant
$vlaBas = Get-Content (Join-Path $root 'src\VLA.bas')
$constLine = $vlaBas | Where-Object { $_ -match 'Public Const VLA_RELEASE_VERSION As String = "([^"]+)"' } | Select-Object -First 1
$inSource = if ($constLine -match '"([^"]+)"') { $Matches[1] } else { '' }
if ($inSource -ne $Version) { $fail.Add("VLA_RELEASE_VERSION in src/VLA.bas is '$inSource', not '$Version' - bump it, rebuild, retest") }

# 3. tag not taken
if (git tag -l $tag) { $fail.Add("tag $tag already exists locally") }
if (git ls-remote --tags origin $tag) { $fail.Add("tag $tag already exists on origin") }

# 4. release notes section
$notesPath = Join-Path $root 'docs\RELEASES.md'
$notes = @()
if (Test-Path $notesPath) {
    $lines = Get-Content $notesPath
    $start = -1
    for ($i = 0; $i -lt $lines.Count; $i++) { if ($lines[$i] -match ('^## ' + [regex]::Escape($Version) + '(\s|$)')) { $start = $i; break } }
    if ($start -lt 0) { $fail.Add("docs/RELEASES.md has no '## $Version' section - write the notes first") }
    else {
        $end = $lines.Count
        for ($j = $start + 1; $j -lt $lines.Count; $j++) { if ($lines[$j] -match '^## ') { $end = $j; break } }
        $notes = $lines[($start + 1)..($end - 1)]
        $body = ($notes -join "`n").Trim()
        # The placeholder RELEASES.md keeps for the next version ("*(next: write
        # this before running the release)*") is longer than the emptiness
        # floor, so it is refused by name - otherwise it would publish as the
        # release's own notes.
        if ($body.Length -lt 40 -or $body -match '\(next:') { $fail.Add("the '## $Version' section in docs/RELEASES.md is empty or still the placeholder - write the notes first") }
    }
} else { $fail.Add('docs/RELEASES.md missing') }

# 5. static ratchets. A check that writes to stderr (Write-Error) must not
# become a terminating error here - PowerShell 5.1 wraps native stderr in
# ErrorRecords - so the preference is relaxed for the call and the verdict
# is the exit code alone.
$prevEap = $ErrorActionPreference
$ErrorActionPreference = 'Continue'
foreach ($chk in Get-ChildItem (Join-Path $root 'tools') -Filter 'check_*.ps1' | Sort-Object Name) {
    $out = & powershell -NoProfile -ExecutionPolicy Bypass -File $chk.FullName 2>&1 | ForEach-Object { "$_" }
    $code = $LASTEXITCODE
    if ($code -ne 0) { $fail.Add("$($chk.Name) failed (exit $code):`n" + (($out | Select-Object -Last 8) -join "`n")) }
    else { Write-Host ("  ok  " + $chk.Name) }
}
$ErrorActionPreference = $prevEap

# 6. build artifacts present and fresh
$assets = @('Frazaro_English.xlam', 'Frazaro_Espanol.xlam') | ForEach-Object { Join-Path $root $_ }
$newestSource = Get-ChildItem (Join-Path $root 'src'), (Join-Path $root 'scripts') -Recurse -File | Sort-Object LastWriteTime -Descending | Select-Object -First 1
foreach ($a in $assets) {
    if (-not (Test-Path $a)) { $fail.Add("missing build artifact: $a (run VlaBuildAddin)"); continue }
    $age = Get-Item $a
    if ($newestSource -and $age.LastWriteTime -lt $newestSource.LastWriteTime) {
        $fail.Add("$($age.Name) ($($age.LastWriteTime)) is older than $($newestSource.FullName) ($($newestSource.LastWriteTime)) - rebuild")
    }
}

# 7. the lock attestation
if (-not $Locked) { $fail.Add('pass -Locked once both .xlam VBA projects are locked in the VBE (DEPLOY.md, Building the add-in, step 4); VBA has no API to verify this, so the switch is the record') }

if ($fail.Count -gt 0) {
    Write-Host "NOT RELEASED - $($fail.Count) problem(s):"
    $fail | ForEach-Object { Write-Host "  - $_" }
    exit 1
}

Write-Host "Preflight clean for $tag ($ahead local commit(s) to push)."
$notesFile = Join-Path $env:TEMP "frazaro-release-notes-$Version.md"
Set-Content -Path $notesFile -Value ($notes -join "`n") -Encoding UTF8

if ($DryRun) {
    Write-Host "DRY RUN - would: git push origin main; git tag -a $tag; git push origin $tag; gh release create $tag <assets> --notes-file $notesFile"
    exit 0
}

git push origin main
git tag -a $tag -m "Frazaro $Version"
git push origin $tag
& gh release create $tag @assets --repo $Repo --title "Frazaro $Version" --notes-file $notesFile
if ($LASTEXITCODE -ne 0) { Write-Host 'gh release create failed - tag and main are pushed; re-run gh release create by hand'; exit 1 }
Write-Host "Released ${tag}: https://github.com/$Repo/releases/tag/$tag"
exit 0
