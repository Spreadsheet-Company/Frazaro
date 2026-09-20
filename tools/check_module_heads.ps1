<#
=== every source module still begins with its own header line ===

WHY THIS EXISTS, and it is a fresh scar rather than a hypothetical.
While adding a message improvement during OPTIMIZE.1's live pass, a
PowerShell one-liner meant to replace a table inside
`src\VLA_Tests_Query.bas` matched nothing, called `Write-Error` - which,
with the default `$ErrorActionPreference`, does not stop a script - and
carried on to insert 180 generated lines at offset 0, ABOVE
`Attribute VB_Name`. The file was then a module with no header.

Every one of the 23 checks then passed, including the one that counts the
very lines that had been duplicated: it reads them from inside the
function it expects them in, so it saw the right number in the right
place and never looked at the top of the file. The defect was found only
because a DIFFERENT count, printed for an unrelated reason, read double.

That is the worst shape a defect can have here: silent to every
mechanical check, and cashed in later as a VBIDE import failure or a
compile error with no obvious cause - on the owner's machine, in the
owner's time, in the middle of a live pass. It costs fifteen lines to
make it impossible.

WHAT IS CHECKED, over every file under src\:

  1. `.bas` - the FIRST line is exactly `Attribute VB_Name = "<basename>"`.
     The name must match the file's own name, so a module renamed on disk
     but not inside cannot ship (VBComponents.Import takes the name from
     the ATTRIBUTE, not from the path, so the two disagreeing is how a
     module quietly imports under the wrong name).
  2. `.cls` - the first line is `VERSION 1.0 CLASS`, and an
     `Attribute VB_Name = "<basename>"` appears in its header block.
  3. `.frm` - the first line begins `VERSION `.
  4. `Attribute VB_Name` appears EXACTLY ONCE in any of them. More than
     one is the signature of exactly the accident above.
  5. `.bas` and `.cls` carry `Option Explicit`. Without it the whole
     class of "assigned but never declared" defects compiles silently.

NOT CHECKED: anything about the code below the header - that is what
every other check in this directory is for.

House shape, per the standing convention for this project's static
scans: PowerShell, host-independent, hardcoded and reviewable baseline,
never wired into VlaSelfTest.

Usage:  powershell -File tools\check_module_heads.ps1
Exit 0 clean, exit 1 with every failure listed.
#>

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$srcDir = Join-Path $root 'src'
$failures = New-Object System.Collections.Generic.List[string]

Write-Output '=== SOURCE MODULE HEADERS (src\*.bas, *.cls, *.frm) ==='

$files = @(Get-ChildItem -LiteralPath $srcDir -File | Where-Object { $_.Extension -in @('.bas', '.cls', '.frm') } | Sort-Object Name)
if ($files.Count -eq 0) { Write-Error "No source modules found under $srcDir" }

foreach ($f in $files) {
    $text = [System.IO.File]::ReadAllText($f.FullName)
    $lines = $text -split "`r`n|`n"
    $first = ''
    if ($lines.Count -gt 0) { $first = $lines[0] }
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($f.Name)
    $wantAttr = 'Attribute VB_Name = "' + $baseName + '"'
    $problems = New-Object System.Collections.Generic.List[string]

    $attrCount = ([regex]::Matches($text, 'Attribute VB_Name')).Count
    if ($attrCount -ne 1) {
        $problems.Add("'Attribute VB_Name' appears $attrCount time(s), expected exactly 1")
    }

    switch ($f.Extension) {
        '.bas' {
            if ($first -ne $wantAttr) {
                $problems.Add("first line is '$first', expected '$wantAttr'")
            }
        }
        '.cls' {
            if ($first -ne 'VERSION 1.0 CLASS') {
                $problems.Add("first line is '$first', expected 'VERSION 1.0 CLASS'")
            }
            if ($text.IndexOf($wantAttr) -lt 0) {
                $problems.Add("no '$wantAttr' anywhere in the file")
            }
        }
        '.frm' {
            if (-not $first.StartsWith('VERSION ')) {
                $problems.Add("first line is '$first', expected one beginning 'VERSION '")
            }
        }
    }

    if ($f.Extension -in @('.bas', '.cls')) {
        if ($text -notmatch '(?m)^Option Explicit\s*$') {
            $problems.Add("no 'Option Explicit'")
        }
    }

    if ($problems.Count -eq 0) {
        Write-Output ("  ok    {0}" -f $f.Name)
    } else {
        foreach ($pb in $problems) {
            Write-Output ("  FAIL  {0}: {1}" -f $f.Name, $pb)
            $failures.Add("$($f.Name): $pb")
        }
    }
}

Write-Output ''
if ($failures.Count -eq 0) {
    Write-Output ("=== CHECK: clean - all {0} source modules carry their own header, once, with Option Explicit ===" -f $files.Count)
    exit 0
} else {
    foreach ($f in $failures) { Write-Output "FAIL: $f" }
    Write-Output ''
    Write-Output "=== CHECK: FAILED - $($failures.Count) problem(s) ==="
    Write-Output "A module whose header is missing, duplicated, or names the wrong module does not import into VBA - and no other check in this directory looks at the top of the file. If a generated block landed above the header, delete everything before 'Attribute VB_Name'."
    exit 1
}
