<#
=== SEC.11 - the two hash implementations are pinned to the same vectors ===

WHY THIS EXISTS. Frazaro hashes a phrasebook's content in two languages.
VBA does it in VLA_Digest.bas (VlaSha256HexSkippingWhitespace, reached
through EnglishSourceHash) because that is where consent is decided;
PowerShell does it in check_rule_coverage.ps1 (Get-SourceHash) because
that is where an exported artifact's freshness is verified. The two MUST
produce the same digest for the same file, and until 0.5.3 the only thing
holding them together was a comment reading "Change one side and you must
change the other."

That is precisely the shape that has now bitten this project nine times
through the two module arrays - see tools/check_devrig_mods_parity.ps1's
own header, and note that when that script was finally written it found a
defect latent since 0.5.0 on its very first run. A comment is not a
mechanism. This script is the mechanism.

HOW IT PINS, and why it does not simply "run both and compare". It
cannot: there is no way to execute VBA from here without launching
Excel, which this project does not do for verification. So the pin is
built the other way round, and is stronger for it - BOTH sides are
pinned to the same third thing:

  1. A hardcoded, reviewable baseline of vectors and their SHA-256
     digests, below. The first three are the published FIPS 180-4 test
     vectors, so the baseline itself is checkable against a standard
     rather than against this codebase's own opinion.
  2. This script verifies that check_rule_coverage.ps1's REAL
     Get-SourceHash - extracted from that file and evaluated, never
     re-typed here, so there is no third twin to drift - reproduces
     them.
  3. This script verifies that the pure test modules (src/VLA_Tests*.bas)
     assert the SAME digest strings, character for character. VlaSelfTest
     runs those assertions against the VBA implementation on every pass.

So if either implementation drifts, its own side goes red: the VBA side
in VlaSelfTest, the PowerShell side here. And if someone "fixes" a
failing side by editing its expected value, step 3 goes red because the
two baselines no longer agree. There is no edit that silently separates
them. A shared bug cannot cancel out either, because the expected values
come from FIPS, not from either implementation.

WHAT IT DELIBERATELY DOES NOT DO. It does not run VBA, so it cannot
prove VLA_Digest.bas computes SHA-256 correctly - only that it is being
asserted against the right answers. Proving the VBA arithmetic is
VlaSelfTest's job, and the vectors below are chosen to make that job
real: block-boundary cases (55/56/64/65 bytes) exercise the padding
and length-encoding logic that a naive implementation gets wrong, and
the whitespace pair proves the skip filter rather than the digest.

House shape, per the standing convention for this project's static
scans: PowerShell, host-independent, hardcoded and reviewable baseline,
never wired into VlaSelfTest.

Usage:  powershell -File tools\check_hash_twin.ps1
Exit 0 clean, exit 1 with every failure listed.
#>

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$failures = New-Object System.Collections.Generic.List[string]

Write-Output '=== SEC.11 HASH TWIN PIN (VLA_Digest.bas <-> check_rule_coverage.ps1) ==='

# ---- the baseline -----------------------------------------------------
# Name, the exact ASCII input, and its SHA-256 over NON-WHITESPACE bytes
# only (which for the whitespace-free vectors is just SHA-256 of the
# input). Count is how many bytes survive the skip filter.
$vectors = @(
    @{ Name = 'fips-abc';        Text = 'abc';
       Count = 3;
       Sha = 'BA7816BF8F01CFEA414140DE5DAE2223B00361A396177A9CB410FF61F20015AD' },
    @{ Name = 'fips-448bit';     Text = 'abcdbcdecdefdefgefghfghighijhijkijkljklmklmnlmnomnopnopq';
       Count = 56;
       Sha = '248D6A61D20638B8E5C026930C3E6039A33CE45964FF2167F6ECEDD419DB06C1' },
    @{ Name = 'empty';           Text = '';
       Count = 0;
       Sha = 'E3B0C44298FC1C149AFBF4C8996FB92427AE41E4649B934CA495991B7852B855' },
    @{ Name = 'block-55';        Text = ('x' * 55);
       Count = 55;
       Sha = 'D5E285683CD4EFC02D021A5C62014694958901005D6F71E89E0989FAC77E4072' },
    @{ Name = 'block-56';        Text = ('x' * 56);
       Count = 56;
       Sha = '04C26261370EE7541549D16DEE320C723E3FD14671E66A099AFE0A377C16888E' },
    @{ Name = 'block-64';        Text = ('x' * 64);
       Count = 64;
       Sha = '7CE100971F64E7001E8FE5A51973ECDFE1CED42BEFE7EE8D5FD6219506B5393C' },
    @{ Name = 'block-65';        Text = ('x' * 65);
       Count = 65;
       Sha = '9537C5FDF120482F7D58D25E9ED583F52C02B4E304EA814DB1633AD565AED7E9' }
)

# The whitespace-skip property, stated as two spellings of ONE phrasebook
# line: CRLF-and-tabs against LF-and-spaces. Same digest, same count, and
# that digest is the plain SHA-256 of the packed bytes (say"hello").
# This is the portability property SEC.11 deliberately preserved - see
# EnglishSourceHash's header for why a raw-byte digest was rejected.
$whitespacePair = @{
    Name  = 'whitespace-skip';
    Crlf  = "(say`t`"hello`")`r`n";
    Lf    = " ( say `"hello`" ) `n";
    Count = 12;
    Sha   = 'F5885BDE1D136B3F76E2F8392A31D8EBF4F84AEA0445CF23F251B9F27A98A728'
}

# ---- 1. the REAL Get-SourceHash, lifted out of the checker ------------
$coveragePath = Join-Path $PSScriptRoot 'check_rule_coverage.ps1'
if (-not (Test-Path -LiteralPath $coveragePath)) {
    Write-Error "Missing sibling script: $coveragePath"
}
$coverageSrc = [System.IO.File]::ReadAllText($coveragePath)

# Extract by name rather than by line number so this survives the file
# moving around, and fail loudly if either function is gone - a rename
# there must not silently turn this check into a no-op.
foreach ($fn in @('Get-NonWhitespaceBytes', 'Get-SourceHash')) {
    $m = [regex]::Match($coverageSrc, "(?ms)^function\s+$([regex]::Escape($fn))\s*\(.*?^\}")
    if (-not $m.Success) {
        $failures.Add("check_rule_coverage.ps1 no longer defines '$fn' - this pin cannot verify a function it cannot find. If it was renamed, rename it here too.")
    } else {
        Invoke-Expression $m.Value
    }
}
if ($failures.Count -gt 0) {
    $failures | ForEach-Object { Write-Output "  FAIL  $_" }
    Write-Output ''
    Write-Error '=== CHECK: FAILED - the twin could not be located ==='
}

$tmpDir = Join-Path ([System.IO.Path]::GetTempPath()) ("vla_hash_twin_" + [Guid]::NewGuid().ToString('N'))
[void][System.IO.Directory]::CreateDirectory($tmpDir)
function Get-TextHash([string]$text) {
    $p = Join-Path $tmpDir ([Guid]::NewGuid().ToString('N'))
    [System.IO.File]::WriteAllBytes($p, [System.Text.Encoding]::ASCII.GetBytes($text))
    try { return Get-SourceHash $p } finally { Remove-Item -LiteralPath $p -Force }
}

Write-Output ''
Write-Output '--- PowerShell side: check_rule_coverage.ps1 Get-SourceHash ---'
try {
    foreach ($v in $vectors) {
        $got = Get-TextHash $v.Text
        if ($got.Hex -ne $v.Sha) {
            $failures.Add("Get-SourceHash('$($v.Name)') = $($got.Hex), baseline says $($v.Sha)")
            Write-Output ("  FAIL  {0,-16} {1}" -f $v.Name, $got.Hex)
        } elseif ($got.Count -ne $v.Count) {
            $failures.Add("Get-SourceHash('$($v.Name)') counted $($got.Count) bytes, baseline says $($v.Count)")
            Write-Output ("  FAIL  {0,-16} count {1}" -f $v.Name, $got.Count)
        } else {
            Write-Output ("  ok    {0,-16} {1}" -f $v.Name, $got.Hex)
        }
    }

    $a = Get-TextHash $whitespacePair.Crlf
    $b = Get-TextHash $whitespacePair.Lf
    if ($a.Hex -ne $b.Hex) {
        $failures.Add("whitespace-skip BROKEN: the CRLF spelling hashes to $($a.Hex), the LF spelling to $($b.Hex). The same phrasebook would have two identities on two machines - this is the portability property SEC.11 preserved on purpose.")
        Write-Output ("  FAIL  {0,-16} CRLF {1} vs LF {2}" -f $whitespacePair.Name, $a.Hex, $b.Hex)
    } elseif ($a.Hex -ne $whitespacePair.Sha) {
        $failures.Add("whitespace-skip digest = $($a.Hex), baseline says $($whitespacePair.Sha)")
        Write-Output ("  FAIL  {0,-16} {1}" -f $whitespacePair.Name, $a.Hex)
    } elseif ($a.Count -ne $whitespacePair.Count) {
        $failures.Add("whitespace-skip counted $($a.Count) bytes, baseline says $($whitespacePair.Count)")
        Write-Output ("  FAIL  {0,-16} count {1}" -f $whitespacePair.Name, $a.Count)
    } else {
        Write-Output ("  ok    {0,-16} {1} (CRLF and LF agree)" -f $whitespacePair.Name, $a.Hex)
    }
} finally {
    Remove-Item -LiteralPath $tmpDir -Recurse -Force -ErrorAction SilentlyContinue
}

# ---- 2. the VBA side asserts the same strings -------------------------
# Not "does VBA compute this" - that is VlaSelfTest's job - but "is
# VlaSelfTest asserting against the same numbers this baseline holds".
# Every pure-suite module, not one named file: the vectors live in
# VLA_Tests.bas (TestSec11Digest) and the stamp-shape assertions beside
# the 0.5.1 ones they replace in VLA_Tests_Grammar.bas, and moving an
# assertion between them must not silently unpin it.
$testsPaths = @(Get-ChildItem -LiteralPath (Join-Path $root 'src') -Filter 'VLA_Tests*.bas' | ForEach-Object { $_.FullName })
if ($testsPaths.Count -eq 0) {
    Write-Error "No src\VLA_Tests*.bas found under $root"
}
$testsSrc = ($testsPaths | ForEach-Object { [System.IO.File]::ReadAllText($_) }) -join "`n"

Write-Output ''
Write-Output "--- VBA side: the digests VlaSelfTest pins ($($testsPaths.Count) test modules scanned) ---"
$allExpected = @()
$allExpected += $vectors | ForEach-Object { @{ Name = $_.Name; Sha = $_.Sha } }
$allExpected += @{ Name = $whitespacePair.Name; Sha = $whitespacePair.Sha }
foreach ($e in $allExpected) {
    if ($testsSrc.IndexOf($e.Sha, [StringComparison]::OrdinalIgnoreCase) -ge 0) {
        Write-Output ("  ok    {0,-16} pinned in VBA" -f $e.Name)
    } else {
        $failures.Add("no VLA_Tests*.bas module contains the digest for '$($e.Name)' ($($e.Sha)). The two sides' baselines have separated - one of them was edited without the other.")
        Write-Output ("  FAIL  {0,-16} NOT pinned in VBA" -f $e.Name)
    }
}

# ---- 3. neither side still carries the retired polynomial as live -----
# EnglishSourceHash must no longer BE the 31-polynomial. The PowerShell
# side legitimately keeps one, named Get-SourceHashLegacy32, to verify
# pre-0.5.3 artifacts; anything else multiplying an accumulator by 31 is
# the old algorithm coming back.
Write-Output ''
Write-Output '--- the retired 32-bit polynomial is not live on either side ---'
$digestPath = Join-Path $root 'src\VLA_Digest.bas'
$enginePath = Join-Path $root 'src\VLA_SentenceEngine.bas'
foreach ($p in @($digestPath, $enginePath)) {
    $txt = [System.IO.File]::ReadAllText($p)
    if ($txt -match '(?m)^\s*h\s*=\s*h\s*\*\s*31\s*\+') {
        $failures.Add("$(Split-Path -Leaf $p) still computes h = h * 31 + b outside a comment - SEC.11 retired that as a content key.")
        Write-Output ("  FAIL  {0}" -f (Split-Path -Leaf $p))
    } else {
        Write-Output ("  ok    {0}" -f (Split-Path -Leaf $p))
    }
}
if ($coverageSrc -match '(?m)^\s*function\s+Get-SourceHashLegacy32') {
    Write-Output '  ok    check_rule_coverage.ps1 keeps Get-SourceHashLegacy32 (reads pre-0.5.3 stamps only)'
} else {
    $failures.Add('check_rule_coverage.ps1 no longer defines Get-SourceHashLegacy32 - artifacts stamped before 0.5.3 can no longer be verified, which breaks the migration path SEC.11 promised.')
    Write-Output '  FAIL  check_rule_coverage.ps1 lost Get-SourceHashLegacy32'
}

# ---- verdict ----------------------------------------------------------
Write-Output ''
if ($failures.Count -eq 0) {
    Write-Output '=== CHECK: clean - both hash implementations are pinned to the same vectors ==='
    exit 0
} else {
    foreach ($f in $failures) { Write-Output "FAIL: $f" }
    Write-Output ''
    Write-Output "=== CHECK: FAILED - $($failures.Count) problem(s) ==="
    exit 1
}
