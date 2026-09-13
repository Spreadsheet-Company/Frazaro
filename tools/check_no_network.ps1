<#
check_no_network.ps1 - SD-13's mechanical pin, and SIG.1's evidence.

SD-13: Frazaro makes no outbound network call, ever, without a dedicated
re-litigation of that decision. docs/IT_REVIEW.md (SIG.1) states it to a
stranger as a fact about the shipped add-in - and states, beside it, the
short list of ways a PROGRAM can still reach outward through the user's own
tools (an Outlook draft, "Refresh everything", a written formula, a process
the uninstall button starts). A sentence like that is worth no more than the
last time someone checked it. This script is that check, run on every push
by CI and before every release by tools/release.ps1.

WHY A STATIC SCAN AND NOT A TEST: "no call was made" cannot be proven by
running anything - a test observes one run, and the property is about every
path. What CAN be pinned is the absence of every mechanism VBA has for
making one, in the code that ships. That is a text property of src/, and a
scan is the right instrument for it.

WHAT IS CHECKED, four properties:
  1. FORBIDDEN - no shipped module names a network mechanism: an HTTP
     client ProgID (WinHttp, XMLHTTP, ServerXMLHTTP, DOMDocument,
     InternetExplorer.Application), a download/socket API
     (URLDownloadToFile, InternetOpen, ...), FollowHyperlink, a new web
     query or data connection (QueryTables.Add, Connections.Add,
     Workbooks.OpenXML), or VBA's own Shell. No baseline, no exemption:
     one hit fails.
  2. DECLARES - every `Declare` (a call into a Windows DLL, any of which
     could reach a network) is in $declares below. Today that is frmCLI's
     four user32 window calls, twice (the PtrSafe and legacy branches of
     one #If). A new one fails until it is reviewed and listed.
  3. URL LITERALS - every "http://" / "https://" string in shipped code is
     in $urls below with the reason it is not a fetch. A new one fails.
  4. OUTWARD REACH - every site where a program can reach outside Excel
     through the user's own tools is in $reach below, and each one is
     named in docs/IT_REVIEW.md. A new one FAILS (stricter than
     check_sec8_provenance_gate.ps1, where a new guarded site only warns),
     because this list is pinned to a document a stranger will act on: a
     reach the document does not name is the document being wrong.
Plus the installer: installer/Frazaro.iss downloads nothing and runs no
program ([Run]/[UninstallRun] sections, DownloadTemporaryFile,
CreateDownloadPage, any URL literal).

WHAT IS NOT CHECKED, and SIG.1 says so rather than leaning on this script:
  * A phrasebook rule marked `raw` splices arbitrary VBA - which can do
    anything VBA can, network included - behind SEC.2's consent dialog. Its
    payload is opaque text by design (THREAT_MODEL.md section 1.4); no scan
    can bound it, which is why it is consent-gated rather than analysed.
  * The Compile path's emitted code calls what the phrasebook templates
    say (SEC.12, accepted: Compile needs VBA-project trust, off by
    default). This script reads the engine, not the phrasebooks.
  * Where Excel or Windows sends a PATH a program names - Workbooks.Open on
    a UNC share - is file I/O to the operating system, gated for
    internet-marked workbooks by SEC.8 (check_sec8_provenance_gate.ps1),
    not a network mechanism in Frazaro's code.
  * What a formula does once written (WEBSERVICE, DDE) is Excel's - SEC.15,
    open. The two sinks that write formulas ARE in $reach, so they stay
    named.

SCANNED SET: the SHIPPED modules, read from VLA_Build.bas's own `mods` array
exactly as check_sec8_provenance_gate.ps1 does - the claim is about what a
downloader receives. VLA_Build.bas itself (it writes a PowerShell script and
a customUI namespace URL at BUILD time, on the maintainer's machine) and the
test/dev-rig modules never ship, so they are out of scope by construction.

BASELINES: hand-maintained below, in the reviewable shape of
check_raise_ratchet.ps1's $ceilings. Adding an entry is a reviewed act: say
why in the commit message, and add the reach to docs/IT_REVIEW.md in the
same commit.

NOT wired into VlaSelfTest, same reasoning F.12/F.14/AS.8/SEC.13 gave for the
other tools/check_*.ps1 scans.

Usage:  pwsh -File tools/check_no_network.ps1
Exit code: 0 if all four properties and the installer hold; 1 otherwise.
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

# --- 1. Forbidden: a network mechanism, anywhere in shipped code ---
# Matched case-insensitively on non-comment lines. Each is a way VBA (or a
# COM object VBA can create) opens a connection or hands a URL to something
# that will.
$forbidden = [ordered]@{
    'HTTP client (WinHttp)'              = 'WinHttp'
    'HTTP client (XMLHTTP)'              = 'XMLHTTP'
    'XML document (can .Load a URL)'     = 'DOMDocument'
    'browser automation'                 = 'InternetExplorer\.Application'
    'download API'                       = 'URLDownloadTo(File|CacheFile)'
    'WinINet API'                        = '\bInternet(Open|Connect|ReadFile|OpenUrl)\b|\bHttp(OpenRequest|SendRequest)\b'
    '.NET web client'                    = 'WebClient|Invoke-WebRequest|Invoke-RestMethod'
    'hyperlink follow (opens a URL)'     = 'FollowHyperlink'
    'new web query'                      = 'QueryTables\s*\.\s*Add'
    'new data connection'                = 'Connections\s*\.\s*Add'
    'XML from a URL'                     = '\.OpenXML\b|XmlMaps\s*\.\s*Add'
    'VBA Shell (starts a process)'       = '(^|[^.\w])Shell\s*(\(|")'
    'ShellExecute'                       = 'ShellExecute'
}

# --- 2. Declares: every DLL call, by "<Module>::<Name>::<Lib>" ---
# 2026-09-10, SIG.1: frmCLI's window-style calls - GetWindowLong/
# SetWindowLong/SetWindowPos/FindWindowA, user32 only, declared once per
# branch of the VBA7 #If. Nothing else in the shipped set calls a DLL.
$declares = @(
    'frmCLI::GetWindowLong::user32'
    'frmCLI::SetWindowLong::user32'
    'frmCLI::SetWindowPos::user32'
    'frmCLI::FindWindowA::user32'
)

# --- 3. URL literals: "<Module>::<Procedure>" and why it is not a fetch ---
$urls = @{
    'VLA_Provenance::VlaPathIsDemonstrablyLocal' =
        'SEC.8: compares a workbook PATH against "http://" / "https://" so that a workbook opened from a URL is NOT counted as local. A string comparison; nothing is fetched.'
}

# --- 4. Outward reach through the user's own tools ---
# "<Module>::<Procedure>::<label>". Each is named in docs/IT_REVIEW.md
# (SIG.1) section 3. 2026-09-10: seven sites.
$reach = @(
    'VLA_Runtime::VlaSendMail::Outlook'                 # a draft, displayed, never sent (SEC.8-gated)
    'VLA_IDE::ReadWordFile::Word'                       # Import Program File, macros force-disabled (SEC.13)
    'VLA_IDE::VlaIdeUninstall::process'                 # starts the installer's own unins000.exe
    'VLA_IDE::ScheduleSelfDelete::process'              # hidden PowerShell that deletes the standalone .xlam
    'VLA_Interpreter::DynamicCall::RefreshAll'          # "Refresh everything." - the workbook's existing connections
    'VLA_Interpreter::DynamicSet::formula-write'        # Interpret's formula sink (SEC.15, open)
    'VLA::EmitStmt::formula-write'                      # Compile's formula sink (SEC.15, open)
)

$reachPatterns = [ordered]@{
    'Outlook'       = '"Outlook\.Application"'
    'Word'          = '"Word\.Application"'
    'process'       = '(?<!Application)\.Run\s'
    'RefreshAll'    = '\.RefreshAll\b'
    'formula-write' = '\.Formula(2|R1C1|Local|Array)?\s*='
}

$procPattern    = '^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+([A-Za-z_][A-Za-z0-9_]*)'
$declarePattern = '\bDeclare\s+(?:PtrSafe\s+)?(?:Function|Sub)\s+([A-Za-z_][A-Za-z0-9_]*)\s+Lib\s+"([^"]+)"'
$urlPattern     = '"https?://'

function Test-IsComment([string]$line) {
    $t = $line.TrimStart()
    return ($t.Length -eq 0 -or $t.StartsWith("'"))
}

$failed       = New-Object System.Collections.Generic.List[string]
$foundDeclare = New-Object System.Collections.Generic.List[string]
$foundUrl     = New-Object System.Collections.Generic.List[string]
$foundReach   = New-Object System.Collections.Generic.List[string]

Write-Output '=== NO-NETWORK PIN (SD-13, SIG.1) ==='
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
    $currentProc = '(module level)'

    for ($i = 0; $i -lt $lines.Count; $i++) {
        $line = $lines[$i]
        if (Test-IsComment $line) { continue }
        $where = "src/$($file.Name):$($i + 1)"

        # Declares first: a Declare line is also where a DLL name appears,
        # and it is module-level, so it must not reset $currentProc.
        $d = [regex]::Match($line, $declarePattern, 'IgnoreCase')
        if ($d.Success) {
            $lib  = [IO.Path]::GetFileNameWithoutExtension($d.Groups[2].Value).ToLowerInvariant()
            $site = "$name::$($d.Groups[1].Value)::$lib"
            $foundDeclare.Add($site)
            if ($declares -contains $site) {
                Write-Output ("  {0,-52} line {1,5}  declare    (known)" -f $site, ($i + 1))
            } else {
                Write-Output ("  {0,-52} line {1,5}  DECLARE    (NEW)  FAIL" -f $site, ($i + 1))
                $failed.Add("new Declare $site ($where)")
            }
            continue
        }

        $m = [regex]::Match($line, $procPattern)
        if ($m.Success) {
            $currentProc = $m.Groups[1].Value
            continue
        }

        foreach ($label in $forbidden.Keys) {
            if ($line -match "(?i)$($forbidden[$label])") {
                Write-Output ("  {0,-52} line {1,5}  FORBIDDEN  ({2})  FAIL" -f "$name::$currentProc", ($i + 1), $label)
                $failed.Add("forbidden $label in $name::$currentProc ($where)")
            }
        }

        if ($line -match "(?i)$urlPattern") {
            $site = "$name::$currentProc"
            $foundUrl.Add($site)
            if ($urls.ContainsKey($site)) {
                Write-Output ("  {0,-52} line {1,5}  url        (known)" -f $site, ($i + 1))
            } else {
                Write-Output ("  {0,-52} line {1,5}  URL        (NEW)  FAIL" -f $site, ($i + 1))
                $failed.Add("new URL literal in $site ($where)")
            }
        }

        foreach ($label in $reachPatterns.Keys) {
            if ($line -notmatch $reachPatterns[$label]) { continue }
            $site = "$name::$currentProc::$label"
            if ($foundReach -contains $site) { continue }   # one site, several lines (GetObject, then CreateObject)
            $foundReach.Add($site)
            if ($reach -contains $site) {
                Write-Output ("  {0,-52} line {1,5}  reach      (known)" -f $site, ($i + 1))
            } else {
                Write-Output ("  {0,-52} line {1,5}  REACH      (NEW)  FAIL" -f $site, ($i + 1))
                $failed.Add("new outward-reach site $site ($where)")
            }
        }
    }
}

# --- The installer: downloads nothing, runs nothing ---
Write-Output ''
Write-Output '--- installer/Frazaro.iss ---'
$issFile  = Join-Path $repoRoot 'installer\Frazaro.iss'
$issLines = [IO.File]::ReadAllText($issFile) -split "`r?`n"
$issBad = [ordered]@{
    '[Run] section (starts a program)'          = '^\s*\[Run\]'
    '[UninstallRun] section (starts a program)' = '^\s*\[UninstallRun\]'
    'download API'                              = 'DownloadTemporaryFile|CreateDownloadPage'
    'URL literal'                               = 'https?://'
}
$issClean = $true
for ($i = 0; $i -lt $issLines.Count; $i++) {
    $t = $issLines[$i].TrimStart()
    if ($t.StartsWith(';') -or $t.StartsWith('//')) { continue }
    foreach ($label in $issBad.Keys) {
        if ($issLines[$i] -match "(?i)$($issBad[$label])") {
            Write-Output ("  Frazaro.iss                                          line {0,5}  {1}  FAIL" -f ($i + 1), $label)
            $failed.Add("installer: $label (installer/Frazaro.iss:$($i + 1))")
            $issClean = $false
        }
    }
}
if ($issClean) { Write-Output '  no [Run]/[UninstallRun] section, no download API, no URL' }

# --- Stale baseline entries: reported, not failed (the document then
# names a reach that is gone - the safe direction to be wrong in) ---
$stale = @()
$stale += $declares | Where-Object { $foundDeclare -notcontains $_ }
$stale += $urls.Keys | Where-Object { $foundUrl -notcontains $_ }
$stale += $reach | Where-Object { $foundReach -notcontains $_ }
if ($stale) {
    Write-Output ''
    Write-Output '=== BASELINE ENTRIES WITH NO LIVE SITE (removed or renamed - drop them above, and from docs/IT_REVIEW.md) ==='
    $stale | ForEach-Object { Write-Output "  $_" }
}

Write-Output ''
if ($failed.Count -eq 0) {
    Write-Output "=== CHECK: clean - no network mechanism in $($modNames.Count) shipped modules; $(@($foundDeclare | Select-Object -Unique).Count) Declare(s), $(@($foundUrl | Select-Object -Unique).Count) URL-literal site(s) and $($foundReach.Count) outward-reach site(s), every one reviewed and listed ==="
} else {
    Write-Output "=== CHECK: $($failed.Count) problem(s) ==="
    $failed | ForEach-Object { Write-Output "  $_" }
    Write-Output 'A network mechanism in shipped code breaks SD-13 outright - remove it, or re-litigate SD-13 first (docs/BETA_ROADMAP1.md). A NEW Declare, URL literal or outward-reach site may be fine, but it changes what docs/IT_REVIEW.md tells a reviewer: review it, add it to the baseline above with its reason, and name it in IT_REVIEW.md section 3 in the same commit.'
}
exit ([Math]::Min($failed.Count, 1))
