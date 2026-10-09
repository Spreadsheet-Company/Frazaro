<#
check_wasm_exports.ps1 - the wasm core and the wasm language crate export exactly their lists.

WHY (KERNEL.22, 2026-10-09): the language crate's C surface, vla_load,
vla_write, vla_step and vla_view with the memory pair and the versions, is
compiled always and exported only under the `c-abi` feature, because a C
export of a library crate is exported again by every module built on it
(measured 2026-10-09 on rustc 1.99 for wasm32-unknown-unknown: a crate
depending on another exported the other's #[no_mangle] function beside its
own, and a feature turned on by one member of a --workspace build turned it
on for every dependent). An export added by accident is a second door into
a module: into the grid, past an engine's device checks, or into the core.
check_core_imports.ps1 holds what a module may ask of its host; this check
is its twin and holds what a host may ask of the module, name for name.

WHAT IT READS: the binary format's section table (magic "\0asm", version 1,
then sections of id byte, LEB128 size, payload). Section id 7 is the export
section: a LEB128 count, then each entry's name, a kind byte (0 a function,
1 a table, 2 a memory, 3 a global) and an index. The names are compared with
the baseline below as sets, and every name missing or added is printed.

WHERE THE ARTIFACTS COME FROM: target/wasm32-unknown-unknown/release/,
frazaro_core.wasm from `cargo build --release -p frazaro-core --target
wasm32-unknown-unknown` and vla_lang.wasm from `cargo build --release -p
vla-lang --features c-abi --target wasm32-unknown-unknown`; CI builds both on
every push. A tree with no Rust toolchain has no artifact and the check says
SKIPPED for it, as the import check does, because CI is where the pin bites.
A language module built without the feature exports its memory alone, and
the check says so and names the command.

-Control proves the reader and the comparison on modules built in memory:
a module exporting a memory and a function reads as those two names, the
same list compares clean, a list one short reports the extra name, and a
list one longer reports the missing one.

House style (tools/check_*.ps1): PowerShell 5.1, host-independent, a
hardcoded and reviewable baseline. Exit 0 clean, exit 1 with every name.

Usage:  powershell -File tools\check_wasm_exports.ps1
        powershell -File tools\check_wasm_exports.ps1 -Control
#>
param(
    [switch]$Control
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot

# ---- the baseline: each module's exports, name for name ----
# The core's C surface (core/src/abi.rs, core/src/lib.rs) and its memory: 13.
$coreExports = @(
    'memory', 'frazaro_abi_version', 'frazaro_alloc', 'frazaro_free', 'frazaro_version_text',
    'frazaro_translate_vla', 'frazaro_translate_vba', 'frazaro_vocab_gate', 'frazaro_build_xlsx',
    'frazaro_view', 'frazaro_reflect', 'frazaro_audit', 'frazaro_diff'
)
# The language's C surface (vla-lang/src/abi.rs, KERNEL.22) and its memory: 10.
$langExports = @(
    'memory', 'vla_abi_version', 'vla_version_text', 'vla_alloc', 'vla_free',
    'vla_load', 'vla_write', 'vla_step', 'vla_view', 'vla_unload'
)

function Read-Leb128([byte[]]$b, [ref]$i) {
    $result = 0; $shift = 0
    do {
        if ($i.Value -ge $b.Length) { throw 'truncated LEB128' }
        $byte = [int]$b[$i.Value]; $i.Value++
        $result = $result -bor (($byte -band 0x7F) -shl $shift)
        $shift += 7
    } while (($byte -band 0x80) -ne 0)
    return $result
}
function Read-Name([byte[]]$b, [ref]$i) {
    $len = Read-Leb128 $b $i
    $s = [System.Text.Encoding]::UTF8.GetString($b, $i.Value, $len)
    $i.Value += $len
    return $s
}

# Returns the export names of a module, or throws if the bytes are not a module.
function Get-Exports([byte[]]$b) {
    if ($b.Length -lt 8) { throw 'not a wasm module: shorter than its header' }
    if (-not ($b[0] -eq 0 -and $b[1] -eq 0x61 -and $b[2] -eq 0x73 -and $b[3] -eq 0x6D)) { throw 'not a wasm module: bad magic' }
    $names = New-Object System.Collections.Generic.List[string]
    $i = 8
    while ($i -lt $b.Length) {
        $id = [int]$b[$i]; $i++
        $size = Read-Leb128 $b ([ref]$i)
        if ($id -eq 7) {
            $j = $i
            $count = Read-Leb128 $b ([ref]$j)
            for ($k = 0; $k -lt $count; $k++) {
                $names.Add((Read-Name $b ([ref]$j)))
                $j++                                  # the kind byte
                [void](Read-Leb128 $b ([ref]$j))      # the index
            }
        }
        $i += $size
    }
    return ,$names
}

# Every name missing from the module and every name it adds, as lines; none when equal.
function Compare-Exports($names, [string[]]$want) {
    $problems = New-Object System.Collections.Generic.List[string]
    foreach ($w in $want) { if (-not ($names -contains $w)) { $problems.Add("missing: $w") } }
    foreach ($n in $names) { if (-not ($want -contains $n)) { $problems.Add("added: $n") } }
    return ,$problems
}

if ($Control) {
    # A header, then an export section (id 7, size 14) of two entries:
    # "memory", kind 2, index 0; "f", kind 0, index 0.
    $two = [byte[]](0x00, 0x61, 0x73, 0x6D, 0x01, 0x00, 0x00, 0x00,
                    0x07, 0x0E, 0x02,
                    0x06, 0x6D, 0x65, 0x6D, 0x6F, 0x72, 0x79, 0x02, 0x00,
                    0x01, 0x66, 0x00, 0x00)
    $names = Get-Exports $two
    $same = Compare-Exports $names @('memory', 'f')
    $short = Compare-Exports $names @('memory')
    $long = Compare-Exports $names @('memory', 'f', 'g')
    $ok = ($names.Count -eq 2 -and $names[0] -eq 'memory' -and $names[1] -eq 'f' -and
           $same.Count -eq 0 -and
           $short.Count -eq 1 -and $short[0] -eq 'added: f' -and
           $long.Count -eq 1 -and $long[0] -eq 'missing: g')
    if ($ok) {
        Write-Host 'OK: control: the module reads as memory and f; the same list compares clean, a list one short reports added: f, one longer reports missing: g'
        exit 0
    }
    Write-Host "FAIL: control: read '$($names -join ', ')'; same '$($same -join '; ')'; short '$($short -join '; ')'; long '$($long -join '; ')'"
    exit 1
}

# Forward slashes: this runs under pwsh on the ubuntu job too, where '\' is a character in a name.
$artifacts = @(
    @{ Path = (Join-Path $root 'target/wasm32-unknown-unknown/release/frazaro_core.wasm'); Want = $coreExports;
       Build = 'cargo build --release -p frazaro-core --target wasm32-unknown-unknown' },
    @{ Path = (Join-Path $root 'target/wasm32-unknown-unknown/release/vla_lang.wasm'); Want = $langExports;
       Build = 'cargo build --release -p vla-lang --features c-abi --target wasm32-unknown-unknown' }
)
$failed = $false
foreach ($a in $artifacts) {
    if (-not (Test-Path $a.Path)) {
        Write-Host "SKIPPED: no wasm artifact at $($a.Path) ($($a.Build) writes it; CI checks it on every push)"
        continue
    }
    $bytes = [System.IO.File]::ReadAllBytes($a.Path)
    $names = Get-Exports $bytes
    $problems = Compare-Exports $names $a.Want
    if ($problems.Count -ne 0) {
        if ($names.Count -eq 1 -and $names[0] -eq 'memory') {
            Write-Host "FAIL: $($a.Path) exports its memory alone: it was built without its exports; $($a.Build) builds it as the baseline holds it"
        } else {
            Write-Host "FAIL: $($a.Path) exports $($names.Count) name(s); the baseline holds $($a.Want.Count), name for name"
        }
        $problems | ForEach-Object { Write-Host "  - $_" }
        $failed = $true
        continue
    }
    Write-Host "OK: $($a.Path) exports exactly its $($a.Want.Count) names ($($bytes.Length) bytes)"
}
if ($failed) { exit 1 }
exit 0
