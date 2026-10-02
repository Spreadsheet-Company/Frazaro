<#
check_core_imports.ps1 - the wasm core imports nothing. Pinned at 0.

WHY: SD-13 says no outbound network call, ever. For the VBA add-in that is a
promise kept by check_no_network.ps1's scan of the source. For the core built
to WebAssembly it can be a property of the artifact instead: a wasm module can
only reach the outside world through functions its host hands it, listed in
the module's import section, so a module whose import section is EMPTY cannot
open a socket, read a clock, or touch a file, whatever the page or shell
around it can do. This check reads that section and holds its count at zero,
which turns the wall from a claim an IT reviewer takes on faith into a fact
verified on the file (docs/HORIZON.md section 12).

WHAT IT READS: the binary format's section table (magic "\0asm", version 1,
then sections of id byte, LEB128 size, payload). Section id 2 is the import
section; its payload begins with a LEB128 count of entries, and each entry
is a module name, a field name, a kind byte and that kind's description.
The names are printed when the count is not zero, so a failure says what was
imported, not just that something was.

WHERE THE ARTIFACT COMES FROM: target/wasm32-unknown-unknown/release/
frazaro_core.wasm, which `cargo build --release -p frazaro-core --target
wasm32-unknown-unknown` writes and CI builds on every push. A tree with no
Rust toolchain has no artifact; the check then says SKIPPED and exits 0,
because a check that fails on every machine without cargo would be ignored on
all of them, and CI, which always has the artifact, is where the pin bites.

-Control proves the reader itself on two modules built in memory: an empty
module (0 imports) and one whose import section names a.b (1 import). The
reader must count both right or it is not trusted to count the real thing.

House style (tools/check_*.ps1): PowerShell 5.1, host-independent, a
hardcoded and reviewable baseline. Exit 0 clean, exit 1 with every import named.

Usage:  powershell -File tools\check_core_imports.ps1
        powershell -File tools\check_core_imports.ps1 -Path x.wasm
        powershell -File tools\check_core_imports.ps1 -Control
#>
param(
    [string]$Path = '',
    [switch]$Control
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$expectedImports = 0

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
function Skip-Limits([byte[]]$b, [ref]$i) {
    $flag = [int]$b[$i.Value]; $i.Value++
    [void](Read-Leb128 $b $i)
    if (($flag -band 1) -ne 0) { [void](Read-Leb128 $b $i) }
}

# Returns the import names of a module, or throws if the bytes are not a module.
function Get-Imports([byte[]]$b) {
    if ($b.Length -lt 8) { throw 'not a wasm module: shorter than its header' }
    if (-not ($b[0] -eq 0 -and $b[1] -eq 0x61 -and $b[2] -eq 0x73 -and $b[3] -eq 0x6D)) { throw 'not a wasm module: bad magic' }
    $names = New-Object System.Collections.Generic.List[string]
    $i = 8
    while ($i -lt $b.Length) {
        $id = [int]$b[$i]; $i++
        $size = Read-Leb128 $b ([ref]$i)
        if ($id -eq 2) {
            $j = $i
            $count = Read-Leb128 $b ([ref]$j)
            for ($k = 0; $k -lt $count; $k++) {
                $mod = Read-Name $b ([ref]$j)
                $fld = Read-Name $b ([ref]$j)
                $kind = [int]$b[$j]; $j++
                switch ($kind) {
                    0 { [void](Read-Leb128 $b ([ref]$j)); $what = 'func' }
                    1 { $j++; Skip-Limits $b ([ref]$j); $what = 'table' }
                    2 { Skip-Limits $b ([ref]$j); $what = 'memory' }
                    3 { $j += 2; $what = 'global' }
                    default { $what = "kind $kind"; $k = $count }   # unknown kind: stop walking entries, keep the count
                }
                $names.Add("$mod.$fld ($what)")
            }
            while ($names.Count -lt $count) { $names.Add('(unparsed entry)') }
        }
        $i += $size
    }
    return ,$names
}

if ($Control) {
    # An empty module: header only.
    $empty = [byte[]](0x00, 0x61, 0x73, 0x6D, 0x01, 0x00, 0x00, 0x00)
    # The same header, then an import section (id 2, size 7) holding one entry:
    # module "a", field "b", kind 0 (func), type index 0.
    $one = [byte[]](0x00, 0x61, 0x73, 0x6D, 0x01, 0x00, 0x00, 0x00,
                    0x02, 0x07, 0x01, 0x01, 0x61, 0x01, 0x62, 0x00, 0x00)
    $n0 = (Get-Imports $empty).Count
    $r1 = Get-Imports $one
    if ($n0 -eq 0 -and $r1.Count -eq 1 -and $r1[0] -eq 'a.b (func)') {
        Write-Host 'OK: control: the empty module counts 0 imports; the mutant counts 1, named a.b (func)'
        exit 0
    }
    Write-Host "FAIL: control: empty module counted $n0 (want 0); mutant counted $($r1.Count) '$($r1 -join ', ')' (want 1, a.b (func))"
    exit 1
}

# Forward slashes: this runs under pwsh on the ubuntu job too, where '\' is a character in a name.
$wasm = if ($Path -ne '') { $Path } else { Join-Path $root 'target/wasm32-unknown-unknown/release/frazaro_core.wasm' }
if (-not (Test-Path $wasm)) {
    Write-Host "SKIPPED: no wasm artifact at $wasm (cargo build --release -p frazaro-core --target wasm32-unknown-unknown writes it; CI checks it on every push)"
    exit 0
}

$bytes = [System.IO.File]::ReadAllBytes($wasm)
$imports = Get-Imports $bytes
if ($imports.Count -ne $expectedImports) {
    Write-Host "FAIL: $wasm imports $($imports.Count) thing(s); the core must import nothing (pinned at $expectedImports)"
    $imports | ForEach-Object { Write-Host "  - $_" }
    exit 1
}
Write-Host "OK: $wasm has an empty import section ($($bytes.Length) bytes, $expectedImports imports)"
exit 0
