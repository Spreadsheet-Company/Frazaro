# CLI.0 -- (re)generates src\frmCLI.frm / src\frmCLI.frx, the Ctrl+Shift+`
# dialog VLA_IDE.VlaOpenCli shows.
#
# Why this script exists at all: a UserForm's .frm carries a paired,
# genuinely binary .frx (control layout/designer stream) that cannot be
# hand-authored as text the way every other module in this repo is. The
# only reliable way to produce a valid pair is to have Excel itself build
# the form and Export it -- exactly what this script automates, once,
# via COM against a throwaway hidden Excel instance (never the dev
# workbook, never anything the developer has open). Needs "Trust access
# to the VBA project object model" enabled locally, the same standing
# requirement VLA_Build.bas's own VlaBuildAddin already carries.
#
# Run this again ONLY when the form's LAYOUT changes (new/moved/resized
# controls). Logic-only changes (what a button's Click handler does) can
# be hand-edited directly in src\frmCLI.frm afterward -- everything past
# the "Attribute VB_Name" line is plain text, same as a .bas module; only
# the "Begin ... End" designer header above it and the paired .frx are
# not meant to be hand-edited.
#
# CLI.5 -- two changes. The layout, as the owner settled it over four
# live looks: Run and Clear History along the top (Run at the left,
# Clear History at the right), the box under them, and the transcript --
# a read-only txtTranscript -- along the very bottom of the window, the
# form opening taller to hold it. No hint line (its words now open the
# box's sample, as ; comments), no status line (the transcript carries
# every status) and no Close button (Esc and the window's own X close
# it). And the form's CODE is no longer kept here: it is
# read from the current src\frmCLI.frm -- every line after its
# "Attribute VB_" lines -- and put back into the rebuilt form unchanged.
# Until CLI.5 this script carried its own copy of that code, the CLI.0
# version, so running it would have silently thrown away everything
# CLI.1 to CLI.4 added (the font, the hint, the resizable frame, the
# history). Nothing in the code has to be copied here again.
#
# Positions below are only where each control starts: the form's own
# UserForm_Resize lays everything out again from its CLI_* constants the
# moment the form is shown (UserForm_Activate nudges the width to make
# it), so these need to be sane, not exact.
#
# Usage: powershell -File tools\build_cli_form.ps1

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$outPath = Join-Path $repoRoot 'src\frmCLI.frm'

# CLI.5: the code comes from the form as it stands.
if (-not (Test-Path $outPath)) {
    throw "src\frmCLI.frm is missing - this script rebuilds the form's layout around the code already in that file, and keeps no code of its own."
}
$existing = [System.IO.File]::ReadAllLines($outPath)
$lastAttr = -1
for ($i = 0; $i -lt $existing.Length; $i++) {
    if ($existing[$i] -like 'Attribute VB_*') { $lastAttr = $i }
}
if ($lastAttr -lt 0 -or $lastAttr -ge $existing.Length - 1) {
    throw "src\frmCLI.frm has no code after its Attribute VB_ lines - not the file this script expects, so nothing has been changed."
}
$codeLines = [System.Collections.Generic.List[string]]::new()
foreach ($l in $existing[($lastAttr + 1)..($existing.Length - 1)]) { $codeLines.Add($l) }
# Excel's Export ends the code with one blank line of its own, so a form
# this script exported carries it; handing it back would let each run add
# another. Trailing blank lines are dropped here, and every run's output
# is then the same as the last.
while ($codeLines.Count -gt 0 -and $codeLines[$codeLines.Count - 1].Trim() -eq '') { $codeLines.RemoveAt($codeLines.Count - 1) }
$code = $codeLines -join "`r`n"
Write-Host ("Carrying over {0} lines of form code from {1}" -f $codeLines.Count, $outPath)

Write-Host "Starting a hidden Excel instance..."
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $wb = $excel.Workbooks.Add()

    try {
        $null = $wb.VBProject.VBComponents.Count
    } catch {
        throw "Trust access to the VBA project object model is not enabled for this Excel install. Enable it (File > Options > Trust Center > Trust Center Settings > Macro Settings) and re-run this script."
    }

    Write-Host "Building frmCLI..."
    $comp = $wb.VBProject.VBComponents.Add(3)   # vbext_ct_MSForm
    $comp.Name = 'frmCLI'
    $form = $comp.Designer
    # Top-level form size/caption go through the VBComponent's extender
    # Properties collection, not direct dot access on Designer - late
    # binding from outside VBA does not reliably expose Width/Height/
    # Caption/StartUpPosition as settable properties any other way
    # (confirmed live: direct .Width assignment throws "cannot be found
    # on this object"). Individual controls added below don't have this
    # problem - Controls.Add returns the real runtime control instance.
    $comp.Properties.Item('Caption').Value = 'Frazaro CLI'
    $comp.Properties.Item('Width').Value = 490
    $comp.Properties.Item('Height').Value = 520
    $comp.Properties.Item('StartUpPosition').Value = 2   # fmStartUpScreen (CenterScreen)

    # CLI.5: added first, though the buttons sit above it, so that it is
    # first in the Tab order - the box is where the typing is - with the
    # two buttons after it.
    $txt = $form.Controls.Add('Forms.TextBox.1', 'txtCommand')
    $txt.Left = 12; $txt.Top = 48; $txt.Width = 466; $txt.Height = 268
    $txt.MultiLine = $true
    $txt.EnterKeyBehavior = $true
    $txt.WordWrap = $true
    $txt.ScrollBars = 2   # fmScrollBarsVertical

    $cmdRun = $form.Controls.Add('Forms.CommandButton.1', 'cmdRun')
    $cmdRun.Left = 12; $cmdRun.Top = 12; $cmdRun.Width = 110; $cmdRun.Height = 28
    $cmdRun.Caption = 'Run  (Ctrl+Enter)'

    # CLI.5 (the owner's request): the right end of the button row. It
    # forgets the kept commands after asking, and its being there at all
    # is what says the history is kept. Its caption names its chord, as
    # Run's names Ctrl+Enter (the owner's call) - about 108 points of the
    # default Tahoma 8, so the button is 140 wide where Run's is 110.
    $cmdClear = $form.Controls.Add('Forms.CommandButton.1', 'cmdClearHistory')
    $cmdClear.Left = 338; $cmdClear.Top = 12; $cmdClear.Width = 140; $cmdClear.Height = 28
    $cmdClear.Caption = 'Clear History  (Ctrl+Shift+Del)'

    # CLI.5: the transcript - a fixed strip along the very bottom of the
    # window, under the box, which takes the rest of the height.
    # Read-only (Locked), out of the Tab order, and scrolling; its font,
    # colour and flat borderless look are set in the form's own code,
    # like txtCommand's font, since setting a font from here fails over COM
    # (CLI.1). Added last because it is the last thing on the form.
    $tr = $form.Controls.Add('Forms.TextBox.1', 'txtTranscript')
    $tr.Left = 12; $tr.Top = 324; $tr.Width = 466; $tr.Height = 156
    $tr.MultiLine = $true
    $tr.WordWrap = $true
    $tr.ScrollBars = 2   # fmScrollBarsVertical
    $tr.Locked = $true
    $tr.TabStop = $false

    $comp.CodeModule.AddFromString($code)

    # CLI.5: exported to a folder of its own first, and copied over src\
    # only once both files exist - so a failed export leaves the form, and
    # its code, exactly as they were, instead of deleted.
    $tmpDir = Join-Path ([System.IO.Path]::GetTempPath()) ('frmCLI_' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $tmpDir | Out-Null
    $tmpFrm = Join-Path $tmpDir 'frmCLI.frm'
    $tmpFrx = Join-Path $tmpDir 'frmCLI.frx'
    $comp.Export($tmpFrm)
    if (-not (Test-Path $tmpFrm) -or -not (Test-Path $tmpFrx)) {
        throw "Excel did not export both frmCLI.frm and frmCLI.frx - src\ has not been changed."
    }
    $frxPath = [System.IO.Path]::ChangeExtension($outPath, '.frx')
    Copy-Item -LiteralPath $tmpFrm -Destination $outPath -Force
    Copy-Item -LiteralPath $tmpFrx -Destination $frxPath -Force
    Remove-Item -LiteralPath $tmpDir -Recurse -Force
    Write-Host "Exported: $outPath"
    Write-Host "Exported: $frxPath"

    $wb.VBProject.VBComponents.Remove($comp)
} finally {
    $wb.Close($false)
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Write-Host "Done. Re-import frmCLI into the dev workbook (or re-run VlaBuildAddin) to pick it up."
