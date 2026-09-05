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
# Usage: powershell -File tools\build_cli_form.ps1

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$outPath = Join-Path $repoRoot 'src\frmCLI.frm'

$code = @'
Private Sub UserForm_Initialize()
    cmdCancel.Cancel = True
    lblStatus.Caption = ""
End Sub

Private Sub UserForm_Activate()
    txtCommand.SetFocus
End Sub

Private Sub cmdRun_Click()
    RunCurrentText
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub txtCommand_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn And (Shift And 2) = 2 Then   ' Ctrl+Enter
        KeyCode = 0
        RunCurrentText
    End If
End Sub

Private Sub RunCurrentText()
    Dim src As String
    src = txtCommand.Text
    If Len(Trim$(src)) = 0 Then Exit Sub
    lblStatus.Caption = "Running..."
    Me.Repaint
    lblStatus.Caption = VLA_IDE.VlaCliRun(src)
    txtCommand.SetFocus
End Sub
'@

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
    $comp.Properties.Item('Width').Value = 480
    $comp.Properties.Item('Height').Value = 300
    $comp.Properties.Item('StartUpPosition').Value = 2   # fmStartUpScreen (CenterScreen)

    $lblHint = $form.Controls.Add('Forms.Label.1', 'lblHint')
    $lblHint.Left = 12; $lblHint.Top = 8; $lblHint.Width = 456; $lblHint.Height = 14
    $lblHint.Caption = 'VLA or English, one entry point - Ctrl+Enter runs, Esc closes.'

    $txt = $form.Controls.Add('Forms.TextBox.1', 'txtCommand')
    $txt.Left = 12; $txt.Top = 26; $txt.Width = 456; $txt.Height = 170
    $txt.MultiLine = $true
    $txt.EnterKeyBehavior = $true
    $txt.WordWrap = $true
    $txt.ScrollBars = 2   # fmScrollBarsVertical

    $lblStatus = $form.Controls.Add('Forms.Label.1', 'lblStatus')
    $lblStatus.Left = 12; $lblStatus.Top = 204; $lblStatus.Width = 340; $lblStatus.Height = 18

    $cmdRun = $form.Controls.Add('Forms.CommandButton.1', 'cmdRun')
    $cmdRun.Left = 232; $cmdRun.Top = 226; $cmdRun.Width = 110; $cmdRun.Height = 28
    $cmdRun.Caption = 'Run  (Ctrl+Enter)'

    $cmdCancel = $form.Controls.Add('Forms.CommandButton.1', 'cmdCancel')
    $cmdCancel.Left = 346; $cmdCancel.Top = 226; $cmdCancel.Width = 110; $cmdCancel.Height = 28
    $cmdCancel.Caption = 'Close  (Esc)'

    $comp.CodeModule.AddFromString($code)

    if (Test-Path $outPath) { Remove-Item $outPath }
    $frxPath = [System.IO.Path]::ChangeExtension($outPath, '.frx')
    if (Test-Path $frxPath) { Remove-Item $frxPath }

    $comp.Export($outPath)
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
