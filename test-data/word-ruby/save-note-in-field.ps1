# Opens note-in-field-v1.rtf (or -Name note-in-field-code-v1: locked fields, so their results stay) in Word 16 (COM, hidden, no alerts, no field update), saves Word's PDF of it (<name>.rtf.pdf)
# and Word's own .docx (<name>.docx), then reopens that .docx and saves its PDF (<name>.docx.pdf): where Word
# draws a note mark whose reference sits in a field's code. Run from this folder: powershell -File save-note-in-field.ps1
param([string]$Name = 'note-in-field-v1')
$ErrorActionPreference = 'Stop'
$base = Join-Path $PSScriptRoot $Name
$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
$w = New-Object -ComObject Word.Application
$mine = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id } | ForEach-Object Id)
try {
    $w.Visible = $false
    $w.DisplayAlerts = 0
    $w.Options.DoNotPromptForConvert = $true
    $w.Options.UpdateFieldsAtPrint = $false
    $d = $w.Documents.Open($base + '.rtf', $false, $true, $false)
    $d.ActiveWindow.View.Type = 3
    $d.SaveAs2($base + '.rtf.pdf', 17)
    $d.SaveAs2($base + '.docx', 16)
    $d.Close(0)
    $d = $w.Documents.Open($base + '.docx', $false, $true, $false)
    $d.ActiveWindow.View.Type = 3
    $d.SaveAs2($base + '.docx.pdf', 17)
    $d.Close(0)
    Write-Output 'saved'
}
finally {
    try { $w.Quit(0) } catch { }
    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($w) } catch { }
    foreach ($id in $mine) { try { Stop-Process -Id $id -Force -ErrorAction Stop } catch { } }
}
