# Opens note-in-field-v1.rtf (or -Name note-in-field-code-v1: locked fields, so their results stay) in Word 16 (COM, hidden, no alerts, no field update), saves Word's PDF of it (<name>.rtf.pdf)
# and Word's own .docx (<name>.docx), then reopens that .docx and saves its PDF (<name>.docx.pdf): where Word
# draws a note mark whose reference sits in a field's code. Run from this folder: powershell -File save-note-in-field.ps1
# -Doc also saves Word's .doc (<name>.doc) and its PDF (<name>.doc.pdf), and lists each field's Type / Locked / Result (Word's view).
param([string]$Name = 'note-in-field-v1', [switch]$Doc)
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
    if ($Doc) { $d.SaveAs2($base + '.doc', 0) }
    $d.Close(0)
    $kinds = @('.docx'); if ($Doc) { $kinds += '.doc' }
    foreach ($kind in $kinds) {
        $d = $w.Documents.Open($base + $kind, $false, $true, $false)
        $d.ActiveWindow.View.Type = 3
        if ($Doc) { foreach ($f in $d.Fields) { Write-Output ("$kind type=$($f.Type) locked=$($f.Locked) result=[$($f.Result.Text)]") } }
        $d.SaveAs2($base + $kind + '.pdf', 17)
        $d.Close(0)
    }
    Write-Output 'saved'
}
finally {
    try { $w.Quit(0) } catch { }
    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($w) } catch { }
    foreach ($id in $mine) { try { Stop-Process -Id $id -Force -ErrorAction Stop } catch { } }
}
