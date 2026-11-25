# Builds the reference rtf-east-asian-runs-v1.pdf: an invisible Word 16 (DisplayAlerts 0, print layout, Quit(0)) opens
# rtf-east-asian-runs-v1.rtf (make-rtf-east-asian-runs.py) and exports it. Run from this folder:
# python make-rtf-east-asian-runs.py; powershell -File make-rtf-east-asian-fixture.ps1
$ErrorActionPreference = 'Stop'
$dir = $PSScriptRoot
$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
$w = New-Object -ComObject Word.Application
$started = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id } | ForEach-Object Id)
$w.Visible = $false
$w.DisplayAlerts = 0
try {
    $d = $w.Documents.Open("$dir\rtf-east-asian-runs-v1.rtf", $false, $true, $false)
    try { $d.ActiveWindow.View.Type = 3 } catch {}
    $d.ExportAsFixedFormat("$dir\rtf-east-asian-runs-v1.pdf", 17)
    $d.Close(0)
    $d = $w.Documents.Open("$dir\..\word-east-asian\east-asian-font-v1.rtf", $false, $true, $false)
    try { $d.ActiveWindow.View.Type = 3 } catch {}
    $d.ExportAsFixedFormat("$dir\east-asian-font-v1-rtf.pdf", 17)
    $d.Close(0)
}
finally { $w.Quit([ref]0); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($w) | Out-Null; Start-Sleep 1; foreach ($id in $started) { if (Get-Process -Id $id -ErrorAction SilentlyContinue) { Stop-Process -Id $id } } }
