# Builds east-asian-font-v1.docx / .doc / .rtf / .odt and the reference east-asian-font-v1.pdf with an invisible Word 16
# (DisplayAlerts 0, print layout, Quit(0)) from east-asian-font-source-v1.docx (make-east-asian-font-docx.py).
# Run from this folder: python make-east-asian-font-docx.py; powershell -File make-east-asian-font-fixture.ps1
$ErrorActionPreference = 'Stop'
$dir = $PSScriptRoot
$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
$w = New-Object -ComObject Word.Application
$started = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id } | ForEach-Object Id)
$w.Visible = $false
$w.DisplayAlerts = 0
try {
    $d = $w.Documents.Open("$dir\east-asian-font-source-v1.docx", $false, $true, $false)
    try { $d.ActiveWindow.View.Type = 3 } catch {}
    $d.SaveAs2("$dir\east-asian-font-v1.docx", 16)
    $d.SaveAs2("$dir\east-asian-font-v1.doc", 0)
    $d.SaveAs2("$dir\east-asian-font-v1.rtf", 6)
    $d.SaveAs2("$dir\east-asian-font-v1.odt", 23)
    $d.ExportAsFixedFormat("$dir\east-asian-font-v1.pdf", 17)
    $d.Close(0)
}
finally { $w.Quit([ref]0); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($w) | Out-Null; Start-Sleep 1; foreach ($id in $started) { if (Get-Process -Id $id -ErrorAction SilentlyContinue) { Stop-Process -Id $id } } }

