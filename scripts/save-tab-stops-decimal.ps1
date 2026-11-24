# Opens test-data/tab-stops/tab-stops-decimal-v1.docx in Word 16 (hidden, no alerts, print layout) and saves it as PDF (the
# reference for decimal stops and underlined tabs). Only the Word process this script started is stopped.
$ErrorActionPreference = 'Stop'
$folder = (Resolve-Path (Join-Path $PSScriptRoot '..\test-data\tab-stops')).Path
$base = Join-Path $folder 'tab-stops-decimal-v1'
$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
$word = New-Object -ComObject Word.Application
$mine = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id } | ForEach-Object Id)
try {
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.Options.DoNotPromptForConvert = $true
    $doc = $word.Documents.Open($base + '.docx', $false, $true, $false)
    $doc.ActiveWindow.View.Type = 3
    $doc.SaveAs2($base + '.pdf', 17)
    $doc.Close(0)
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc)
}
finally {
    try { $word.Quit(0) } catch { }
    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) } catch { }
    foreach ($id in $mine) { try { Stop-Process -Id $id -Force -ErrorAction Stop } catch { } }
}
