# Opens test-data/word-charts-binary/word-charts-3d-rows-v1.docx and word-charts-crosses-v1.docx in Word 16 (hidden, no alerts)
# and saves each as PDF (chart text stays text) and .doc. No chart is inserted or edited, so Excel is never driven.
# Word is quit in finally; only the Word process this script started is stopped.
$ErrorActionPreference = 'Stop'
$folder = (Resolve-Path (Join-Path $PSScriptRoot '..\test-data\word-charts-binary')).Path
$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
$word = New-Object -ComObject Word.Application
$mine = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id } | ForEach-Object Id)
Write-Host ('word pid ' + ($mine -join ','))
try {
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.Options.DoNotPromptForConvert = $true
    foreach ($name in 'word-charts-3d-rows-v1', 'word-charts-crosses-v1') {
        $base = Join-Path $folder $name
        $doc = $word.Documents.Open($base + '.docx', $false, $true, $false)
        # the PDF first: once saved as .doc / .rtf the charts are OLE pictures and print their text as outlines
        $doc.SaveAs2($base + '.pdf', 17)
        $doc.SaveAs2($base + '.doc', 0)
        $doc.Close(0)
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc)
    }
}
finally {
    try { $word.Quit(0) } catch { }
    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) } catch { }
    foreach ($id in $mine) { try { Stop-Process -Id $id -Force -ErrorAction Stop } catch { } }
}
