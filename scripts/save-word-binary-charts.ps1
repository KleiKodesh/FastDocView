# Opens test-data/word-charts-binary/word-charts-binary-v2.docx in Word 16 (hidden, no alerts) and saves it as .doc, .rtf and
# PDF (ExportAsFixedFormat keeps chart text as text). No chart is inserted or edited here, so Excel is never driven.
# Word is quit in finally; only the Word process this script started is stopped.
$ErrorActionPreference = 'Stop'
$folder = (Resolve-Path (Join-Path $PSScriptRoot '..\test-data\word-charts-binary')).Path
$base = Join-Path $folder 'word-charts-binary-v2'
$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
$word = New-Object -ComObject Word.Application
$mine = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id } | ForEach-Object Id)
Write-Host ('word pid ' + ($mine -join ','))
try {
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.Options.DoNotPromptForConvert = $true
    $doc = $word.Documents.Open($base + '.docx', $false, $true, $false)
    # the PDF first: once saved as .doc / .rtf the charts are OLE pictures and print their text as outlines
    $doc.ExportAsFixedFormat($base + '.pdf', 17, $false, 0, 0, 1, 1, 0, $true, $true, 0, $true, $true, $false)
    $doc.SaveAs2($base + '.doc', 0)
    $doc.SaveAs2($base + '.rtf', 6)
    $doc.Close(0)
}
finally {
    try { $word.Quit() } catch { }
    foreach ($id in $mine) { try { Stop-Process -Id $id -Force -ErrorAction Stop } catch { } }
}
