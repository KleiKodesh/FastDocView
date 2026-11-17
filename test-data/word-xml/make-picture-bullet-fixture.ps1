# Opens picture-bullet-source-v1.docx (make-picture-bullet-source.py: a two-item list with a picture bullet) in Word 16
# (hidden, no alerts, read-only) and saves it as word-xml-picture-bullet-v1 .docx, Word 2003 XML (11) and PDF.
# Only the Word process this script started is stopped.
$ErrorActionPreference = 'Stop'
$base = Join-Path $PSScriptRoot 'word-xml-picture-bullet-v1'
$picture = (Resolve-Path (Join-Path $PSScriptRoot '..\word-picture-crop\quadrants.png')).Path
$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
$word = New-Object -ComObject Word.Application
$mine = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id } | ForEach-Object Id)
Write-Host ('word pid ' + ($mine -join ','))
try {
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $doc = $word.Documents.Open((Join-Path $PSScriptRoot 'picture-bullet-source-v1.docx'), $false, $true, $false)
    $doc.SaveAs2($base + '.docx', 16)
    $doc.SaveAs2($base + '-2003.xml', 11)
    $doc.ExportAsFixedFormat($base + '.pdf', 17)
    $doc.Close(0)
}
finally {
    try { $word.Quit(0) } catch { }
    foreach ($id in $mine) { try { Stop-Process -Id $id -Force -ErrorAction Stop } catch { } }
}
