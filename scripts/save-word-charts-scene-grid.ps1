# Prints each test-data/word-charts-scene-grid/grid-*.docx to PDF with a hidden Word 16 (DisplayAlerts 0, print layout).
# Run once per file (one Word per call: a second Open in the same hidden instance hung). Quits Word in finally and stops only
# the Word process this script started.
param([string]$Folder, [string]$Filter = 'grid-*.docx')
$ErrorActionPreference = 'Stop'
$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
$word = New-Object -ComObject Word.Application
$mine = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id } | ForEach-Object Id)
Write-Host ('word pid ' + ($mine -join ','))
try {
    $word.Visible = $false
    $word.DisplayAlerts = 0
    foreach ($f in (Get-ChildItem $Folder -Filter $Filter)) {
        $pdf = [IO.Path]::ChangeExtension($f.FullName, '.pdf')
        if (Test-Path $pdf) { continue }
        $doc = $word.Documents.Open($f.FullName, $false, $true, $false)
        $doc.ActiveWindow.View.Type = 3
        $doc.SaveAs2($pdf, 17)
        $doc.Close(0)
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc)
        Write-Host ('pdf ' + $f.Name)
    }
}
finally {
    try { $word.Quit(0) } catch { }
    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) } catch { }
    foreach ($id in $mine) { try { Stop-Process -Id $id -Force -ErrorAction Stop } catch { } }
}
