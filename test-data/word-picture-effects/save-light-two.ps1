# Prints effects-light-two-v1.docx to PDF with a hidden Word 16 (DisplayAlerts 0, not visible, print layout, Quit(0) in finally;
# only the Word instance this script started is stopped, by PID).
$d = $PSScriptRoot
$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
$w = New-Object -ComObject Word.Application
$w.DisplayAlerts = 0; $w.Visible = $false
$mine = @(Get-Process WINWORD | Where-Object { $before -notcontains $_.Id } | ForEach-Object Id)
try {
    $doc = $w.Documents.Open("$d\effects-light-two-v1.docx", $false, $true)
    $doc.ActiveWindow.View.Type = 3
    $doc.SaveAs2("$d\effects-light-two-v1.pdf", 17)
    $doc.Close([ref]0)
} finally {
    $w.Quit([ref]0)
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($w)
    Start-Sleep -Milliseconds 500
    foreach ($id in $mine) { Stop-Process -Id $id -Force -ErrorAction SilentlyContinue }
}
