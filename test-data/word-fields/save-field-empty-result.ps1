# Opens field-empty-result-v1.rtf and .docx in Word 16 (COM, hidden, no alerts, no field update), saves Word's PDF of each
# (<name>.rtf.pdf, <name>.docx.pdf) and, saved first (Word's PDF export recomputes unlocked fields), Word's .doc of each (<name>.from-rtf.doc, <name>.from-docx.doc) with their PDFs,
# and lists every field's Type / Locked / Result and every paragraph as Word holds it. Run: powershell -File save-field-empty-result.ps1
$ErrorActionPreference = 'Stop'
$base = Join-Path $PSScriptRoot 'field-empty-result-v1'
$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
$w = New-Object -ComObject Word.Application
$mine = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id } | ForEach-Object Id)
function Show-Doc($d, $label) {
    foreach ($f in $d.Fields) { Write-Output ("$label field type=$($f.Type) locked=$($f.Locked) code=[$($f.Code.Text)] result=[$($f.Result.Text)]") }
    foreach ($p in $d.Paragraphs) { Write-Output ("$label para [" + $p.Range.Text.TrimEnd([char]13) + "]") }
}
try {
    $w.Visible = $false
    $w.DisplayAlerts = 0
    $w.Options.DoNotPromptForConvert = $true
    $w.Options.UpdateFieldsAtPrint = $false
    foreach ($kind in @('rtf', 'docx')) {
        $d = $w.Documents.Open("$base.$kind", $false, $true, $false)
        $d.ActiveWindow.View.Type = 3
        $d.ActiveWindow.View.ShowFieldCodes = $false
        Show-Doc $d $kind
        $d.SaveAs2("$base.from-$kind.doc", 0)
        $d.SaveAs2("$base.$kind.pdf", 17)
        $d.Close(0)
        $d = $w.Documents.Open("$base.from-$kind.doc", $false, $true, $false)
        $d.ActiveWindow.View.Type = 3
        Show-Doc $d "doc-from-$kind"
        $d.SaveAs2("$base.from-$kind.doc.pdf", 17)
        $d.Close(0)
    }
    Write-Output 'saved'
}
finally {
    try { $w.Quit(0) } catch { }
    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($w) } catch { }
    foreach ($id in $mine) { try { Stop-Process -Id $id -Force -ErrorAction Stop } catch { } }
}
