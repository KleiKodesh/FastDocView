# Builds body-script-v1.docx / .doc / .rtf with Word 16 (COM) and body-script-v1.odt from the .docx with LibreOffice:
# right-to-left paragraphs mixing Hebrew, English words and digits, Latin size 11 / Courier New and complex-script size 16 / Arial
# (Word sets each run's own right-to-left flag by its script), and a bullet whose Wingdings font is set only for complex script.
# Word's PDF export goes to body-script-v1.pdf (the reference for the sizes and fonts).
# Run from this folder: powershell -File make-body-script-fixture.ps1
$ErrorActionPreference = 'Stop'
$dir = $PSScriptRoot
$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
$w = New-Object -ComObject Word.Application
$started = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id } | ForEach-Object Id)
$w.Visible = $false
$w.DisplayAlerts = 0
$heb = -join ([char[]](0x05E9,0x05DC,0x05D5,0x05DD))   # a Hebrew word
try {
    $d = $w.Documents.Add()
    $d.Paragraphs.Item(1).Range.Text = "$heb English $heb 123 $heb"
    $d.Content.InsertParagraphAfter()
    $d.Paragraphs.Item(2).Range.Text = "$heb wingdings $heb"
    foreach ($p in $d.Paragraphs) { $p.ReadingOrder = 0 }  # wdReadingOrderRtl
    $r = $d.Content
    $r.Font.Size = 11; $r.Font.SizeBi = 16
    $r.Font.Name = "Courier New"; $r.Font.NameBi = "Arial"
    # a bullet whose Wingdings font is set only for complex script (ListLevel.Font.NameBi); its Latin font stays the text's
    $t = $d.ListTemplates.Add($false)
    $l = $t.ListLevels.Item(1)
    $l.NumberFormat = [string][char]0xF0A7
    $l.NumberStyle = 23
    $l.Font.NameBi = "Wingdings"
    $d.Paragraphs.Item(2).Range.ListFormat.ApplyListTemplate($t, $false) | Out-Null
    $d.SaveAs2("$dir\body-script-v1.docx", 16)
    $d.SaveAs2("$dir\body-script-v1.doc", 0)
    $d.SaveAs2("$dir\body-script-v1.rtf", 6)
    $d.ExportAsFixedFormat("$dir\body-script-v1.pdf", 17)
    $d.Close(0)
}
finally { $w.Quit(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($w) | Out-Null; Start-Sleep 1; foreach ($id in $started) { if (Get-Process -Id $id -ErrorAction SilentlyContinue) { Stop-Process -Id $id } } }
& "C:\Program Files\LibreOffice\program\soffice.exe" --headless --convert-to odt --outdir $dir "$dir\body-script-v1.docx" | Out-Null
