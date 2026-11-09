# Builds list-bullets-v1.docx / .doc / .rtf with Word 16 (COM) and list-bullets-v1.odt from the .docx with LibreOffice:
# a picture-bullet list (a 40 x 20 px PNG put in by add-picture-bullet.py) at 11 pt and at 24 pt, a Symbol bullet, a Wingdings bullet and a numbered list
# whose label is bold red. Word's own size of the picture bullet (ListLevel.PictureBullet Width / Height, points) per font
# size goes to list-bullets-v1.expected.txt. Run from this folder: powershell -File make-list-bullets-fixture.ps1
$ErrorActionPreference = 'Stop'
$dir = $PSScriptRoot
$w = New-Object -ComObject Word.Application
$w.Visible = $false
$w.DisplayAlerts = 0
try {
    $d = $w.Documents.Add()
    $d.Content.LanguageID = 1033
    $d.Paragraphs.Item(1).Range.Text = "list bullets"
    function AddItem($text, $template, $size) {
        $d.Content.InsertParagraphAfter()
        $p = $d.Paragraphs.Item($d.Paragraphs.Count).Range
        $p.Text = $text
        $p = $d.Paragraphs.Item($d.Paragraphs.Count).Range
        $p.Font.Size = $size
        $p.ListFormat.ApplyListTemplate($template, $false) | Out-Null
        $d.Paragraphs.Item($d.Paragraphs.Count).Range
    }
    $lines = @()
    foreach ($size in @(11, 24)) {
        $t = $d.ListTemplates.Add($false)
        $l = $t.ListLevels.Item(1)
        $l.NumberFormat = "o"
        $l.NumberStyle = 23
        $l.Font.Name = "Courier New"
        AddItem "picture $size" $t $size | Out-Null
    }
    foreach ($case in @(@('Symbol', 0xF0B7), @('Wingdings', 0xF0A7))) {
        $t = $d.ListTemplates.Add($false)
        $l = $t.ListLevels.Item(1)
        $l.NumberFormat = [string][char]$case[1]
        $l.NumberStyle = 23
        $l.Font.Name = $case[0]
        AddItem ("symbol " + $case[0]) $t 11 | Out-Null
    }
    $t = $d.ListTemplates.Add($false)
    $l = $t.ListLevels.Item(1)
    $l.NumberStyle = 0
    $l.NumberFormat = "%1."
    $l.Font.Bold = $true
    $l.Font.Color = 255
    AddItem "styled label" $t 11 | Out-Null
    $d.SaveAs2("$dir\list-bullets-v1-plain.docx", 16)
    $d.Close(0)
    python "$dir\add-picture-bullet.py" "$dir\list-bullets-v1-plain.docx" "$dir\list-bullets-v1-input.docx"
    if ($LASTEXITCODE -ne 0) { throw "add-picture-bullet.py failed" }
    $d = $w.Documents.Open("$dir\list-bullets-v1-input.docx")
    $lines = foreach ($p in $d.Paragraphs) { $pb = $p.Range.ListFormat.ListPictureBullet; if ($pb) { "picture`t{0}`t{1}`t{2}" -f $p.Range.Font.Size, $pb.Width, $pb.Height } }
    [System.IO.File]::WriteAllLines("$dir\list-bullets-v1.expected.txt", [string[]]$lines)
    $d.SaveAs2("$dir\list-bullets-v1.docx", 16)
    $d.SaveAs2("$dir\list-bullets-v1.doc", 0)
    $d.SaveAs2("$dir\list-bullets-v1.rtf", 6)
    $d.Close(0)
}
finally { $w.Quit(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($w) | Out-Null }
& "C:\Program Files\LibreOffice\program\soffice.exe" --headless --convert-to odt --outdir $dir "$dir\list-bullets-v1.docx" | Out-Null
Remove-Item "$dir\list-bullets-v1-plain.docx", "$dir\list-bullets-v1-input.docx"
