# Builds ruby-v1.docx / .doc / .rtf / .pdf with Word 16 (COM): one paragraph per phonetic-guide alignment (Range.PhoneticGuide,
# WdPhoneticGuideAlignmentType 0-4), the guide over the word "base<N>". Word saves a phonetic guide to .doc and RTF as an
# EQ field ("EQ \* jcN \* "Font:..." \* hpsN \o\aX(\s\up N(guide),base)").
# Run from this folder: powershell -File make-ruby-fixtures.ps1   (the harness reads the saved files; Word is not needed to test)
$ErrorActionPreference = 'Stop'
$dir = $PSScriptRoot
$w = New-Object -ComObject Word.Application
$w.Visible = $false
$w.DisplayAlerts = 0
try {
    $d = $w.Documents.Add()
    $sizes = @(10, 8, 10, 12, 10)
    $d.Content.Text = (0..4 | ForEach-Object { "ruby case ${_}: before base$_ after" }) -join "`r"
    for ($i = 4; $i -ge 0; $i--) {
        $r = $d.Paragraphs.Item($i + 1).Range
        $find = $r.Find
        if ($find.Execute("base$i")) { $r.PhoneticGuide("guide$i", $i, 0, $sizes[$i], "Arial") }
    }
    $d.SaveAs2("$dir\ruby-v1.docx", 16)
    $d.SaveAs2("$dir\ruby-v1.rtf", 6)
    $d.SaveAs2("$dir\ruby-v1.pdf", 17)
    $d.SaveAs2("$dir\ruby-v1.doc", 0)
    $d.Close([ref]0)
}
finally { $w.Quit([ref]0) }   # 0: do not save (no dialog)
