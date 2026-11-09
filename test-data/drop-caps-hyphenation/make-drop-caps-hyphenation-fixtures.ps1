# Builds drop-caps-hyphenation-v1.doc / .docx / .rtf / .odt with Word 16 (COM): automatic hyphenation on, an English
# paragraph with a three-line drop cap (Word puts the initial in a framed paragraph of its own), an English paragraph
# with hyphenation turned off ("Don't hyphenate"), and a second drop cap in the margin over two lines.
# Word's own view of the paragraphs is printed. Run from this folder: powershell -File make-drop-caps-hyphenation-fixtures.ps1
$ErrorActionPreference = 'Stop'
$dir = $PSScriptRoot
$w = New-Object -ComObject Word.Application
$w.Visible = $false
$w.DisplayAlerts = 0
try {
    $d = $w.Documents.Add()
    $d.Content.LanguageID = 1033
    $d.AutoHyphenation = $true
    $long = "Lorem incomprehensibilities notwithstanding, the extraordinarily characteristic representations of internationalization accompany " +
            "uncharacteristically counterrevolutionary institutionalization across thoroughly interdisciplinary considerations."
    $d.Paragraphs.Item(1).Range.Text = $long
    $d.Content.InsertParagraphAfter()
    $d.Paragraphs.Item(2).Range.Text = "Unhyphenated: " + $long
    $d.Content.InsertParagraphAfter()
    $d.Paragraphs.Item(3).Range.Text = "Margin " + $long
    $d.Paragraphs.Item(2).Format.Hyphenation = $false
    $first = $d.Paragraphs.Item(1).DropCap
    $first.Position = 1                 # wdDropNormal
    $first.LinesToDrop = 3
    $first.DistanceFromText = 6         # points
    $third = $d.Paragraphs.Item($d.Paragraphs.Count).DropCap
    $third.Position = 2                 # wdDropMargin
    $third.LinesToDrop = 2
    $i = 0
    foreach ($p in $d.Paragraphs) { "p{0}: hyphenation={1} [{2}]" -f $i, $p.Format.Hyphenation, $p.Range.Text.TrimEnd("`r"); $i++ }
    $d.SaveAs2("$dir\drop-caps-hyphenation-v1.docx", 16)   # wdFormatXMLDocument
    $d.SaveAs2("$dir\drop-caps-hyphenation-v1.doc", 0)     # wdFormatDocument
    $d.SaveAs2("$dir\drop-caps-hyphenation-v1.rtf", 6)     # wdFormatRTF
    $d.SaveAs2("$dir\drop-caps-hyphenation-v1.odt", 23)    # wdFormatOpenDocumentText
    $d.Close(0)
}
finally { $w.Quit() }
