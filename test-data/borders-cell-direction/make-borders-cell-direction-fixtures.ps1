# Builds borders-cell-direction-v1.doc / .docx / .rtf (+ Word's own PDF of it) with Word 16 (COM):
# a one-row table whose cells run horizontal, downward (top to bottom), upward (bottom to top) and East-Asian vertical;
# two paragraphs with the same box border (4 pt from the text at the top, 10 pt at the sides, shadowed) that Word joins
# into one box; a third paragraph with a different (unshadowed) border that gets its own box.
# Run from this folder: powershell -File make-borders-cell-direction-fixtures.ps1
$ErrorActionPreference = 'Stop'
$dir = $PSScriptRoot
$w = New-Object -ComObject Word.Application
$w.Visible = $false
$w.DisplayAlerts = 0
try {
    $d = $w.Documents.Add()
    $d.Content.LanguageID = 1033
    $d.Paragraphs.Item(1).Range.Text = "Cells:"
    $d.Content.InsertParagraphAfter()
    $r = $d.Paragraphs.Item(2).Range
    $t = $d.Tables.Add($r, 1, 4)
    $t.Borders.Enable = $true
    $t.Rows.Item(1).Height = 90
    $t.Rows.Item(1).HeightRule = 2            # wdRowHeightExactly
    $texts = @("Flat", "Down", "Up", "Asian")
    $orientations = @(0, 3, 2, 1)             # wdTextOrientationHorizontal, Downward, Upward, VerticalFarEast
    for ($i = 0; $i -lt 4; $i++) {
        $cell = $t.Cell(1, $i + 1)
        $cell.Range.Text = $texts[$i]
        $cell.Range.Orientation = $orientations[$i]
    }
    $after = $d.Content
    $after.InsertParagraphAfter()
    $n = $d.Paragraphs.Count
    $d.Paragraphs.Item($n).Range.Text = "Boxed one: the first paragraph of a border group."
    $d.Content.InsertParagraphAfter()
    $d.Paragraphs.Item($d.Paragraphs.Count).Range.Text = "Boxed two: the second paragraph, inside the same box."
    $d.Content.InsertParagraphAfter()
    $d.Paragraphs.Item($d.Paragraphs.Count).Range.Text = "Alone: a different border, a box of its own."
    $count = 0; foreach ($p in $d.Paragraphs) { $count++ }
        foreach ($k in @(($count - 2), ($count - 1))) {
        $b = $d.Paragraphs.Item($k).Borders
        $b.Enable = $true
        $b.Shadow = $true
        $b.DistanceFromTop = 4
        $b.DistanceFromBottom = 4
        $b.DistanceFromLeft = 10
        $b.DistanceFromRight = 10
    }
    $alone = $d.Paragraphs.Item($count).Borders
    $alone.Enable = $true
    $alone.OutsideLineWidth = 12              # wdLineWidth150pt
    $i = 0
    foreach ($p in $d.Paragraphs) { "p{0}: shadow={1} [{2}]" -f $i, $p.Borders.Shadow, $p.Range.Text.TrimEnd("`r", [char]7); $i++ }
    $d.SaveAs2("$dir\borders-cell-direction-v1.docx", 16)   # wdFormatXMLDocument
    $d.SaveAs2("$dir\borders-cell-direction-v1.doc", 0)     # wdFormatDocument
    $d.SaveAs2("$dir\borders-cell-direction-v1.rtf", 6)     # wdFormatRTF
    $d.SaveAs2("$dir\borders-cell-direction-v1.word.pdf", 17) # wdFormatPDF
    $d.Close(0)
}
catch { "FAILED: $_ at $($_.InvocationInfo.ScriptLineNumber)" }
finally { $w.Quit([ref]0) }





