# Builds borders-cell-groups-v1.rtf (+ Word's own PDF) with Word 16 (COM): a one-cell table holding two paragraphs with the
# same box border (Word joins them into one box) and a third whose thicker lines Word applies to the cell itself (its PDF draws them as the cell box), and a text box holding
# two paragraphs with the same box border. Run from this folder: powershell -File make-borders-cell-groups-fixtures.ps1
$ErrorActionPreference = 'Stop'
$dir = $PSScriptRoot
$w = New-Object -ComObject Word.Application
$w.Visible = $false
$w.DisplayAlerts = 0
try {
    $d = $w.Documents.Add()
    $d.Content.LanguageID = 1033
    $d.Paragraphs.Item(1).Range.Text = "Table:"
    $d.Content.InsertParagraphAfter()
    $t = $d.Tables.Add($d.Paragraphs.Item(2).Range, 1, 1)
    $cell = $t.Cell(1, 1).Range
    $cell.Text = "Cell one: first of a border group.`rCell two: same box.`rCell alone: a box of its own."
    $count = 0; foreach ($p in $cell.Paragraphs) { $count++ }
    $k = 0
    foreach ($p in $cell.Paragraphs) {
        $k++
        $b = $p.Borders
        $b.Enable = $true
        if ($k -eq $count) { foreach ($side in -1, -2, -3, -4) { $b.Item($side).LineWidth = 12 } }   # wdLineWidth150pt: a different border
    }
    $box = $d.Shapes.AddTextbox(1, 72, 300, 300, 120)       # msoTextOrientationHorizontal, points
    $box.Line.Visible = 0
    $tr = $box.TextFrame.TextRange
    $tr.Text = "Box one: first of a border group.`rBox two: same box."
    foreach ($p in $tr.Paragraphs) { $p.Borders.Enable = $true }
    $i = 0
    foreach ($p in $d.StoryRanges.Item(1).Paragraphs) { "p{0}: [{1}]" -f $i, $p.Range.Text.TrimEnd("`r", [char]7); $i++ }
    $d.SaveAs2("$dir\borders-cell-groups-v1.rtf", 6)       # wdFormatRTF
    $d.SaveAs2("$dir\borders-cell-groups-v1.word.pdf", 17) # wdFormatPDF
    $d.Close(0)
}
catch { "FAILED: $_ at $($_.InvocationInfo.ScriptLineNumber)" }
finally { $w.Quit([ref]0) }
