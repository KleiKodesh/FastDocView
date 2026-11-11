# Builds math-v2.docx / .doc / .rtf / .pdf with Word 16 (COM) from math-base.docx (scripts/make-math-base-docx.py: one
# paragraph per OMML construct), adding an equation in a table cell, a footnote and a text box.
# Run from this folder: powershell -File make-math-fixtures.ps1   (the harness reads the saved files; Word is not needed to test)
$ErrorActionPreference = 'Stop'
$dir = $PSScriptRoot
$w = New-Object -ComObject Word.Application
$w.Visible = $false
$w.DisplayAlerts = 0
try {
    $d = $w.Documents.Open("$dir\math-base.docx", $false, $true)
    $d.SetCompatibilityMode(15)   # Word 2013+ mode (the base file has no settings part: Word opens it in an older mode)
    function MathAt($range, [string]$linear) {
        $range.Text = $linear
        $m = $range.OMaths.Add($range)
        $m.OMaths.Item(1).BuildUp()
    }
    $end = $d.Content; $end.Collapse(0)
    $end.InsertParagraphAfter()
    $t = $d.Tables.Add($d.Paragraphs.Item($d.Paragraphs.Count).Range, 1, 2)
    $t.Cell(1, 1).Range.Text = "table cell"
    MathAt $t.Cell(1, 2).Range "tablemath=a/b"
    $p = $d.Content; $p.Collapse(0); $p.InsertAfter("note anchor"); $p = $d.Content; $p.Collapse(0)
    $fn = $d.Footnotes.Add($p)
    MathAt $fn.Range "notemath=x^2"
    $box = $d.Shapes.AddTextbox(1, 100, 100, 200, 60, $d.Paragraphs.Item($d.Paragraphs.Count).Range)
    MathAt $box.TextFrame.TextRange ("boxmath=" + [char]0x221A + "y")
    $d.SaveAs2("$dir\math-v2.docx", 16)
    $d.SaveAs2("$dir\math-v2.rtf", 6)
    $d.SaveAs2("$dir\math-v2.pdf", 17)
    $d.SaveAs2("$dir\math-v2.doc", 0)   # last: saving as .doc turns the open document's equations into pictures
    $d.Close(0)
}
finally { $w.Quit(0) }   # 0: do not save (no dialog)
