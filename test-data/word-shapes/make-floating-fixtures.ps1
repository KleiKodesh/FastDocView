# Builds the floating-object fixtures with Word 16 (COM, invisible) and Word's own print of each saved file:
#   floating-v1.doc / .rtf (+ floating-v1-doc-word.pdf, floating-v1-rtf-word.pdf): text boxes placed by offset from the
#     column, page and character, by alignment, a rotated shape, a text box in a paragraph with a footnote and one in a
#     right-to-left paragraph; wrap none / top-and-bottom (Word prints these where the file says) and square (with an 18 pt gap).
#   floating-group-v1.doc / .rtf (+ PDFs): a group of two rectangles with text, aligned right, top and bottom. Word 16
#     refuses Group() on Word 2013 shapes through COM (E_ACCESSDENIED, also through the selection), so this document is
#     made in Word 2003 compatibility mode.
# The section and paragraphs are set left-to-right: the Word here defaults to a right-to-left section.
# Run from this folder: powershell -File make-floating-fixtures.ps1
$ErrorActionPreference = 'Stop'
$dir = $PSScriptRoot
$w = New-Object -ComObject Word.Application
$w.DisplayAlerts = 0
$w.Visible = $false
$d = $null
$filler = ("filler words wrap around the floating object here " * 8).Trim()

function NewDocument([int]$mode) {
    $doc = $w.Documents.Add()
    if ($mode -gt 0) { $doc.SetCompatibilityMode($mode) }
    $doc.Content.LanguageID = 1033
    $doc.Sections.Item(1).PageSetup.SectionDirection = 1          # wdSectionDirectionLtr
    $doc.Content.ParagraphFormat.ReadingOrder = 1                  # wdReadingOrderLtr
    return $doc
}
function Para([string]$text) {
    $r = $script:d.Content; $r.Collapse(0); $r.InsertParagraphAfter()
    $r = $script:d.Content; $r.Collapse(0); $r.Text = $text
    $p = $script:d.Paragraphs.Item($script:d.Paragraphs.Count)
    $p.ReadingOrder = 1
    return $p.Range
}
# relative: wdRelativeHorizontalPosition margin 0 / page 1 / column 2 / character 3; left: points or wdShapeLeft -999998,
# wdShapeCenter -999995, wdShapeRight -999996; wrap: wdWrapSquare 0, wdWrapNone 3 (in front of text), wdWrapTopBottom 4
function Place($shape, [string]$text, [single]$left, [int]$relative, [int]$wrap) {
    $shape.TextFrame.TextRange.Text = $text
    $shape.TextFrame.TextRange.ParagraphFormat.ReadingOrder = 1
    $shape.RelativeHorizontalPosition = $relative
    $shape.RelativeVerticalPosition = 2                            # wdRelativeVerticalPositionParagraph
    $shape.Left = $left
    $shape.Top = 0
    $shape.WrapFormat.Type = $wrap
    $shape.Left = $left                                            # again: a new wrap can move it
    $shape.Top = 0
}
function Save([string]$name) {
    foreach ($shape in $script:d.Shapes) { Write-Output ("  {0}: wrap={1} rel={2} left={3} distL={4} rot={5}" -f $(if ($shape.Type -eq 6) { "group" } else { $shape.TextFrame.TextRange.Text.Trim() }), $shape.WrapFormat.Type, $shape.RelativeHorizontalPosition, $shape.Left, $shape.WrapFormat.DistanceLeft, $shape.Rotation) }
    $script:d.SaveAs2("$dir\$name.rtf", 6)
    $script:d.SaveAs2("$dir\$name.doc", 0)
    $script:d.Close(0); $script:d = $null
    foreach ($file in "$name.doc", "$name.rtf") {                 # Word's own print of what each file holds
        $script:d = $w.Documents.Open("$dir\$file", $false, $true)
        $script:d.SaveAs2("$dir\" + $file.Replace(".", "-") + "-word.pdf", 17)
        $script:d.Close(0); $script:d = $null
    }
}

try {
    $d = NewDocument 0
    $d.Paragraphs.Item(1).Range.Text = "Floating fixture"
    # 1: 200 pt from the column's left edge, in front of text
    $p = Para "colanchor $filler"
    Place ($d.Shapes.AddTextbox(1, 0, 0, 100, 30, $p)) "tbcolumn" 200 2 3
    # 2: 40 pt from the page edge, above and below the text, in a paragraph with a footnote
    $p = Para "noteanchor $filler"
    $d.Footnotes.Add($d.Paragraphs.Item($d.Paragraphs.Count).Range.Words.Item(1)).Range.Text = "notebody"
    Place ($d.Shapes.AddTextbox(1, 0, 0, 100, 30, $d.Paragraphs.Item($d.Paragraphs.Count).Range)) "tbnote" 40 1 4
    # 3: a rectangle with text rotated 30 degrees, centred in the column, above and below
    $p = Para "rotanchor $filler"
    $s = $d.Shapes.AddShape(1, 0, 0, 120, 50, $p)
    Place $s "rotbox" -999995 2 4
    $s.Rotation = 30
    # 4: aligned right of the margin, above and below, in a right-to-left paragraph (Hebrew text)
    $hebrew = -join ([char[]](0x05E9, 0x05DC, 0x05D5, 0x05DD, 0x0020, 0x05E2, 0x05D5, 0x05DC, 0x05DD))
    $p = Para "rtlanchor $hebrew $filler"
    $d.Paragraphs.Item($d.Paragraphs.Count).ReadingOrder = 0
    Place ($d.Shapes.AddTextbox(1, 0, 0, 100, 30, $d.Paragraphs.Item($d.Paragraphs.Count).Range)) "tbrtl" -999996 0 4
    # 5: square wrap with an 18 pt gap on the left, 300 pt from the column (nearer the right edge)
    $p = Para "squareanchor $filler"
    $s = $d.Shapes.AddTextbox(1, 0, 0, 100, 30, $p)
    Place $s "tbsquare" 300 2 0
    $s.WrapFormat.DistanceLeft = 18
    # 6: 20 pt from the anchor character, square
    $p = Para "charanchor $filler"
    Place ($d.Shapes.AddTextbox(1, 0, 0, 100, 30, $p)) "tbchar" 20 3 0
    Para "end" | Out-Null
    Write-Output "floating-v1"
    Save "floating-v1"

    $d = NewDocument 11
    $d.Paragraphs.Item(1).Range.Text = "Group fixture"
    $p = Para "groupanchor $filler"
    $a = $d.Shapes.AddShape(1, 0, 0, 80, 30, $p); $a.TextFrame.TextRange.Text = "grpone"
    $b = $d.Shapes.AddShape(1, 90, 0, 80, 30, $p); $b.TextFrame.TextRange.Text = "grptwo"
    $n = [int]$d.Shapes.Count
    $g = $d.Shapes.Range([object[]]@(($n - 1), $n)).Group()
    $g.RelativeHorizontalPosition = 0
    $g.RelativeVerticalPosition = 2
    $g.WrapFormat.Type = 4
    $g.Left = -999996
    $g.Top = 0
    Para "end" | Out-Null
    Write-Output "floating-group-v1"
    Save "floating-group-v1"
}
catch { Write-Output ("failed: 0x" + $_.Exception.HResult.ToString("X8") + " " + $_.Exception.GetType().Name + " at " + $_.InvocationInfo.ScriptLineNumber) }
finally {
    if ($d) { $d.Close(0) }
    $doNotSave = 0
    $w.Quit([ref]$doNotSave)   # Quit(0): SaveChanges is a by-ref Variant from PowerShell
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($w) | Out-Null
}
