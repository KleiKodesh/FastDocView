# Builds tables-v1.doc / tables-v1.rtf with Word 16 (COM): merges both ways, nesting 3 deep, a right-to-left table,
# per-cell borders and colours, vertical alignment. Also records whether Word accepts a footnote inside a text box.
# Run from this folder: powershell -File make-tables-fixtures.ps1
$ErrorActionPreference = 'Stop'
$dir = $PSScriptRoot
$w = New-Object -ComObject Word.Application
$w.Visible = $false
$w.DisplayAlerts = 0
try {
    $d = $w.Documents.Add()
    $d.Content.LanguageID = 1033
    $d.Content.ParagraphFormat.ReadingOrder = 1                     # left-to-right document (the Word UI here defaults to Hebrew)
    function EndRange { $r = $script:d.Content; $r.Collapse(0); return $r }
    function Ltr($t) { $t.TableDirection = 1; $t.Borders.Enable = $true }
    function Para([string]$text) { $r = EndRange; $r.InsertParagraphAfter(); $r = EndRange; $r.Text = $text; return $r }
    $d.Paragraphs.Item(1).Range.Text = "Tables fixture"
    # 1: merges - A1:B1 horizontally, C2:C3 vertically
    Para "merges" | Out-Null; $r = EndRange; $r.InsertParagraphAfter(); $r = EndRange
    $t = $d.Tables.Add($r, 3, 3)
    for ($i = 1; $i -le 3; $i++) { for ($j = 1; $j -le 3; $j++) { $t.Cell($i, $j).Range.Text = "m$i$j" } }
    Ltr $t
    $t.Cell(2, 3).Merge($t.Cell(3, 3))
    $t.Cell(1, 1).Merge($t.Cell(1, 2))
    # 2: nested 3 deep
    Para "nested" | Out-Null; $r = EndRange; $r.InsertParagraphAfter(); $r = EndRange
    $t = $d.Tables.Add($r, 1, 2)
    Ltr $t
    $t.Cell(1, 2).Range.Text = "outer2"
    $inner = $t.Cell(1, 1).Range.Tables.Add($t.Cell(1, 1).Range, 2, 2)
    Ltr $inner
    $deep = $inner.Cell(1, 1).Range.Tables.Add($inner.Cell(1, 1).Range, 1, 2)
    Ltr $deep
    $deep.Cell(1, 1).Range.Text = "deep1"; $deep.Cell(1, 2).Range.Text = "deep2"
    $inner = $t.Cell(1, 1).Tables.Item(1)
    while ($inner.Rows.Count -lt 2) { $inner.Rows.Add() | Out-Null }
    $inner.Cell(1, 2).Range.Text = "inner12"; $inner.Cell(2, 1).Range.Text = "inner21"; $inner.Cell(2, 2).Range.Text = "inner22"

    # 3: right-to-left table
    Para "rtl" | Out-Null; $r = EndRange; $r.InsertParagraphAfter(); $r = EndRange
    $t = $d.Tables.Add($r, 1, 3)
    $t.Cell(1, 1).Range.Text = "first"; $t.Cell(1, 2).Range.Text = "second"; $t.Cell(1, 3).Range.Text = "third"
    $t.TableDirection = 0
    # 4: per-cell borders and colours (no table borders)
    Para "borders" | Out-Null; $r = EndRange; $r.InsertParagraphAfter(); $r = EndRange
    $t = $d.Tables.Add($r, 1, 3)
    Ltr $t
    $t.Borders.Enable = $false
    $t.Cell(1, 1).Range.Text = "red-top"; $t.Cell(1, 2).Range.Text = "yellow"; $t.Cell(1, 3).Range.Text = "blue-box"
    $b = $t.Cell(1, 1).Borders.Item(-1); $b.LineStyle = 1; $b.LineWidth = 18; $b.Color = 255            # top, 2.25 pt, red
    $t.Cell(1, 2).Shading.BackgroundPatternColor = 65535                                                 # yellow
    foreach ($side in -1, -2, -3, -4) { $b = $t.Cell(1, 3).Borders.Item($side); $b.LineStyle = 7; $b.LineWidth = 6; $b.Color = 16711680 }   # double, blue
    # 5: vertical alignment in a 40 pt row
    Para "valign" | Out-Null; $r = EndRange; $r.InsertParagraphAfter(); $r = EndRange
    $t = $d.Tables.Add($r, 1, 3)
    Ltr $t
    $t.Rows.Item(1).HeightRule = 2; $t.Rows.Item(1).Height = 40
    $t.Cell(1, 1).Range.Text = "top"; $t.Cell(1, 2).Range.Text = "center"; $t.Cell(1, 3).Range.Text = "bottom"
    $t.Cell(1, 1).VerticalAlignment = 0; $t.Cell(1, 2).VerticalAlignment = 1; $t.Cell(1, 3).VerticalAlignment = 3
    Para "end" | Out-Null
    $d.SaveAs2("$dir\tables-v1.doc", 0)
    $d.SaveAs2("$dir\tables-v1.rtf", 6)
    $d.Close(0)
    # footnote inside a text box: does Word allow it?
    $d = $w.Documents.Add()
    $s = $d.Shapes.AddTextbox(1, 50, 50, 200, 80)
    $s.TextFrame.TextRange.Text = "box text"
    $result = "allowed"
    try { $f = $d.Footnotes.Add($s.TextFrame.TextRange); $result = "allowed (footnotes: " + $d.Footnotes.Count + ")" } catch { $result = "refused (HRESULT 0x" + $_.Exception.HResult.ToString("X8") + ")" }
    $e = "allowed"
    try { $f = $d.Endnotes.Add($s.TextFrame.TextRange); $e = "allowed (endnotes: " + $d.Endnotes.Count + ")" } catch { $e = "refused (HRESULT 0x" + $_.Exception.HResult.ToString("X8") + ")" }
    "footnote in text box: $result" | Out-File "$dir\footnote-in-textbox.txt" -Encoding ascii
    "endnote in text box: $e" | Out-File "$dir\footnote-in-textbox.txt" -Encoding ascii -Append
    $d.Close(0)
}
finally { $w.Quit() }
