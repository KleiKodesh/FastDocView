# Builds tables-extra-v1.doc / .rtf with Word 16 (COM): row heights (exact 30 pt, at least 50 pt) with vertical alignment, an
# explicit "no border" on one cell of a bordered table, a border in a colour outside Word's 16-colour palette (Word saves it with
# sprmTSetBrc in .doc), and two tables made to touch by deleting the paragraph between them. Writes what Word reports to
# tables-extra-word.txt (table count after the touch, table directions in tables-v1.doc, the no-border cell's line style).
# Run from this folder: powershell -File make-tables-extra-fixtures.ps1
$ErrorActionPreference = 'Stop'
$dir = $PSScriptRoot
$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
$w = New-Object -ComObject Word.Application
$started = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id } | ForEach-Object Id)
$w.Visible = $false
$w.DisplayAlerts = 0
$report = @()
try {
    $d = $w.Documents.Add()
    $d.Content.LanguageID = 1033
    $d.Content.ParagraphFormat.ReadingOrder = 1
    function EndRange { $r = $script:d.Content; $r.Collapse(0); return $r }
    function Para([string]$text) { $r = EndRange; $r.InsertParagraphAfter(); $r = EndRange; $r.Text = $text; return $r }
    $d.Paragraphs.Item(1).Range.Text = "Tables extra fixture"
    # 1: row heights + vertical alignment, explicit no border, a custom-colour border
    Para "heights" | Out-Null; $r = EndRange; $r.InsertParagraphAfter(); $r = EndRange
    $t = $d.Tables.Add($r, 2, 3)
    $t.TableDirection = 1; $t.Borders.Enable = $true
    $t.Rows.Item(1).HeightRule = 2; $t.Rows.Item(1).Height = 30                                       # exact
    $t.Rows.Item(2).HeightRule = 1; $t.Rows.Item(2).Height = 50                                       # at least
    $t.Cell(1, 1).Range.Text = "exact-top"; $t.Cell(1, 2).Range.Text = "exact-mid"; $t.Cell(1, 3).Range.Text = "custom-colour"
    $t.Cell(2, 1).Range.Text = "least-top"; $t.Cell(2, 2).Range.Text = "no-top-line"; $t.Cell(2, 3).Range.Text = "least-bottom"
    $t.Cell(1, 2).VerticalAlignment = 1; $t.Cell(2, 3).VerticalAlignment = 3
    $t.Cell(2, 2).Borders.Item(-1).LineStyle = 0                                                      # top: none (explicit)
    $t.Cell(2, 2).Borders.Item(-3).LineStyle = 0                                                      # bottom: none (explicit)
    foreach ($side in -1, -2, -3, -4) { $b = $t.Cell(1, 3).Borders.Item($side); $b.LineStyle = 1; $b.LineWidth = 12; $b.Color = 0x563412 }   # RGB 12 34 56
    # 2: two tables, then the paragraph between them deleted
    Para "touching" | Out-Null; $r = EndRange; $r.InsertParagraphAfter(); $r = EndRange
    $a = $d.Tables.Add($r, 2, 2); $a.TableDirection = 1; $a.Borders.Enable = $true
    $a.Cell(1, 1).Range.Text = "a11"; $a.Cell(1, 2).Range.Text = "a12"; $a.Cell(2, 1).Range.Text = "a21"; $a.Cell(2, 2).Range.Text = "a22"
    $r = EndRange; $r.InsertParagraphAfter(); $r = EndRange
    $b2 = $d.Tables.Add($r, 2, 2); $b2.TableDirection = 1; $b2.Borders.Enable = $true
    $b2.Borders.OutsideLineStyle = 1; $b2.Borders.OutsideLineWidth = 18; $b2.Borders.OutsideColor = 255
    $b2.Cell(1, 1).Range.Text = "b11"; $b2.Cell(1, 2).Range.Text = "b12"; $b2.Cell(2, 1).Range.Text = "b21"; $b2.Cell(2, 2).Range.Text = "b22"
    $count = $d.Tables.Count
    $between = $a.Range; $between.Collapse(0); $between.Paragraphs.Item(1).Range.Delete() | Out-Null
    $report += "tables before the touch: $count, after: " + $d.Tables.Count
    $report += "after the touch, table 2 rows: " + $d.Tables.Item(2).Rows.Count
    Para "end" | Out-Null
    $d.SaveAs2("$dir\tables-extra-v1.doc", 0)
    $d.SaveAs2("$dir\tables-extra-v1.rtf", 6)
    $d.Close(0)
    $d = $w.Documents.Open("$dir\tables-extra-v1.doc", $false, $true)
    $t = $d.Tables.Item(1)
    $report += "reopened .doc: no-top-line cell top LineStyle " + $t.Cell(2, 2).Borders.Item(-1).LineStyle + ", custom-colour cell top Color 0x" + $t.Cell(1, 3).Borders.Item(-1).Color.ToString("X6")
    $report += "reopened .doc: tables " + $d.Tables.Count
    $d.Close(0)
    $d = $w.Documents.Open("$dir\tables-v1.doc", $false, $true)
    for ($i = 1; $i -le $d.Tables.Count; $i++) { $report += "tables-v1.doc table $i TableDirection " + $d.Tables.Item($i).TableDirection + " (0 rtl, 1 ltr) first cell: " + $d.Tables.Item($i).Range.Cells.Item(1).Range.Text.Trim([char]7, [char]13) }
    $d.Close(0)
}
finally {
    $w.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($w) | Out-Null
    foreach ($id in $started) { Stop-Process -Id $id -Force -ErrorAction SilentlyContinue }
}
$report | Out-File "$dir\tables-extra-word.txt" -Encoding ascii
