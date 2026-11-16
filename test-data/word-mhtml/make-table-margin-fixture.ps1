param([string]$out = $PSScriptRoot)  # Word 16, never visible; writes table-margin-v1.mht (SaveAs2 9) and table-margin-v1.pdf (SaveAs2 17)
$w = New-Object -ComObject Word.Application
try {
  $w.DisplayAlerts = 0; $w.Visible = $false
  $d = $w.Documents.Add()
  $d.Content.ParagraphFormat.ReadingOrder = 1          # left-to-right
  $d.Content.ParagraphFormat.Alignment = 0
  $r = $d.Content
  $r.InsertAfter("Body line at the left margin."); $r.InsertParagraphAfter()
  $end = $d.Content.End - 1; $t = $d.Tables.Add($d.Range($end, $end), 2, 2)
  $t.Borders.Enable = 1
  $t.Rows.LeftIndent = -5.4                             # the cell text lines up with the body text (Word 2003 and older default; margin-left:-5.4pt)
  $t.Cell(1,1).Range.Text = "CellA"; $t.Cell(1,2).Range.Text = "CellB"; $t.Cell(2,1).Range.Text = "CellC"; $t.Cell(2,2).Range.Text = "CellD"
  $d.SaveAs2("$out\table-margin-v1.mht", 9)
  $d.SaveAs2("$out\table-margin-v1.pdf", 17)
  $d.Close([ref]0)
} finally { $w.Quit([ref]0) }
