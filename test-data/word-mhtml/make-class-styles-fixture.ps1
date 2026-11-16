param([string]$out = $PSScriptRoot)  # Word 16, never visible; writes class-styles-v1.mht (SaveAs2 9) and class-styles-v1.pdf (SaveAs2 17)
$w = New-Object -ComObject Word.Application
try {
  $w.DisplayAlerts = 0; $w.Visible = $false
  $d = $w.Documents.Add()
  $s = $d.Styles.Add("FancyPara", 1); $s.Font.Name = "Courier New"; $s.Font.Size = 20; $s.Font.Bold = 1; $s.Font.Color = 255; $s.ParagraphFormat.Alignment = 1
  $c = $d.Styles.Add("GreenChar", 2); $c.Font.Color = 32768; $c.Font.Italic = 1; $c.Font.Size = 16
  $r = $d.Content
  $r.InsertAfter("Normal text line."); $r.InsertParagraphAfter()
  $r.InsertAfter("Fancy styled paragraph."); $r.InsertParagraphAfter()
  $r.InsertAfter("Heading one text"); $r.InsertParagraphAfter()
  $r.InsertAfter("Before green after."); $r.InsertParagraphAfter()
  $d.Paragraphs(2).Style = "FancyPara"
  $d.Paragraphs(3).Style = -2
  $p4 = $d.Paragraphs(4).Range; $f = $p4.Duplicate; $f.SetRange($p4.Start + 7, $p4.Start + 12); $f.Style = "GreenChar"
  $d.SaveAs2("$out\class-styles-v1.mht", 9)
  $d.SaveAs2("$out\class-styles-v1.pdf", 17)
  $d.Close([ref]0)
} finally { $w.Quit([ref]0) }
