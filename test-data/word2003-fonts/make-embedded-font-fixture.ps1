param([string]$out = $PSScriptRoot)  # Word 16, never visible: one document in Caladea (free font) regular + bold with a footnote, "Embed fonts in the file" on;
# saved as embedded-font-v1.docx (SaveAs2 16), as Word 2003 XML embedded-font-v1.xml (SaveAs2 11) and printed to embedded-font-v1.pdf
$word = New-Object -ComObject Word.Application
try {
  $word.DisplayAlerts = 0; $word.Visible = $false
  $d = $word.Documents.Add()
  $d.EmbedTrueTypeFonts = $true; $d.SaveSubsetFonts = $true
  foreach ($style in @(-1, -30, -38)) { $d.Styles($style).Font.Name = "Caladea" }   # Normal, Footnote Text, Footnote Reference
  $r = $d.Content
  $r.Text = "Embedded font sample text"
  $r.Font.Name = "Caladea"; $r.Font.Size = 20
  $r.InsertParagraphAfter()
  $p = $d.Paragraphs(2).Range; $p.Text = "Second line with a note"; $p.Font.Name = "Caladea"; $p.Font.Bold = $true
  $d.Footnotes.Add($d.Paragraphs(2).Range.Characters.Last, [Type]::Missing, "Note body") | Out-Null
  $d.SaveAs2("$out\embedded-font-v1.xml", 11)
  $d.SaveAs2("$out\embedded-font-v1.docx", 16)
  $d.ExportAsFixedFormat("$out\embedded-font-v1.pdf", 17)
  $d.Close([ref]0)
} finally { $word.Quit([ref]0) }
