# Builds w14-arabic-calt-v1.docx with Word 16 (COM): Calibri Arabic beh-beh-beh / beh-hah-beh (Calibri arab calt changes them),
# then David Libre Hebrew after Latin "fi"; first of each pair with ligatures none + no contextual alternates, second with all +
# contextual alternates. Writes the .xps beside it for the glyph check (NOTES.md). Run: powershell -File make-w14-arabic-calt-fixture.ps1
$out = "$PSScriptRoot\w14-arabic-calt-v1"
function U([int[]]$c) { -join ($c | ForEach-Object { [char]$_ }) }
$arb = U 0x0628,0x0628,0x0628,0x0020,0x0628,0x062D,0x0628
$heb = U 0x05D5,0x05C2
$w = New-Object -ComObject Word.Application
$w.Visible = $false
$w.DisplayAlerts = 0
try {
  $d = $w.Documents.Add()
  $d.Content.Text = (@($arb,$arb,("fi " + $heb),("fi " + $heb)) -join "`r")
  $p = $d.Paragraphs
  for ($i=1; $i -le 4; $i++) { $r=$p.Item($i).Range; $r.Font.Size = 28; $r.Font.SizeBi = 28
    if ($i -le 2) { $r.ParagraphFormat.ReadingOrder = 0; $r.Font.Name = "Calibri"; $r.Font.NameBi = "Calibri" } else { $r.ParagraphFormat.ReadingOrder = 1; $r.Font.Name = "David Libre"; $r.Font.NameBi = "David Libre" }
    if ($i % 2 -eq 1) { $r.Font.Ligatures = 0; $r.Font.ContextualAlternates = 0 } else { $r.Font.Ligatures = 15; $r.Font.ContextualAlternates = 1 } }
  $d.SaveAs2("$out.docx", 16)
  $d.SaveAs2("$out.xps", 18)   # the XPS is for the check only; it is not committed
  $d.Close(0)
} finally { $w.Quit([ref]0) }
