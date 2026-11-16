# Builds w14-scripts-v1.docx with Word 16 (COM): David Libre Hebrew (vav + U+05C2, a liga pair in that font; shin with points)
# and Arial Arabic (the Allah ligature), each with ligatures none + no contextual alternates and with all + contextual
# alternates. Word's XPS glyphs are the same for both settings (NOTES.md). Run: powershell -File make-w14-scripts-fixture.ps1
$dir = $PSScriptRoot
function U([int[]]$c) { -join ($c | ForEach-Object { [char]$_ }) }
$heb = (U 0x05D5,0x05C2) + " " + (U 0x05E9,0x05C1,0x05B8,0x05DC,0x05D5,0x05B9,0x05DD)
$arb = (U 0x0627,0x0644,0x0644,0x0647) + " " + (U 0x0628,0x0633,0x0645,0x0020,0x0643,0x062A,0x0627,0x0628)
$w = New-Object -ComObject Word.Application
$w.Visible = $false
$w.DisplayAlerts = 0
try {
    $d = $w.Documents.Add()
    $d.Content.Text = (@($heb,$heb,$arb,$arb) -join "`r")
    $p = $d.Paragraphs
    for ($i=1; $i -le 4; $i++) { $r=$p.Item($i).Range; $r.ParagraphFormat.ReadingOrder = 0; $r.Font.SizeBi = 28; $r.Font.Size = 28
        if ($i -le 2) { $r.Font.NameBi = "David Libre"; $r.Font.Name = "David Libre" } else { $r.Font.NameBi = "Arial"; $r.Font.Name = "Arial" }
        if ($i % 2 -eq 1) { $r.Font.Ligatures = 0; $r.Font.ContextualAlternates = 0 } else { $r.Font.Ligatures = 15; $r.Font.ContextualAlternates = 1 } }
    $d.SaveAs2("$dir\w14-scripts-v1.docx", 16)
    $d.Close(0)
} finally { $w.Quit([ref]0) }
