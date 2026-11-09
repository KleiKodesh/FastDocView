# Builds marks-borders-v1.doc and marks-borders-v1.rtf with Word 16 (COM): one paragraph per feature - emphasis marks
# (dot, comma, circle, under dot), underline colour, a character border, a boxed + shaded paragraph, a paragraph with a
# bottom border only, and character shading with a 25 % pattern. The harness (FormattingMarksFidelityTest) checks both files.
# Run from this folder: powershell -File make-marks-borders-fixture.ps1
$ErrorActionPreference = 'Stop'
$dir = $PSScriptRoot
$w = New-Object -ComObject Word.Application
$w.Visible = $false
$w.DisplayAlerts = 0
try {
    $d = $w.Documents.Add()
    $d.Content.LanguageID = 1033
    $texts = @("dot marks", "comma marks", "circle marks", "under dot marks", "red underline", "bordered word", "boxed paragraph", "bottom rule", "patterned shade")
    $d.Paragraphs.Item(1).Range.Text = $texts[0]
    for ($i = 1; $i -lt $texts.Count; $i++) { $d.Content.InsertParagraphAfter(); $d.Paragraphs.Item($d.Paragraphs.Count).Range.Text = $texts[$i] }
    function Text($i) { $r = $script:d.Paragraphs.Item($i).Range; $r.MoveEnd(1, -1) | Out-Null; return $r }
    (Text 1).Font.EmphasisMark = 1        # wdEmphasisMarkOverSolidCircle
    (Text 2).Font.EmphasisMark = 2        # wdEmphasisMarkOverComma
    (Text 3).Font.EmphasisMark = 3        # wdEmphasisMarkOverWhiteCircle
    (Text 4).Font.EmphasisMark = 4        # wdEmphasisMarkUnderSolidCircle
    $u = Text 5; $u.Font.Underline = 1; $u.Font.UnderlineColor = 255          # wdColorRed (BGR 0x0000FF)
    $b = Text 6; $b.Font.Borders.Enable = 1; $b.Font.Borders.Item(1).LineStyle = 7; $b.Font.Borders.Item(1).LineWidth = 12; $b.Font.Borders.Item(1).Color = 16711680   # double, 1.5 pt, blue
    $p = $d.Paragraphs.Item(7)
    $p.Borders.Enable = 1
    foreach ($side in -1, -2, -3, -4) { $p.Borders.Item($side).LineStyle = 1; $p.Borders.Item($side).LineWidth = 8; $p.Borders.Item($side).Color = 32768 }   # single, 1 pt, green
    $p.Shading.BackgroundPatternColor = 65535                                 # yellow fill
    $q = $d.Paragraphs.Item(8); $q.Borders.Item(-3).LineStyle = 3; $q.Borders.Item(-3).LineWidth = 4   # bottom only, dashed 0.5 pt
    $s = Text 9; $s.Font.Shading.Texture = 250; $s.Font.Shading.ForegroundPatternColor = 0; $s.Font.Shading.BackgroundPatternColor = 16777215   # 25 %, black on white
    $d.SaveAs2("$dir\marks-borders-v1.doc", 0)
    $d.SaveAs2("$dir\marks-borders-v1.rtf", 6)
    $d.Close(0)
}
finally { $w.Quit() }
