# Builds w14-effects-v1.docx with Word 16 (COM): one paragraph each for ligatures standard / none, number form + spacing,
# Gabriola stylistic sets (StylisticSet 6 = sets 2 and 3) and contextual alternates, a reflection preset, and 3-D text
# (bevel, extrusion, lighting). On a machine with a right-to-left default Word marks the digit runs w:rtl; the committed
# file has those w:rtl removed (and the 3-D run coloured 2E74B5) so its digits are Latin. Word's XPS output of it is the
# reference in NOTES.md "w14 OpenType typography, reflection, 3-D text". Run: powershell -File make-w14-effects-fixture.ps1
$ErrorActionPreference = 'Stop'
$dir = $PSScriptRoot
$w = New-Object -ComObject Word.Application
$w.Visible = $false
$w.DisplayAlerts = 0
try {
    $d = $w.Documents.Add()
    $lines = @("office fluffy final standard","office fluffy final none","0123456789 oldstyle tabular","0123456789 lining proportional","Gabriola stylistic set six","Gabriola contextual alternates","Reflected words wrap across lines here and keep going a long while","Bevelled text with depth")
    $d.Content.Text = ($lines -join "`r"); $d.Content.ParagraphFormat.ReadingOrder = 1; $d.Content.ParagraphFormat.Alignment = 0; $d.Content.LanguageID = 1033; $d.Content.Font.SizeBi = 28
    $p = $d.Paragraphs
    for ($i=1; $i -le $p.Count; $i++) { $p.Item($i).Range.Font.Name = "Calibri"; $p.Item($i).Range.Font.Size = 28 }
    $p.Item(1).Range.Font.Ligatures = 1        # wdLigaturesStandard
    $p.Item(2).Range.Font.Ligatures = 0
    $p.Item(3).Range.Font.NumberForm = 2       # oldStyle
    $p.Item(3).Range.Font.NumberSpacing = 2    # tabular
    $p.Item(4).Range.Font.NumberForm = 1
    $p.Item(4).Range.Font.NumberSpacing = 1
    $p.Item(5).Range.Font.Name = "Gabriola"; $p.Item(5).Range.Font.StylisticSet = 6
    $p.Item(6).Range.Font.Name = "Gabriola"; $p.Item(6).Range.Font.ContextualAlternates = 1
    $r = $p.Item(7).Range.Font.Reflection; $r.Type = 2
    $t = $p.Item(8).Range.Font.ThreeD; $t.BevelTopType = 3; $t.BevelTopDepth = 4; $t.BevelTopInset = 6; $t.Depth = 10; $t.ExtrusionColor.RGB = 255; $t.PresetLighting = 15
    "saving"
    $d.SaveAs2("$dir\w14-effects-v1.docx", 16)
    $d.Close(0)
}
finally { $w.Quit([ref]0) }
