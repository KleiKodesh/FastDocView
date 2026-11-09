# Builds text-effects-v1.doc with Word 16 (COM): one paragraph per character scale
# (Font.Scaling 80 / 100 / 150).
# The harness (TextEffectsFidelityTest) checks the renderer against these. Run from this folder:
# powershell -File make-text-effects-fixture.ps1
$ErrorActionPreference = 'Stop'
$dir = $PSScriptRoot
$w = New-Object -ComObject Word.Application
$w.Visible = $false
$w.DisplayAlerts = 0
try {
    $d = $w.Documents.Add()
    $d.Content.LanguageID = 1033
    $d.Paragraphs.Item(1).Range.Text = "scaled eighty words"
    $d.Content.InsertParagraphAfter()
    $d.Paragraphs.Item(2).Range.Text = "normal words"
    $d.Content.InsertParagraphAfter()
    $d.Paragraphs.Item(3).Range.Text = "scaled wide words"
    $d.Paragraphs.Item(1).Range.Font.Scaling = 80
    $d.Paragraphs.Item(3).Range.Font.Scaling = 150
    $i = 0
    foreach ($p in $d.Paragraphs) { "p{0}: scaling {1} [{2}]" -f $i, $p.Range.Font.Scaling, $p.Range.Text.TrimEnd("`r"); $i++ }
    $d.SaveAs2("$dir\text-effects-v1.doc", 0)
    $d.Close(0)
}
finally { $w.Quit() }
