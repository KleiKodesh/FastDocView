# Builds ruby-edges-v1.docx / .doc / .rtf and Word 16's PDF of each (COM). Word's UI keeps only the EQ result of a phonetic
# guide (no tab, note or bookmark survives inside it), so Word makes a skeleton with one ruby per paragraph and
# make-ruby-edges.py rewrites the rubies into the forms a .docx may still hold: a tab, a footnote reference or a bookmark after a
# ")" in the base; a footnote reference and a bookmark in a base that a comma guide replaces; a ruby with a ")" ending its
# paragraph; a ";" in guide and base. Word then opens that .docx, prints it, saves it as .doc and RTF and prints those.
# Run from this folder: powershell -File make-ruby-edges-fixtures.ps1 (the harness reads the saved files; Word is not needed to test)
$ErrorActionPreference = 'Stop'
$dir = $PSScriptRoot
$w = New-Object -ComObject Word.Application
$w.Visible = $false
$w.DisplayAlerts = 0
try {
    $d = $w.Documents.Add()
    $d.Content.Text = ((0..6 | ForEach-Object { "ruby edge ${_}: before base$_ after" }) -join "`r")
    for ($i = 6; $i -ge 0; $i--) {
        $r = $d.Paragraphs.Item($i + 1).Range; $s = $r.Start + $r.Text.IndexOf("base$i")
        $d.Range($s, $s + 5).PhoneticGuide("g$i", 0, 0, 10, 'Courier New')
    }
    # two footnotes (ids 1 and 2) for make-ruby-edges.py to reference from inside rubies; it removes these references
    $holder = $d.Content; $holder.InsertParagraphAfter(); $holder = $d.Paragraphs.Item(8).Range
    $holder.InsertAfter('notes holder'); $end = $d.Paragraphs.Item(8).Range
    $d.Footnotes.Add($d.Range($end.End - 1, $end.End - 1), [Type]::Missing, 'note one') | Out-Null
    $end = $d.Paragraphs.Item(8).Range
    $d.Footnotes.Add($d.Range($end.End - 1, $end.End - 1), [Type]::Missing, 'note two') | Out-Null
    $d.SaveAs2("$dir\ruby-edges-skeleton.docx", 16)
    $d.Close([ref]0)
    & python -I "$dir\make-ruby-edges.py"
    Remove-Item "$dir\ruby-edges-skeleton.docx"
    $e =$w.Documents.Open("$dir\ruby-edges-v1.docx", $false, $true)
    $e.SaveAs2("$dir\ruby-edges-v1.docx.pdf", 17)
    $e.SaveAs2("$dir\ruby-edges-v1.rtf", 6)
    $e.SaveAs2("$dir\ruby-edges-v1.doc", 0)
    $e.Close([ref]0)
    foreach ($x in 'doc', 'rtf') {
        $p = $w.Documents.Open("$dir\ruby-edges-v1.$x", $false, $true)
        $p.SaveAs2("$dir\ruby-edges-v1.$x.pdf", 17)
        $p.Close([ref]0)
    }
    # ";" in place of the comma between guide and base (make-ruby-edges-semicolon.py), printed by Word
    & python -I "$dir\make-ruby-edges-semicolon.py"
    $p = $w.Documents.Open("$dir\ruby-edges-semicolon-v1.rtf", $false, $true)
    $p.SaveAs2("$dir\ruby-edges-semicolon-v1.pdf", 17)
    $p.Close([ref]0)
}
finally { $w.Quit([ref]0) }   # 0: do not save (no dialog)
