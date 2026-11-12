# Has Word 16 (COM) open note-numbering-v1.docx (written by make-note-numbering-docx.py) and save the .doc, the .rtf and
# Word's own PDF, the reference for the printed note numbers. Run from this folder: powershell -File make-note-numbering-fixtures.ps1
$ErrorActionPreference = 'Stop'
$dir = $PSScriptRoot
$w = New-Object -ComObject Word.Application
$w.Visible = $false
$w.DisplayAlerts = 0
try {
    $d = $w.Documents.Open([IO.Path]::Combine($dir, "note-numbering-v1.docx"), $false, $true)
    try {
        $d.SaveAs2([IO.Path]::Combine($dir, "note-numbering-v1.doc"), 0)
        $d.SaveAs2([IO.Path]::Combine($dir, "note-numbering-v1.rtf"), 6)
        $d.SaveAs2([IO.Path]::Combine($dir, "note-numbering-v1-word.pdf"), 17)
    } finally { $d.Close(0) }
} finally { $w.Quit() }
