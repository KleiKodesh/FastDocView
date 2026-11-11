# Saves brackets-cambria-v1.docx / .pdf and brackets-dejavu-v1.docx / .pdf with Word 16 (COM) from brackets-base.docx
# (scripts/make-math-brackets-docx.py): the same tall brackets and growing integrals in each math font Word offers here.
$ErrorActionPreference = 'Stop'
$dir = $PSScriptRoot
$w = New-Object -ComObject Word.Application
$w.Visible = $false
$w.DisplayAlerts = 0
try {
    foreach ($pair in @(@('cambria', 'Cambria Math'), @('dejavu', 'DejaVu Math TeX Gyre'))) {
        $d = $w.Documents.Open("$dir\brackets-base.docx", $false, $true)
        $d.SetCompatibilityMode(15)
        $d.OMathFontName = $pair[1]
        foreach ($m in $d.OMaths) { $m.Range.Font.Name = $pair[1] }   # the runs name their font: the setting alone changes only new equations
        $d.SaveAs2("$dir\brackets-$($pair[0])-v1.docx", 16)
        $d.SaveAs2("$dir\brackets-$($pair[0])-v1.pdf", 17)
        $d.Close(0)
    }
}
finally { $w.Quit() }
