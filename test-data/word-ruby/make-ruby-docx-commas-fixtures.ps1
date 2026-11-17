# Builds ruby-docx-commas-v1.docx / .doc / .pdf / .doc.pdf with Word 16 (COM): phonetic guides whose guide or base holds
# commas or brackets in different places, guide in Courier New over a Calibri base - how Word prints them from a .docx (its
# own w:ruby) and from the .doc it saves (an EQ field).
# Run from this folder: powershell -File make-ruby-docx-commas-fixtures.ps1   (the harness reads the saved files; Word is not needed to test)
$ErrorActionPreference = 'Stop'
$dir = $PSScriptRoot
$cases = @(
    @{ Base = 'base';    Guide = 'g,1' },
    @{ Base = 'b,1';     Guide = 'g' },
    @{ Base = 'base';    Guide = 'g,1,2' },
    @{ Base = 'b,2';     Guide = 'g,1' },
    @{ Base = 'ef';      Guide = 'ab)cd' },
    @{ Base = 'cd)ef';   Guide = 'ab' },
    @{ Base = 'c(d';     Guide = 'a(b' },
    @{ Base = 'base,1';  Guide = 'g,1' }
)
$w = New-Object -ComObject Word.Application
$w.Visible = $false
$w.DisplayAlerts = 0
try {
    $d = $w.Documents.Add()
    $d.Content.Text = (0..($cases.Count - 1) | ForEach-Object { "ruby comma ${_}: before " + $cases[$_].Base + " after" }) -join "`r"
    for ($i = $cases.Count - 1; $i -ge 0; $i--) {
        $c = $cases[$i]
        $r = $d.Paragraphs.Item($i + 1).Range
        $start = $r.Start + $r.Text.IndexOf($c.Base)
        $d.Range($start, $start + $c.Base.Length).PhoneticGuide($c.Guide, 0, 0, 10, 'Courier New')
    }
    $d.SaveAs2("$dir\ruby-docx-commas-v1.docx", 16)
    $d.SaveAs2("$dir\ruby-docx-commas-v1.doc", 0)
    $d.Close([ref]0)
    foreach ($e in 'docx', 'doc') {
        $p = $w.Documents.Open("$dir\ruby-docx-commas-v1.$e", $false, $true)
        $p.SaveAs2("$dir\ruby-docx-commas-v1.$e.pdf", 17)
        $p.Close([ref]0)
    }
}
finally { $w.Quit([ref]0) }   # 0: do not save (no dialog)
