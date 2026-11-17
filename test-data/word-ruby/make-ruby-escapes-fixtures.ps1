# Builds ruby-escapes-v1.docx / .doc / .rtf / .pdf with Word 16 (COM): phonetic guides whose guide or base holds brackets,
# commas, semicolons or backslashes, and guide fonts with spaces in their names - how Word writes them into the EQ field of
# .doc and RTF (escapes, quoting) and how its PDF draws them.
# Run from this folder: powershell -File make-ruby-escapes-fixtures.ps1   (the harness reads the saved files; Word is not needed to test)
$ErrorActionPreference = 'Stop'
$dir = $PSScriptRoot
$cases = @(
    @{ Base = 'base(0)';  Guide = 'g(0)';   Font = 'Times New Roman' },
    @{ Base = 'base,1';   Guide = 'g,1';    Font = 'Courier New' },
    @{ Base = 'base;2';   Guide = 'g;2';    Font = 'Arial' },
    @{ Base = 'base\3';   Guide = 'g\3';    Font = 'Arial' },
    @{ Base = 'base)4(';  Guide = 'g)4(';   Font = 'Comic Sans MS' },
    @{ Base = 'base 5';   Guide = 'g 5 x';  Font = 'Segoe UI' }
)
$w = New-Object -ComObject Word.Application
$w.Visible = $false
$w.DisplayAlerts = 0
try {
    $d = $w.Documents.Add()
    $d.Content.Text = (0..($cases.Count - 1) | ForEach-Object { "ruby escape ${_}: before " + $cases[$_].Base + " after" }) -join "`r"
    for ($i = $cases.Count - 1; $i -ge 0; $i--) {
        $c = $cases[$i]
        $r = $d.Paragraphs.Item($i + 1).Range
        $start = $r.Start + $r.Text.IndexOf($c.Base)
        $target = $d.Range($start, $start + $c.Base.Length)
        $target.PhoneticGuide($c.Guide, 0, 0, 10, $c.Font)
    }
    $d.SaveAs2("$dir\ruby-escapes-v1.docx", 16)
    $d.SaveAs2("$dir\ruby-escapes-v1.rtf", 6)
    $d.SaveAs2("$dir\ruby-escapes-v1.doc", 0)
    $d.Close([ref]0)
    # Word's print of the .doc (the .rtf prints the same; the .docx differs - see docs/UNSUPPORTED-FEATURES.md)
    $p = $w.Documents.Open("$dir\ruby-escapes-v1.doc", $false, $true)
    $p.SaveAs2("$dir\ruby-escapes-v1.pdf", 17)
    $p.Close([ref]0)
    # the probe: backslash "escapes" and unquoted multi-word fonts, written by make-ruby-escapes-probe.py from the .rtf above
    & python -I "$dir\make-ruby-escapes-probe.py"
    $p = $w.Documents.Open("$dir\ruby-escapes-probe-v1.rtf", $false, $true)
    $p.SaveAs2("$dir\ruby-escapes-probe-v1.pdf", 17)
    $p.Close([ref]0)
}
finally { $w.Quit([ref]0) }   # 0: do not save (no dialog)
