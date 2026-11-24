# Builds eq-layout-v1.docx / .rtf / .doc / .pdf with Word 16 (COM, hidden, no alerts): one paragraph per EQ field case,
# "case <name>: <EQ field> end" in Times New Roman 20 pt, so Word's PDF gives each case's glyph positions
# (\o alignment, integral limits with and without \in, \b with one bracket option, bracket sizes, line heights).
# Run from this folder: powershell -File make-eq-layout-fixture.ps1   (the harness reads the saved files; Word is not needed to test)
$ErrorActionPreference = 'Stop'
$dir = $PSScriptRoot
$base = Join-Path $dir 'eq-layout-v1'
$cases = @(
    @('o', 'EQ \o(W,i)'), @('oal', 'EQ \o\al(W,i)'), @('oac', 'EQ \o\ac(W,i)'), @('oar', 'EQ \o\ar(W,i)'),
    @('oi', 'EQ \o(i,W)'), @('oari', 'EQ \o\ar(i,W)'),
    @('int', 'EQ \i(0,1,x)'), @('intin', 'EQ \i\in(0,1,x)'), @('sum', 'EQ \i\su(0,1,x)'), @('sumin', 'EQ \i\su\in(0,1,x)'),
    @('fc', 'EQ \i\fc\Q(0,1,x)'), @('vc', 'EQ \i\vc\Q(0,1,x)'),
    @('blc', 'EQ \b\lc\[(x)'), @('brc', 'EQ \b\rc\](x)'), @('blcbrace', 'EQ \b\lc\{(x)'), @('blcbar', 'EQ \b\lc\|(x)'),
    @('bfrac', 'EQ \b(\f(1,2))'), @('ba2', 'EQ \b(\a(1,2))'), @('ba3', 'EQ \b(\a(1,2,3))'), @('ba5', 'EQ \b(\a(1,2,3,4,5))'),
    @('plain', 'EQ x'), @('frac', 'EQ \f(1,2)'), @('a3', 'EQ \a(1,2,3)')
)
$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
$w = New-Object -ComObject Word.Application
$mine = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id } | ForEach-Object Id)
try {
    $w.Visible = $false
    $w.DisplayAlerts = 0
    $d = $w.Documents.Add()
    $d.ActiveWindow.View.Type = 3
    $d.Content.LanguageID = 1033
    $d.Content.Font.Name = 'Times New Roman'
    $d.Content.Font.Size = 20
    $d.Content.ParagraphFormat.SpaceAfter = 0
    $d.Content.ParagraphFormat.LineSpacingRule = 0
    $d.Content.ParagraphFormat.ReadingOrder = 1
    $d.Content.ParagraphFormat.Alignment = 0
    $first = $true
    foreach ($c in $cases) {
        if (-not $first) { $d.Content.InsertParagraphAfter() }
        $first = $false
        $p = $d.Paragraphs.Item($d.Paragraphs.Count).Range
        $p.Text = 'case ' + $c[0] + ': '
        $at = $d.Paragraphs.Item($d.Paragraphs.Count).Range
        $at.MoveEnd(1, -1) | Out-Null
        $at.Collapse(0)
        $f = $d.Fields.Add($at, -1, $c[1], $false)
        $tail = $d.Paragraphs.Item($d.Paragraphs.Count).Range
        $tail.MoveEnd(1, -1) | Out-Null
        $tail.Collapse(0)
        $tail.InsertAfter(' end')
    }
    $d.Fields.Update() | Out-Null
    $d.SaveAs2($base + '.docx', 16)
    $d.SaveAs2($base + '.pdf', 17)
    $d.SaveAs2($base + '.rtf', 6)
    $d.SaveAs2($base + '.doc', 0)
    $d.Close(0)
    Write-Output ('saved ' + $cases.Count + ' cases')
}
finally {
    try { $w.Quit(0) } catch { }
    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($w) } catch { }
    foreach ($id in $mine) { try { Stop-Process -Id $id -Force -ErrorAction Stop } catch { } }
}
