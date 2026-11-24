# Builds eq-search-text-v1.docx / .rtf / .doc with Word 16 (COM, hidden, no alerts): one paragraph per EQ case,
# "case <name>: <EQ field> tail<name>", and writes eq-search-text-v1.word-find.txt: for each case what Word itself gives as
# text - Range.Text with field codes hidden, and whether Find matches the EQ's argument text (alone and joined to the text
# around the field). DocumentTextTest checks GetText and the search against it in all three formats.
# Run from this folder: powershell -File make-eq-search-text-fixture.ps1   (the harness reads the saved files; Word is not needed to test)
$ErrorActionPreference = 'Stop'
$dir = $PSScriptRoot
$base = Join-Path $dir 'eq-search-text-v1'
# name, code, words to Find
$cases = @(
    @('f', 'EQ \f(ab,cd)', @('ab', 'cd', 'abcd')),
    @('r', 'EQ \r(3,xy)', @('xy', '3xy')),
    @('o', 'EQ \o(ef,gh)', @('ef', 'gh')),
    @('a', 'EQ \a(pq,rs)', @('pq', 'rs')),
    @('b', 'EQ \b(mn)', @('mn')),
    @('s', 'EQ \s\up8(up)', @('up')),
    @('x', 'EQ \x(bx)', @('bx')),
    @('i', 'EQ \i(lo,hi,ig)', @('lo', 'ig')),
    @('l', 'EQ \l(l1,l2)', @('l1', 'l2')),
    @('ruby', 'EQ \* jc2 \* "Font:Arial" \* hps10 \o\ad(\s\up9(gd),base)', @('gd', 'base', 'gdbase', 'basegd', ': gd', 'base tail'))
)
$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
$w = New-Object -ComObject Word.Application
$mine = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id } | ForEach-Object Id)
$report = @()
try {
    $w.Visible = $false
    $w.DisplayAlerts = 0
    $d = $w.Documents.Add()
    $d.ActiveWindow.View.Type = 3
    $d.Content.LanguageID = 1033
    foreach ($c in $cases) {
        $r = $d.Content; $r.Collapse(0); $r.InsertAfter("case $($c[0]): ")
        $r = $d.Content; $r.Collapse(0); $r.Collapse(0)
        $d.Fields.Add($r, -1, $c[1], $false) | Out-Null
        $r = $d.Content; $r.Collapse(0); $r.Collapse(0); $r.InsertAfter(" tail$($c[0])")
        if ($c -ne $cases[-1]) { $d.Content.InsertParagraphAfter() }
    }
    $w.ActiveWindow.View.ShowFieldCodes = $false
    $i = 0
    foreach ($p in $d.Paragraphs) {
        if ($i -ge $cases.Count) { break }
        $c = $cases[$i]; $i++
        $rg = $p.Range; $rg.TextRetrievalMode.IncludeFieldCodes = $false
        $text = (($rg.Text.ToCharArray() | Where-Object { [int]$_ -ge 32 }) -join '')
        $finds = foreach ($word in $c[2]) {
            $fr = $p.Range; $fd = $fr.Find; $fd.ClearFormatting(); $fd.Text = $word; $fd.Forward = $true; $fd.Wrap = 0
            "$word=" + [int]$fd.Execute()
        }
        $report += "case $($c[0]) | Range.Text [$text] | Find $($finds -join ' ')"
    }
    $d.SaveAs2("$base.docx", 16)
    $d.SaveAs2("$base.rtf", 6)
    $d.SaveAs2("$base.doc", 0)
    $d.Close(0)
} finally {
    $w.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($w) | Out-Null
    Start-Sleep -Milliseconds 500
    foreach ($id in $mine) { Stop-Process -Id $id -Force -ErrorAction SilentlyContinue }
}
$report | Set-Content -Encoding ascii "$base.word-find.txt"
$report
