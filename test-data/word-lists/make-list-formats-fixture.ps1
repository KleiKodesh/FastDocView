# Builds list-formats-v1-a/-b.docx / .doc / .rtf (two documents: Word runs out of memory past ~1000 lists in one) with Word 16 (COM): for every WdListNumberStyle (= MSONFC code) a one-level list per
# start value, "%1." each, plus Word's own ListFormat.ListString per paragraph in list-formats-v1-$part.expected.txt
# ("code<TAB>value<TAB>label" with non-ASCII as \uXXXX). The harness compares FastDocView's labels with that file.
# Run from this folder: powershell -File make-list-formats-fixture.ps1
$ErrorActionPreference = 'Stop'
$dir = $PSScriptRoot
$codes = @(0,1,2,3,4,5,6,7,10,11,12,13,14,15,16,17,18,19,20,21,22,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,61,62,63,64,65,66,67,68,69,70,71,72)
$values = @(1,2,9,10,11,15,20,21,27,50,100,101,1234)
function Esc([string]$s) { $b = New-Object System.Text.StringBuilder; foreach ($c in $s.ToCharArray()) { if ([int]$c -ge 32 -and [int]$c -lt 127 -and $c -ne '\') { [void]$b.Append($c) } else { [void]$b.Append(('\u{0:X4}' -f [int]$c)) } }; $b.ToString() }
foreach ($part in @('a', 'b')) {
$partCodes = if ($part -eq 'a') { $codes[0..34] } else { $codes[35..($codes.Count - 1)] }
$w = New-Object -ComObject Word.Application
$w.Visible = $false
$w.DisplayAlerts = 0
try {
    $d = $w.Documents.Add()
    $d.Content.LanguageID = 1033
    $d.Paragraphs.Item(1).Range.Text = "list formats"
    $cases = New-Object System.Collections.ArrayList
    foreach ($code in $partCodes) {
        foreach ($v in $values) {
            $t = $d.ListTemplates.Add($false)
            $l = $t.ListLevels.Item(1)
            try { $l.NumberStyle = $code } catch { Write-Host "code $code rejected"; break }
            $l.NumberFormat = "%1."
            try { $l.StartAt = $v } catch { continue }
            $d.Content.InsertParagraphAfter()
            $p = $d.Paragraphs.Item($d.Paragraphs.Count).Range
            $p.Text = "c$code v$v"
            $p = $d.Paragraphs.Item($d.Paragraphs.Count).Range
            $p.ListFormat.ApplyListTemplate($t, $false) | Out-Null
            [void]$cases.Add(@($code, $v, $d.Paragraphs.Count))
        }
        Write-Host "code $code done"
    }
    $lines = foreach ($c in $cases) { "{0}`t{1}`t{2}" -f $c[0], $c[1], (Esc $d.Paragraphs.Item($c[2]).Range.ListFormat.ListString) }
    [System.IO.File]::WriteAllLines("$dir\list-formats-v1-$part.expected.txt", [string[]]$lines)
    $d.SaveAs2("$dir\list-formats-v1-$part.docx", 16)
    $d.SaveAs2("$dir\list-formats-v1-$part.doc", 0)
    $d.SaveAs2("$dir\list-formats-v1-$part.rtf", 6)
    $d.Close(0)
}
finally { $w.Quit() ; [System.Runtime.InteropServices.Marshal]::ReleaseComObject($w) | Out-Null }
}
