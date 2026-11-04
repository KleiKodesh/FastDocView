# Builds fields-lists-v1.doc with Word 16 (COM): a three-level list "%1)" "%2)" "%3)" and REF \n \w \r to its deepest item,
# the same layout as the harness's .docx fixture (WordFieldsFidelityTest.DocxListContexts). Word's own results are printed;
# the harness computes the same codes through the .doc's field context and checks them against these.
# Run from this folder: powershell -File make-fields-lists-fixture.ps1
$ErrorActionPreference = 'Stop'
$dir = $PSScriptRoot
$w = New-Object -ComObject Word.Application
$w.Visible = $false
$w.DisplayAlerts = 0
try {
    $d = $w.Documents.Add()
    $d.Content.LanguageID = 1033
    $template = $d.ListTemplates.Add($true)
    $styles = @(0, 4, 2)          # wdListNumberStyleArabic, LowercaseLetter, LowercaseRoman
    for ($i = 1; $i -le 3; $i++) { $l = $template.ListLevels.Item($i); $l.NumberFormat = "%$i)"; $l.NumberStyle = $styles[$i - 1]; $l.StartAt = 1 }
    $texts = @(@("one", 1), @("two", 1), @("sub a", 2), @("sub b ", 2), @("subsub", 3))
    $d.Paragraphs.Item(1).Range.Text = "before"
    foreach ($t in $texts) {
        $d.Content.InsertParagraphAfter()
        $p = $d.Paragraphs.Item($d.Paragraphs.Count).Range
        $p.Text = $t[0]
        $p = $d.Paragraphs.Item($d.Paragraphs.Count).Range
        $p.ListFormat.ApplyListTemplateWithLevel($template, $true, 0) | Out-Null
        $p.ListFormat.ListLevelNumber = $t[1]
    }
    $deep = $d.Paragraphs.Item($d.Paragraphs.Count).Range; $deep.MoveEnd(1, -1) | Out-Null
    $d.Bookmarks.Add("deep", $deep) | Out-Null
    function FieldAtEnd($paragraph, $code) {
        $end = $script:d.Paragraphs.Item($paragraph).Range; $end.MoveEnd(1, -1) | Out-Null; $end.Collapse(0)
        return $script:d.Fields.Add($end, -1, $code, $false)
    }
    FieldAtEnd 5 "REF deep \r" | Out-Null                    # inside 2)b)
    $d.Content.InsertParagraphAfter()
    $last = $d.Paragraphs.Item($d.Paragraphs.Count).Range
    $last.ListFormat.RemoveNumbers() | Out-Null
    $last.Text = "refs "
    $n = $d.Paragraphs.Count
    FieldAtEnd $n "REF deep \n" | Out-Null
    $e = $d.Paragraphs.Item($n).Range; $e.MoveEnd(1, -1) | Out-Null; $e.Collapse(0); $e.InsertAfter(" | ")
    FieldAtEnd $n "REF deep \w" | Out-Null
    $e = $d.Paragraphs.Item($n).Range; $e.MoveEnd(1, -1) | Out-Null; $e.Collapse(0); $e.InsertAfter(" | ")
    FieldAtEnd $n "REF deep \r" | Out-Null
    $d.Fields.Update() | Out-Null
    $i = 0
    foreach ($p in $d.Paragraphs) { "p{0}: [{1}] [{2}]" -f $i, $p.Range.ListFormat.ListString, $p.Range.Text.TrimEnd("`r"); $i++ }
    $i = 0
    foreach ($f in $d.Fields) { "{0}: {1} => [{2}]" -f $i, $f.Code.Text, $f.Result.Text; $i++ }
    $d.SaveAs2("$dir\fields-lists-v1.doc", 0)
    $d.Close(0)
}
finally { $w.Quit() }
