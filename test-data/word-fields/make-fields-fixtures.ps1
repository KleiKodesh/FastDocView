# Builds fields-v1.doc and fields-v1.rtf with Word 16 (COM): one paragraph per field case, results as Word computes them.
# Run from this folder: powershell -File make-fields-fixtures.ps1   (the harness reads the saved files; Word is not needed to test)
$ErrorActionPreference = 'Stop'
$dir = $PSScriptRoot
$w = New-Object -ComObject Word.Application
$w.Visible = $false
$w.DisplayAlerts = 0
try {
    $d = $w.Documents.Add()
    $d.Content.LanguageID = 1033
    # document properties through late binding (PowerShell cannot reach Office's property collections directly)
    function SetBuiltIn($name, $value) {
        $props = $script:d.BuiltInDocumentProperties
        $prop = [System.__ComObject].InvokeMember("Item", [System.Reflection.BindingFlags]::GetProperty, $null, $props, @($name))
        [System.__ComObject].InvokeMember("Value", [System.Reflection.BindingFlags]::SetProperty, $null, $prop, @($value)) | Out-Null
    }
    SetBuiltIn "Author" "Moshe Cohen"
    SetBuiltIn "Title" "Fields fixture"
    $custom = $d.CustomDocumentProperties
    [System.__ComObject].InvokeMember("Add", [System.Reflection.BindingFlags]::InvokeMethod, $null, $custom, @("Client", $false, 4, "Acme Ltd")) | Out-Null
    $r = $d.Content
    function Para([string]$text) { $script:r.InsertParagraphAfter(); $p = $script:d.Paragraphs.Item($script:d.Paragraphs.Count).Range; $p.Text = $text; return $p }
    function FieldAt($text, $code, [int]$type = -1) {
        $p = Para $text
        $end = $script:d.Paragraphs.Item($script:d.Paragraphs.Count).Range
        $end.MoveEnd(1, -1) | Out-Null
        $end.Collapse(0)
        return $script:d.Fields.Add($end, $type, $code, $false)
    }
    # 0: the bookmark target
    $p = $d.Paragraphs.Item(1).Range
    $p.Text = "Target text"
    $t = $d.Paragraphs.Item(1).Range; $t.MoveEnd(1, -1) | Out-Null
    $d.Bookmarks.Add("Target", $t) | Out-Null
    FieldAt "cached " "REF Target \h" | Out-Null                       # 1
    FieldAt "page " "PAGE" | Out-Null                                 # 2
    FieldAt "see " "PAGEREF Target \h" | Out-Null                     # 3
    FieldAt "sum " "= 2+3*4" | Out-Null                               # 4
    FieldAt "Figure " "SEQ Figure" | Out-Null                         # 5
    FieldAt "Figure " "SEQ Figure \* roman" | Out-Null                # 6
    $locked = FieldAt "locked " "= 1+1"                               # 7
    $locked.Locked = $true
    FieldAt "author " "AUTHOR" | Out-Null                             # 8
    FieldAt "client " "DOCPROPERTY Client" | Out-Null                 # 9
    $p = Para "box "                                                  # 10
    $end = $d.Paragraphs.Item($d.Paragraphs.Count).Range; $end.MoveEnd(1, -1) | Out-Null; $end.Collapse(0)
    $box = $d.FormFields.Add($end, 71); $box.CheckBox.Value = $true
    $p = Para "drop "                                                 # 11
    $end = $d.Paragraphs.Item($d.Paragraphs.Count).Range; $end.MoveEnd(1, -1) | Out-Null; $end.Collapse(0)
    $drop = $d.FormFields.Add($end, 83); $drop.DropDown.ListEntries.Add("Red") | Out-Null; $drop.DropDown.ListEntries.Add("Blueberry") | Out-Null; $drop.DropDown.Value = 2
    FieldAt "macro " "MACROBUTTON NoMacro Press here" | Out-Null      # 12
    FieldAt "eq " "EQ \f(1,2)" | Out-Null                             # 13
    FieldAt "auto " "AUTONUM" | Out-Null                              # 14
    FieldAt "list " "LISTNUM" | Out-Null                              # 15
    FieldAt "index" "XE ""hiddenentry""" | Out-Null                   # 16
    FieldAt "merge " "MERGEFIELD FirstName" | Out-Null                # 17
    FieldAt "quote " "QUOTE ""quoted text"" \* Upper" | Out-Null      # 18
    $nested = FieldAt "nested " "IF 1 = 1 ""nestedyes"" ""nestedno"""   # 19
    FieldAt "link " "HYPERLINK \l ""Target""" | Out-Null              # 20
    $d.Fields.Update() | Out-Null
    $d.SaveAs2("$dir\fields-v1.doc", 0)
    $d.SaveAs2("$dir\fields-v1.rtf", 6)
    $i = 0
    foreach ($f in $d.Fields) { "{0}: {1} => [{2}]" -f $i, $f.Code.Text, $f.Result.Text; $i++ }
    $d.Close(0)
}
finally { $w.Quit() }
