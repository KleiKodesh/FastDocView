# Builds word-xml-v1 in Word 16 (hidden, no alerts) and saves the same document as .docx, Word XML Document
# (Flat OPC, wdFormatFlatXML = 19), Word 2003 XML Document (wdFormatXML = 11) and PDF (the reference).
# Content: heading, bold / italic, numbered and bulleted lists, a footnote and an endnote, a table, a hyperlink,
# a Hebrew right-to-left paragraph and an inline picture. Only the Word process this script started is stopped.
$ErrorActionPreference = 'Stop'
$folder = $PSScriptRoot
$base = Join-Path $folder 'word-xml-v1'
$picture = (Resolve-Path (Join-Path $folder '..\word-picture-crop\quadrants.png')).Path
function Hebrew([int[]]$codes) { -join ($codes | ForEach-Object { [char]$_ }) }
$shalom = Hebrew 0x05E9,0x05DC,0x05D5,0x05DD,0x20,0x05E2,0x05D5,0x05DC,0x05DD   # "peace, world"
$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
$word = New-Object -ComObject Word.Application
$mine = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id } | ForEach-Object Id)
Write-Host ('word pid ' + ($mine -join ','))
try {
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $doc = $word.Documents.Add()
    $s = $word.Selection
    $s.Style = $doc.Styles.Item(-2)                  # Heading 1
    $s.TypeText('Flat XML heading'); $s.TypeParagraph()
    $s.Style = $doc.Styles.Item(-1)                  # Normal
    $s.TypeText('Plain text, ')
    $s.Font.Bold = 1; $s.TypeText('bold'); $s.Font.Bold = 0
    $s.TypeText(' and ')
    $s.Font.Italic = 1; $s.TypeText('italic'); $s.Font.Italic = 0
    $s.TypeText(' with a footnote')
    $doc.Footnotes.Add($s.Range, [Type]::Missing, 'The footnote text.') | Out-Null
    $s.EndKey(6) | Out-Null
    $s.TypeText(' and an endnote')
    $doc.Endnotes.Add($s.Range, [Type]::Missing, 'The endnote text.') | Out-Null
    $s.EndKey(6) | Out-Null
    $s.TypeText('.'); $s.TypeParagraph()
    $s.Range.ListFormat.ApplyNumberDefault() | Out-Null
    $s.TypeText('First numbered'); $s.TypeParagraph()
    $s.TypeText('Second numbered'); $s.TypeParagraph()
    $s.Range.ListFormat.RemoveNumbers() | Out-Null
    $s.Range.ListFormat.ApplyBulletDefault() | Out-Null
    $s.TypeText('A bullet'); $s.TypeParagraph()
    $s.Range.ListFormat.RemoveNumbers() | Out-Null
    $doc.Hyperlinks.Add($s.Range, 'https://example.com/', [Type]::Missing, [Type]::Missing, 'A link') | Out-Null
    $s.EndKey(6) | Out-Null
    $s.TypeParagraph()
    $table = $doc.Tables.Add($s.Range, 2, 2)
    $table.Borders.Enable = 1
    $table.Cell(1, 1).Range.Text = 'A1'; $table.Cell(1, 2).Range.Text = 'B1'
    $table.Cell(2, 1).Range.Text = 'A2'; $table.Cell(2, 2).Range.Text = 'B2'
    $s.EndKey(6) | Out-Null
    $s.ParagraphFormat.ReadingOrder = 0              # wdReadingOrderRtl
    $s.ParagraphFormat.Alignment = 2
    $s.TypeText($shalom); $s.TypeParagraph()
    $s.ParagraphFormat.ReadingOrder = 1
    $s.ParagraphFormat.Alignment = 0
    $s.InlineShapes.AddPicture($picture) | Out-Null
    $s.TypeParagraph()
    $s.TypeText('Last paragraph.')
    $doc.SaveAs2($base + '.docx', 16)
    $doc.SaveAs2($base + '.xml', 19)                 # Word XML Document (Flat OPC)
    $doc.SaveAs2($base + '-2003.xml', 11)            # Word 2003 XML Document
    $doc.ExportAsFixedFormat($base + '.pdf', 17)
    $doc.Close(0)
}
finally {
    try { $word.Quit(0) } catch { }
    foreach ($id in $mine) { try { Stop-Process -Id $id -Force -ErrorAction Stop } catch { } }
}
