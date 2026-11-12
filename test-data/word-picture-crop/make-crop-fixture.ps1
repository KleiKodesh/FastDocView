# Word 16 (hidden, no alerts) builds picture-crop-v1: quadrants.png (200 x 100 px = 150 x 75 pt) three times -
# inline cropped on each side (left 15, top 7.5, right 30, bottom 15 pt), inline with a negative crop (left -15, top -7.5 pt:
# Word's fill padding) and floating (square wrap) cropped left 30, bottom 15 pt - saved as .docx, .doc, .rtf and Word's PDF.
# Word is quit in finally; only the Word process this script started is stopped.
$ErrorActionPreference = 'Stop'
$folder = $PSScriptRoot
$base = Join-Path $folder 'picture-crop-v1'
$png = Join-Path $folder 'quadrants.png'
$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
$word = New-Object -ComObject Word.Application
$mine = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id } | ForEach-Object Id)
try {
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $doc = $word.Documents.Add()
    $doc.Content.Text = "cropsides`rcropnegative`rcropfloat`rafter"
    function Inline($index, $l, $t, $r, $b) {
        $range = $doc.Paragraphs($index).Range
        $range.Collapse(0); $range.Move(1, -1) | Out-Null
        $picture = $doc.InlineShapes.AddPicture($png, $false, $true, $range)
        $picture.PictureFormat.CropLeft = [single]$l; $picture.PictureFormat.CropTop = [single]$t
        $picture.PictureFormat.CropRight = [single]$r; $picture.PictureFormat.CropBottom = [single]$b
        return $picture
    }
    $one = Inline 1 15 7.5 30 15
    $two = Inline 2 -15 -7.5 0 0
    $three = Inline 3 30 0 0 15
    $shape = $three.ConvertToShape()
    $shape.WrapFormat.Type = 0
    Write-Host ('sizes ' + $one.Width + 'x' + $one.Height + ' ' + $two.Width + 'x' + $two.Height + ' ' + $shape.Width + 'x' + $shape.Height)
    $doc.SaveAs2($base + '.docx', 16)
    $doc.ExportAsFixedFormat($base + '.pdf', 17)
    $doc.SaveAs2($base + '.doc', 0)
    $doc.SaveAs2($base + '.rtf', 6)
    $doc.Close(0)
}
finally {
    try { $word.Quit(0) } catch { }
    foreach ($id in $mine) { try { Stop-Process -Id $id -Force -ErrorAction Stop } catch { } }
}
