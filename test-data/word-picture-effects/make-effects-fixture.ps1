# Builds the picture-effect and web-archive-shape fixtures with a hidden Word (never shown; DisplayAlerts 0; Quit in finally).
#   python make_png.py swatch-v1.png                        -> the colour swatch (240 x 160 px)
#   this script, step 1                                     -> effects-com-v1.docx (Word's own effects through its object model)
#   python make_variants.py effects-com-v1.docx effects-xml-v1.docx  -> XML-only variants added
#   python make_3d.py effects-com-v1.docx effects-3d-v1.docx; this script -Step 3 -> effects-3d-v1.pdf
#   this script -Step 4                                     -> effects-com-v1.doc / .rtf (Word's own .doc / RTF save), effects-float-v1.doc / .rtf (floating), their PDFs
#   this script, step 2                                     -> effects-xml-v1.pdf / .mht, vml-shapes-v1.docx / .pdf / .mht / .mht.pdf
param([int]$Step = 1)
$d = $PSScriptRoot
$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
$w = New-Object -ComObject Word.Application
$w.DisplayAlerts = 0; $w.Visible = $false
$mine = @(Get-Process WINWORD | Where-Object { $before -notcontains $_.Id } | ForEach-Object Id)
try {
    if ($Step -eq 4) {
        $doc = $w.Documents.Open("$d\effects-com-v1.docx", $false, $true)
        $doc.SaveAs2("$d\effects-com-v1.doc", 0)
        $doc.SaveAs2("$d\effects-com-v1.rtf", 6)
        $doc.Close([ref]0)
        # floating pictures (OfficeArt shapes anchored to a paragraph): grey, black and white, washout
        $doc = $w.Documents.Add()
        $doc.Content.Text = 'Floating pictures with effects.'
        $a = $doc.Paragraphs.First.Range
        $types = 2, 3, 4
        for ($i = 0; $i -lt 3; $i++) {
            $s = $doc.Shapes.AddPicture("$d\swatch-v1.png", $false, $true, 72, 100 + 130 * $i, 180, 120, $a)
            $s.WrapFormat.Type = 4
            $s.PictureFormat.ColorType = $types[$i]
        }
        $doc.SaveAs2("$d\effects-float-v1.doc", 0)
        $doc.SaveAs2("$d\effects-float-v1.rtf", 6)
        $doc.Close([ref]0)
        foreach ($name in 'effects-com-v1', 'effects-float-v1') {
            foreach ($ext in 'doc', 'rtf') {
                $doc = $w.Documents.Open("$d\$name.$ext", $false, $true)
                $doc.ActiveWindow.View.Type = 3
                $doc.SaveAs2("$d\$name-$ext.pdf", 17)
                $doc.Close([ref]0)
            }
        }
    } elseif ($Step -eq 3) {
        $doc = $w.Documents.Open("$d\effects-3d-v1.docx", $false, $true)
        $doc.ActiveWindow.View.Type = 3
        $doc.SaveAs2("$d\effects-3d-v1.pdf", 17)
        $doc.Close([ref]0)
    } elseif ($Step -eq 1) {
        $doc = $w.Documents.Add()
        function Add($label) {
            $doc.Content.InsertParagraphAfter(); $doc.Paragraphs.Last.Range.Text = $label
            $doc.Content.InsertParagraphAfter(); $r = $doc.Paragraphs.Last.Range
            return $doc.InlineShapes.AddPicture("$d\swatch-v1.png", $false, $true, $r)
        }
        $p = Add 'none'
        $p = Add 'grayscale'; $p.PictureFormat.ColorType = 2
        $p = Add 'blackwhite'; $p.PictureFormat.ColorType = 3
        $p = Add 'washout'; $p.PictureFormat.ColorType = 4
        $p = Add 'bright70-contrast30'; $p.PictureFormat.Brightness = 0.7; $p.PictureFormat.Contrast = 0.3
        $p = Add 'bright30-contrast80'; $p.PictureFormat.Brightness = 0.3; $p.PictureFormat.Contrast = 0.8
        $p = Add 'transparent-red'; $p.PictureFormat.TransparentBackground = -1; $p.PictureFormat.TransparencyColor = 0x0000FF
        $p = Add 'pencilsketch'; $null = $p.Fill.PictureEffects.Insert(21)
        $p = Add 'blur'; $null = $p.Fill.PictureEffects.Insert(2)
        $p = Add 'saturation0'; $null = $p.Fill.PictureEffects.Insert(24)
        $p = Add 'softedge'; $p.SoftEdge.Radius = 10
        $p = Add 'rotx40'   # no 3-D rotation through the object model: make_variants.py writes it
        $doc.SaveAs2("$d\effects-com-v1.docx", 16)
        $doc.Close([ref]0)
    } else {
        $doc = $w.Documents.Open("$d\effects-xml-v1.docx", $false, $true)
        $doc.ActiveWindow.View.Type = 3
        $doc.SaveAs2("$d\effects-xml-v1.pdf", 17)
        $doc.SaveAs2("$d\effects-xml-v1.mht", 9)
        $doc.Close([ref]0)
        $doc = $w.Documents.Add()
        $doc.Content.Text = 'Shapes in a web archive. First paragraph.'
        $doc.Content.InsertParagraphAfter()
        $a = $doc.Paragraphs.Last.Range
        $s = $doc.Shapes.AddShape(1, 72, 120, 150, 60, $a); $s.Fill.ForeColor.RGB = 0x0080FF; $s.TextFrame.TextRange.Text = 'Rectangle text'
        $s = $doc.Shapes.AddShape(9, 260, 120, 100, 80, $a); $s.Fill.ForeColor.RGB = 0x00C000
        $s = $doc.Shapes.AddTextbox(1, 72, 220, 200, 50, $a); $s.TextFrame.TextRange.Text = 'Text box content'
        $s = $doc.Shapes.AddLine(300, 240, 450, 300, $a); $s.Line.Weight = 3
        $doc.Content.InsertParagraphAfter(); $r = $doc.Paragraphs.Last.Range
        $p = $doc.InlineShapes.AddPicture("$d\swatch-v1.png", $false, $true, $r); $p.PictureFormat.ColorType = 2
        $doc.Content.InsertParagraphAfter(); $doc.Paragraphs.Last.Range.Text = 'Last paragraph.'
        $doc.ActiveWindow.View.Type = 3
        $doc.SaveAs2("$d\vml-shapes-v1.docx", 16)
        $doc.SaveAs2("$d\vml-shapes-v1.pdf", 17)
        $doc.SaveAs2("$d\vml-shapes-v1.mht", 9)
        $doc.Close([ref]0)
        $doc = $w.Documents.Open("$d\vml-shapes-v1.mht", $false, $true)
        $doc.SaveAs2("$d\vml-shapes-v1.mht.pdf", 17)
        $doc.Close([ref]0)
    }
} finally {
    $w.Quit([ref]0)
    Start-Sleep 2
    foreach ($i in $mine) { if (Get-Process -Id $i -ErrorAction SilentlyContinue) { Stop-Process -Id $i -Force } }
}
