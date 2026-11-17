# Builds shapes-3d-v1.docx / .doc / .rtf and Word's PDF of each with a hidden Word (DisplayAlerts 0, never shown, Quit in finally).
# Shapes with 3-D formats set through Word's object model (ThreeDFormat): legacy oblique / perspective extrusions, modern cameras
# with depth, bevels, materials and light rigs. Word writes a:scene3d / a:sp3d in .docx and the c3D* properties in .doc / RTF
# (with f3D off: Word prints those flat). -Step 2 prints shapes-3d-on-v1.rtf (make_rtf_3d.py: f3D on) to its PDF.
param([int]$Step = 1)
$d = $PSScriptRoot
$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
$w = New-Object -ComObject Word.Application
$w.DisplayAlerts = 0; $w.Visible = $false
$mine = @(Get-Process WINWORD | Where-Object { $before -notcontains $_.Id } | ForEach-Object Id)
try {
  if ($Step -eq 2) {
    $doc = $w.Documents.Open("$d\shapes-3d-on-v1.rtf", $false, $true)
    $doc.ActiveWindow.View.Type = 3
    $doc.SaveAs2("$d\shapes-3d-on-v1-rtf-word.pdf", 17)
    $doc.Close([ref]0)
  } else {
    $doc = $w.Documents.Add()
    $doc.Content.Text = '3-D shapes'
    $a = $doc.Paragraphs.First.Range
    # (type, left, top, width, height, setup) - left / top in points on page 1 and 2
    $n = 0
    function Add($type, $x, $y, $wd, $ht, [scriptblock]$setup, $page = 0) {
        $s = $doc.Shapes.AddShape($type, $x, $y, $wd, $ht, $a)
        $s.RelativeHorizontalPosition = 1; $s.RelativeVerticalPosition = 1   # page
        $s.Left = $x; $s.Top = $y
        $s.Fill.ForeColor.RGB = 0xC07040; $s.Line.Visible = 0
        & $setup $s.ThreeD
        $script:n++
        $s.Name = "s3d$script:n"
    }
    Add 1 60 80 120 80 { param($t) $t.SetThreeDFormat(1) }                          # legacy oblique (3-D style 1)
    Add 1 250 80 120 80 { param($t) $t.SetPresetCamera(14); $t.Depth = 36 }         # legacy perspective front
    Add 1 430 80 120 80 { param($t) $t.SetPresetCamera(11); $t.Depth = 36 }         # legacy perspective top
    Add 1 60 240 120 80 { param($t) $t.SetPresetCamera(25); $t.Depth = 20 }         # isometric left down
    Add 9 250 240 120 80 { param($t) $t.SetPresetCamera(48); $t.RotationY = 30; $t.Depth = 20 }   # perspective, turned
    Add 1 430 240 120 80 { param($t) $t.SetPresetCamera(19); $t.RotationX = 25; $t.RotationY = 35; $t.Depth = 30; $t.ExtrusionColor.RGB = 0x30A030 }
    Add 1 60 400 120 80 { param($t) $t.BevelTopType = 3; $t.BevelTopInset = 8; $t.BevelTopDepth = 8 }           # circle bevel only
    Add 9 250 400 120 80 { param($t) $t.BevelTopType = 4; $t.BevelTopInset = 10; $t.BevelTopDepth = 6 }          # slope bevel on an oval
    Add 1 430 400 120 80 { param($t) $t.BevelTopType = 6; $t.BevelTopInset = 6; $t.BevelTopDepth = 6; $t.PresetMaterial = 3 }   # angle, metal
    Add 5 60 560 120 80 { param($t) $t.SetPresetCamera(48); $t.RotationX = 20; $t.Depth = 24; $t.BevelTopType = 3; $t.BevelTopInset = 6; $t.BevelTopDepth = 6 }
    Add 7 250 560 120 80 { param($t) $t.SetPresetCamera(41); $t.Depth = 24; $t.PresetLighting = 14 }           # oblique top, balanced rig
    Add 1 430 560 120 80 { param($t) $t.SetPresetCamera(19); $t.RotationY = 50; $t.Depth = 30; $t.PresetMaterial = 14; $t.ContourWidth = 2; $t.ContourColor.RGB = 0x202020 }
    $doc.ActiveWindow.View.Type = 3
    $doc.SaveAs2("$d\shapes-3d-v1.docx", 16)
    $doc.SaveAs2("$d\shapes-3d-v1-docx-word.pdf", 17)
    $doc.SaveAs2("$d\shapes-3d-v1.doc", 0)
    $doc.SaveAs2("$d\shapes-3d-v1.rtf", 6)
    $doc.Close([ref]0)
    foreach ($ext in 'doc', 'rtf') {
        $doc = $w.Documents.Open("$d\shapes-3d-v1.$ext", $false, $true)
        $doc.ActiveWindow.View.Type = 3
        $doc.SaveAs2("$d\shapes-3d-v1-$ext-word.pdf", 17)
        $doc.Close([ref]0)
    }
  }
} finally {
    $w.Quit([ref]0)
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($w) | Out-Null
    $mine | ForEach-Object { Stop-Process -Id $_ -Force -ErrorAction SilentlyContinue }
}
