# Builds shapes-3d-bevels-v1.docx (every bevel preset, front camera, threePt rig) and shapes-3d-parts-v1.docx (contours and bevels on
# outlines with holes or several parts), and Word's PDF of each, with a hidden Word (DisplayAlerts 0, never shown, Quit in finally).
$d = $PSScriptRoot
$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
$w = New-Object -ComObject Word.Application
$w.DisplayAlerts = 0; $w.Visible = $false
$mine = @(Get-Process WINWORD | Where-Object { $before -notcontains $_.Id } | ForEach-Object Id)
try {
  function NewDoc() { $doc = $w.Documents.Add(); $doc.Content.Text = '3-D'; return $doc }
  function Add($doc, $type, $x, $y, $wd, $ht, [scriptblock]$setup) {
      $s = $doc.Shapes.AddShape($type, $x, $y, $wd, $ht, $doc.Paragraphs.First.Range)
      $s.RelativeHorizontalPosition = 1; $s.RelativeVerticalPosition = 1
      $s.Left = [single]$x; $s.Top = [single]$y
      $s.Fill.ForeColor.RGB = 0xC07040; $s.Line.Visible = 0
      & $setup $s.ThreeD
  }
  function Finish($doc, $name) {
      $doc.ActiveWindow.View.Type = 3
      $doc.SaveAs2("$d\$name.docx", 16)
      $doc.SaveAs2("$d\$name-docx-word.pdf", 17)
      $doc.Close([ref]0)
  }
  # bevel presets (MsoBevelType 2-13) in reading order: 3 columns, 4 rows, 150 x 110 pt, inset and depth 24 pt
  $doc = NewDoc
  $i = 0
  foreach ($type in 2..13) {
      $x = 50 + ($i % 3) * 180; $y = 70 + [int][math]::Floor($i / 3) * 170
      Add $doc 1 $x $y 150 110 ([scriptblock]::Create("param(`$t) `$t.BevelTopType = $type; `$t.BevelTopInset = 24; `$t.BevelTopDepth = 24"))
      $i++
  }
  Finish $doc 'shapes-3d-bevels-v1'
  # the same presets seen edge on (orthographic, turned 90 degrees about x): the picture's height at each x is the profile's highest point so far
  $doc = NewDoc
  $i = 0
  foreach ($type in 2..13) {
      $x = 50 + ($i % 3) * 180; $y = 70 + [int][math]::Floor($i / 3) * 170
      Add $doc 1 $x $y 150 110 ([scriptblock]::Create("param(`$t) `$t.SetPresetCamera(19); `$t.RotationX = 90; `$t.BevelTopType = $type; `$t.BevelTopInset = 24; `$t.BevelTopDepth = 24"))
      $i++
  }
  Finish $doc 'shapes-3d-bevels-edge-v1'
  # outlines with holes / several parts: donut 18, frame 158, no symbol 19, equal 166; contour turned, then bevels
  $doc = NewDoc
  Add $doc 18 50 70 150 110 { param($t) $t.SetPresetCamera(19); $t.RotationY = 40; $t.Depth = 20; $t.ContourWidth = 3; $t.ContourColor.RGB = 0x202020 }
  Add $doc 158 230 70 150 110 { param($t) $t.SetPresetCamera(19); $t.RotationY = 40; $t.Depth = 20; $t.ContourWidth = 3; $t.ContourColor.RGB = 0x202020 }
  Add $doc 19 410 70 150 110 { param($t) $t.SetPresetCamera(19); $t.RotationX = 30; $t.Depth = 20; $t.ContourWidth = 3; $t.ContourColor.RGB = 0x202020 }
  Add $doc 166 50 240 150 110 { param($t) $t.SetPresetCamera(19); $t.RotationY = 40; $t.Depth = 20; $t.ContourWidth = 3; $t.ContourColor.RGB = 0x202020 }
  Add $doc 18 230 240 150 110 { param($t) $t.BevelTopType = 3; $t.BevelTopInset = 8; $t.BevelTopDepth = 8 }
  Add $doc 158 410 240 150 110 { param($t) $t.BevelTopType = 6; $t.BevelTopInset = 6; $t.BevelTopDepth = 6 }
  Add $doc 166 50 410 150 110 { param($t) $t.BevelTopType = 3; $t.BevelTopInset = 5; $t.BevelTopDepth = 5 }
  Add $doc 18 230 410 150 110 { param($t) $t.BevelTopType = 3; $t.BevelTopInset = 8; $t.BevelTopDepth = 8; $t.ContourWidth = 3; $t.ContourColor.RGB = 0x202020 }
  Finish $doc 'shapes-3d-parts-v1'
} finally {
    $w.Quit([ref]0)
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($w) | Out-Null
    $mine | ForEach-Object { Stop-Process -Id $_ -Force -ErrorAction SilentlyContinue }
}
