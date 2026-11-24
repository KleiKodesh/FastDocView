# Builds shapes-3d-bevels-oval-v1.docx (every bevel preset on an oval: a curved outline of many edges, front camera, threePt rig)
# and Word's PDF of it, with a hidden Word (DisplayAlerts 0, never shown, Quit in finally). Sampled by sample_bevels_oval.py.
$d = $PSScriptRoot
$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
$w = New-Object -ComObject Word.Application
$w.DisplayAlerts = 0; $w.Visible = $false
$mine = @(Get-Process WINWORD | Where-Object { $before -notcontains $_.Id } | ForEach-Object Id)
try {
  $doc = $w.Documents.Add(); $doc.Content.Text = '3-D'
  $i = 0
  foreach ($type in 2..13) {
      $x = 50 + ($i % 3) * 180; $y = 70 + [int][math]::Floor($i / 3) * 170
      $s = $doc.Shapes.AddShape(9, $x, $y, 150, 110, $doc.Paragraphs.First.Range)
      $s.RelativeHorizontalPosition = 1; $s.RelativeVerticalPosition = 1
      $s.Left = [single]$x; $s.Top = [single]$y
      $s.Fill.ForeColor.RGB = 0xC07040; $s.Line.Visible = 0
      $s.ThreeD.BevelTopType = $type; $s.ThreeD.BevelTopInset = 24; $s.ThreeD.BevelTopDepth = 24
      $i++
  }
  $doc.ActiveWindow.View.Type = 3
  $doc.SaveAs2("$d\shapes-3d-bevels-oval-v1.docx", 16)
  $doc.SaveAs2("$d\shapes-3d-bevels-oval-v1-docx-word.pdf", 17)
  $doc.Close([ref]0)
} finally {
    $w.Quit([ref]0)
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($w) | Out-Null
    $mine | ForEach-Object { Stop-Process -Id $_ -Force -ErrorAction SilentlyContinue }
}
