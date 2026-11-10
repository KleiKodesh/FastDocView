# Word 16 re-save of make-legacy-lists.py output + text start positions (pt from page edge; margin 72) in Word 16, print layout
param([string]$folder = $PSScriptRoot)
$src = Join-Path $env:FIXTURE_SOURCE "legacy-lists-src.docx"; $outBase = Join-Path $folder "legacy-lists-v1"
$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | % Id)
$word = New-Object -ComObject Word.Application
$mine = @(Get-Process WINWORD -ErrorAction SilentlyContinue | % Id | ? { $before -notcontains $_ })
function X($doc,$pos){ $doc.Range($pos,$pos).Select(); return $word.Selection.Information(5) }
try {
  $word.Visible = $true; $word.DisplayAlerts = 0
  $doc = $word.Documents.Open($src, $false, $false)
  $doc.Content.ParagraphFormat.ReadingOrder = 1
  if ($outBase) { $doc.SaveAs2($outBase + ".docx", 16); $doc.SaveAs2($outBase + ".doc", 0); $doc.SaveAs2($outBase + ".rtf", 6) }
  $doc.ActiveWindow.View.Type = 3; $doc.Repaginate()
  foreach ($p in $doc.Paragraphs) {
    $r=$p.Range; $t=$r.Text; $s=$r.Start
    $ti=$t.IndexOf("legacy"); $tb=$t.IndexOf("`t")
    "text x={0} | {1} | {2}" -f (X $doc ($s+$ti)), $t.Trim(), $p.Range.ListFormat.ListString
  }
  $doc.Close(0)
} finally { $word.Quit(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null; Start-Sleep 1; foreach ($id in $mine) { Stop-Process -Id $id -Force -ErrorAction SilentlyContinue } }
