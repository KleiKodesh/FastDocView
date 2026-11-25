param([string]$out = $PSScriptRoot)  # Word 16, never visible: opens css-selectors-v1.mht (hand-written archive) and prints it to css-selectors-v1.pdf (Word's print)
$w = New-Object -ComObject Word.Application
try {
  $w.DisplayAlerts = 0; $w.Visible = $false
  $d = $w.Documents.Open("$out\css-selectors-v1.mht", $false, $true, $false)
  $w.ActiveWindow.View.Type = 3
  $d.ExportAsFixedFormat("$out\css-selectors-v1.pdf", 17)
  $d.Close([ref]0)
} finally { $w.Quit([ref]0) }
