# Writes list-formats-words-v1.expected.txt with Word 16 (COM): for each WdListNumberStyle (= MSONFC code) below, one "%1" list
# of 129 paragraphs (values 1..129) and then large start values, each with Word's own ListFormat.ListString
# ("code<TAB>value<TAB>label", non-ASCII as \uXXXX; start values Word refuses are left out). ListNumberFormatsFidelityTest
# compares WordListNumberScripts with it. Quits Word in finally and stops only the WINWORD process it started.
# Run from this folder: powershell -File make-list-formats-words.ps1
$ErrorActionPreference = 'Stop'
$out = Join-Path $PSScriptRoot 'list-formats-words-v1.expected.txt'
$codes = @(11,12,13,16,20,21,24,25,33,34,35,37,38,39,41,42,43,52,55,56,60,61,65,66,67,68)
$big = @(130,199,200,999,1000,1001,1010,1100,2000,9999,10000,10001,10010,10100,11000,12345,20000,100000,100001,1000000,32767)
function Esc([string]$s) { $b = New-Object System.Text.StringBuilder; foreach ($c in $s.ToCharArray()) { if ([int]$c -ge 32 -and [int]$c -lt 127 -and $c -ne '\') { [void]$b.Append($c) } else { [void]$b.Append(('\u{0:X4}' -f [int]$c)) } }; $b.ToString() }
$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
$w = New-Object -ComObject Word.Application
$mine = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id | Where-Object { $before -notcontains $_ })
$w.Visible = $false; $w.DisplayAlerts = 0
$lines = New-Object System.Collections.ArrayList
try {
  foreach ($code in $codes) {
    $d = $w.Documents.Add()
    $t = $d.ListTemplates.Add($false); $l = $t.ListLevels.Item(1); $l.NumberStyle = $code; $l.NumberFormat = "%1"; $l.StartAt = 1
    $d.Content.Text = (1..129 | ForEach-Object { "x$_" }) -join "`r"
    $d.Content.ListFormat.ApplyListTemplate($t, $false) | Out-Null
    for ($i = 1; $i -le $d.Paragraphs.Count; $i++) { [void]$lines.Add("$code`t$i`t" + (Esc ([string]$d.Paragraphs.Item($i).Range.ListFormat.ListString))) }
    foreach ($v in $big) {
      $t2 = $d.ListTemplates.Add($false); $l2 = $t2.ListLevels.Item(1); $l2.NumberStyle = $code; $l2.NumberFormat = "%1"
      try { $l2.StartAt = $v } catch { continue }
      $p = $d.Paragraphs.Item(1).Range; $p.ListFormat.ApplyListTemplate($t2, $false) | Out-Null
      [void]$lines.Add("$code`t$v`t" + (Esc ([string]$d.Paragraphs.Item(1).Range.ListFormat.ListString)))
    }
    $d.Close(0); "code $code done"
  }
  [System.IO.File]::WriteAllLines($out, [string[]]$lines)
}
finally {
  try { $w.Quit() } catch {}
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($w) | Out-Null
  Start-Sleep -Seconds 2
  foreach ($id in $mine) { Stop-Process -Id $id -ErrorAction SilentlyContinue }
}
