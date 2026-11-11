# Makes test-data/word-charts-binary/word-charts-binary-v1.{docx,doc,rtf,pdf} with Word 16 through COM:
# one chart per type (column, bar, line, pie, area, scatter, 3-D column) with ASCII data and titles,
# saved as .docx, then as .doc (Word 97-2003) and .rtf, plus Word's PDF for position comparisons.
# Quits Word in finally and stops only the Word/Excel processes this script started.
param([string]$OutDirectory = (Join-Path $PSScriptRoot '..\test-data\word-charts-binary'))
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force $OutDirectory | Out-Null
$OutDirectory = (Resolve-Path $OutDirectory).Path
$base = Join-Path $OutDirectory 'word-charts-binary-v1'
$before = @(Get-Process WINWORD, EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)
$charts = @(
    @{ Type = 51; Title = 'Column Chart'; Note = 'ColumnChart paragraph' },
    @{ Type = 57; Title = 'Bar Chart'; Note = 'BarChart paragraph' },
    @{ Type = 4; Title = 'Line Chart'; Note = 'LineChart paragraph' },
    @{ Type = 5; Title = 'Pie Chart'; Note = 'PieChart paragraph' },
    @{ Type = 1; Title = 'Area Chart'; Note = 'AreaChart paragraph' },
    @{ Type = -4169; Title = 'Scatter Chart'; Note = 'ScatterChart paragraph' },
    @{ Type = 54; Title = 'Depth Column Chart'; Note = 'DepthChart paragraph' }
)
$word = New-Object -ComObject Word.Application
try {
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $doc = $word.Documents.Add()
    foreach ($c in $charts) {
        $range = $doc.Content
        $range.InsertParagraphAfter()
        $range = $doc.Paragraphs.Last.Range
        $range.Text = $c.Note
        $range.InsertParagraphAfter()
        $at = $doc.Paragraphs.Last.Range
        $shape = $doc.InlineShapes.AddChart2(-1, $c.Type, $at)
        $chart = $shape.Chart
        $rows = @(@('', 'North', 'South', 'West'), @('Q1', 4.3, 2.4, 2), @('Q2', 2.5, 4.4, 2), @('Q3', 3.5, 1.8, 3), @('Q4', 4.5, 2.8, 5))
        for ($attempt = 0; $attempt -lt 20; $attempt++) {
            try {
                $chart.ChartData.Activate()
                $book = $chart.ChartData.Workbook
                $sheet = $book.Worksheets.Item(1)
                for ($r = 0; $r -lt 5; $r++) { for ($k = 0; $k -lt 4; $k++) { $cell = [string][char](65 + $k) + ($r + 1); $sheet.Range($cell).Formula = [string]$rows[$r][$k] } }
                $chart.SetSourceData("'" + $sheet.Name + "'!`$A`$1:`$D`$5")
                try { $book.Application.DisplayAlerts = $false; $book.Close($true) } catch { }
                break
            } catch { Start-Sleep -Milliseconds 500 }
        }
        Start-Sleep -Milliseconds 500
        $chart.HasTitle = $true
        $chart.ChartTitle.Text = $c.Title
        $doc.Content.InsertParagraphAfter()
    }
    $doc.SaveAs2($base + '.docx', 16)
    $doc.SaveAs2($base + '.doc', 0)
    $doc.SaveAs2($base + '.rtf', 6)
    $doc.SaveAs2($base + '.pdf', 17)
    $doc.Close(0)
}
catch { Write-Host ($_.Exception.Message); Write-Host ($_.InvocationInfo.ScriptLineNumber); throw }
finally {
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    Get-Process WINWORD, EXCEL -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id } | ForEach-Object { try { Stop-Process -Id $_.Id -Force -ErrorAction Stop } catch { } }
}
Get-ChildItem $OutDirectory | Select-Object Name, Length
