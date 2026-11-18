# Builds test-data/smartart-layouts/: one .docx per built-in SmartArt layout (Word 16, hidden, no alerts), each with a known
# node text set, saved by Word with its drawing part, plus Word's PDF of it (print layout). strip-smartart-drawing.py then
# makes the "-nodrawing" copies the layout engine is tested on. Only the Word instance this script started is quit.
$ErrorActionPreference = 'Stop'
$folder = Join-Path $PSScriptRoot '..\test-data\smartart-layouts'
New-Item -ItemType Directory -Force $folder | Out-Null
$folder = (Resolve-Path $folder).Path
# (Word refuses child nodes in hList1 / list1 through the object model, so those two are flat lists.)
# layout id (urn:microsoft.com/office/officeart/2005/8/layout/<id>) -> file name, node texts (a leading '-' = a child of the previous top node)
$layouts = [ordered]@{
    'default'   = @('block-list',  @('Alpha', 'Beta', 'Gamma', 'Delta', 'Epsilon'))
    'process1'  = @('process',     @('Plan', 'Build', 'Test'))
    'cycle2'    = @('cycle',       @('One', 'Two', 'Three', 'Four', 'Five'))
    'orgChart1' = @('org-chart',   @('Head', '-Left', '-Middle', '-Right'))
    'hierarchy1'= @('hierarchy',   @('Root', '-Branch A', '-Branch B'))
    'pyramid1'  = @('pyramid',     @('Top', 'Middle', 'Base'))
    'vList2'    = @('vertical-list', @('First', 'Second', 'Third'))
    'hList1'    = @('horizontal-list', @('North', 'East', 'South', 'West', 'Centre'))
    'chevron1'  = @('chevron',     @('Step A', 'Step B', 'Step C'))
    'list1'     = @('stacked-list', @('Red', 'Green', 'Blue', 'White'))
}
$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
$word = New-Object -ComObject Word.Application
$mine = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id } | ForEach-Object Id)
try {
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $probe = $word.Documents.Add()
    $byId = @{}
    foreach ($l in $word.SmartArtLayouts) { $byId[($l.Id -replace '^.*/', '')] = $l }
    Write-Output ("layouts: " + $byId.Count)
    foreach ($id in $layouts.Keys) {
        if (-not $byId.ContainsKey($id)) { Write-Output "missing layout $id"; continue }
        $name = $layouts[$id][0]; $texts = $layouts[$id][1]
        $doc = $word.Documents.Add()
        try {
        $doc.ActiveWindow.View.Type = 3
        $shape = $doc.InlineShapes.AddSmartArt($byId[$id], $doc.Paragraphs(1).Range)
        $art = $shape.SmartArt
        while ($art.AllNodes.Count -gt 0) { $art.AllNodes.Item(1).Delete() }
        $top = $null; $child = $null; $childParent = $null
        foreach ($t in $texts) {
            if ($t.StartsWith('-')) {
                # the first child goes below its parent (msoSmartArtNodeBelow = 5), the next ones after the previous child (msoSmartArtNodeAfter = 2)
                if ($child -ne $null -and $childParent -eq $top) { $node = $child.AddNode(2) } else { $node = $top.AddNode(5) }
                $node.TextFrame2.TextRange.Text = $t.Substring(1); $child = $node; $childParent = $top
            }
            else { $top = $art.Nodes.Add(); $top.TextFrame2.TextRange.Text = $t }
        }
        $base = Join-Path $folder ($name + '-v1')
        $doc.SaveAs2($base + '.docx', 16)
        $doc.SaveAs2($base + '.pdf', 17)
        Write-Output "saved $name"
        } catch { Write-Output ("failed $name : " + ("$_" -replace '[^ -~]', '?')) }
        $doc.Close([ref]0)
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc)
    }
}
finally {
    try { $probe.Close([ref]0) } catch { }
    try { $word.Quit([ref]0) } catch { }
    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) } catch { }
    Start-Sleep -Milliseconds 1500
    foreach ($id in $mine) { try { Stop-Process -Id $id -Force -ErrorAction Stop } catch { } }
}
