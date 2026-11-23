# Builds test-data/smartart-layouts/: one .docx per built-in SmartArt layout (Word 16, hidden, no alerts), each with a known
# node text set, saved by Word with its drawing part, plus Word's PDF of it (print layout). strip-smartart-drawing.py then
# makes the "-nodrawing" copies the layout engine is tested on. Only the Word instance this script started is quit.
# -Fixtures <json>: build only the entries of that file instead (scripts/make-smartart-hebrew-fixtures.py writes Hebrew ones:
# id, name, texts, optional minorLatin / minorComplex / minorEastAsian = the theme's minor Latin / complex-script / East Asian
# typeface; make-smartart-font-fixtures.py adds bold / italic (every node's text) and complexFont (the runs' a:cs typeface)).
param([string]$Fixtures)
$ErrorActionPreference = 'Stop'
$folder = Join-Path $PSScriptRoot '..\test-data\smartart-layouts'
New-Item -ItemType Directory -Force $folder | Out-Null
$folder = (Resolve-Path $folder).Path
# (Word refuses child nodes in hList1 / list1 through the object model, so those two are flat lists.)
# layout id (urn:microsoft.com/office/officeart/2005/8/layout/<id>), file name, node texts (a leading '-' = a child of the previous top node)
$layouts = @(
    @('default',    'block-list',  @('Alpha', 'Beta', 'Gamma', 'Delta', 'Epsilon')),
    @('process1',   'process',     @('Plan', 'Build', 'Test')),
    @('cycle2',     'cycle',       @('One', 'Two', 'Three', 'Four', 'Five')),
    @('orgChart1',  'org-chart',   @('Head', '-Left', '-Middle', '-Right')),
    @('hierarchy1', 'hierarchy',   @('Root', '-Branch A', '-Branch B')),
    @('pyramid1',   'pyramid',     @('Top', 'Middle', 'Base')),
    @('pyramid1',   'pyramid-4-levels', @('Apex', 'Second level', 'Third', 'Foundation')),
    @('vList2',     'vertical-list', @('First', 'Second', 'Third')),
    @('hList1',     'horizontal-list', @('North', 'East', 'South', 'West', 'Centre')),
    @('chevron1',   'chevron',     @('Step A', 'Step B', 'Step C')),
    @('list1',      'stacked-list', @('Red', 'Green', 'Blue', 'White'))
)
$themeFonts = @{}
$runFormats = @{}
if ($Fixtures) {
    $layouts = @()
    foreach ($f in (Get-Content -Raw $Fixtures | ConvertFrom-Json)) {
        $layouts += ,@($f.id, $f.name, @($f.texts))
        $themeFonts[$f.name] = @($f.minorLatin, $f.minorComplex, $f.minorEastAsian)
        $runFormats[$f.name] = @($f.bold, $f.italic, $f.complexFont)
    }
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
    foreach ($entry in $layouts) {
        $id = $entry[0]; $name = $entry[1]; $texts = $entry[2]
        if (-not $byId.ContainsKey($id)) { Write-Output "missing layout $id"; continue }
        $doc = $word.Documents.Add()
        try {
        $doc.ActiveWindow.View.Type = 3
        if ($themeFonts.ContainsKey($name)) {
            # MsoFontLanguageIndex: msoThemeLatin = 1, msoThemeComplexScript = 2, msoThemeEastAsian = 3
            $minor = $doc.DocumentTheme.ThemeFontScheme.MinorFont
            if ($themeFonts[$name][0]) { $minor.Item(1).Name = $themeFonts[$name][0] }
            if ($themeFonts[$name][1]) { $minor.Item(2).Name = $themeFonts[$name][1] }
            if ($themeFonts[$name][2]) { $minor.Item(3).Name = $themeFonts[$name][2] }
        }
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
        if ($runFormats.ContainsKey($name)) {
            foreach ($node in $art.AllNodes) {
                $font = $node.TextFrame2.TextRange.Font
                if ($runFormats[$name][0]) { $font.Bold = -1 }
                if ($runFormats[$name][1]) { $font.Italic = -1 }
                if ($runFormats[$name][2]) { $font.NameComplexScript = $runFormats[$name][2] }
            }
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
