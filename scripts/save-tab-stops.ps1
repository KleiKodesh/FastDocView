# Opens test-data/tab-stops/tab-stops-v2.docx in Word 16 (hidden, no alerts) and saves it as PDF, .doc and .rtf; then LibreOffice
# (headless, private profile, no printers) converts the .docx to .odt and the .odt to PDF. Only the processes this script started
# are stopped.
$ErrorActionPreference = 'Stop'
$folder = (Resolve-Path (Join-Path $PSScriptRoot '..\test-data\tab-stops')).Path
$base = Join-Path $folder 'tab-stops-v2'
$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
$word = New-Object -ComObject Word.Application
$mine = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id } | ForEach-Object Id)
try {
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.Options.DoNotPromptForConvert = $true
    $doc = $word.Documents.Open($base + '.docx', $false, $true, $false)
    $doc.SaveAs2($base + '.pdf', 17)
    $doc.SaveAs2($base + '.doc', 0)
    $doc.SaveAs2($base + '.rtf', 6)
    $doc.Close(0)
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc)
}
finally {
    try { $word.Quit(0) } catch { }
    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) } catch { }
    foreach ($id in $mine) { try { Stop-Process -Id $id -Force -ErrorAction Stop } catch { } }
}

$env:SAL_DISABLE_PRINTERLIST = '1'
$env:SAL_DISABLE_DEFAULTPRINTER = '1'
$soffice = 'C:\Program Files\LibreOffice\program\soffice.exe'
$profile = 'file:///' + ((Join-Path $folder '.lo-profile') -replace '\\', '/')
$common = @('--headless', '--invisible', '--norestore', '--nologo', "-env:UserInstallation=$profile")
$p = Start-Process -FilePath $soffice -ArgumentList ($common + @('--convert-to', 'odt', '--outdir', "`"$folder`"", "`"$base.docx`"")) -PassThru -Wait -WindowStyle Hidden
$p = Start-Process -FilePath $soffice -ArgumentList ($common + @('--convert-to', 'pdf', '--outdir', "`"$folder\lo`"", "`"$base.odt`"")) -PassThru -Wait -WindowStyle Hidden
Remove-Item (Join-Path $folder '.lo-profile') -Recurse -Force -ErrorAction SilentlyContinue
Get-ChildItem $folder -Recurse | Select-Object FullName, Length
