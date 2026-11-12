# LibreOffice (headless, invisible, private profile, no printers) converts picture-crop-v1.docx to picture-crop-v1.odt:
# LibreOffice writes each crop as fo:clip on the picture's graphic style. Only the soffice process started here is stopped.
$ErrorActionPreference = 'Stop'
$folder = $PSScriptRoot
$profile = Join-Path $env:USERPROFILE ('.fdv-lo-profile-' + [guid]::NewGuid().ToString('N'))
$env:SAL_DISABLE_PRINTERLIST = '1'
$env:SAL_DISABLE_DEFAULTPRINTER = '1'
$soffice = 'C:\Program Files\LibreOffice\program\soffice.exe'
$arguments = @('--headless', '--invisible', '--norestore', '--nologo', ('-env:UserInstallation=file:///' + $profile.Replace('\', '/')),
               '--convert-to', 'odt', '--outdir', ('"' + $folder + '"'), ('"' + (Join-Path $folder 'picture-crop-v1.docx') + '"'))
$process = Start-Process -FilePath $soffice -ArgumentList $arguments -PassThru -WindowStyle Hidden
try { if (-not $process.WaitForExit(120000)) { throw 'soffice timed out' } }
finally {
    if (-not $process.HasExited) { Stop-Process -Id $process.Id -Force }
    Remove-Item -Recurse -Force $profile -ErrorAction SilentlyContinue
}
