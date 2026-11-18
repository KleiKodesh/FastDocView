# Stops what agents leave running when they die or are stopped mid-task: dev servers, Vite,
# Word automation, private-profile LibreOffice, idle MSBuild nodes. Never touches the user's
# own processes: only processes whose command line or executable sits under an agent work
# folder (..\KleiKodeshProject-work, the scratchpad) or that are automation-only instances.
#
#   powershell -NoProfile -File scripts\stop-agent-rigs.ps1          # stop them
#   powershell -NoProfile -File scripts\stop-agent-rigs.ps1 -List    # only list
param([switch]$List)

$workMarkers = @('\KleiKodeshProject-work\', '\AppData\Local\Temp\claude\')

function Test-AgentProcess($p) {
    $text = "$($p.CommandLine) $($p.ExecutablePath)"
    foreach ($m in $workMarkers) { if ($text -like "*$m*") { return $true } }
    # Word started by COM automation (agents never use an interactive Word).
    if ($p.Name -eq 'WINWORD.EXE' -and $p.CommandLine -match '/Automation') { return $true }
    # LibreOffice with a private profile is always an agent run (the user's has none).
    if ($p.Name -like 'soffice*' -and $p.CommandLine -match '-env:UserInstallation') { return $true }
    return $false
}

$names = @('node.exe', 'dotnet.exe', 'FastDocView.DevServer.exe', 'WINWORD.EXE', 'soffice.exe', 'soffice.bin',
           'KitveiHakodeshService.exe')   # never shells: the calling session's own shell may sit there
$targets = Get-CimInstance Win32_Process | Where-Object { $names -contains $_.Name -and (Test-AgentProcess $_) } |
    Where-Object { $_.ProcessId -ne $PID }

foreach ($t in $targets) {
    $line = "{0,6} {1,-28} {2}" -f $t.ProcessId, $t.Name, ($t.CommandLine -replace '\s+', ' ').Substring(0, [Math]::Min(120, ($t.CommandLine -replace '\s+', ' ').Length))
    if ($List) { "would stop: $line" } else {
        try { Stop-Process -Id $t.ProcessId -Force -ErrorAction Stop; "stopped:    $line" } catch { "gone:       $line" }
    }
}
if (-not $targets) { 'no agent processes found' }

# Idle MSBuild worker nodes (node reuse) hold RAM after agent builds; shutting them down is safe.
if (-not $List) { & dotnet build-server shutdown *> $null; 'msbuild build servers shut down' }
