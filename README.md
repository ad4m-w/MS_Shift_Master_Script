# Powershell file automating functions at MS Shift
Run this in Powershell using:

$ScriptFromGitHub = Invoke-WebRequest https://raw.githubusercontent.com/ad4m-w/powershell_msshift/refs/heads/main/MSShell.ps1
Invoke-Expression $($ScriptFromGitHub.Content)
