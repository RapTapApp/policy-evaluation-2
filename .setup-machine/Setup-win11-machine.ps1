#Requires -Version 5.1
#Requires -RunAsAdministrator

param()

$InformationPreference = 'Continue'

Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Force

New-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\FileSystem' -Name 'LongPathsEnabled' -Value 1 -PropertyType DWORD -Force



Set-Location $PSScriptRoot

# Executing step: 0-bootstrap-setup.ps1
Write-Information "`nExecuting step: 0-bootstrap-setup.ps1 ..."

& cmd.exe /s /c 'PowerShell.exe -NoProfile -NonInteractive -File .\steps\0-bootstrap-setup.ps1 -ExecutionPolicy Unrestricted'



# Executing step: 1-install-apps.ps1
Write-Information "`nExecuting step: 1-install-apps.ps1 ..."

& cmd.exe /s /c 'pwsh.exe -NoProfile -NonInteractive -File .\steps\1-install-apps.ps1 -ExecutionPolicy Unrestricted'



# Executing step: 2-setup-apps.ps1
Write-Information "`nExecuting step: 2-setup-apps.ps1 ..."

& cmd.exe /s /c 'pwsh.exe -NoProfile -NonInteractive -File .\steps\2-setup-apps.ps1 -ExecutionPolicy Unrestricted'
