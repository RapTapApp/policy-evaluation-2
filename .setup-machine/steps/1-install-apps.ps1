#Requires -Version 7.3
#Requires -RunAsAdministrator

param(
    [switch] $Force
)

$InformationPreference = 'Continue'



# Setup target-apps
Write-Information 'Setup: target-apps...'

$TargetApps = [ordered] @{

    '(Microsoft) VS-code' = 'Microsoft.VisualStudioCode'
    'Python'              = 'Python.Python.3.13'
    'Git'                 = 'Git.Git'
    'Oh my posh'          = 'JanDeDobbeleer.OhMyPosh'
}



# Install target-apps
Write-Information "`nInstalling target-apps..."

$TargetApps.GetEnumerator() | ForEach-Object {

    Write-Information "`n - $($PSItem.Key)"



    if ($Force.IsPresent) {

        # Uninstall target-app
        Write-Verbose ' Uninstalling...' -Verbose

        winget uninstall --id $PSItem.Value --exact --purge --accept-source-agreements --silent --disable-interactivity | Out-Null



        # Install target-app
        Write-Verbose '   Installing...' -Verbose
    }

    winget install --id $PSItem.Value --exact --ignore-security-hash --dependency-source --accept-source-agreements --accept-package-agreements --silent
}



# Done!
Write-Information "`n[DONE]"
