using namespace Microsoft.PowerShell.Commands
using namespace System.Net


$InformationPreference = 'Continue'
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'Ignore'

$PSDefaultParameterValues = @{
    'Write-Verbose:Verbose' = $true

    'Add-Member:MemberType' = 'ScriptMethod'
    'Add-Member:PassThru'   = $true
}

$PSStyle.OutputRendering = 'Ansi'



#region -- Declare: $Logging --
$global:Logging = [PSCustomObject]::new()
& {
    function Add-LoggingMethod([string] $Name, [scriptblock] $Method) {

        $global:Logging = $global:Logging | Add-Member -Name $Name -Value $Method
    }



    Add-LoggingMethod 'Info_BatchProcessInvoked' -Method {

        param([int] $BatchNr, [int] $BatchedItems)

        $S = $PSStyle.Foreground.BrightBlue + $PSStyle.Bold
        $F = $PSStyle.Foreground.White + $PSStyle.BoldOff + $PSStyle.Italic
        $R = $PSStyle.Reset

        Write-Information $($S + " ! Batch process invoked: $F [BatchNr: $BatchNr; BatchedItems: $BatchedItems]..." + $R)
    }

    Add-LoggingMethod 'Info_StartedCommittingRawTexts' -Method {

        $S = $PSStyle.Foreground.BrightWhite + $PSStyle.Bold + $PSStyle.Italic
        $R = $PSStyle.Reset

        Write-Information $($S + 'Started Committing raw text from all pdf files...' + $R)
    }

}
#endregion

#region -- Declare: Initialize-BatchProcess --
function Initialize-BatchProcess ([int] $Size = 30, [scriptblock] $OnProcess) {

    $Batch = [PSCustomObject] @{
        Size      = $Size
        OnProcess = $OnProcess

        Number    = 0
        Items     = [System.Collections.Generic.List[PSCustomObject]]::new($Size)
    }

    $Batch = $Batch | Add-Member -Name 'ForEachItem' -Value {

        param([PSCustomObject] $Item)

        $this.Items.Add($Item)

        if ($this.Items.Count -ge $this.Size) {
            $this.FlushItems()
        }
    }

    $Batch = $Batch | Add-Member -Name 'FlushItems' -Value {

        if ($this.Items.Count -ge 0) {

            $this.Number++

            $global:Logging.Info_BatchProcessInvoked($this.Number, $this.Items.Count)
            $this.OnProcess.Invoke($this.Number, $this.Items)

            $this.Items.Clear()
        }
    }

    return $Batch
}
#endregion



try {
    Push-Location "$PSScriptRoot\.." # Set to repo root

    $BatchProc = Initialize-BatchProcess -Size 300 -OnProcess {

        param([int] $BatchNr, [object[]] $BatchedItems)

        foreach ($File in $BatchedItems) {
            git add $File
        }
        git commit -m "Added batch #$BatchNr of raw text extracted policy-docs (total: $($BatchedItems.Count)"
        git push
    }



    $Logging.Info_StartedCommittingRawTexts()
    $SourceFiles = @(
        Get-ChildItem -Path "$PSScriptRoot\Archive" -Filter *.raw.txt -File -Recurse |
            ForEach-Object {
                "$(Resolve-Path $PSItem.FullName -Relative)".Substring(2)
            }
    )

    $StartAt = $SourceFiles.IndexOf('policy-documents\Archive\kamerstukken\2010\05\21\bijlage-evaluatie-fol-verdrag-kamerbrief-inzake-evaluatie-fol-verdrag\133-889dwhak-bijlage.raw.txt')

    $SourceFiles | Select-Object -Skip $StartAt | ForEach-Object {
        $BatchProc.ForEachItem($PSItem)
    }

    $BatchProc.FlushItems()

} finally {
    Pop-Location
}
