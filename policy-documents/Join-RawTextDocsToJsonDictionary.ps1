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

    Add-LoggingMethod 'Info_StartedJoiningRawTexts' -Method {

        $S = $PSStyle.Foreground.BrightWhite + $PSStyle.Bold + $PSStyle.Italic
        $R = $PSStyle.Reset

        Write-Information $($S + 'Started joining raw text from all pdf files...' + $R)
    }

    Add-LoggingMethod 'Info_JoiningRawTextFromFile' -Method {

        param($file)

        $S = $PSStyle.Foreground.Green + $PSStyle.Bold
        $F = $PSStyle.Foreground.White + $PSStyle.BoldOff + $PSStyle.Italic
        $R = $PSStyle.Reset

        Write-Information $($S + " * Joining raw-text from file: $F[ $($file.LoggedDir) ]> $($file.Name)" + $R)
    }

    Add-LoggingMethod 'Warn_TargetFileAllreadyExists' -Method {

        param($file)

        $S = $PSStyle.Foreground.BrightYellow + $PSStyle.Bold
        $F = $PSStyle.Foreground.White + $PSStyle.BoldOff + $PSStyle.Italic
        $R = $PSStyle.Reset

        Write-Warning $($S + " ! Target-file allready exists: $F[ $($file.LoggedDir) ]> $($file.Name)" + $R)
    }

    Add-LoggingMethod 'Warn_JoinRawTextFailed' -Method {

        param($file, $err)

        $S = $PSStyle.Foreground.BrightYellow + $PSStyle.Bold
        $F = $PSStyle.Foreground.White + $PSStyle.BoldOff + $PSStyle.Italic
        $R = $PSStyle.Reset

        Write-Warning $($S + " ! Failed Joining raw-text from file: $F[ $($file.LoggedDir) ]> $($file.Name)" + $R + "`n$err")
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



#region -- Declare: Get-SourceFileInfo --
function Get-SourceFileInfo {

    [CmdletBinding()]
    param(
        [Alias('FullName')]
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [string] $SourcePath
    )

    begin {
        $RootPath = Resolve-Path "$PSScriptRoot\Archive"
    }

    process {
        $File_Name = [System.IO.Path]::GetFileName($SourcePath)

        $Dir_AbsPath = [System.IO.Path]::GetDirectoryName($SourcePath)
        $Dir_RelPath = [System.IO.Path]::GetRelativePath($RootPath, $Dir_AbsPath)

        return [PSCustomObject] @{
            Name       = $File_Name
            Path       = $SourcePath
            LoggedDir  = $Dir_RelPath
            LoggedFile = Join-Path $Dir_RelPath -ChildPath $File_Name
        }
    }
}
#endregion

#region -- Declare: Join-RawTextData --
function Join-RawTextData {

    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [PSCustomObject] $SourceFile
    )

    process {
        $Logging.Info_JoiningRawTextFromFile($SourceFile)

        try {
            $SourceData = [PSCustomObject]@{
                File = $SourceFile.LoggedFile
                Type = ''
                Year = [datetime]::Now.Year
                Date = [datetime]::Now.Date
                Tags = @()
                Text = Get-Content $SourceFile.Path -Raw
            }

            if ($SourceFile.LoggedDir -match '(?<TYPE>.*?)\\(?<YEAR>.*?)\\(?<MONTH>.*?)\\(?<DAY>.*?)(\\(?<REST>.*))?') {
                $SourceData.Type = $Matches.TYPE
                $SourceData.Year = [int] $Matches.YEAR
                $SourceData.Date = [datetime]::new($Matches.YEAR, $Matches.MONTH, $Matches.DAY)
                $SourceData.Tags = @($Matches.TAGS -split '\\')
            }

            return [PSCustomObject] @{
                SourceFile = $SourceFile
                SourceData = $SourceData
            }

        } catch {
            $Logging.Warn_JoinRawTextFailed($SourceFile, $_)
        }
    }
}
#endregion

#region -- Declare: Select-SubSetBy --
function Select-SubSetBy {

    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [PSCustomObject] $SourceFile,

        [Parameter(Mandatory)]
        [int] $Percentage
    )

    begin {
        [double] $Increment = $Percentage / 100.0
        [int] $Previous = 0
        [double] $Index = 0
    }

    process {
        $Index += $Increment

        $Current = [System.Math]::Floor($Index)
        if ($Current -gt $Previous) {
            $Previous = $Current

            Write-Output $SourceFile
        }
    }
}
#endregion

#region -- Declare: Join-ToJsonDictionary --
function Join-ToJsonDictionary {

    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [PSCustomObject] $SourceFile,

        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [PSCustomObject] $SourceData,

        [Parameter(Mandatory)]
        [string] $TargetPath
    )

    begin {
        $Target_Dict = [ordered] @{}
    }

    process {
        $Target_Dict.Add($SourceFile.LoggedFile, $SourceData)
    }

    end {
        ConvertTo-Json $Target_Dict -Depth 20 -EscapeHandling Default |
            Set-Content -LiteralPath $TargetPath -Force
    }
}
#endregion



try {
    Push-Location "$PSScriptRoot\.." # Set to repo root


    $Logging.Info_StartedJoiningRawTexts()

    $AllSourceFilesWithData = @(
        Get-ChildItem -Path "$PSScriptRoot\Archive" -Directory |
            Where-Object Name -In @(
                'kamerstukken',
                'rapporten',
                'publicaties',
                'jaarverslagen',
                'beleidsnotas'
            ) |
            Get-ChildItem -Directory |
            Where-Object {
                $Year = [int] $PSItem.Name

                $Year -ge 2015 -and $Year -lt 2024
            } |
            Get-ChildItem -Filter *.raw.txt -File -Recurse |
            Get-SourceFileInfo |
            Join-RawTextData
    )

    $AllSourceFilesWithData |
        Select-SubSetBy -Percentage 1 |
        Join-ToJsonDictionary -TargetPath "$PSScriptRoot\policy-docset-001.json"

    $AllSourceFilesWithData |
        Select-SubSetBy -Percentage 10 |
        Join-ToJsonDictionary -TargetPath "$PSScriptRoot\policy-docset-010.json"

    $AllSourceFilesWithData |
        Select-SubSetBy -Percentage 100 |
        Join-ToJsonDictionary -TargetPath "$PSScriptRoot\policy-docset-100.json"

} finally {
    Pop-Location
}
