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

    Add-LoggingMethod 'Info_StartedExportingRawTexts' -Method {

        $S = $PSStyle.Foreground.BrightWhite + $PSStyle.Bold + $PSStyle.Italic
        $R = $PSStyle.Reset

        Write-Information $($S + 'Started Exporting raw text from all pdf files...' + $R)
    }

    Add-LoggingMethod 'Info_RawTextExportingFromFile' -Method {

        param([string] $dir_relpath, [string] $file)

        $S = $PSStyle.Foreground.Green + $PSStyle.Bold
        $F = $PSStyle.Foreground.White + $PSStyle.BoldOff + $PSStyle.Italic
        $R = $PSStyle.Reset

        Write-Information $($S + " * Raw-text Exporting from file: $F[ $dir_relpath ]> $file" + $R)
    }

    Add-LoggingMethod 'Warn_TargetFileAllreadyExists' -Method {

        param([string] $dir_relpath, [string] $file)

        $S = $PSStyle.Foreground.BrightYellow + $PSStyle.Bold
        $F = $PSStyle.Foreground.White + $PSStyle.BoldOff + $PSStyle.Italic
        $R = $PSStyle.Reset

        Write-Warning $($S + " ! Target-file allready exists: $F[ $dir_relpath ]> $file" + $R)
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



#region -- Declare: Get-TargetInfoFromSourceFile --
function Get-TargetInfoFromSourceFile {

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
        $Dir_AbsPath = [System.IO.Path]::GetDirectoryName($SourcePath)
        $Dir_RelPath = [System.IO.Path]::GetRelativePath($RootPath, $Dir_AbsPath)

        $Source_Name = [System.IO.Path]::GetFileName($SourcePath)
        $Source_Base = [System.IO.Path]::GetFileNameWithoutExtension($SourcePath)
        $Logging.Info_RawTextExportingFromFile($Dir_RelPath, $Source_Name)

        $Target_Name = $Source_Base + '.raw.txt'
        $Target_Path = Join-Path $Dir_AbsPath -ChildPath $Target_Name

        if ($(Test-Path $Target_Path -PathType Leaf)) {
            $Logging.Warn_FileAllreadyExists($Dir_RelPath, $Target_Name)

        } else {
            [PSCustomObject] @{
                SourcePath = $SourcePath
                TargetPath = $Target_Path
            }
        }
    }
}
#endregion

#region -- Declare: Export-RawTextFromSourceFile --
function Export-RawTextFromSourceFile {

    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [string] $SourcePath,

        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [string] $TargetPath
    )

    process {
        $Source_Pdf = [iTextSharp.text.pdf.PdfReader]::new($SourcePath)
        $Source_MaxPages = $Source_Pdf.NumberOfPages

        try {
            $Target_Text = [System.IO.File]::CreateText($TargetPath)

            for ($PageNr = 1; $PageNr -le $Source_MaxPages; $PageNr++) {

                $PageText = [iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($Source_Pdf, $PageNr)

                $Target_Text.WriteLine($PageText)
                $Target_Text.WriteLine()

                $PageNr++
            }

        } finally {
            $Target_Text.Flush()
            $Target_Text.Close()
        }

        return $TargetPath
    }
}
#endregion



try {
    Push-Location "$PSScriptRoot\.." # Set to repo root

    $BatchProc = Initialize-BatchProcess -Size 100 -OnProcess {

        param([int] $BatchNr, [object[]] $BatchedItems)

        git add --all
        git commit -m "Added batch #$BatchNr of policy-docs (total: $($BatchedItems.Count)"
        git push
    }



    Add-Type -Path "$PSScriptRoot\lib\BouncyCastle.Cryptography\lib\net6.0\BouncyCastle.Cryptography.dll"
    Add-Type -Path "$PSScriptRoot\lib\iTextSharp\itextsharp.dll"



    $Logging.Info_StartedExportingRawTexts()
    Get-ChildItem -Filter *.pdf -File -Recurse |
        Get-TargetInfoFromSourceFile |
        Export-RawTextFromSourceFile |
        ForEach-Object {
            $BatchProc.ForEachItem($PSItem)
        }

    $BatchProc.FlushItems()

} finally {
    Pop-Location
}
