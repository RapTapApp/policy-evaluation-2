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

    Add-LoggingMethod 'Info_ExportingRawTextFromFile' -Method {

        param($file)

        $S = $PSStyle.Foreground.Green + $PSStyle.Bold
        $F = $PSStyle.Foreground.White + $PSStyle.BoldOff + $PSStyle.Italic
        $R = $PSStyle.Reset

        Write-Information $($S + " * Exporting raw-text from file: $F[ $($file.LoggedDir) ]> $($file.Name)" + $R)
    }

    Add-LoggingMethod 'Warn_TargetFileAllreadyExists' -Method {

        param($file)

        $S = $PSStyle.Foreground.BrightYellow + $PSStyle.Bold
        $F = $PSStyle.Foreground.White + $PSStyle.BoldOff + $PSStyle.Italic
        $R = $PSStyle.Reset

        Write-Warning $($S + " ! Target-file allready exists: $F[ $($file.LoggedDir) ]> $($file.Name)" + $R)
    }

    Add-LoggingMethod 'Warn_ExportRawTextFailed' -Method {

        param($file, $err)

        $S = $PSStyle.Foreground.BrightYellow + $PSStyle.Bold
        $F = $PSStyle.Foreground.White + $PSStyle.BoldOff + $PSStyle.Italic
        $R = $PSStyle.Reset

        Write-Error $($S + " ! Failed exporting raw-text from file: $F[ $($file.LoggedDir) ]> $($file.Name)" + $R + "`n$err") -ea Continue
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

        function Get-FileInfo([string] $FilePath, [string] $LoggedDir) {
            return [PSCustomObject] @{
                Name      = [System.IO.Path]::GetFileName($FilePath)
                Path      = $FilePath
                LoggedDir = $LoggedDir
            }
        }
    }

    process {
        $Dir_AbsPath = [System.IO.Path]::GetDirectoryName($SourcePath)
        $Dir_RelPath = [System.IO.Path]::GetRelativePath($RootPath, $Dir_AbsPath)

        $Target_Name = [System.IO.Path]::GetFileNameWithoutExtension($SourcePath) + '.raw.txt'
        $TargetPath = Join-Path $Dir_AbsPath -ChildPath $Target_Name

        [PSCustomObject] @{
            SourceFile = Get-FileInfo $SourcePath -LoggedDir $Dir_RelPath
            TargetFile = Get-FileInfo $TargetPath -LoggedDir $Dir_RelPath
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
        [PSCustomObject] $SourceFile,

        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [PSCustomObject] $TargetFile
    )

    begin {
        function Export-RawText($Source_PdfReader, $TargetPath) {

            $Target_TextWriter = [System.IO.File]::CreateText($TargetPath)
            try {

                $PageMax = $Source_PdfReader.NumberOfPages
                for ($PageNr = 1; $PageNr -le $PageMax; $PageNr++) {

                    $PageText = [iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($Source_PdfReader, $PageNr)

                    $Target_TextWriter.WriteLine($PageText)
                    $Target_TextWriter.WriteLine()

                    $PageNr++
                }

            } finally {
                $Target_TextWriter.Flush()
                $Target_TextWriter.Close()
            }
        }
    }

    process {
        if ($(Test-Path $TargetFile.Path -PathType Leaf)) {
            # $Logging.Warn_TargetFileAllreadyExists($TargetFile)

        } else {
            $Logging.Info_ExportingRawTextFromFile($SourceFile)

            try {
                $Source_PdfReader = [iTextSharp.text.pdf.PdfReader]::new($SourcePath)

                Export-RawText $Source_PdfReader -TargetPath $TargetFile.Path

                return $TargetPath
            } catch {
                $Logging.Warn_ExportRawTextFailed($TargetFile, $_)

                if ($(Test-Path $TargetFile.Path -PathType Leaf)) {
                    Remove-Item $TargetFile.Path -Force
                }
            }
        }
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
