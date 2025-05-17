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

    Add-LoggingMethod 'Info_StartedDownloadingAllFiles' -Method {

        $S = $PSStyle.Foreground.BrightWhite + $PSStyle.Bold + $PSStyle.Italic
        $R = $PSStyle.Reset

        Write-Information $($S + 'Started downloading all files...' + $R)
    }

    Add-LoggingMethod 'Info_GotResultsFromSearchPage' -Method {

        param([int] $CurrentPageNr, [int] $MaximumPageNr)

        $S = $PSStyle.Foreground.White + $PSStyle.Italic
        $R = $PSStyle.Reset

        Write-Information $("`n$S - Search[$MaximumPageNr] => Got results from page $CurrentPageNr" + $R)
    }

    Add-LoggingMethod 'Info_FileDownloaded' -Method {

        param([string] $dir_relpath, [string] $file)

        $S = $PSStyle.Foreground.Green + $PSStyle.Bold
        $F = $PSStyle.Foreground.White + $PSStyle.BoldOff + $PSStyle.Italic
        $R = $PSStyle.Reset

        Write-Information $($S + " * File downloaded: $F[ $dir_relpath ]> $file" + $R)
    }

    Add-LoggingMethod 'Warn_FileAllreadyExists' -Method {

        param([string] $dir_relpath, [string] $file)

        $S = $PSStyle.Foreground.BrightYellow + $PSStyle.Bold
        $F = $PSStyle.Foreground.White + $PSStyle.BoldOff + $PSStyle.Italic
        $R = $PSStyle.Reset

        Write-Warning $($S + " ! File allready exists: $F[ $dir_relpath ]> $file" + $R)
    }

    Add-LoggingMethod 'Warn_ConnectionFailed' -Method {

        $S = $PSStyle.Foreground.BrightYellow + $PSStyle.Bold
        $F = $PSStyle.Foreground.White + $PSStyle.BoldOff + $PSStyle.Italic
        $R = $PSStyle.Reset

        Write-Warning $($S + " ! Connection error: $F retrying after sleep..." + $R)
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



#region -- Declare: Initialize-Search --
function Initialize-Search ([int] $StartFromPageNr = 1) {

    $Search = [PSCustomObject] @{

        CurrentPageNr = $StartFromPageNr
        MaximumPageNr = -1

        Headers       = @{
            'Accept'             = '*/*'
            'Accept-Encoding'    = 'gzip, deflate, br, zstd'
            'Accept-Language'    = 'en-GB,en;q=0.9,en-US;q=0.8,nl;q=0.7'
            'Origin'             = 'https://rijksoverheid.sitearchief.nl'
            'Referer'            = 'https://rijksoverheid.sitearchief.nl/'
            'Sec-Fetch-Dest'     = 'empty'
            'Sec-Fetch-Mode'     = 'cors'
            'Sec-Fetch-Site'     = 'same-origin'
            'X-Requested-With'   = 'XMLHttpRequest'
            'sec-ch-ua'          = "`"Microsoft Edge`";v=`"135`", `"Not-A.Brand`";v=`"8`", `"Chromium`";v=`"135`""
            'sec-ch-ua-mobile'   = '?0'
            'sec-ch-ua-platform' = "`"Windows`""
        }

        BodyFormat    = @(
            'type=advsearch',
            'search-words-and=',
            'search-words-or=Cara%C3%AFbisch+Caribisch+Carabisch',
            'search-words-exact=',
            'search-words-except=',
            'search-site=all',
            'search-site-url=http%3A%2F%2F',
            'search-fields=title%2Ccontent',
            'search-types=pdf%2Cword%2Cexcel',
            'search-date-from=',
            'search-date-to=',
            'page={0}'
        ) -join '&'
    }



    $Search = $Search | Add-Member -MemberType ScriptProperty -Name 'HasFinishedPages' -Value {

        return ($this.MaximumPageNr -gt 0) -and ($this.CurrentPageNr -gt $this.MaximumPageNr)
    }



    $Search = $Search | Add-Member -Name 'GetResultsFromNextPage' -Value {

        if ($this.HasFinishedPages) {
            throw [System.InvalidOperationException] $('The search has finished getting the results from all pages. Current: {0}; Maximum: {1}' -f $this.CurrentPageNr, $this.MaximumPageNr)
        }

        $WebSession = [WebRequestSession]::new()
        $WebSession.UserAgent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/135.0.0.0 Safari/537.36 Edg/135.0.0.0'

        $SearchArgs = @{
            Uri         = 'https://rijksoverheid.sitearchief.nl/search.php'
            Method      = 'POST'
            ContentType = 'application/x-www-form-urlencoded; charset=UTF-8'
            WebSession  = $WebSession
            Headers     = $this.Headers
            Body        = $this.BodyFormat -f $this.CurrentPageNr
        }

        $Response = Invoke-WebRequest @SearchArgs
        if ($Response.StatusCode -ne 200) {
            throw [System.InvalidOperationException] $("Invalid response {0} - {1}`n for page nr: {2}`n from uri: {3}" -f $Response.StatusCode, $Response.StatusDescription, $this.CurrentPageNr, $SearchArgs.Uri)
        }

        if ($this.MaximumPageNr -lt 0) {
            if (($Response.Links[-1].id -eq 'lnkLastPage') -and ($Response.Links[-1].href -match 'javascript:archiefweb\.switch_to_page\((?<MAX_PAGE>\d+)\)')) {
                $this.MaximumPageNr = [int] $Matches.MAX_PAGE
            }
        }

        $global:Logging.Info_GotResultsFromSearchPage($this.CurrentPageNr, $this.MaximumPageNr)
        $this.CurrentPageNr++

        return $Response.Links
    }

    return $Search
}
#endregion

#region -- Declare: Get-TargetInfoFromSourceUri --
function Get-TargetInfoFromSourceUri {

    [CmdletBinding()]
    param(
        [Alias('href')]
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [String] $SourceUri
    )

    begin {
        function Limit-Path([string] $Path, [int] $MaxLength = 120) {

            if (-not $Path -or $Path.Length -le $MaxLength) {
                return $Path

            } else {
                $ChopPosition = $MaxLength - 9 # ~ char + 8 Hex chars => hashcode of chopped part

                $Uncut = $Path.Substring(0, $ChopPosition)
                $Chopped = $Path.Substring($ChopPosition)

                return $('{0}-{1:X8}' -f $Uncut, $Chopped.GetHashCode())
            }
        }

        $repo_root = "$(Get-Location)"
    }

    process {
        if ($SourceUri -notmatch '^http(s)?:/') {
            return
        }

        $SourceUri = $SourceUri -replace '[\+]+', '-'
        $SourceUri = $SourceUri -replace '[\-]+', '-'

        if ($SourceUri -ne $PSBoundParameters.'SourceUri') {
            return
        }

        if ($SourceUri -match 'https://www\.rijksoverheid\.nl/binaries/rijksoverheid/documenten/(?<DIR>.*)/(?<File>.*)$') {
            $parsed = @{
                dir  = $Matches.DIR
                file = $Matches.FILE
            }

            $changed = @{}
            foreach ($item in $parsed.GetEnumerator()) {
                $value = [uri]::UnescapeDataString($item.Value)
                $value = $value -replace 'beleidsnota-s', 'beleidsnotas'

                $changed.Add($item.Key, $value)
            }

            $file_name = $changed.File

            $file_extn = [System.IO.Path]::GetExtension($file_name)
            $file_base = [System.IO.Path]::GetFileNameWithoutExtension($file_name)



            $dir_relpath = $changed.Dir -replace "/$([regex]::Escape($file_base))", ''
            $dir_relpath = Limit-Path $dir_relpath
            $dir_abspath = Join-Path $PSScriptRoot -ChildPath 'Archive' -AdditionalChildPath $dir_relpath

            $TargetDir = [PSCustomObject] @{
                RelPath = $dir_relpath
                AbsPath = $dir_abspath
            }



            $file_base = Limit-Path $file_base
            $file_name = $file_base + $file_extn
            $file_path = Join-Path $dir_abspath -ChildPath $file_name

            $TargetFile = [PSCustomObject] @{
                Name    = $file_name
                RelPath = [System.IO.Path]::GetRelativePath($repo_root, $file_path)
                AbsPath = $file_path
            }

            return [PSCustomObject] @{
                SourceUri  = $SourceUri
                TargetDir  = $TargetDir
                TargetFile = $TargetFile
            }
        }
    }
}
#endregion

#region -- Declare: Save-TargetFileFromArchiveSite --
function Save-TargetFileFromArchiveSite {

    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [PSCustomObject] $SourceUri,

        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [PSCustomObject] $TargetDir,

        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [PSCustomObject] $TargetFile
    )

    process {
        try {
            if (-not $(Test-Path $TargetDir.AbsPath)) {
                mkdir $TargetDir.AbsPath -ErrorAction Ignore | Out-Null
            }

            if ($(Test-Path $TargetFile.AbsPath -PathType Leaf)) {
                $Logging.Warn_FileAllreadyExists($TargetDir.RelPath, $TargetFile.Name)

            } else {
                try {
                    Invoke-WebRequest -Uri $SourceUri -OutFile $TargetFile.AbsPath -TimeoutSec 0
                    $Logging.Info_FileDownloaded($TargetDir.RelPath, $TargetFile.Name)
                    return $TargetFile

                } catch [Http.HttpRequestException] {
                    if ($_.Exception.HttpRequestError -eq 'ConnectionError') {
                        $Logging.Warn_ConnectionFailed()

                        Start-Sleep -Seconds 3

                        Invoke-WebRequest -Uri $SourceUri -OutFile $TargetFile.AbsPath -TimeoutSec 0
                        $Logging.Info_FileDownloaded($TargetDir.RelPath, $TargetFile.Name)
                        return $TargetFile

                    } else {
                        throw $_
                    }
                }
            }

        } catch {
            Write-Error $_ -ErrorAction Continue
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



    $Logging.Info_StartedDownloadingAllFiles()
    $Search = Initialize-Search -StartFromPageNr 1

    while (-not $Search.HasFinishedPages) {

        $Search.GetResultsFromNextPage() |
            Get-TargetInfoFromSourceUri |
            Save-TargetFileFromArchiveSite |
            ForEach-Object {
                $BatchProc.ForEachItem($PSItem)
            }
    }

    $BatchProc.FlushItems()

} finally {
    Pop-Location
}
