param(
    [switch]$IncludeOneDrives,
    [switch]$IncludeSystemLibraries,

    # Optional filters (handy for testing or partial runs)
    [string]$SiteUrlLike = "",        # e.g. "*sharepoint.com/sites/*" or "*sites/Test*"
    [string]$LibraryNameLike = "",    # e.g. "*Documentos*" or "*Documents*"

    [string]$OutputRoot = (
        Join-Path -Path (Get-Location) `
        -ChildPath ("AllFilesReport_{0}" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
    ),

    # App-only parameters (recommended)
    [Parameter(Mandatory=$true)]
    [string]$TenantId,

    [Parameter(Mandatory=$true)]
    [Alias('ClientId')]
    [string]$AppId,

    [Parameter(Mandatory=$true)]
    [string]$CertificateThumbprint,

    [int]$PageSize = 200,

    # Throttling friendliness
    [int]$SleepMsBetweenRequests = 0
)

function Connect-GraphAppOnly {
    Write-Host "Connecting to Microsoft Graph (App-only)..." -ForegroundColor Cyan
    Disconnect-MgGraph -ErrorAction SilentlyContinue

    # App-only certificate auth supported by Connect-MgGraph parameter set. 
    Connect-MgGraph -TenantId $TenantId -ClientId $AppId -CertificateThumbprint $CertificateThumbprint -NoWelcome | Out-Null

    $ctx = Get-MgContext
    if (-not $ctx) { throw "No Microsoft Graph session found after Connect-MgGraph." }
    if ($ctx.AuthType -ne "AppOnly") { throw "Not connected as AppOnly. AuthType: $($ctx.AuthType)" }

    Write-Host "TenantId : $($ctx.TenantId)" -ForegroundColor Yellow
    Write-Host "AuthType : $($ctx.AuthType)" -ForegroundColor Yellow
}

function Invoke-GetJson {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Uri,

        [int]$MaxRetries = 6
    )

    $attempt = 0
    while ($true) {
        try {
            $resp = Invoke-MgGraphRequest -Method GET -Uri $Uri -ErrorAction Stop
            if ($SleepMsBetweenRequests -gt 0) { Start-Sleep -Milliseconds $SleepMsBetweenRequests }
            return $resp
        } catch {
            $attempt++
            if ($attempt -gt $MaxRetries) { throw }

            $delay = [Math]::Min(60000, (1000 * [Math]::Pow(2, $attempt)))
            Write-Warning "Graph request failed (attempt $attempt/$MaxRetries). Sleeping ${delay}ms. $($_.Exception.Message)"
            Start-Sleep -Milliseconds $delay
        }
    }
}

function Extract-FolderPath {
    param([string]$parentPath)
    if ([string]::IsNullOrWhiteSpace($parentPath)) { return "/" }

    # parentReference.path often looks like: /drives/{driveId}/root:/Folder/Subfolder
    if ($parentPath -match 'root:(.*)$') {
        $p = $Matches[1]
        if ([string]::IsNullOrWhiteSpace($p)) { return "/" }
        return $p
    }
    return "/"
}

# Common system libraries to skip unless requested
$DefaultExcludedLibraries = @(
    'Form Templates','Preservation Hold Library','Site Assets','Style Library',
    'Site Pages','Pages','Images','Site Collection Documents','Site Collection Images'
)

# Prepare output folder + CSV
$null = New-Item -ItemType Directory -Force -Path $OutputRoot
$outCsv = Join-Path $OutputRoot "AllFiles.csv"
$script:wroteHeader = $false

function Write-Row {
    param([pscustomobject]$Row)
    if (-not $script:wroteHeader) {
        $Row | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $outCsv
        $script:wroteHeader = $true
    } else {
        $Row | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $outCsv -Append
    }
}

# Counters
$TotalSites  = 0
$TotalDrives = 0
$TotalItems  = 0
$TotalFiles  = 0

Connect-GraphAppOnly

Write-Host "Enumerating all SharePoint sites..." -ForegroundColor Cyan
$sites = Get-MgSite -All

# Optional filters
if (-not $IncludeOneDrives) {
    $sites = $sites | Where-Object { $_.WebUrl -notlike "*-my.sharepoint.com*" }
}
if ($SiteUrlLike) {
    $sites = $sites | Where-Object { $_.WebUrl -like $SiteUrlLike }
}

$siteIndex = 0
$totalSitesCount = ($sites | Measure-Object).Count

foreach ($site in $sites) {
    $siteIndex++
    $TotalSites++

    Write-Progress -Activity "Processing sites" `
        -Status "[$siteIndex/$totalSitesCount] $($site.WebUrl)" `
        -PercentComplete (($siteIndex / [math]::Max($totalSitesCount,1)) * 100)

    # List drives (document libraries) for the site: /sites/{siteId}/drives. 
    $drivesResp = Invoke-GetJson -Uri ("https://graph.microsoft.com/v1.0/sites/{0}/drives" -f $site.Id)
    $drives = @($drivesResp.value) | Where-Object { $_.driveType -eq "documentLibrary" }

    if (-not $IncludeSystemLibraries) {
        $drives = $drives | Where-Object { $_.name -notin $DefaultExcludedLibraries }
    }
    if ($LibraryNameLike) {
        $drives = $drives | Where-Object { $_.name -like $LibraryNameLike }
    }

    foreach ($d in $drives) {
        $TotalDrives++

        Write-Progress -Activity "Processing libraries" `
            -Status "$($site.WebUrl) ? $($d.name)" `
            -PercentComplete 0

        # Delta enumeration (recursive) for the drive. 
        $uri = "https://graph.microsoft.com/v1.0/drives/{0}/root/delta?`$top={1}&`$select=id,name,webUrl,size,file,folder,fileSystemInfo,parentReference,deleted" -f $d.id, $PageSize

        while ($uri) {
            $resp = Invoke-GetJson -Uri $uri
            $uri = $null

            foreach ($it in $resp.value) {
                $TotalItems++

                # ? Correct file detection:
                # Graph driveItem facet model => folders have folder facet; files have file facet. 
                $isDeleted = ($null -ne $it.deleted)
                $isFolder  = ($null -ne $it.folder)
                $isFile    = ($null -ne $it.file) -and (-not $isFolder) -and (-not $isDeleted)

                if ($isFile) {
                    $TotalFiles++

                    $folderPath = Extract-FolderPath -parentPath $it.parentReference.path
                    $qxh = $null
                    if ($it.file.hashes -and $it.file.hashes.quickXorHash) { $qxh = $it.file.hashes.quickXorHash }

                    Write-Row ([pscustomobject]@{
                        SiteUrl      = $site.WebUrl
                        Library      = $d.name
                        FolderPath   = $folderPath
                        FileName     = $it.name
                        FileUrl      = $it.webUrl
                        SizeBytes    = $it.size
                        LastModified = $it.fileSystemInfo.lastModifiedDateTime
                        DriveId      = $d.id
                        ItemId       = $it.id
                        QuickXorHash = $qxh
                    })
                }
            }

            if ($resp.'@odata.nextLink') { $uri = $resp.'@odata.nextLink' }
        }
    }
}

Write-Host "======================================" -ForegroundColor Yellow
Write-Host "Sites processed   : $TotalSites" -ForegroundColor Yellow
Write-Host "Libraries scanned : $TotalDrives" -ForegroundColor Yellow
Write-Host "Items returned    : $TotalItems" -ForegroundColor Yellow
Write-Host "Files discovered  : $TotalFiles" -ForegroundColor Yellow
Write-Host "AllFiles.csv      : $outCsv" -ForegroundColor Yellow
Write-Host "Output folder     : $OutputRoot" -ForegroundColor Yellow
Write-Host "======================================" -ForegroundColor Yellow