#Requires -Version 7.4.2
#Requires -Modules @{ModuleName='PnP.PowerShell';ModuleVersion='2.4.0'}

[CmdletBinding(DefaultParameterSetName = 'History')]
param(
    [Parameter(Mandatory, ParameterSetName = 'History')]
    [Parameter(Mandatory, ParameterSetName = 'Age')]
    [string]$SiteUrl,

    [Parameter(ParameterSetName = 'History')]
    [int]$MaxVersionHistory,

    [Parameter(ParameterSetName = 'Age')]
    [int]$MaxVersionAge,

    [switch]$NoTranscript
)

if ($NoTranscript.IsPresent -eq $false) { Start-Transcript "$PSScriptRoot\logs\Remove-PnPObsoleteVersion_$(Get-Date -Format 'yyyy-MM-dd_HHmmss').txt" }
$start = Get-Date

# Connect to SharePoint site
$appSettings = Get-Content -Path "$PSScriptRoot\registeredAppSettings.json" | ConvertFrom-Json -AsHashtable
$appSettings.Add('Url', $SiteUrl)
Connect-PnPOnline @appSettings

# Get all files
Write-Host "Get all files..."
$ListItems = Get-PnPListItem -List 'Documents' -Includes Versions, FileSystemObjectType -PageSize 1000 -ScriptBlock {
    param($items) $items.Context.ExecuteQuery() | ForEach-Object { $_ }
}
$ListItems = $ListItems | Where-Object { $_.FileSystemObjectType -eq 'File' }

# Prepare progress bar & processing
if ($MaxVersionHistory) {
    $MaxVersionHistory++ # To avoid catching the current version
    $filter1 = { $_.Versions.Count -gt $MaxVersionHistory }
    $filter2 = { $true }
    $splat   = @{ SkipLast = $MaxVersionHistory }
}
if ($MaxVersionAge) {
    $filter1 = { $_.Versions.Count -gt 1 -and $_.Versions.Created.Date -lt (Get-Date).AddYears(-$MaxVersionAge) }
    $filter2 = { $_.Created.Date -lt (Get-Date).AddYears(-$MaxVersionAge) }
    $splat   = $null
}

$ListItems = $ListItems | Where-Object $filter1
$total = ($ListItems | Measure-Object).Count

# Show statistics of the audit
$versionsCount = $ListItems.Versions.Count
if ($MaxVersionHistory) { Write-Host "Files with more than $MaxVersionHistory versions: $total" }
if ($MaxVersionAge) { Write-Host "Files with versions older than $MaxVersionAge year(s): $total" }
Write-Host "Versions count: $versionsCount"

# Removing old versions
$i = 0
$removed = $ListItems | ForEach-Object {
    $fileUrl = $_['FileRef']
    $fileName = ($fileUrl -split '/')[-1]

    # Processing file progress bar
    Write-Progress -Activity "$i/$total" -Status "Processing file: $fileUrl" -PercentComplete ($i / $total * 100) -CurrentOperation 'fileLoop' -ID 1
    $i++

    # Get versions to remove
    $versionsToDelete = Get-PnPFileVersion -Url $fileUrl | Where-Object $filter2 | Select-Object @splat

    # Write progress
    Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Processing '$fileName' with $($versionsToDelete.Count) version(s)" -ForegroundColor Yellow

    # Remove old versions
    $j = 0
    $versionsToDelete | ForEach-Object {
        $id = $_.ID

        # Processing version progress bar
        Write-Progress -Activity "[$j/$($versionsToDelete.Count)]" -Status "Remove version: $id" -PercentComplete ($j / $($versionsToDelete.Count) * 100) -CurrentOperation 'versionLoop' -ID 2
        $j++

        try {
            Write-Host "    Remove version: $id"
            Remove-PnPFileVersion -Url $fileUrl -Identity $_.ID -Force -EA Stop
            $_.Size
        }
        catch {
            $errorMessage = $_.Exception
            Write-Host "    Error removing version $id`: $($errorMessage.Message)" -ForegroundColor Red
            0
        }
    }
}

$end = Get-Date
$timespan = New-TimeSpan -Start $start -End $end

# Show results
$results = [PSCustomObject]@{
    Site           = $SiteUrl
    FilesCount     = $total
    GBRecovered    = [math]::Round((($removed | Measure-Object -Sum).Sum / 1GB),2)
    VersionDeleted = (($removed | Where-Object { $_ -ne 0 } | Measure-Object).Count)
    Errors         = (($removed | Where-Object { $_ -eq 0 } | Measure-Object).Count)
    StartTime      = (Get-Date $start -Format 'yyyy-MM-dd HH:mm:ss')
    EndTime        = (Get-Date $end -Format 'yyyy-MM-dd HH:mm:ss')
    TimeElapsedSec = [int]($timespan.TotalSeconds)
}

# Export results
$results | Format-List
$results | Export-Csv -Path "$PSScriptRoot\results.csv" -Append -Encoding UTF8 -Delimiter ';'

if ($NoTranscript.IsPresent -eq $false) { Stop-Transcript }