#Requires -Version 7.4.2
#Requires -Modules @{ModuleName='PnP.PowerShell';ModuleVersion='2.4.0'}

param(
    [string]$AdminUrl = 'https://contoso-admin.sharepoint.com',
    [int]$SiteMinimumSizeGB = 100,
    [int]$DaysBetweenCleanup = 30,
    [int]$SitesPerCleanup = 1,
    [int]$MaxVersionAge = 2
)

Start-Transcript "$PSScriptRoot\logs\Start-PnPSiteCleanup_$(Get-Date -Format 'yyyy-MM-dd_HHmmss').txt"

# Connect to SharePoint Admin
$appSettings = Get-Content -Path "$PSScriptRoot\registeredAppSettings.json" | ConvertFrom-Json -AsHashtable
$appSettings.Add('Url', $AdminUrl)
Connect-PnPOnline @appSettings

# Get all sites
$sites = Get-PnPTenantSite

# Limit results to site larger than X GB
$limit = $SiteMinimumSizeGB * 1000
$sites = $sites | Where-Object { $_.StorageUsageCurrent -gt $limit } |
Sort-Object StorageUsageCurrent -Descending |
Select-Object Url, Title, StorageUsageCurrent
Write-Host "$(($sites | Measure-Object).Count) site(s) with over $SiteMinimumSizeGB GB of storage usage"

# Exclude all SharePoint site that have been done recently (in the last X days)
$results = Import-Csv -Path "$PSScriptRoot\results.csv" -Delimiter ';' -Encoding UTF8 |
Where-Object { (Get-Date $_.EndTime) -gt (Get-Date).AddDays(-$DaysBetweenCleanup) }
$sites = $sites | Where-Object { $_.Url -notin $results.Site }
Write-Host "$(($sites | Measure-Object).Count) site(s) left to clean"

# Pick X sites randomly
$sites = try { $sites | Get-Random -Count $SitesPerCleanup -EA Stop } catch { $sites }
Write-Host "The following site(s) are going to be cleaned:"
$sites | Format-List

# Remove obsolete versions
$sites | ForEach-Object {
    Write-Host "Processing site: $($_.Url)"
    .\Remove-PnPObsoleteVersion -MaxVersionAge $MaxVersionAge -SiteURL $_.Url -NoTranscript
}

Stop-Transcript