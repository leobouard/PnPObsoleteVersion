#Requires -Version 7.4.2
#Requires -Modules @{ModuleName='PnP.PowerShell';ModuleVersion='2.4.0'}

param(
    [string]$AdminUrl = 'https://contoso-admin.sharepoint.com',
    [int]$SiteMinimumSizeGB = 100,
    [int]$DaysBetweenCleanup = 30,
    [int]$SitesPerCleanup = 1,
    [int]$MaxVersionAge = 2
)

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

# Exclude all SharePoint site that have been done recently (in the last X days)
$results = Import-Csv -Path "$PSScriptRoot\results.csv" -Delimiter ';' -Encoding UTF8 |
Where-Object { (Get-Date $_.EndTime) -lt (Get-Date).AddDays(-$DaysBetweenCleanup) }

# Remove obsolete versions
$sites | Where-Object { $_.Url -notin $results.Site } | Get-Random -Count $SitesPerCleanup | ForEach-Object {
    .\Remove-PnPObsoleteVersion -MaxVersionAge $MaxVersionAge -SiteURL $_.Url
}
