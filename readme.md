# PnPObsoleteVersion

## Remove-PnPObsoleteVersion

Remove every obsolete versions from all files in a SharePoint site. Versions are considerate obsolete if they are either too old or too numerous.

- `-SiteUrl`: Indicate the URL of the SharePoint site that you want to cleanup
- `-MaxVersionHistory`: Maximum number of versions to keep on each file
- `-MaxVersionAge`: Delete all versions older than X year(s)
- `-NoTranscript`: Doesn't create a transcript file

## Start-PnPSiteCleanup

This is the script to set and forget in a scheduled task. It will get all large SharePoint sites and force a cleanup using the first script.

- `-AdminUrl`: URL of the SharePoint admin center of your tenant
- `-SiteMinimumSizeGB`: Threshold in GB to filter out SharePoint sites that are too small
- `-DaysBetweenCleanup`: Number of days before the SharePoint site can be cleaned again
- `-SitesPerCleanup`: Number of SharePoint sites that can be cleaned on the same execution
- `-MaxVersionAge`: Delete all versions older than X year(s)

## Entra ID app registration

The usage of an Entra ID app registration is strongly recommended to avoid being throttled by SharePoint Online. You can find more information about throttling here: [Avoid getting throttled or blocked in SharePoint Online \| Microsoft Learn](https://learn.microsoft.com/sharepoint/dev/general-development/how-to-avoid-getting-throttled-or-blocked-in-sharepoint-online).

You must complete the file `registeredAppSettings.json` to add necessary information to connect to the app registration:

JSON property  | Information
-------------- | -----------
thumbprint     | Thumbprint of your certificate
clientId       | Application (client) ID
tenant         | Directory (tenant) ID

> [!WARNING]
> You must connect to the app registration using a certificate. The scripts **will not work** if you are using a secret to connect to the app.

### Permissions

If you plan to use an Entra ID app registration, you'll need the following SharePoint API permissions:

- `Sites.FullControl.All`: to use the command `Get-PnPTenantSite` in the script `Start-PnPSiteCleanup.ps1`
- `Sites.ReadWrite.All`: for the script `Remove-PnPObsoleteVersion.ps1`

You can use [Register-PnPAzureADApp \| PnP PowerShell](https://pnp.github.io/powershell/cmdlets/Register-PnPAzureADApp.html) to do this in PowerShell.

## Result

Scripts will output their results in `results.csv` file.

### Output

Property       | Value
--------       | -----
Site           | <https://contoso.sharepoint.com/sites/leadership-connection>
FilesCount     | 30
GBRecovered    | 2.25
VersionDeleted | 60
Errors         | 0
StartTime      | 2024-07-31 10:15:51
EndTime        | 2024-07-31 10:18:51
TimeElapsedSec | 179

### Transcript

```plaintext
Get all files...
Files with versions older than 2 year(s): 30
Versions count: 90
[2024-07-31 10:18:19] Processing '*************************************************.xlsx' with 7 version(s)
    Remove version: 512
    Remove version: 1024
    Remove version: 1536
    Remove version: 2048
    Remove version: 2560
    Remove version: 3072
    Remove version: 3584
[2024-07-31 10:18:21] Processing '*************************************************************.pdf' with 1 version(s)
    Remove version: 512
[2024-07-31 10:18:21] Processing '**************************************************************.pdf' with 1 version(s)
    Remove version: 512
[2024-07-31 10:18:22] Processing '************************.xlsm' with 1 version(s)
    Remove version: 512
[2024-07-31 10:18:22] Processing '***********************************************.jpg' with 1 version(s)
    Remove version: 512
[2024-07-31 10:18:23] Processing '**********************************************************************.pdf' with 1 version(s)
    Remove version: 512
[2024-07-31 10:18:23] Processing '*******************************************************************.pdf' with 1 version(s)
    Remove version: 512
[2024-07-31 10:18:24] Processing '*************************************************.pdf' with 1 version(s)
    Remove version: 512
[2024-07-31 10:18:24] Processing '***************************************************************************.pdf' with 2 version(s)
    Remove version: 512
    Remove version: 1024
[2024-07-31 10:18:28] Processing '*****************************************.pdf' with 1 version(s)
    Remove version: 512
[2024-07-31 10:18:29] Processing '*****************************************.pdf' with 1 version(s)
    Remove version: 512
[2024-07-31 10:18:29] Processing '***************************************.pdf' with 1 version(s)
    Remove version: 512
[2024-07-31 10:18:30] Processing '************************************************.pdf' with 1 version(s)
    Remove version: 512
[2024-07-31 10:18:30] Processing '**************.xlsx' with 13 version(s)
    Remove version: 512
    Remove version: 1024
    Remove version: 1536
    Remove version: 2048
    Remove version: 2560
    Remove version: 3072
    Remove version: 3584
    Remove version: 4096
    Remove version: 4608
    Remove version: 5120
    Remove version: 5632
    Remove version: 6144
    Remove version: 6656
[2024-07-31 10:18:40] Processing '***************************************************.pdf' with 1 version(s)
    Remove version: 512
[2024-07-31 10:18:41] Processing '****************************************.pdf' with 1 version(s)
    Remove version: 512
[2024-07-31 10:18:41] Processing '********************************************************.pdf' with 1 version(s)
    Remove version: 512
[2024-07-31 10:18:42] Processing '******************************************.pdf' with 1 version(s)
    Remove version: 512
[2024-07-31 10:18:42] Processing '*************************************************.pdf' with 1 version(s)
    Remove version: 512
[2024-07-31 10:18:42] Processing '*********************************************************************.pdf' with 1 version(s)
    Remove version: 512
[2024-07-31 10:18:43] Processing '**************************************.pdf' with 2 version(s)
    Remove version: 512
    Remove version: 1024
[2024-07-31 10:18:44] Processing '**********************************.pdf' with 2 version(s)
    Remove version: 512
    Remove version: 1024
[2024-07-31 10:18:44] Processing '****************************************.pdf' with 2 version(s)
    Remove version: 512
    Remove version: 1024
[2024-07-31 10:18:45] Processing '********************************************.pdf' with 2 version(s)
    Remove version: 512
    Remove version: 1024
[2024-07-31 10:18:46] Processing '*************************************.pdf' with 2 version(s)
    Remove version: 512
    Remove version: 1024
[2024-07-31 10:18:47] Processing '************************.xlsx' with 7 version(s)
    Remove version: 512
    Remove version: 1024
    Remove version: 1536
    Remove version: 2048
    Remove version: 2560
    Remove version: 3072
    Remove version: 3584
[2024-07-31 10:18:49] Processing '*******************************.xlsx' with 1 version(s)
    Remove version: 512
[2024-07-31 10:18:49] Processing '*****************.xlsx' with 1 version(s)
    Remove version: 512
[2024-07-31 10:18:50] Processing '***********************.pptx' with 1 version(s)
    Remove version: 512
[2024-07-31 10:18:50] Processing '********************.pptx' with 1 version(s)
    Remove version: 512
```
