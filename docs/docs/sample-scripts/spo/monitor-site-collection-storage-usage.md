# Monitor Site Collections Storage Usage

Inspired by [Salaudeen Rajack](https://www.sharepointdiary.com/2020/08/sharepoint-online-monitor-site-storage-usage-with-powershell.html)

```powershell tab="PowerShell Core"
<#
.SYNOPSIS
    Monitor Site Collections storage usage and send an email.
.DESCRIPTION
    Monitor Site Collections storage usage and send an email with sites over the designated storage threshold.
.EXAMPLE
    PS C:\> Send-SiteCollectionStorageReport -storageThreshold 60 -sendTo john.smith@contoso.com
    Running this function with send an email to John Smith with all the sites over 60% storage used.
.INPUTS
    Inputs (if any)
.OUTPUTS
    Output (if any)
.NOTES
    None.
#>
function Send-SiteCollectionStorageReport {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, HelpMessage = "Number for the threshold percentage (i.e.: 50 for every sites above 50% storage used")]
        [int]$storageThreshold,
        [Parameter(Mandatory = $true, HelpMessage = "User email address to send the report to")]
        [string]$sendTo
    )
    #Declare variables
    $allSites = m365 spo site classic list -o json | ConvertFrom-Json
    $largeSites = $allSites | Where-Object { $_.StorageUsage -gt 1 }
    $results = @()

    # Format the headers for the email
    $style = "<style>BODY{font-family: Arial; font-size: 10pt;}"
    $style = $style + "TABLE{border: 1px solid black; border-collapse: collapse;}"
    $style = $style + "TH{border: 1px solid black; background: #dddddd; padding: 5px; }"
    $style = $style + "TD{border: 1px solid black; padding: 5px; }"
    $style = $style + "</style>"

    #Format the data 
    foreach ($site in $largeSites) {
        $results += [pscustomobject][ordered]@{
            SiteName             = $site.Title
            "StorageUsed (GB)"   = (“{0:N2}” -f ($site.StorageUsage / 1024))
            "StorageLimit (GB)"  = ($site.StorageMaximumLevel / 1024)
            StorageUsedInPercent = (“{0:P2}” -f ($site.StorageUsage / $site.StorageMaximumLevel))
        }
    }

    #Filter to only sites above the threshold limit
    $siteExceeding = $results | Where-Object { $_.StorageUsedInPercent -gt $storageThreshold } 

    #Send an email if sites are over the designated threshold limit
    if ($null -ne $siteExceeding) {
        $emailBody = $siteExceeding | ConvertTo-Html -Head $style 
        $emailBody = $emailBody -replace '"', '\"'  ## To be parsed correctly by Node.js when sending the email
        m365 outlook mail send --to $sendTo --subject "Site Collections Storage Report" --bodyContents "$emailBody" --bodyContentType HTML --saveToSentItems false
    }
}
```

Keywords

- SharePoint Online
- Governance
