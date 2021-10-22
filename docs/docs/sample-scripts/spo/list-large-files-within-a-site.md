# List all the large files within a SharePoint Site

Author: [Veronique Lengelle](https://twitter.com/veronicageek)

```powershell tab="PowerShell"
<#
.SYNOPSIS
  List all the large files.
.DESCRIPTION
  List all the files with the size (in MB) you define.
.EXAMPLE
  PS C:\> List-SPOLargeFilesInSite -SiteUrl "https://contoso.sharepoint.com/sites/Marketing" -SizeInMB 500
  This script will retrieve all the files that are greater or equal to 500MB with the site provided.
.EXAMPLE
  PS C:\> List-SPOLargeFilesInSite -SiteUrl "https://contoso.sharepoint.com/sites/IT" -SizeInMB 1000
  This script will retrieve all the files that are greater or equal to 1GB with the site provided.
.INPUTS
  Inputs (if any)
.OUTPUTS
  Output (if any)
.NOTES
  This script will look recursively into nested folders too. So be patient as it might take time!
#>
[CmdletBinding()]
param (
  [Parameter(Mandatory = $true, HelpMessage = "URL of the target site", Position = 0)]
  [string]$SiteUrl,
  [Parameter(Mandatory = $true, HelpMessage = "Size in MB", Position = 1)]
  [int]$SizeInMB
)
#Declare variables
$site = $SiteUrl
$results = @()
$allLibs = m365 spo list list --webUrl $site --query "[?BaseTemplate == ``101``]" -o json | ConvertFrom-Json

foreach($lib in $allLibs){
$allFiles = m365 spo file list --webUrl $site --folder $lib.Title --recursive -o json | ConvertFrom-Json

    foreach($file in $allFiles){
        $largeFiles = m365 spo file get --webUrl $site --id $file.UniqueId -o json | ConvertFrom-Json

        foreach($f in $largeFiles){
            #Cast as [long] in case some are above 1GB
            $fileSizeToNum = [long]($f.Length)

            if(($fileSizeToNum -ge ($SizeInMB * 1000000)) -and ($f.name -like "*.*")){
                $results += [pscustomobject]@{
                    libraryName = $lib.Title
                    fileName = $f.Name
                    filePath = $f.ServerRelativeUrl
                    normalSize = $f.Length
                    SizeInMB = ($fileSizeToNum / 1MB).ToString("N")
                    LastModifiedDate = [DateTime]$file.TimeLastModified
                }
            }
        }
    }

}
$results
```

Keywords

-   SharePoint Online
-   Governance
