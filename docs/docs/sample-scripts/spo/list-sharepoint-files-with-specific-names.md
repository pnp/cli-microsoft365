# List SharePoint files with specific names

Author: [Veronique Lengelle](https://twitter.com/veronicageek)

!!! note The script will retrieve the files even inside nested folders. In the below script sample (within the `if` statement) we'll retrieve the files where the file name contains `cli` or `report`.

```PowerShell tab="PowerShell"
<#
.SYNOPSIS
    Lsit all files with specific names.
.DESCRIPTION
    This script will list all files in a site with the specific word in their file name (defined in the condition).
.EXAMPLE
    PS C:\> .\List-FilesWithSpecificNames.ps1
    This script will retrieve all the files within a site with a specific name.
.INPUTS
    Inputs (if any)
.OUTPUTS
    Output (if any)
.NOTES
    General notes
#>

$site = "https://<TENANT-NAME>.sharepoint.com/sites/<YOUR-SITE>"
$results = @()
$allLibs = m365 spo list list --webUrl $site --query "[?BaseTemplate == ``101``]" -o json | ConvertFrom-Json

foreach($lib in $allLibs){
    $allFiles = m365 spo file list --webUrl $site --folder $lib.Title --recursive -o json | ConvertFrom-Json

    foreach($file in $allFiles){
        $specificFiles = m365 spo file get --webUrl $site --id $file.UniqueId -o json | ConvertFrom-Json

        foreach($f in $specificFiles){
            if(($f.Name -like "*cli*") -or ($f.name -like "*report*")){
                $results += [pscustomobject]@{
                    libraryName = $lib.Title
                    fileName = $f.Name
                    filePath = $f.ServerRelativeUrl
                    fileLastModified = $f.TimeLastModified
                    fileVersion = $f.UIVersionLabel
                }
            }
        }
    }
}
$results
```

Keywords

-   SharePoint Online
-   Autditing
