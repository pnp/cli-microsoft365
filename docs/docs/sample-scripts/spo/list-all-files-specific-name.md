# List all documents with a specific name within a SharePoint site

Author: [Veronique Lengelle](https://veronicageek.com/2019/get-files-with-specific-names/)

This script will retrieve all the files in a site that have a specific word (i.e.: search all documents where the word "CLI" is part of the file name).

=== "PowerShell"

    ```powershell
    param (
        [Parameter(Mandatory = $true, HelpMessage = "URL of the target site", Position = 0)]
        [string]$SiteUrl,
        [Parameter(Mandatory = $true, HelpMessage = "Filename ", Position = 1)]
        [string]$FileName,
        [Parameter(HelpMessage = "Show progress messages", Position = 2)]
        [switch]$ShowProgress
    )

    $m365Status = m365 status --output text
    if ($m365Status -eq "Logged Out") {
      # Connection to Microsoft 365
      m365 login
    }

    #Declare variables
    $site = $SiteUrl
    $results = @()
    $allLibs = m365 spo list list --webUrl $site --query "[?BaseTemplate == ``101``]" -o json | ConvertFrom-Json

    foreach ($lib in $allLibs) {
        # Counters
        $i++
        if ($ShowProgress) { Write-Host "Processing '$($lib.Title)' - ($i/$($allLibs.length))" }

        $allFiles = m365 spo file list --webUrl $site --folder $lib.Url --recursive -o json | ConvertFrom-Json

        foreach ($file in $allFiles) {
            if ($ShowProgress) { Write-Host "Processing file '$($file.ServerRelativeUrl)'" }

            if (($file.name -like "*$($FileName)*")) {
                $results += [pscustomobject]@{
                    libraryName      = $lib.Title
                    fileName         = $file.Name
                    filePath         = $file.ServerRelativeUrl
                    normalSize       = $file.Length
                    SizeInMB         = ($fileSizeToNum / 1MB).ToString("N")
                    LastModifiedDate = [DateTime]$file.TimeLastModified
                }
            }
        }
    }
    $results
    ```

Keywords

-   SharePoint Online
-   Governance
