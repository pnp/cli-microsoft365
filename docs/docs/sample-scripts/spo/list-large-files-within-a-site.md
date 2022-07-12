# List all large files within a SharePoint Site

Author: [Veronique Lengelle](https://veronicageek.com/2019/get-files-bigger-50mb/)

The following script will help you find every files in a specific SharePoint Online site that are over a certain size. It iterates through all libraries and identifies all files larger than the set threshold.

=== "PowerShell"

    ```powershell
    param (
        [Parameter(Mandatory = $true, HelpMessage = "URL of the target site", Position = 0)]
        [string]$SiteUrl,
        [Parameter(Mandatory = $true, HelpMessage = "Size in MB", Position = 1)]
        [int]$SizeInMB,
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

            #Cast as [long] in case some are above 1GB
            $fileSizeToNum = [long]($file.Length)

            if (($fileSizeToNum -ge ($SizeInMB * 1000000)) -and ($file.name -like "*.*")) {
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

- SharePoint Online
- Governance
