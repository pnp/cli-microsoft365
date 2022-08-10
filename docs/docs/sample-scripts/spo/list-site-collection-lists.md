# List site collections and their lists

Author: [Albert-Jan Schot](https://www.cloudappie.nl/migration-report-climicrosoft365)

This script helps you to list and export all site collection and their lists SharePoint Online sites, ideal for getting insights into the size of your environment.

=== "PowerShell"

    ```powershell
    $fileExportPath = "<PUTYOURPATHHERE.csv>"

    $m365Status = m365 status --output text

    if ($m365Status -eq "Logged Out") {
      # Connection to Microsoft 365
      m365 login
    }

    $results = @()
    Write-host "Retrieving all sites..."
    $allSPOSites = m365 spo site classic list -o json | ConvertFrom-Json
    $siteCount = $allSPOSites.Count

    Write-Host "Processing $siteCount sites..."
    #Loop through each site
    $siteCounter = 0

    foreach ($site in $allSPOSites) {
      $siteCounter++
      Write-Host "Processing $($site.Url)... ($siteCounter/$siteCount)"

      $results += [pscustomobject][ordered]@{
        Type         = "site"
        Title        = $site.Title
        Url          = $site.Url
        StorageUsage = $site.StorageUsage
        Template     = $site.Template
      }

      Write-host "Retrieving all lists..."

      $allLists = m365 spo list list -u $site.url -o json | ConvertFrom-Json
      foreach ($list in $allLists) {

        $results += [pscustomobject][ordered]@{
          Type     = "list"
          Title    = $list.Title
          Url      = $list.Url
          Template = $list.BaseTemplate
        }
      }
    }

    Write-Host "Exporting file to $fileExportPath..."
    $results | Export-Csv -Path $fileExportPath -NoTypeInformation
    Write-Host "Completed."
    ```

Keywords:

- SharePoint Online
- Governance
