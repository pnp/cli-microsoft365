# Lists number of files in all lists and folders for the given site

Author: [Albert-Jan Schot](https://www.cloudappie.nl/lists-file-count-cli-microsoft-365/)

List all Lists, the folders and sub folders in a given site, and output the item count. Each folder is processed recursively. By default only non hidden document libraries are processed. As specified with the filter `$false -eq $list.Hidden -and $list.BaseTemplate -eq "101"`. The output is a CSV that contains the itemcount for each list and folder found in the specified site collection.

=== "PowerShell"

    ```powershell
    $siteUrl = "<PUTYOURURLHERE>"
    $fileExportPath = "<PUTYOURPATHHERE.csv>"

    $m365Status = m365 status

    if ($m365Status -eq "Logged Out") {
      # Connection to Microsoft 365
      m365 login
    }

    [System.Collections.ArrayList]$results = @()

    function Get-Folders($webUrl, $folderUrl) {
      $folders = m365 spo folder list -u $webUrl --parentFolderUrl $folderUrl -o json | ConvertFrom-Json

      foreach ($folder in $folders) {
        $folderStats = m365 spo folder get -u $webUrl --folderUrl $folder.ServerRelativeUrl -o json | ConvertFrom-Json

        Write-Output "Processing folder: $($folder.ServerRelativeUrl);"
        [void]$results.Add([pscustomobject]@{ Url = $folder.ServerRelativeUrl; ItemCount = $folderStats.ItemCount; Type = "Folder"; })

        Get-Folders $webUrl $folder.ServerRelativeUrl
      }
    }

    $allLists = m365 spo list list -u $siteUrl -o json | ConvertFrom-Json

    foreach ($list in $allLists) {
      if ($false -eq $list.Hidden -and $list.BaseTemplate -eq "101") {
        Write-Output "Processing $($list.Url)"
        [void]$results.Add([PSCustomObject]@{ Url = $list.Url; ItemCount = $list.ItemCount; Type = "List"; })

        Get-Folders $siteUrl $list.Url
      }
    }

    $results | Export-Csv -Path $fileExportPath -NoTypeInformation
    ```

Keywords:

- SharePoint Online
- Governance
