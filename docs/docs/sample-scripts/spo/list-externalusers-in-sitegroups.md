# List all external users in site groups across all site collections

Author: [Martin Lingstuyl](https://www.blimped.nl)

This script shows how you can check if external users are added to site groups. It will show all external users across all site collections and the site groups they where added to.

=== "PowerShell"

    ```powershell
    $m365Status = m365 status --output text

    if ($m365Status -eq "Logged Out") {
      m365 login
    }

    Write-Host "Retrieving all sites and check external users..." -ForegroundColor Green

    $sites = m365 spo site list --type All | ConvertFrom-Json
    $siteCount = $sites.Count
    $siteCounter = 0
    $results = [System.Collections.ArrayList]::new()

    $spoAccessToken = m365 util accesstoken get --resource sharepoint --new | ConvertFrom-Json

    Write-Host "Processing $siteCount sites..."

    foreach ($site in $sites) {
      $siteCounter++  
      Write-Host "$siteCounter/$siteCount - Get external users in site groups for $($site.Url)..." -ForegroundColor Green

      $response = Invoke-WebRequest -Uri "$($site.Url)/_api/web/siteusers?`$filter=IsShareByEmailGuestUser eq true&`$expand=Groups&`$select=Title,LoginName,Email,Groups/LoginName" -Method Get -Headers @{ Authorization = "Bearer $spoAccessToken"; Accept = "application/json;odata=nometadata" }
      $users = $response.Content | ConvertFrom-Json  

      foreach($user in $users.value) {
        foreach($group in $user.Groups) {
          $obj = [PSCustomObject][ordered]@{
              Title = $user.Title;
              Email = $user.Email;
              LoginName = $user.LoginName;
              Group = $group.LoginName;
          }
          $results.Add($obj) | Out-Null
        }
      }
    }

    Write-Host "Exporting list..." -ForegroundColor Green
    $results | Export-Csv -Path "./cli-external-users-in-sitegroups.csv" -NoTypeInformation
    ```

Keywords:

- SharePoint Online
- Governance
- External Users
