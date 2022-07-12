# Replace site collection admin with another user

Author: [Patrick Lamber](https://www.nubo.eu/Replace-Site-Collection-Admin-Using-CLI/)
Inspired By: [Salaudeen Rajack](https://www.sharepointdiary.com/2015/08/sharepoint-online-add-site-collection-administrator-using-powershell.html)

The script removes a user from a site collection and adds a new one as site collection admin.

=== "PowerShell"

    ```powershell
    $userToAdd = "<upnOfUserToAdd>"
    $userToRemove = "<upnOfUserToRemove>"
    $webUrl = "<spoUrl>"

    $m365Status = m365 status --output text
    Write-Host $m365Status
    if ($m365Status -eq "Logged Out") {
      # Connection to Microsoft 365
      m365 login
      $m365Status = m365 status --output text
    }

    m365 spo user remove --webUrl $webUrl --loginName "i:0#.f|membership|$userToRemove" --confirm
    m365 spo site classic set --url $webUrl --owners $userToAdd
    ```

Keywords

- SharePoint Online
- Governance
