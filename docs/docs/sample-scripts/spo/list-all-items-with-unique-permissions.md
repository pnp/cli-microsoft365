# List all items with unique permissions

Author: [Veronique Lengelle](https://twitter.com/veronicageek), Inspired by [Salaudeen Rajack](https://www.sharepointdiary.com/2017/03/sharepoint-online-get-all-list-items-with-unique-permissions-using-powershell.html)

## List all items for a specific SharePoint list on a site

=== "PowerShell"

    ```powershell
    #Declare variables
    $siteURL = "https://<TENANT-NAME>.sharepoint.com/sites/<YOUR-SITE>"
    $listName = "<YOUR-LIST-NAME>"
    $allItems = m365 spo listitem list --webUrl $siteUrl --title $listName --fields "ID, HasUniqueRoleAssignments, Title" | ConvertFrom-Json
    $results = @()

    #Loop through each item in the list
    foreach($item in $allItems){
        $results += [pscustomobject][ordered]@{
            ListName = $listName
            ItemID = $item.Id
            ItemTitle = $item.Title
            UniquePermissions = $item.HasUniqueRoleAssignments
        }
    }
    $results
    ```

## List all items for multiple SharePoint lists on a site

=== "PowerShell"

    ```powershell
    #Declare variables
    $siteURL = "https://<TENANT-NAME>.sharepoint.com/sites/<YOUR-SITE>"
    $allLists = m365 spo list list --webUrl $siteUrl --query "[?BaseTemplate == ``100``]" | ConvertFrom-Json
    $results = @()

    foreach($list in $allLists){
        if ($list.Hidden -eq $false){ 
            
            $allItems = m365 spo listitem list --webUrl $siteURL --id $list.Id --fields "ID, HasUniqueRoleAssignments, Title" | ConvertFrom-Json
            
            foreach($item in $allItems){
                $results += [pscustomobject][ordered]@{
                    ListName = $list.Title
                    ItemID = $item.Id
                    ItemTitle = $item.Title
                    UniquePermissions = $item.HasUniqueRoleAssignments
                }
            }
        }
    }
    $results
    ```

Keywords

- SharePoint Online
- Governance