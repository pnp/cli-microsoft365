# Copy list items between SharePoint lists

Author: [SekThang](https://github.com/SekThang), Inspired by [Ruud](https://lazyadmin.nl/it/copy-sharepoint-list-items-to-another-list-with-powershell-and-pnp/)

This script helps you to copy list items from one list to another list. I have written script logics to migrate list items from one site collection to another site collection.

- Prerequisites: List and metadata should be created in the destination site collection site as it's in the source site collection
- Update columns to match your list in lines 12 and 14

=== "PowerShell"

    ```powershell
    $SourceSite = Read-Host -Prompt 'Source site Url'
    $DestinationSite = Read-Host -Prompt 'Destination site Url'
    $SourceList = Read-Host -Prompt 'Source list name'
    $DestinationList = Read-Host -Prompt 'Destination list name'
    
    $listItems = m365 spo listitem list --title $SourceList --webUrl $SourceSite --output json | ConvertFrom-Json
    Write-Host 'Total count in the source list is'-> -fore Green $listItems.Count
    $count = 0
    $reportLocation = Get-Location
    foreach($item in $listItems)
    {
      $count++
      m365 spo listitem add --listTitle $DestinationList --webUrl $DestinationSite --Title $item.Title --Firstname $item.Firstname --Lastname $item.Lastname
      Write-Host $count 'item has been migrated to destination list. Reference item id is' $item.Id -fore Magenta
      Write-output "Id:" $item.ID " - Firstname: " $item.Firstname "________________" | Out-File $reportLocation\Reports.txt -Append
    }
    Write-Host 'Report has been generated in .txt format, please check your drive' -fore Cyan
    ```

Keywords:

- SharePoint Online
- Lists
