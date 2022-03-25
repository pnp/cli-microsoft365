# List Attachment Names From SharePoint Lists For A Site

Author: [Veronique Lengelle](https://twitter.com/veronicageek)

=== "PowerShell"

    ```powershell
    $siteUrl = "https://<TENANT-NAME>.sharepoint.com/sites/<YOUR-SITE>"
    $allLists = m365 spo list list --webUrl $siteUrl --query "[?BaseTemplate == ``100``]" | ConvertFrom-Json
    $results = @()

    foreach($list in $allLists){
        if ($list.Hidden -eq $false){ 
            $allItems = m365 spo listitem list --id $list.Id --webUrl $siteUrl | ConvertFrom-Json
            
            foreach($item in $allItems){
                $allAttachments = m365 spo listitem attachment list --webUrl $siteUrl --listTitle $list.Title --itemId $item.Id | ConvertFrom-Json
                
                foreach($attachment in $allAttachments){
                    $results += [pscustomobject][ordered]@{
                        itemID = $item.Id
                        itemTitle = $item.Title
                        fileName = $attachment.FileName
                        attachmentPath = $attachment.ServerRelativeUrl
                    }
                }
            }
        }
    }
    $results
    ```

Keywords

- SharePoint Online
