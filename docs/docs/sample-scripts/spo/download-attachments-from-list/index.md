---
tags:
  - attachments
  - download file
  - migration
---

# Download attachments from a SharePoint Online list

Download attachments from a SharePoint Online list.

=== "PowerShell"

    ```powershell
   param
(
    [Parameter(Mandatory = $true)] [string] $SiteURL,
    [Parameter(Mandatory = $true)] [string] $ListTitle,
    [Parameter(Mandatory = $true)] [string] $DownloadDirectory
)   
 
$m365Status = m365 status
if ($m365Status -match "Logged Out") {
  Write-Host "Logging in the User!"
  m365 login --authType browser
}

    Try {
 
        #Get All Items from the List
        $ListItems = m365 spo listitem list --webUrl $SiteURL --listTitle $ListTitle | ConvertFrom-Json
         
        #Create download directory if it doesn't exist
        If (!(Test-Path -path $DownloadDirectory)) {           
            New-Item $DownloadDirectory -type directory         
        }
         
        #Iterate through each list item
        Foreach ($Item in $ListItems) {
 
            #Get All attachments from the List Item
            $Attachments = m365 spo listitem attachment list --webUrl $SiteURL --listTitle $ListTitle --itemId $Item.Id | ConvertFrom-Json
            foreach ($Attachment in $Attachments) {
                $TargetFilePath = "$($DownloadDirectory)/$($Item.Id)_$($Attachment.FileName)"
                #Download attachment
                m365 spo file get --webUrl $SiteURL --url $Attachment.ServerRelativeUrl --asFile --path $TargetFilePath
            }
        }
 
        write-host  -f Green "List Attachments Downloaded Successfully!"
    }
    Catch {
        write-host -f Red "Error Downloading List Attachments!" $_.Exception.Message
    }

    ```
