# How to perform operations if a command is not covered by the CLI for Microsoft 365

Author: [Joseph Velliah](https://blog.josephvelliah.com/spol-download-attachments-from-list-items-using-cli-for-microsoft-365)

One of the most powerful tools a Microsoft 365 user has is the CLI for Microsoft 365. The command line allows any user to get a lot of things done in a fast way. There is no boundary to the number of things a seasoned user can do by merely using the CLI for Microsoft 365.

This script shows how to perform operations if a command is not covered by the CLI for Microsoft 365.

Right now, AttachmentFiles property associated with a SharePoint list item is not available in CLI for Microsoft 365, so we need to execute a separate query to ```/_api/web/lists/getByTitle('list-title')/items(item-id)/AttachmentFiles``` endpoint to get the item attachments.

To call AttachmentFiles endpoint, we must acquire an access token from the Microsoft identity platform. To do this we can use ```m365 util accesstoken get``` command and attach the access token with AttachmentFiles endpoint as shown in this script.

Prerequisites:

- [CLI for Microsoft 365](https://pnp.github.io/cli-microsoft365/)
- SharePoint Online site with list item attachments

=== "PowerShell"

    ```powershell
    Function Get-ListAttachments() {
        param
        (
            [Parameter(Mandatory = $true)] [string] $AccessToken,
            [Parameter(Mandatory = $true)] [string] $SiteURL,
            [Parameter(Mandatory = $true)] [string] $ListTitle,
            [Parameter(Mandatory = $true)] [int] $ItemId
        )   
        Try {
            $ListItemAttachmentsEndPoint = "$($SiteURL)/_api/web/lists/getbytitle('$($ListTitle)')/items($($ItemId))/AttachmentFiles"
            $Header = @{
                "Authorization" = "Bearer $($AccessToken)"
                "Accept"        = "application/json; odata=verbose" 
                "Content-Type"  = "application/json "
            }
            $ListItemAttachments = Invoke-RestMethod -Uri $ListItemAttachmentsEndPoint -Headers $Header -Method Get  
            return $ListItemAttachments.d.results
        }
        Catch {
            throw "Error Getting List Item Attachments! $($_.Exception.Message)" 
        }
    }
    Function Download-ListAttachments() {
        param
        (
            [Parameter(Mandatory = $true)] [string] $TenantName,
            [Parameter(Mandatory = $true)] [string] $SiteURL,
            [Parameter(Mandatory = $true)] [string] $ListTitle,
            [Parameter(Mandatory = $true)] [string] $DownloadDirectory
        )   
        Try {
    
            #Get All Items from the List
            $ListItems = m365 spo listitem list --webUrl $SiteURL --title $ListTitle -o json | ConvertFrom-Json -AsHashtable
             
            #Iterate through each list item
            Foreach ($Item in $ListItems) {
                Try {
                    Write-Output "Processing Item Id $($Item.Id)"
    
                    # Right now AttachmentFiles property is not available in cli-microsoft365 so we need to execute a separate query to /_api/web/lists/getByTitle('list-title')/items(item-id)/AttachmentFiles to get the item attachments. 
                    # AttachmentFiles endpoint requires access token 
                    $AccessToken = m365 util accesstoken get --resource "https://$($TenantName).sharepoint.com" --new 
    
                    #Get All attachments from the List Item
                    $Attachments = Get-ListAttachments -AccessToken $AccessToken -SiteURL $SiteURL -ListTitle $ListTitle -ItemId $Item.Id
    
                    If ($Attachments.Length -gt 0) {
                        #Create directory for each list item if it doesn't exist
                        $TargetDownloadDirectory = "$($DownloadDirectory)/$($Item.Id)"
                        If (!(Test-Path -path $TargetDownloadDirectory)) { New-Item $TargetDownloadDirectory -type Directory | Out-Null }
    
                        foreach ($Attachment in $Attachments) {
                            Try {
                                Write-Output "Downloading $($Attachment.FileName)"
                                $TargetFilePath = "$($TargetDownloadDirectory)/$($Attachment.FileName)"
                                #Download attachment
                                m365 spo file get --webUrl $SiteURL --url $Attachment.ServerRelativeUrl --asFile --path $TargetFilePath
                            }
                            Catch {
                                Write-Error "Error Downloading This Attachment! $($_.Exception.Message)" 
                            }
                        }
                    }
                    else {
                        Write-Warning "Attachments Not Found For This List Item!"
                    }
                }
                Catch {
                    Write-Error "Error Downloading This List Item Attachments! $($_.Exception.Message)"
                }
            }
        }
        Catch {
            Write-Error "Error Downloading List Attachments! $($_.Exception.Message)"
        }
    }
    
    #Set Parameters
    $TenantName = "tenant-name"
    $SiteRelativePath = "site-relative-path"
    $ListTitle = "list-title"
    
    $DownloadDirectory = "$($PSScriptRoot)/$($ListTitle)"
    $SiteURL = "https://$($TenantName).sharepoint.com/$($SiteRelativePath)"
    
    #Call the function to download list items attachments
    Download-ListAttachments -TenantName $TenantName -SiteURL $SiteURL -ListTitle $ListTitle -DownloadDirectory $DownloadDirectory
    ```

Keywords:

- CLI for Microsoft 365
- SharePoint Online
