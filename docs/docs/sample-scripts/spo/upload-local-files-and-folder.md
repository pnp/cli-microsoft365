# Upload local files and folders to SharePoint Online

Author: [Patrick Lamber](https://github.com/plamber), [Adam](https://github.com/Adam-it)

This script shows how you can use the CLI to upload files located on a local folder to a SharePoint Online library or subfolder. This is a simple script that could be used for simple data migration scenarios. The given example uploads to the given site to Shared Documents library all files and sub folders of ./import local folder

```powershell tab="PowerShell Core"
Write-host 'upload files and folders from directory example'

function Import-FilesAndFolders(
  [Parameter(Mandatory = $True)][string] $folderPath,
  [Parameter(Mandatory = $True)][string] $sPFolderPath,
  [Parameter(Mandatory = $True)][string] $siteUrl) {
    $items = Get-ChildItem -Path $folderPath
    foreach ($item in $items) {
        if ((Get-Item $item.FullName) -is [System.IO.DirectoryInfo]) {
          Write-host "creating folder $item"
          $folderCreated = m365 spo folder add --webUrl $siteUrl --parentFolderUrl $sPFolderPath --name $item.Name

          Write-host "importing folder $item"
          Import-FilesAndFolders  -folderPath $item.FullName -sPFolderPath "$sPFolderPath/$item" -siteUrl $siteUrl
        }
        else {
          Write-host "importing file $item"

          m365 spo file add --webUrl $siteUrl --folder $sPFolderPath --path $item.FullName
        }
    }
}

Write-host 'ensure logged in'
$m365Status = m365 status
if ($m365Status -eq "Logged Out") {
    m365 login
}

$importFolderPath = './import'
$sPFolderPath = '/Shared Documents'
$siteUrl = 'https://contoso.sharepoint.com/sites/TestFileImport'
Import-FilesAndFolders -folderPath $importFolderPath -sPFolderPath $sPFolderPath -siteUrl $siteUrl

```

Keywords:

- SharePoint Migration