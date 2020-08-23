# Insert pictures in a SharePoint Document Library into a Word document

Author: [Joseph Velliah](https://sprider.blog/insert-pictures-in-a-sharepoint-document-library-into-a-word-document)

This PowerShell script shows how to download and insert many pictures in a SharePoint Document Library into a Word document in a two-column table with file name using CLI for Microsoft 365 commands and PowerShell Script.

Customers have multiple pictures in a SharePoint Document Library, and they need to automatically insert the pictures in Word as it will take a lot of time if operating from UI. So, they need a script to accomplish that.

Prerequisites:

- Windows 10
- Windows PowerShell
- CLI for Microsoft 365
- Office 2007 or Higher version
- SharePoint Online Site
- Document Library with some images
- Folder to download the images
- Blank Word document to add the images

```powershell tab="PowerShell Core"
Write-Host "Execution started"

$imagesDownloadFolderPath = "C:\Users\username\Downloads\Temp\images"
$targetWordDocumentPath = "C:\Users\username\Downloads\Temp\output\word-document-name.docx"

$siteUrl = "https://tenant-name.sharepoint.com/sites/site-name"
$docLibRootFolderName = "Shared Documents"

# assumption - folder contains only images but feel free to change the filter conditions to limit the items/file types returned from document library
$spolImagesCollection = m365 spo file list --webUrl $siteUrl --folder $docLibRootFolderName -o json | ConvertFrom-Json

if ($spolImagesCollection.Count -gt 0) {
  $numberOfRows = $spolImagesCollection.Count
  $numberOfColumns = 2

  $wordClient = New-Object -comobject word.application
  $wordClient.Visible = $false
  $wordDoc = $wordClient.Documents.Add()
  $range = $wordDoc.Range()
  $wordDoc.Tables.Add($range, $numberOfRows, $numberOfColumns) | Out-Null

  $table = $wordDoc.Tables.item(1)
  $table.Cell(1, 1).Range.Text = "File Name" # column 1 heading
  $table.Cell(1, 2).Range.Text = "Image" # column 2 heading 1

  $rowNumber = 2 # to insert the images from second row

  ForEach ($spolImage in $spolImagesCollection) {
    $targetFilePath = Join-Path $imagesDownloadFolderPath $spolImage.Name
    $docServerRelativeUrl = $spolImage.ServerRelativeUrl

    Write-Host "Processing: $docServerRelativeUrl"

    m365 spo file get --webUrl $siteUrl --url $docServerRelativeUrl --asFile --path $targetFilePath
    Write-Host "File downloaded: " $docServerRelativeUrl

    $table.Cell($rowNumber, 1).Range.Text = $spolImage.Name
    $table.Cell($rowNumber, 2).Range.InlineShapes.AddPicture($targetFilePath) | Out-Null
    Write-Host "Added image in temp document table row " $rowNumber

    $rowNumber++
  }

  [ref]$saveFormat = "microsoft.office.interop.word.WdSaveFormat" -as [type]
  $wordDoc.saveas([ref] $targetWordDocumentPath, [ref]$saveFormat::wdFormatDocumentDefault)
  $wordDoc.close()
  $wordClient.quit()

  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wordDoc) | Out-Null
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wordClient) | Out-Null
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($range) | Out-Null
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($table) | Out-Null
  Remove-Variable wordDoc, wordClient, range, table
  [gc]::collect()
  [gc]::WaitForPendingFinalizers()

  Write-Host "Open the document located in $targetWordDocumentPath and check the images in the table"
}
else {
  Write-Host "No files in this document library"
}

Write-Host "Execution completed"
```

Keywords:

- SharePoint Online
- Windows
- Microsoft Word
