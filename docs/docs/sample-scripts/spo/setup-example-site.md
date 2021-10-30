# Setup example site

Author: [Adam](https://github.com/Adam-it)

This script is a good starting point for a setup script to create site with some assets like columns, content types, lists, navigation etc. The given example:

- creates a site,
- adds a site column and a content type,
- adds list and modifies it's settings (add a content type to it and makes it hidden),
- adds a document library with a custom column and some folder,
- modifies the all items view of the document library,
- modifies the site navigation links

=== "PowerShell"

    ```powershell
    Write-host 'setup script example'

    Write-host 'ensure logged in'
    $m365Status = m365 status
    if ($m365Status -eq "Logged Out") {
      m365 login
    }

    Write-host 'create setup site'
    $siteRelativeUrl = 'sites/setupTestSite'
    $tenantUrl = 'https://contoso.sharepoint.com'
    $siteUrl = "$tenantUrl/$siteRelativeUrl"
    $siteTitle = 'setup test site'
    $siteType = 'CommunicationSite'
    $site = m365 spo site get --url $siteUrl --output 'json'
    $site = $site | ConvertFrom-Json
    if ($null -eq $site) {
      Write-host 'setup site does not exist, I will create it'
      m365 spo site add --url $siteUrl --type $siteType --title $siteTitle
    }
    else {
      Write-host 'setup site already exists'
    }

    Write-host 'add site column'
    $fieldName = 'Sample Text Column'
    $field = m365 spo field get --webUrl $siteUrl --fieldTitle $fieldName --output 'json'
    if ($null -eq $field) {
      Write-host 'sample site column does not exist, I will create it'
      $fieldXml = "<Field ID='{13AFECC0-2454-41F3-85E6-E194458C861C}' Type='Text' Name='SampleTextColumn' DisplayName='Sample Text Column' Indexed='FALSE' Group='Sample Columns' Required='FALSE' SourceID='{4f118c69-66e0-497c-96ff-d7855ce0713d}' StaticName='SampleTextColumn' FromBaseType='TRUE' ></Field>"
      $field = m365 spo field add --webUrl $siteUrl --xml $fieldXml --output 'json'
    }
    else {
      Write-host 'sample site column already exists'
    }
    $field = $field | ConvertFrom-Json

    Write-host 'add site content type'
    $contentTypeName = 'sample content type'
    $contentTypeGroup = 'sample content type group'
    $parentId = '0x01007926A45D687BA842B947286090B8F67D' # list item content type
    $contentType = m365 spo contenttype get --webUrl $siteUrl --id $parentId --output 'json'
    if ($null -eq $contentType) {
      Write-host 'sample site content type does not exist, I will create it'
      $contentType = m365 spo contenttype add --webUrl $siteUrl --name $contentTypeName --id $parentId --group $contentTypeGroup --output 'json'
      $contentType = m365 spo contenttype get --webUrl $siteUrl --id $parentId --output 'json'
    }
    else {
      Write-host 'sample site content type already exists'
    }
    $contentType = $contentType | ConvertFrom-Json

    Write-host 'add field to content type'
    $fieldId = $field.Id
    $contentTypeId = $contentType.StringId
    m365 spo contenttype field set --webUrl $siteUrl --contentTypeId $contentTypeId --fieldId $fieldId --required false

    Write-host 'create generic list'
    $listName = 'setup test list'
    $list = m365 spo list get --title $listName --webUrl $siteUrl --output 'json'
    if ($null -eq $list) {
      Write-host 'sample generic list does not exist, I will create it'
      $list = m365 spo list add --title $listName --baseTemplate 'GenericList' --webUrl $siteUrl --output 'json'
    }
    else {
      Write-host 'sample generic list already exists'
    }
    $list = $list | ConvertFrom-Json

    Write-host 'modify list settings to allow content types'
    m365 spo list set --webUrl $siteUrl --id $list.Id --contentTypesEnabled true

    Write-host 'add content type to list'
    $contentTypeAddedToList = m365 spo list contenttype add --webUrl $siteUrl --listId $list.Id --contentTypeId $contentTypeId

    Write-host 'make list hidden'
    m365 spo list set --webUrl $siteUrl --id $list.Id --hidden true

    Write-host 'create document lib'
    $libName = 'setup test lib'
    $lib = m365 spo list get --title $libName --webUrl $siteUrl --output 'json'
    if ($null -eq $lib) {
      Write-host 'sample document lib does not exist, I will create it'
      $lib = m365 spo list add --title $libName --baseTemplate 'DocumentLibrary' --webUrl $siteUrl --output 'json'
    }
    else {
      Write-host 'sample document lib already exists'
    }
    $lib = $lib | ConvertFrom-Json

    Write-host 'add sample column'
    $columnName = 'Sample Text Column'
    $column = m365 spo field get --webUrl $siteUrl --listUrl "Lists/$libName" --fieldTitle $columnName --output 'json'
    if ($null -eq $column) {
      Write-host 'sample column in lib does not exist, I will create it'
      $columnXml = "<Field ID='{AC827B0C-8B45-4B4F-927B-CDDC4FEEE79E}' Type='Text' Name='SampleTextColumn' DisplayName='Sample Text Column' Required='FALSE' SourceID='http://schemas.microsoft.com/sharepoint/v3' StaticName='SampleTextColumn' FromBaseType='TRUE' />"
      $column = m365 spo field add --webUrl $siteUrl --listTitle $libName --xml $columnXml --output 'json'
    }
    else {
      Write-host 'sample column in lib already exists'
    }
    $column = $column | ConvertFrom-Json

    Write-host 'add sample folder'
    $folderName = 'sample Folder'
    $folder = m365 spo folder get --webUrl $siteUrl --folderUrl "/$libName/$folderName" --output 'json'
    if ($null -eq $folder) {
      Write-host 'sample folder in lib does not exist, I will create it'
      $folder = m365 spo folder add --webUrl $siteUrl --parentFolderUrl "/$libName" --name $folderName --output 'json'
    }
    else {
      Write-host 'sample folder in lib already exists'
    }

    Write-host 'modify list view'
    $views = m365 spo list view list --webUrl $siteUrl --listTitle $libName --output 'json'
    $views = $views | ConvertFrom-Json
    $viewName = $views[0].Title # all items view
    m365 spo list view field add --webUrl $siteUrl --listTitle $libName --viewTitle $viewName --fieldTitle $columnName

    Write-host 'modify site navigation'
    $currentNavigation = m365 spo navigation node list --webUrl $siteUrl --location QuickLaunch --output 'json'
    $currentNavigation = $currentNavigation | ConvertFrom-Json
    Write-host 'clearing old navigation links'
    foreach ($navigationItem in $currentNavigation) {
      m365 spo navigation node remove --webUrl $siteUrl --location QuickLaunch --id $navigationItem.Id --confirm
    }
    Write-host 'adding new navigation'
    $nodeAddedResponse = m365 spo navigation node add --webUrl $siteUrl --location QuickLaunch --title 'Sample Document Library' --url "/$siteRelativeUrl/$libName/Forms/AllItems.aspx"
    $nodeAddedResponse = m365 spo navigation node add --webUrl $siteUrl --location QuickLaunch --title 'Hidden Sample List' --url "/$siteRelativeUrl/Lists/$listName/AllItems.aspx"
    ```

Keywords:

- SharePoint Online
- Create Site
- Create Field
- Create Content type
- Create List
- Modify View
- Modify Navigation
