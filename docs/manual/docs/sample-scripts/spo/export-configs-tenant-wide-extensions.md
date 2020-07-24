# Export Configurations of Tenant Wide Extensions

Author: [Joseph Velliah](https://sprider.blog/export-configs-tenant-wide-extensions)

The SharePoint Admin Center provides various governance features, but there is no way to easily export Configurations of Tenant Wide Extensions from the SharePoint admin center for governance activities. This script retrieves Tenant Wide Extension configurations from the App Catalog and exports the same in a comma-separated values (CSV) file.

```powershell tab="PowerShell Core"
$resultDir = "Output"
$listName = "Tenant Wide Extensions"
$fields = $fields = "Title, Modified, Created, AuthorId, EditorId, TenantWideExtensionComponentId, TenantWideExtensionComponentProperties, TenantWideExtensionListTemplate, TenantWideExtensionWebTemplate, TenantWideExtensionSequence, TenantWideExtensionHostProperties, TenantWideExtensionLocation, TenantWideExtensionDisabled"

$executionDir = $PSScriptRoot
$outputDir = "$executionDir/$resultDir"
$outputFilePath = "$outputDir/$(get-date -f yyyyMMdd-HHmmss)-tenantwideextensions.csv"

if (-not (Test-Path -Path "$outputDir" -PathType Container)) {
    Write-Host "Creating $outputDir folder..."
    New-Item -ItemType Directory -Path "$outputDir"
}

$appCatalogUrl = o365 spo tenant appcatalogurl get

if ($appCatalogUrl) {
    $spolItems = o365 spo listitem list --title $listName --webUrl $appCatalogUrl --fields $fields  -o json | ConvertFrom-Json

    if ($spolItems.Count -gt 0) {
        $configurations = @()

        foreach ($spolItem in $spolItems) {
            $author = o365 spo user get --webUrl $appCatalogUrl --id $spolItem.AuthorId -o json | ConvertFrom-Json
            $editor = o365 spo user get --webUrl $appCatalogUrl --id $spolItem.EditorId -o json | ConvertFrom-Json

            $configurationObject = New-Object -TypeName PSObject

            $configurationObject | Add-Member -MemberType NoteProperty -Name "Title" -Value $spolItem.Title
            $configurationObject | Add-Member -MemberType NoteProperty -Name "Modified" -Value $spolItem.Modified
            $configurationObject | Add-Member -MemberType NoteProperty -Name "Created" -Value $spolItem.Created
            $configurationObject | Add-Member -MemberType NoteProperty -Name "Author" -Value $author.Title
            $configurationObject | Add-Member -MemberType NoteProperty -Name "Editor" -Value $editor.Title
            $configurationObject | Add-Member -MemberType NoteProperty -Name "TenantWideExtensionComponentId" -Value $spolItem.TenantWideExtensionComponentId
            $configurationObject | Add-Member -MemberType NoteProperty -Name "TenantWideExtensionComponentProperties" -Value $spolItem.TenantWideExtensionComponentProperties
            $configurationObject | Add-Member -MemberType NoteProperty -Name "TenantWideExtensionListTemplate" -Value $spolItem.TenantWideExtensionListTemplate
            $configurationObject | Add-Member -MemberType NoteProperty -Name "TenantWideExtensionWebTemplate" -Value $spolItem.TenantWideExtensionWebTemplate
            $configurationObject | Add-Member -MemberType NoteProperty -Name "TenantWideExtensionSequence" -Value $spolItem.TenantWideExtensionSequence
            $configurationObject | Add-Member -MemberType NoteProperty -Name "TenantWideExtensionHostProperties" -Value $spolItem.TenantWideExtensionHostProperties
            $configurationObject | Add-Member -MemberType NoteProperty -Name "TenantWideExtensionLocation" -Value $spolItem.TenantWideExtensionLocation
            $configurationObject | Add-Member -MemberType NoteProperty -Name "TenantWideExtensionDisabled" -Value $spolItem.TenantWideExtensionDisabled

            $configurations += $configurationObject
        }

        $configurations | Export-Csv -Path "$outputFilePath" -NoTypeInformation
        Write-Host "Open $outputFilePath to review Tenant Wide Extensions configurations report."
    }
    else {
        Write-Host "Tenant Wide Extensions list is empty."
    }
}
else {
    Write-Host "Unable to get App Catalog Url."
}
```

Keywords:

- SharePoint Online
- SharePoint Framework Extensions
- Governance
