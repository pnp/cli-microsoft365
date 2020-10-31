# Disable specified Tenant-wide Extension

Author: [Shantha Kumar T](https://www.ktskumar.com/2020/04/manage-tenant-wide-extensions-using-office-365-cli/)

Tenant Wide Extensions list from the App Catalog helps to manage the activation / deactivation of the tenant wide extensions. The below sample script helps to disable the specified tenant wide extension based on the id parameter.

Note: TenantWideExtensionDisabled column denotes the extension is enabled or disabled.

```powershell tab="PowerShell Core"
$extensionName = Read-Host "Enter the Extension Name"
$listName = "Tenant Wide Extensions"

$appCatalogUrl = m365 spo tenant appcatalogurl get
$filterQuery = "Title eq '" + $extensionName + "'"
$appItems = m365 spo listitem list --title $listName --webUrl $appCatalogUrl --fields "Id,Title" --filter $filterQuery --output json
$extItems = $appItems.Replace("Id", "ExtId") | ConvertFrom-JSON

if ($extItems.count -gt 0) {
  m365 spo listitem set --listTitle $listName --id $extItems.ExtId --webUrl $appCatalogUrl --TenantWideExtensionDisabled "true" >$null 2>&1
  Write-Host("Extension disabled.");
}
else {
  Write-Host("No extensions found with the name '" + $extensionName + "'.");
}
```

```bash tab="Bash"
#!/bin/bash

# requires jq: https://stedolan.github.io/jq/

echo "Enter the extension name to disable: "; read extensionName;
listName="Tenant Wide Extensions";

appCatalogUrl=$(m365 spo tenant appcatalogurl get)
filterQuery="Title eq '$extensionName'"
appItemsJson=$(m365 spo listitem list --title "$listName" --webUrl "$appCatalogUrl" --fields "Id,Title" --filter "$filterQuery" --output json)
appItemId=( $(jq -r '.[].Id' <<< $appItemsJson))

if [[ $appItemId -gt 0 ]]
then
 m365 spo listitem set --listTitle "$listName" --id "$appItemId" --webUrl "$appCatalogUrl" --TenantWideExtensionDisabled "true" >/dev/null 2>&1
 echo "Extension disabled."
else
  echo "No extensions found with the name '$extensionName'."
fi
```

Keywords:

- SharePoint Online
- Tenant Wide Extension
