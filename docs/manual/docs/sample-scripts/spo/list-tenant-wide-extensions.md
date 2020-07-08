# List all tenant-wide extensions

Author: [Shantha Kumar T](https://www.ktskumar.com/2020/04/manage-tenant-wide-extensions-using-office-365-cli/)

The following script lists all tenant-wide extensions deployed in the tenant. The sample returns the Id, Title, Extension Location and Extension Disabled status of each extension.

```powershell tab="PowerShell Core"
$listName = "Tenant Wide Extensions"
$fields = "Id, Title, TenantWideExtensionDisabled, TenantWideExtensionLocation"

$appcatalogurl = m365 spo tenant appcatalogurl get
m365 spo listitem list --title $listName --webUrl $appcatalogurl --fields $fields
```

```bash tab="Bash"
#!/bin/bash

listName="Tenant Wide Extensions"
fields="Id, Title, TenantWideExtensionLocation, TenantWideExtensionDisabled"

appcatalogurl=$(m365 spo tenant appcatalogurl get)
m365 spo listitem list --title "$listName" --webUrl $appcatalogurl --fields  "$fields"
```

!!! note
    To view more/different properties of the extensions, adjust the internal names in the `fields` variable.

Column|Internal Name|Description
--|--|--
Title|Title|Title of the extension.
Component Id|TenantWideExtensionComponentId|The manifest ID of the component. It has to be in GUID format and the component must exist in the App Catalog.
Component Properties|TenantWideExtensionComponentProperties|component properties.
Web Template|TenantWideExtensionWebTemplate|It can be used to target extension only to a specific web template.
List template|TenantWideExtensionListTemplate|List type as a number.
Location|TenantWideExtensionLocation|Location of the extension. There are different support locations for application customizers and ListView Command Sets.
Sequence|TenantWideExtensionSequence|The sequence of the extension in rendering.
Host Properties|TenantWideExtensionHostProperties|Additional server-side configuration, like pre-allocated height for placeholders.
Disabled|TenantWideExtensionDisabled|Is the extension enabled or disabled?

Keywords:

- SharePoint Online
- Tenant Wide Extension
