# List all Tenant-wide Extensions

Author: [Shantha Kumar T](https://www.ktskumar.com/2020/04/manage-tenant-wide-extensions-using-office-365-cli/)

The below sample script helps to list out all the tenant-wide extensins deployed in the tenant. The sample output returns the Id, Title, Extension Location and Extension Disabled status of each extension.


```powershell tab="PowerShell Core"
$listName = "Tenant Wide Extensions" 
$fields = "Id, Title, TenantWideExtensionDisabled, TenantWideExtensionLocation"

o365 login
$appcatalogurl = o365 spo tenant appcatalogurl get
o365 spo listitem list -t $listName -u $appcatalogurl -f $fields
```

```bash tab="Bash"
#!/bin/bash

listName="Tenant Wide Extensions"
fields="Id, Title, TenantWideExtensionLocation, TenantWideExtensionDisabled"

o365 login
appcatalogurl=$(o365 spo tenant appcatalogurl get)
o365 spo listitem list -t "$listName" -u $appcatalogurl -f  "$fields"
```

Note: To view more properties of the extensions, use the internal names in fields variable.
Column	|	Internal Name	|	Description
--	|	--	|	--
Title	|	Title	|	Title of the extension.
Component Id	|	TenantWideExtensionComponentId	|	The manifest ID of the component. It has to be in GUID format and the component must exist in the App Catalog.
Component Properties	|	TenantWideExtensionComponentProperties	|	component properties.
Web Template	|	TenantWideExtensionWebTemplate	|	It can be used to target extension only to a specific web template.
List template	|	TenantWideExtensionListTemplate	|	List type as a number.
Location	|	TenantWideExtensionLocation	|	Location of the extension. There are different support locations for application customizers and ListView Command Sets.
Sequence	|	TenantWideExtensionSequence	|	The sequence of the extension in rendering.
Host Properties	|	TenantWideExtensionHostProperties	|	Additional server-side configuration, like pre-allocated height for placeholders.
Disabled	|	TenantWideExtensionDisabled	|	Is the extension enabled or disabled?

Keywords:

- SharePoint Online
- Tenant Wide Extension
