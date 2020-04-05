# Disable specified Tenant-wide Extension

Author: [Shantha Kumar T](https://www.ktskumar.com/2020/04/manage-tenant-wide-extensions-using-office-365-cli/)

Tenant Wide Extensions list from the App Catalog helps to manage the activation / deactivation of the tenant wide extensions. The below sample script helps to disable the specifed tenant wide extension based on the id parameter.

Note: TenantWideExtensionDisabled column denotes the extension is enabled or disabled.


```powershell tab="PowerShell Core"
$listName = "Tenant Wide Extensions" 

o365 login
$appcatalogurl = o365 spo tenant appcatalogurl get
o365 spo listitem set -t $listName -i 2 -u $appcatalogurl --TenantWideExtensionDisabled "true"
```

```bash tab="Bash"
#!/bin/bash

listName="Tenant Wide Extensions"

o365 login
appcatalogurl=$(o365 spo tenant appcatalogurl get)
o365 spo listitem set -t "$listName" -i 2 -u $appcatalogurl --TenantWideExtensionDisabled "true"
```



Keywords:

- SharePoint Online
- Tenant Wide Extension
