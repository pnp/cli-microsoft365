# Add App Catalog to SharePoint site

Author: [David Ramalho](https://sharepoint-tricks.com/tenant-app-catalog-vs-site-collection-app-catalog/)

When you just want to deploy certain SharePoint solution to a specific site, it's required to create an app catalog for that site. The below script will create it for the site. In the article referenced above you can check where you can use App catalog for the site instead of global app catalog.

```powershell tab="PowerShell Core"
$site = "https://contoso.sharepoint.com/sites/site"
m365 login
m365 spo site appcatalog add --url $site
Write-output "App Catalog Created on " $site
```

```bash tab="Bash"
#!/bin/bash

site=https://tricks365.sharepoint.com/sites/site

m365 login
m365 spo site appcatalog add --url $site
echo "App Catalog Created on $site"
```

Keywords:

- SharePoint Online
- Create App Catalog for site