# Hide SharePoint list from Site Contents

Author: [David Ramalho](https://sharepoint-tricks.com/tenant-app-catalog-vs-site-collection-app-catalog/)


When you just want to deploy certain SharePoint solution to a specific site, it's required to create an app catalog for that site, the below script will create it for the site. On the article link above you can check where you can use App catalog for the site instead of global app catalog.

```powershell tab="PowerShell Core"

spo login https://contoso-admin.sharepoint.com
spo site appcatalog add --url https://contoso.sharepoint/sites/site

```

Keywords:

- SharePoint Online
- Create App Catalog for site