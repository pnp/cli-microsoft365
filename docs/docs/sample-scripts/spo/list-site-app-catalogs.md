# Lists active SharePoint site collection application catalogs

Inspired by: [David Ramalho](http://sharepoint-tricks.com/check-all-sharepoint-sites-collection-with-app-catalog-active/)

A sample that shows how to find all installed site collection application catalogs within a tenant. IT Professionals or DevOps can benefit from it when they govern tenants or scan tenant for customizations. Pulling a list with site collection app catalogs can give them valuable information at what scale the tenant site collections are customized. The sample outputs the URL of the site collection, and this can help IT Pros or DevOps to dig deeper and find out what and how many solution packages a site collection app catalog has installed. Check for un-healthy solution packages or such that could be a security risk.

Note, because the sample uses the SharePoint search API to identify the site collection application catalogs, a newly created one might not be indexed right away. The sample output would not list the newly created app catalog until the search crawler indexes it; this usually does not take longer than a few minutes.

```powershell tab="PowerShell Core"
$appCatalogs = m365 spo search --query "contentclass:STS_List_336" --selectProperties SPSiteURL --allResults --output json | ConvertFrom-Json

$appCatalogs | ForEach-Object { Write-Host $_.SPSiteURL }
Write-Host 'Total count:' $appCatalogs.Count
```

```bash tab="Bash"
#!/bin/bash

# requires jq: https://stedolan.github.io/jq/

appCatalogs=$(m365 spo search --query "contentclass:STS_List_336" --selectProperties SPSiteURL --allResults --output json)

echo $appCatalogs | jq -r '.[].SPSiteURL'
echo "Total count:" $(echo $appCatalogs | jq length)
```

Keywords:

- SharePoint Online
- Governance
- Security