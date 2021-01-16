# List all application customizers in a tenant

Author: [Rabia Williams](https://twitter.com/williamsrabia)

List all the application customizers in a tenant. Scope is default `All`. Here we are using the [custom action list](https://pnp.github.io/cli-microsoft365/cmd/spo/customaction/customaction-list/) command to list out all the Application Customizers in all the sites in the tenant.

```powershell tab="PowerShell Core"
$sites = m365 spo search --queryText "contentclass:STS_site -SPSiteURL:personal" --selectProperties "Path,Title" --allResults --output json | ConvertFrom-Json
foreach ($site in $sites) {                                                      
  write-host $site.Title                      
  write-host $site.Path                                             
  m365 spo customaction list --url $site.Path   
} 
```

```bash tab="Bash"
#!/bin/bash

# requires jq: https://stedolan.github.io/jq/

defaultIFS=$IFS
IFS=$'\n'

sites=$(m365 spo search --queryText "contentclass:STS_site -SPSiteURL:personal" --selectProperties "Path,Title" --allResults --output json)

for site in $(echo $sites | jq -c '.[]'); do
  siteUrl=$(echo ${site} | jq -r '.Path')
  siteName=$(echo ${site} | jq -r '.Title')
  echo $siteUrl
  echo $siteName
  m365 spo customaction list --url $siteUrl
done
```

Keywords:

- SharePoint Online
- Governance
