# Hide SharePoint list from Site Contents

Author: [David Ramalho](https://sharepoint-tricks.com/hide-sharepoint-list-from-site-contents/)

If you need to hide the SharePoint list from the UI this simple PowerShell script will hide a specific list from the site contents. This will prevent users from easily accessing the list while, for example, you are still setting it up.

```powershell tab="PowerShell Core"
$listName = "listName"
$site = "https://contoso.sharepoint.com/"

o365 login
$list = o365 spo list get --webUrl $site -t $listName -o json | ConvertFrom-Json
o365 spo list set --webUrl $site -i $list.Id -t $listName --hidden true
```

```bash tab="Bash"
#!/bin/bash

# requires jq: https://stedolan.github.io/jq/

listName="listName"
site=https://contoso.sharepoint.com/

o365 login
listId=$(o365 spo list get --webUrl $site -t "$listName" -o json | jq ".Id")
o365 spo list set --webUrl $site -i $listId -t $listName --hidden true
```

Keywords:

- SharePoint Online
- Hide List
