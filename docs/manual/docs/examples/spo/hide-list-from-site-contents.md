# Hide SharePoint list from Site Contents

Author: [David Ramalho](https://sharepoint-tricks.com/hide-sharepoint-list-from-site-contents/)

If you need to hide the SharePoint list from the UI this simple PowerShell will hide a specific list from the site Conents. This will avoid some users to access the list while for example, you still are setup the list or it is not ready to be used. 

```powershell tab="PowerShell Core"

$listName = "listName"
$site = "https://contoso.sharepoint.com/"

o365 login
$list = o365 spo list get --webUrl $site -t $listName -o json | ConvertFrom-Json
o365 spo list set --webUrl $site -i $list.Id -t $listName --hidden true 

```

Keywords:

- SharePoint Online
- Hide List