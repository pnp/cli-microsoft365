# Add a Site Collection Admin using a csv file

!!! warning
    When you decide to add users for all your Site Collections, you can't simply run through all the sites in your tenant. Why? Because what you "see" in the SPO admin center, is not reflecting what you _really_ have. There a few Site Collections that are not visible. A few examples would be:

    - Search
    - My Site (contoso-my.sharepoint.com)
    - Sites created with Private Channels

    Sites created with Private Channels are not (yet?) visible in the SharePoint Online admin center, so adding users as Site Collection Admins into those could cause chaos! It's safer to get a list of all your sites, and keep the intended ones in your csv file to feed the script.

## Get all the Site Collections in your tenant

To get all the Site Collections in your tenant and export to a .csv file, you can run the following:

```powershell tab="PowerShell Core"
$allSites = m365 spo site classic list --query "[?Template!='SRCHCEN#0']" -o json | ConvertFrom-Json
$results = @()

foreach($site in $allSites){
    $results += [pscustomobject][ordered]@{
        Title = $site.Title
        Url = $site.Url
        Template = $site.Template
    }
}
$results | Export-Csv -Path "<YOUR-CSV-FILE-PATH>"
```

The script above has a query to ignore the _Search_ site collection by filtering with the template code. If for example you wish to only get the sites that have been created with Private Channels, you could amend your query as follows:

```powershell tab="PowerShell Core"
$privateChannelSites = m365 spo site classic list --query "[?Template=='TEAMCHANNEL#0']" -o json | ConvertFrom-Json
```

## Add the user as Site Collection Admin

Once you've got the .csv file from the script above, filter it to your needs to keep only the targeted sites, and use it in the script below.

!!! note
    The script will add the user as a "site admin" on classic and non group-connected sites, or a an "additional admin" in group-connected sites (and not as a group Member).

```powershell tab="PowerShell Core"
$csvSites = Import-Csv -Path "<YOUR-CSV-FILE-PATH>"
$UserToAdd = "john.smith@contoso.com"  ## Change to your user
$siteCount = $csvSites.Count

Write-Host "Processing $siteCount sites..." -f Cyan

#Loop through the sites in the csv file
foreach($site in $csvSites){
    Write-Host "Going through $($site.Title)" 
    
    $users = m365 spo user list --webUrl $site.Url -o json | ConvertFrom-Json
    $admins = $users.value | Where-Object {$_.IsSiteAdmin -eq $true}
        
        if ($admins.Email -eq $UserToAdd) {
            Write-Host "User $($UserToAdd) is already an Admin in $($site.Title)." -f Green
        }
        else{
            Write-Host "Adding $($UserToAdd) to $($site.Title). " -f Magenta
            m365 spo site classic set --url $site.Url --owners $UserToAdd
        }
}
```

Keywords

- SharePoint Online
- Governance
