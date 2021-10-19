# Add users to the Associated SharePoint Groups of a site from a CSV File

Author: [Arjun Menon](https://arjunumenon.com/add-multiple-users-sharepoint-groups-site/)

## Script Details

This is a script which adds multiple users to associated SharePoint groups (Owners, Members, Visitors) of a site from a CSV file.

Typical use case of the script will be a migration scenario where the contents are migrated, while the permissions are not migrated. In that case you will have a  list of users and its equivalent permission level of the source system. The script sample will read the source permission details from a CSV file and will add the users to the associated SharePoint groups of a site

### File Format

Below is an example of the format needed for your .csv file:

```text
Username,PermissionLevel
adelev@contoso.com,Read
alexw@contoso.com,Owner
alland@contoso.com,Member
christiec@contoso.com,Member
```

### Permission Level Mapping

Permission Level mapping assumptions are given below

| Permission Level | Equivalent SharePoint Group | Details |
| --------| ---------- | ---------- |
| Read | Visitors | User will be added to the associated Visitors group of the site
| Owner | Owners | User will be added to the associated Owners group of the site
| Member | Member |User will be added to the associated Owners group of the site

## Complete Script

```powershell tab="PowerShell"
#Check the M365 login status for CLI
$LoginStatus = m365 status
if($LoginStatus -Match "Logged out"){
    #Executing login command for CLI
    m365 login   
}

#Set the URL of the site where the users need to be added
$siteURL = "https://aum365.sharepoint.com/sites/M365CLI"

#Getting the Associated Groups for the specific site
$SiteInformation = m365 spo web get --webUrl $siteURL --withGroups --output json | ConvertFrom-Json

#Importing the Current permission list from CSV. Adding the equivalent SharePoint Group Id to the imported object.
#Object will be grouped with multiple users with , seperator since CLI supports adding multiple users in a single command
$GroupedResult = Import-Csv -Path .\Current-Permission-Migration.csv | Group-Object PermissionLevel | ForEach-Object {
[PsCustomObject]@{
    PermissionLevel = $_.Name
    UsernameValues = $_.Group.Username -join ', '
    SPGroupId = switch ($_.Name){
        "Read" {"$($SiteInformation.AssociatedVisitorGroup.Id)"}#Adding to the default Visitor's Group
        "Member" {"$($SiteInformation.AssociatedMemberGroup.Id)"}#Adding to the default Member's Group
        "Owner" {"$($SiteInformation.AssociatedOwnerGroup.Id)"}#Adding to the default Owner's Group
        default {"$($SiteInformation.AssociatedVisitorGroup.Id)"}
        }
    }
}

#Show the Formatted data table for reference
$GroupedResult | Format-Table

#Read Grouped Permission level and users and add the users to the SharePoint Groups
Foreach ($PermissionLevel in $GroupedResult) {
    Write-Host "Adding $($PermissionLevel.PermissionLevel) Permission users to the SharePoint Group ID: $($PermissionLevel.SPGroupId)"
    #Since the command supports multiple usernames to be added in the single command, script will add users in single command execution
    m365 spo group user add --webUrl $siteURL --groupId $PermissionLevel.SPGroupId --userName $PermissionLevel.UsernameValues
}
```

Keywords

- SharePoint Online
- SharePoint Group
