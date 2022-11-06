# spo folder roleassignment add

Adds a role assignment from the specified folder.

## Usage

```sh
m365 spo folder roleassignment add [options]
```

## Options

`-u, --webUrl <webUrl>`
: The URL of the site where the folder is located.

`-f, --folderUrl <folderUrl>`
: The server- or site-relative URL of the folder.

`--principalId [principalId]`
: The SharePoint principal id. It may be either an user id or group id for which the role assignment will be addd. Specify either upn, groupName or principalId but not multiple.

`--upn [upn]`
: The upn/email of the user. Specify either upn, groupName or principalId but not multiple.

`--groupName [groupName]`
: The Azure AD or SharePoint group name. Specify either upn, groupName or principalId but not multiple.

`--roleDefinitionId [roleDefinitionId]`
: ID of the role definition. Specify either roleDefinitionId or roleDefinitionName but not both.

`--roleDefinitionName [roleDefinitionName]`
: The name of the role definition. E.g. 'Contribute', 'Read'. Specify either roleDefinitionId or roleDefinitionName but not both

--8<-- "docs/cmd/_global.md"

## Examples

Add the role assignment to the specified folder based on the group name and role definition name.

```sh
m365 spo folder roleassignment add --webUrl "https://contoso.sharepoint.com/sites/contoso-sales" --folderUrl  "/Shared Documents/FolderPermission" --groupName "saleGroup" --roleDefinitionName "Edit"
```

Add the role assignment to the specified folder based on the principal Id and role definition id

```sh
m365 spo folder roleassignment add --webUrl "https://contoso.sharepoint.com/sites/contoso-sales" --folderUrl "/Shared Documents/FolderPermission" --principalId 2 --roleDefinitionId 1073741827 
```

Add the role assignment to the specified folder based on the upn and role definition name

```sh
m365 spo folder roleassignment add --webUrl "https://contoso.sharepoint.com/sites/contoso-sales" --folderUrl "/Shared Documents/FolderPermission" --upn "test@contoso.onmicrosoft.com" --roleDefinitionName "Edit"
```
Add the role assignment to the root folder based on the upn and role definition name

```sh
m365 spo folder roleassignment add --webUrl "https://contoso.sharepoint.com/sites/contoso-sales" --folderUrl "/Shared Documents" --upn "test@contoso.onmicrosoft.com" --roleDefinitionName "Edit"
```
