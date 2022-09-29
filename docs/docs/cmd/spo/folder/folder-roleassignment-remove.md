# spo folder roleassignment remove

Removes a role assignment from the specified folder.

## Usage

```sh
m365 spo folder roleassignment remove [options]
```

## Options

`-u, --webUrl <webUrl>`
: The URL of the site where the folder is located.

`-f, --folderUrl <folderUrl>`
: The server- or site-relative URL of the folder.

`--principalId [principalId]`
: The SharePoint principal id. It may be either an user id or group id for which the role assignment will be removed. Specify either upn, groupName or principalId but not multiple.

`--upn [upn]`
: The upn/email of the user. Specify either upn, groupName or principalId but not multiple.

`--groupName [groupName]`
: The Azure AD or SharePoint group name. Specify either upn, groupName or principalId but not multiple.

`--confirm`
: Don't prompt for confirmation when removing the role assignment.

--8<-- "docs/cmd/\_global.md"

## Examples

Remove roleassignment from folder based on group name

```sh
m365 spo folder roleassignment remove --webUrl "https://contoso.sharepoint.com/sites/contoso-sales" --folderUrl  "/Shared Documents/FolderPermission" --groupName "saleGroup"
```

Remove the role assignment from the specified folder based on the principal id.

```sh
m365 spo folder roleassignment remove --webUrl "https://contoso.sharepoint.com/sites/contoso-sales" --folderUrl "/Shared Documents/FolderPermission" --principalId 2
```

Remove the role assignment from the specified folder based on the principal id without prompting for removal confirmation.

```sh
m365 spo folder roleassignment remove --webUrl "https://contoso.sharepoint.com/sites/contoso-sales" --folderUrl "/Shared Documents/FolderPermission" --principalId 2 --confirm
```

Remove the role assignment from the specified folder based on the upn.

```sh
m365 spo folder roleassignment remove --webUrl "https://contoso.sharepoint.com/sites/contoso-sales" --folderUrl "/Shared Documents/FolderPermission" --upn "test@contoso.onmicrosoft.com"
```
