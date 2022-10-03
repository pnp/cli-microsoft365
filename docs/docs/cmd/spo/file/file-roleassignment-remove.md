# spo file roleassignment remove

Removes role assignment from a file.

## Usage

```sh
m365 spo file roleassignment remove [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site where the file is located

`--fileUrl [fileUrl]`
: The server-relative URL of the file from which role assignment will be removed. Specify either `fileUrl` or `fileId` but not both

`-i, --fileId [fileId]`
: The UniqueId (GUID) of the file from which role assignment will be removed. Specify either `fileUrl` or `fileId` but not both

`--principalId [principalId]`
: The SharePoint Id of the principal. It may be either a user id or group id to remove a role assignment for. Specify either `upn`, `groupName`, or `principalId`

`--upn [upn]`
: upn/email of user to remove role assignment of. Specify either `upn`, `groupName`, or `principalId`

`--groupName [groupName]`
: The group name of Azure AD or SharePoint group to remove role assignment of. Specify either `upn`, `groupName`, or `principalId`.`

`--confirm [confirm]`
: Don't prompt for confirming removing the role assignment

--8<-- "docs/cmd/_global.md"

## Examples

Remove role assignment from file with id _b2307a39-e878-458b-bc90-03bc578531d6_ located in site _https://contoso.sharepoint.com/sites/contoso-sales_ based on principal Id _2_

```sh
m365 spo file roleassignment remove --webUrl "https://contoso.sharepoint.com/sites/contoso-sales" --fileId "b2307a39-e878-458b-bc90-03bc578531d6" --principalId 2
```

Remove role assignment from file with server-relative url _/sites/contoso-sales/documents/Test1.docx_ located in site _https://contoso.sharepoint.com/sites/contoso-sales_ based on upn _user1@contoso.onmicrosoft.com_

```sh
m365 spo file roleassignment remove --webUrl "https://contoso.sharepoint.com/sites/contoso-sales" --fileUrl "/sites/contoso-sales/documents/Test1.docx" --upn "user1@contoso.onmicrosoft.com"
```

Remove role assignment from file with id _b2307a39-e878-458b-bc90-03bc578531d6_ located in site _https://contoso.sharepoint.com/sites/contoso-sales_ based on group name _saleGroup_

```sh
m365 spo file roleassignment remove --webUrl "https://contoso.sharepoint.com/sites/contoso-sales" --fileId "b2307a39-e878-458b-bc90-03bc578531d6" --groupName "saleGroup"
```
