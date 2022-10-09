# spo file roleassignment remove

Removes a role assignment from a file.

## Usage

```sh
m365 spo file roleassignment remove [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site where the file is located.

`--fileUrl [fileUrl]`
: The server-relative URL of the file. Specify either `fileUrl` or `fileId` but not both.

`-i, --fileId [fileId]`
: The UniqueId (GUID) of the file. Specify either `fileUrl` or `fileId` but not both.

`--principalId [principalId]`
: The SharePoint Id of the principal. It may be either a user id or group id. Specify either `upn`, `groupName`, or `principalId`.

`--upn [upn]`
: Upn/email of the user. Specify either `upn`, `groupName`, or `principalId`.

`--groupName [groupName]`
: The group name of an Azure AD or SharePoint group. Specify either `upn`, `groupName`, or `principalId`.

`--confirm [confirm]`
: Don't prompt for confirmation.

--8<-- "docs/cmd/_global.md"

## Examples

Remove a role assignment by principal id from a file by id

```sh
m365 spo file roleassignment remove --webUrl "https://contoso.sharepoint.com/sites/contoso-sales" --fileId "b2307a39-e878-458b-bc90-03bc578531d6" --principalId 2
```

Remove a role assignment by upn from a file by url

```sh
m365 spo file roleassignment remove --webUrl "https://contoso.sharepoint.com/sites/contoso-sales" --fileUrl "/sites/contoso-sales/documents/Test1.docx" --upn "user1@contoso.onmicrosoft.com"
```

Remove a role assignment by group name from a file by id

```sh
m365 spo file roleassignment remove --webUrl "https://contoso.sharepoint.com/sites/contoso-sales" --fileId "b2307a39-e878-458b-bc90-03bc578531d6" --groupName "saleGroup"
```
