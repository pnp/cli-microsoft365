# spo file roleassignment add

Adds a role assignment to the specified file.

## Usage

```sh
m365 spo file roleassignment add [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site where the file is located

`--fileUrl [fileUrl]`
: The server-relative URL of the file to retrieve. Specify either `fileUrl` or `fileId` but not both

`i, --fileId [fileId]`
: The UniqueId (GUID) of the file to retrieve. Specify either `fileUrl` or `fileId` but not both

`--principalId [principalId]`
: The SharePoint Id of the principal. It may be either a user id or group id to add a role assignment for. Specify either upn, groupName or principalId.

`--upn [upn]`
: upn/email of user to assign role to. Specify either upn, groupName or principalId.

`--groupName [groupName]`
: The group name of Azure AD or SharePoint group. Specify either upn, groupName or principalId.

`--roleDefinitionId [roleDefinitionId]`
: ID of role definition. Specify either roleDefinitionId or roleDefinitionName but not both

`--roleDefinitionName [roleDefinitionName]`
: Enter the name of a role definition, like 'Contribute', 'Read', etc. Specify either roleDefinitionId or roleDefinitionName but not both

--8<-- "docs/cmd/_global.md"

## Examples

Add role assignment to file with id _b2307a39-e878-458b-bc90-03bc578531d6_ in site _https://contoso.sharepoint.com/sites/project-x_ for a principal with id _11_ and role definition id _1073741829_.

```sh
m365 spo file roleassignment add --webUrl "https://contoso.sharepoint.com/sites/project-x" --fileId "b2307a39-e878-458b-bc90-03bc578531d6" --principalId 11 --roleDefinitionId 1073741829
```

Add role assignment to file with id _b2307a39-e878-458b-bc90-03bc578531d6_ in site _https://contoso.sharepoint.com/sites/project-x_ for upn _testuser@tenant.onmicrosoft.com_ and role definition name _Read_.

```sh
m365 spo file roleassignment add --webUrl "https://contoso.sharepoint.com/sites/project-x" --fileId "b2307a39-e878-458b-bc90-03bc578531d6" --upn "testuser@tenant.onmicrosoft.com" --roleDefinitionName "Read"
```

Add role assignment to file with server-relative url _/sites/project-x/documents/Test1.docx_ in site _https://contoso.sharepoint.com/sites/project-x_ for group _demoGroup__ and role definition name _Read_.

```sh
m365 spo file roleassignment add --webUrl "https://contoso.sharepoint.com/sites/project-x" --fileUrl "/sites/project-x/documents/Test1.docx" --upn "testuser@tenant.onmicrosoft.com" --roleDefinitionName "Read"
```