# spo listitem roleassignment add

Adds a role assignment to a listitem.

## Usage

```sh
m365 spo listitem roleassignment add [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site where the listitem is located

`--listId [listId]`
: ID of the list. Specify either listId, listTitle or listUrl but not multiple.

`--listTitle [listTitle]`
: Title of the list. Specify either listId, listTitle or listUrl but not multiple.

`--listUrl [listUrl]`
: Relative URL of the list. Specify either listId, listTitle or listUrl but not multiple.

`--listItemId <listItemId>`
: Id of the listitem to assign the role to.

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

Add role assignment to specified listitem in specified list located in specified site for specified principal id and specified role definition id.

```sh
m365 spo listitem roleassignment add --webUrl "https://contoso.sharepoint.com/sites/project-x" --listTitle "someList" --listItemId 1 --principalId 11 --roleDefinitionId 1073741829
```

Add role assignment to specified listitem in specified list located in specified site for specified principal id and specified role definition id.

```sh
m365 spo listitem roleassignment add --webUrl "https://contoso.sharepoint.com/sites/project-x" --listId "0CD891EF-AFCE-4E55-B836-FCE03286CCCF" --listItemId 1 --principalId 11 --roleDefinitionId 1073741829
```

Add role assignment to specified listitem in specified list located in specified site for specified principal id and specified role definition id.

```sh
m365 spo listitem roleassignment add --webUrl "https://contoso.sharepoint.com/sites/project-x" --listUrl "sites/documents" --listItemId 1 --principalId 11 --roleDefinitionId 1073741829
```

Add role assignment to specified listitem in specified list located in specified site for specified upn and specified role definition id.

```sh
m365 spo listitem roleassignment add --webUrl "https://contoso.sharepoint.com/sites/project-x" --listTitle "someList" --listItemId 1 --upn "someaccount@tenant.onmicrosoft.com" --roleDefinitionId 1073741829
```

Add role assignment to specified listitem in specified list located in specified site for specified group and specified role definition id.

```sh
m365 spo listitem roleassignment add --webUrl "https://contoso.sharepoint.com/sites/project-x" --listTitle "someList" --listItemId 1 --groupName "someGroup" --roleDefinitionId 1073741829
```

Add role assignment to specified listitem in specified list located in specified site for specified principal id and specified role definition name.

```sh
m365 spo listitem roleassignment add --webUrl "https://contoso.sharepoint.com/sites/project-x" --listTitle "someList" --listItemId 1 --principalId 11 --roleDefinitionName "Full Control"
```

## Response

The command won't return a response on success.
