# spo list roleassignment add

Adds a role assignment to list permissions

## Usage

```sh
m365 spo list roleassignment add [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site where the list is located

`-i, --listId [listId]`
: ID of the list. Specify either listId, listTitle or listUrl but not multiple.

`-t, --listTitle [listTitle]`
: Title of the list. Specify either listId, listTitle or listUrl but not multiple.

`--listUrl [listUrl]`
:	Relative URL of the list. Specify either listId, listTitle or listUrl but not multiple.

`--principalId [principalId]`
: SharePoint ID of principal it may be either user id or group id we want to add permissions to. Specify principalId only when upn or groupName are not used.

`--upn [upn]`
:	Upn/email of user to assign role to. Specify either upn or princpialId

`--groupName [groupName]`
:	Enter group name of Azure AD or SharePoint group.. Specify either groupName or princpialId

`--roleDefinitionId [roleDefinitionId]`
:	ID of role definition. Specify either roleDefinitionId or roleDefinitionName but not both

`--roleDefinitionName [roleDefinitionName]`
:	Enter the name of a role definition, like 'Contribute', 'Read', etc. Specify either roleDefinitionId or roleDefinitionName but not both

--8<-- "docs/cmd/_global.md"

## Examples

// TODO: Add examples