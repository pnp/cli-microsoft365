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
: Relative URL of the list. Specify either listId, listTitle or listUrl but not multiple.

`--principalId [principalId]`
: SharePoint ID of principal it may be either user id or group id we want to add permissions to. Specify principalId only when upn or groupName are not used.

`--upn [upn]`
: Upn/email of user to assign role to. Specify either upn or princpialId

`--groupName [groupName]`
: Enter group name of Azure AD or SharePoint group.. Specify either groupName or princpialId

`--roleDefinitionId [roleDefinitionId]`
: ID of role definition. Specify either roleDefinitionId or roleDefinitionName but not both

`--roleDefinitionName [roleDefinitionName]`
: Enter the name of a role definition, like 'Contribute', 'Read', etc. Specify either roleDefinitionId or roleDefinitionName but not both

--8<-- "docs/cmd/_global.md"

## Examples

add role assignment to list _someList_ located in site _https://contoso.sharepoint.com/sites/project-x_for principal id _11_ and role definition id _1073741829_

```sh
m365 spo list roleassignment add --webUrl "https://contoso.sharepoint.com/sites/project-x" --listTitle "someList" --principalId 11 --roleDefinitionId 1073741829
```

add role assignment to list _0CD891EF-AFCE-4E55-B836-FCE03286CCCF_ located in site _https://contoso.sharepoint.com/sites/project-x_for principal id _11_ and role definition id _1073741829_

```sh
m365 spo list roleassignment add --webUrl "https://contoso.sharepoint.com/sites/project-x" --listId "0CD891EF-AFCE-4E55-B836-FCE03286CCCF" --principalId 11 --roleDefinitionId 1073741829
```

add role assignment to list _sites/documents_ located in site _https://contoso.sharepoint.com/sites/project-x_for principal id _11_ and role definition id _1073741829_

```sh
m365 spo list roleassignment add --webUrl "https://contoso.sharepoint.com/sites/project-x" --listUrl "sites/documents" --principalId 11 --roleDefinitionId 1073741829
```

add role assignment to list _someList_ located in site _https://contoso.sharepoint.com/sites/project-x_for upn _someaccount@tenant.onmicrosoft.com_ and role definition id _1073741829_

```sh
m365 spo list roleassignment add --webUrl "https://contoso.sharepoint.com/sites/project-x" --listTitle "someList" --upn "someaccount@tenant.onmicrosoft.com" --roleDefinitionId 1073741829
```

add role assignment to list _someList_ located in site _https://contoso.sharepoint.com/sites/project-x_for group _someGroup_ and role definition id _1073741829_

```sh
m365 spo list roleassignment add --webUrl "https://contoso.sharepoint.com/sites/project-x" --listTitle "someList" --groupName "someGroup" --roleDefinitionId 1073741829
```

add role assignment to list _someList_ located in site _https://contoso.sharepoint.com/sites/project-x_for principal id _11_ and role definition name _Full Control_

```sh
m365 spo list roleassignment add --webUrl "https://contoso.sharepoint.com/sites/project-x" --listTitle "someList" --principalId 11 --roleDefinitionName "Full Control"
```