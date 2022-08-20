# spo web roleassignment add

Adds a role assignment to web permissions.

## Usage

```sh
m365 spo web roleassignment add [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site

`--principalId [principalId]`
: SharePoint ID of principal it may be either user id or group id we want to add permissions to. Specify principalId only when upn or groupName are not used.

`--upn [upn]`
: upn/email of user to assign role to. Specify either upn or princpialId

`--groupName [groupName]`
: enter group name of Azure AD or SharePoint group.. Specify either groupName or princpialId

`--roleDefinitionId [roleDefinitionId]`
: ID of role definition. Specify either roleDefinitionId or roleDefinitionName but not both

`--roleDefinitionName [roleDefinitionName]`
: enter the name of a role definition, like 'Contribute', 'Read', etc. Specify either roleDefinitionId or roleDefinitionName but not both

--8<-- "docs/cmd/_global.md"

## Examples

add role assignment to site _https://contoso.sharepoint.com/sites/project-x_for principal id _11_ and role definition id _1073741829_

```sh
m365 spo list roleassignment add --webUrl "https://contoso.sharepoint.com/sites/project-x" --principalId 11 --roleDefinitionId 1073741829
```

add role assignment to site _https://contoso.sharepoint.com/sites/project-x_for upn _someaccount@tenant.onmicrosoft.com_ and role definition id _1073741829_

```sh
m365 spo list roleassignment add --webUrl "https://contoso.sharepoint.com/sites/project-x" --upn "someaccount@tenant.onmicrosoft.com" --roleDefinitionId 1073741829
```

add role assignment to site _https://contoso.sharepoint.com/sites/project-x_for group _someGroup_ and role definition id _1073741829_

```sh
m365 spo list roleassignment add --webUrl "https://contoso.sharepoint.com/sites/project-x" --groupName "someGroup" --roleDefinitionId 1073741829
```

add role assignment to site _https://contoso.sharepoint.com/sites/project-x_for principal id _11_ and role definition name _Full Control_

```sh
m365 spo list roleassignment add --webUrl "https://contoso.sharepoint.com/sites/project-x" --principalId 11 --roleDefinitionName "Full Control"
```