# spo web roleassignment remove

Removes a role assignment from web permissions.

## Usage

```sh
m365 spo web roleassignment remove [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site

`--principalId [principalId]`
: SharePoint ID of principal it may be either user id or group id you want to add permissions to. Specify either `principalId`, `upn`, or `groupName` but not multiple.

`--upn [upn]`
: Upn/email of user to assign role to. Specify either `principalId`, `upn`, or `groupName` but not multiple.

`--groupName [groupName]`
: Group name of Azure AD or SharePoint group. Specify either `principalId`, `upn`, or `groupName` but not multiple.

`--confirm [confirm]`
: Don't prompt for confirming removing the roleassignment.

--8<-- "docs/cmd/_global.md"

## Examples

Remove roleassignment from web based on group name

```sh
m365 spo web roleassignment remove --webUrl "https://contoso.sharepoint.com/sites/contoso-sales" --groupName "saleGroup"
```

Remove roleassignment from web based on principal Id

```sh
m365 spo web roleassignment remove --webUrl "https://contoso.sharepoint.com/sites/contoso-sales" --principalId 2
```

Remove roleassignment from web based on upn

```sh
m365 spo web roleassignment remove --webUrl "https://contoso.sharepoint.com/sites/contoso-sales" --upn "someaccount@tenant.onmicrosoft.com"
```

Remove roleassignment from web based on principal Id without prompting for confirmation

```sh
m365 spo web roleassignment remove --webUrl "https://contoso.sharepoint.com/sites/contoso-sales" --principalId 2 --confirm
```

