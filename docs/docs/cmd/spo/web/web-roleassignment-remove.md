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
: SharePoint ID of principal it may be either user id or group id we want to add permissions to. Specify principalId only when upn or groupName are not used.

`--upn [upn]`
: Upn/email of user to assign role to. Specify upn only when principalId or groupName are not used.

`--groupName [groupName]`
: Enter group name of Azure AD or SharePoint group. Specify groupName only when principalId or upn are not used.

`--confirm [confirm]`
: Don't prompt for confirming removing the roleassignment.

--8<-- "docs/cmd/_global.md"

## Examples

Remove roleassignment from web based on group name

```sh
m365 spo list roleassignment remove --webUrl "https://contoso.sharepoint.com/sites/contoso-sales"  --groupName "saleGroup"
```

Remove roleassignment from web based on principal Id

```sh
m365 spo list roleassignment remove --webUrl "https://contoso.sharepoint.com/sites/contoso-sales"  --principalId 2
```

Remove roleassignment from web based on upn

```sh
m365 spo list roleassignment remove --webUrl "https://contoso.sharepoint.com/sites/contoso-sales"  --upn "someaccount@tenant.onmicrosoft.com"
```

Remove roleassignment from web based on principal Id without prompting for confirmation

```sh
m365 spo list roleassignment remove --webUrl "https://contoso.sharepoint.com/sites/contoso-sales"  --principalId 2 --confirm
```