# spo site apppermission add

Adds a specific application permissions to the site

## Usage

```sh
m365 spo site apppermission add [options]
```

## Options

`-u, --siteUrl <siteUrl>`
: URL of the site collection to add the permission

`-p, --permission <permission>`
: Permission to site (`read`, `write`, `read,write`). If multiple permissions have to be granted, they have to be comma separated ex. `read,write`

`-i, --appId [appId]`
: Client ID of the Azure AD app for which to grant permissions

`-n, --appDisplayName [appDisplayName]`
: Display name of the Azure AD app for which to grant permissions

--8<-- "docs/cmd/_global.md"

## Remarks

To set permissions, specify at minimum either `appId` or `addDisplayName`. For best performance specify both values to avoid extra lookup.

## Example

Grants the specified app the _read_ permission to site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo site apppermission add --siteUrl https://contoso.sharepoint.com/sites/project-x --permission read --appDisplayName Foo
```
