# spo site apppermission set

Updates a specific application permission for a site

## Usage

```sh
m365 spo site apppermission set [options]
```

## Options

`-u, --siteUrl <siteUrl>`
: URL of the site collection where the permission to retrieve is located

`-p, --permission <permission>`
: Permission to site (`read`, `write`, `read,write`). If multiple permissions have to be granted, they have to be comma separated ex. `read,write`

`-i, --appId [appId]`
: Client ID of the Azure AD app for which to grant permissions

`-n, --appDisplayName [appDisplayName]`
: Display name of the Azure AD app for which to grant permissions

--8<-- "docs/cmd/_global.md"

## Examples

Updates a specific application permission to _read_ for the _https://contoso.sharepoint.com/sites/project-x_ site collection with an application called _Foo_

```sh
m365 spo site apppermission set --siteUrl https://contoso.sharepoint.com/sites/project-x --appDisplayName Foo --permission read
```

Updates a specific application permission to _read_ for the _https://contoso.sharepoint.com/sites/project-x_ site collection with an application id _89ea5c94-7736-4e25-95ad-3fa95f62b66e_

```sh
m365 spo site apppermission set --siteUrl https://contoso.sharepoint.com/sites/project-x --appId 89ea5c94-7736-4e25-95ad-3fa95f62b66e --permission read
```
