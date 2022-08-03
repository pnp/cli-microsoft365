# spo site apppermission set

Updates a specific application permission for a site

## Usage

```sh
m365 spo site apppermission set [options]
```

## Options

`-u, --siteUrl <siteUrl>`
: URL of the site collection where the permission to retrieve is located

`-i, --permissionId [permissionId]`
: ID of the permission to update. Specify `permissionId`, `appId` or `appDisplayName`

`--appId [appId]`
: Client ID of the Azure AD app for which to update permissions. Specify `permissionId`, `appId` or `appDisplayName`

`-n, --appDisplayName [appDisplayName]`
: Display name of the Azure AD app for which to update permissions. Specify `permissionId`, `appId` or `appDisplayName`

`-p, --permission <permission>`
: Permission to site (`read`, `write`, or `owner`)

--8<-- "docs/cmd/_global.md"

## Examples

Updates a specific application permission to _read_ for the _https://contoso.sharepoint.com/sites/project-x_ site collection with permission id _aTowaS50fG1zLnNwLmV4dHw4OWVhNWM5NC03NzM2LTRlMjUtOTVhZC0zZmE5NWY2MmI2NmVAZGUzNDhiYzctMWFlYi00NDA2LThjYjMtOTdkYjAyMWNhZGI0_

```sh
m365 spo site apppermission set --siteUrl https://contoso.sharepoint.com/sites/project-x --permissionId aTowaS50fG1zLnNwLmV4dHw4OWVhNWM5NC03NzM2LTRlMjUtOTVhZC0zZmE5NWY2MmI2NmVAZGUzNDhiYzctMWFlYi00NDA2LThjYjMtOTdkYjAyMWNhZGI0 --permission read
```

Updates a specific application permission to _read_ for the _https://contoso.sharepoint.com/sites/project-x_ site collection with an application called _Foo_

```sh
m365 spo site apppermission set --siteUrl https://contoso.sharepoint.com/sites/project-x --appDisplayName Foo --permission read
```

Updates a specific application permission to _read_ for the _https://contoso.sharepoint.com/sites/project-x_ site collection with an application id _89ea5c94-7736-4e25-95ad-3fa95f62b66e_

```sh
m365 spo site apppermission set --siteUrl https://contoso.sharepoint.com/sites/project-x --appId 89ea5c94-7736-4e25-95ad-3fa95f62b66e --permission read
```
