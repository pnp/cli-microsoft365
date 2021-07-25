# spo site apppermission remove

Removes a specific application permission from a site

## Usage

```sh
m365 spo site apppermission remove [options]
```

## Options

`-u, --siteUrl <siteUrl>`
: URL of the site collection where the permission to remove is located

`--appId [appId]`
: App Id

`-n, --appDisplayName [appDisplayName]`
: App display name

`-i, --permissionId [permissionId]`
: ID of the permission to remove

`--confirm`
: Don't prompt for confirmation

--8<-- "docs/cmd/_global.md"

## Example

Removes list of application permissions for the _https://contoso.sharepoint.com/sites/project-x_ site collection and filter by an application id _89ea5c94-7736-4e25-95ad-3fa95f62b66e_

```sh
m365 spo site apppermission remove --siteUrl https://contoso.sharepoint.com/sites/project-x --appId 89ea5c94-7736-4e25-95ad-3fa95f62b66e
```

Removes list of application permissions for the _https://contoso.sharepoint.com/sites/project-x_ site collection and filter by an application called _Foo_

```sh
m365 spo site apppermission remove --siteUrl https://contoso.sharepoint.com/sites/project-x --appDisplayName Foo
```

Removes a specific application permissions for the _https://contoso.sharepoint.com/sites/project-x_ site collection with permission id _aTowaS50fG1zLnNwLmV4dHw4OWVhNWM5NC03NzM2LTRlMjUtOTVhZC0zZmE5NWY2MmI2NmVAZGUzNDhiYzctMWFlYi00NDA2LThjYjMtOTdkYjAyMWNhZGI0_

```sh
m365 spo site apppermission remove --siteUrl https://contoso.sharepoint.com/sites/project-x --permissionId aTowaS50fG1zLnNwLmV4dHw4OWVhNWM5NC03NzM2LTRlMjUtOTVhZC0zZmE5NWY2MmI2NmVAZGUzNDhiYzctMWFlYi00NDA2LThjYjMtOTdkYjAyMWNhZGI0
```
