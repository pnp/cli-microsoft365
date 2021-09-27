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

Removes all application permissions for application with id _89ea5c94-7736-4e25-95ad-3fa95f62b66e_ on site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo site apppermission remove --siteUrl https://contoso.sharepoint.com/sites/project-x --appId 89ea5c94-7736-4e25-95ad-3fa95f62b66e
```

Removes all application permissions for application named _Foo_ on site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo site apppermission remove --siteUrl https://contoso.sharepoint.com/sites/project-x --appDisplayName Foo
```

Removes the application permission with the specified ID on site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo site apppermission remove --siteUrl https://contoso.sharepoint.com/sites/project-x --permissionId aTowaS50fG1zLnNwLmV4dHw4OWVhNWM5NC03NzM2LTRlMjUtOTVhZC0zZmE5NWY2MmI2NmVAZGUzNDhiYzctMWFlYi00NDA2LThjYjMtOTdkYjAyMWNhZGI0
```
