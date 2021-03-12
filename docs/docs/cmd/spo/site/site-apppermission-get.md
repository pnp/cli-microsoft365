# spo site apppermission get

Get a specific application permissions for the site

## Usage

```sh
m365 spo site apppermission get [options]
```

## Options

`-u, --siteUrl <siteUrl>`
: URL of the site collection where the permission to retrieve is located

`-i, --permissionId <permissionId>`
: ID of the permission to retrieve

--8<-- "docs/cmd/_global.md"

## Example

Return a specific application permissions for the _https://contoso.sharepoint.com/sites/project-x_ site collection with permission id _aTowaS50fG1zLnNwLmV4dHw4OWVhNWM5NC03NzM2LTRlMjUtOTVhZC0zZmE5NWY2MmI2NmVAZGUzNDhiYzctMWFlYi00NDA2LThjYjMtOTdkYjAyMWNhZGI0_

```sh
m365 spo site apppermission get --siteUrl https://contoso.sharepoint.com/sites/project-x --permissionId aTowaS50fG1zLnNwLmV4dHw4OWVhNWM5NC03NzM2LTRlMjUtOTVhZC0zZmE5NWY2MmI2NmVAZGUzNDhiYzctMWFlYi00NDA2LThjYjMtOTdkYjAyMWNhZGI0
```
