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

`-i, --appId <appId>`
: App Id

`-n, --appDisplayName <appDisplayName>`
: App display name

--8<-- "docs/cmd/_global.md"

## Remarks
Pass in appId and/or addDisplayName, use both for best performance avoiding the extra lookup.

## Example

Adds a specific application permission to the _https://contoso.sharepoint.com/sites/project-x_ site collection with _read_ permission 

```sh
m365 spo site apppermission add --siteUrl https://contoso.sharepoint.com/sites/project-x --permission read --appDisplayName Foo
```
