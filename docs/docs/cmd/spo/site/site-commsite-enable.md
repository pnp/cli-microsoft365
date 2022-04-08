# spo site commsite enable

Enables communication site features on the specified site

## Usage

```sh
m365 spo site commsite enable [options]
```

## Options

`-u, --url <url>`
: The URL of the site to enable communication site features on

`-i, --designPackageId [designPackageId]`
: The ID of the site design to apply when enabling communication site features

--8<-- "docs/cmd/_global.md"

!!! important
    To use this command you have to have permissions to access the tenant admin site.

## Examples

Enable communication site features on an existing site

```sh
m365 spo site commsite enable --url https://contoso.sharepoint.com
```
