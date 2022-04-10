# spo app instance list

Retrieve apps installed in a site

## Usage

```sh
m365 spo app instance list [options]
```

## Options

`-u, --siteUrl <siteUrl>`
: URL of the site collection to retrieve the apps for

--8<-- "docs/cmd/_global.md"

## Examples

Return a list of installed apps on site _https://contoso.sharepoint.com/sites/site1_.

```sh
m365 spo app instance list --siteUrl https://contoso.sharepoint.com/sites/site1
```
