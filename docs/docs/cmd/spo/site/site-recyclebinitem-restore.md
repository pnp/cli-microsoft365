# spo site recyclebinitem restore

Restores given items from the site recycle bin

## Usage

```sh
m365 spo site recyclebinitem restore [options]
```

## Options

`-u, --siteUrl <siteUrl>`
: URL of the site for which to restore the recycle bin items

`-i, --ids <ids>`
: List of ids of items which will be restored from the site recycle bin

--8<-- "docs/cmd/_global.md"

## Examples

Restore specific items by given ids from recycle bin for site _https://contoso.sharepoint.com/site_

```sh
m365 spo site recyclebinitem restore --siteUrl https://contoso.sharepoint.com/site --ids "ae6f97a7-280e-48d6-b481-0ea986c323da,aadbf916-1f71-42ee-abf2-8ee4802ae291"
```
