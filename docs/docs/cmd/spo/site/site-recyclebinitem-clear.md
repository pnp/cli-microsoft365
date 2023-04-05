# spo site recyclebinitem clear

Permanently removes all items in a site recycle bin

## Usage

```sh
m365 spo site recyclebinitem clear [options]
```

## Options

`-u, --siteUrl <siteUrl>`
: URL of the site for which to retrieve the recycle bin items

`--secondary`
: Use this switch to retrieve items from secondary recycle bin

`--confirm`
: Don't prompt for confirmation.

--8<-- "docs/cmd/_global.md"

## Examples

Clear all items from the first-stage recycle bin

```sh
m365 spo site recyclebinitem clear --siteUrl https://contoso.sharepoint.com/sites/sales
```

Clear all items from the second-stage recycle bin

```sh
m365 spo site recyclebinitem clear --siteUrl https://contoso.sharepoint.com/sites/sales --secondary
```
