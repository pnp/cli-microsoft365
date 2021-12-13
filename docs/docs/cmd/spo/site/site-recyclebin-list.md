# spo site recyclebin list

Lists items from recycle bin

## Usage

```sh
m365 spo site recyclebin list [options]
```

## Options

`-u, --siteUrl <siteUrl>`
: URL of the site for which to retrieve the recycle bin items

`--type [type]`
: type of items which should be retreived (1 - list items, 3 - folder, 5 - files)

`--secondary`
: use this switch to retrieve items from secondary recycle bin

--8<-- "docs/cmd/_global.md"

## Remarks

When using the text output type (default), the command lists only items `Title`. When setting the output type to JSON, all available properties are included in the command output.

## Examples

Lists items from recycle bin for site _https://contoso.sharepoint.com/site

```sh
m365 spo site recyclebin list --siteUrl https://contoso.sharepoint.com/site
```