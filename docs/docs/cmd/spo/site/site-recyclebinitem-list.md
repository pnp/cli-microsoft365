# spo site recyclebinitem list

Lists items from recycle bin

## Usage

```sh
m365 spo site recyclebinitem list [options]
```

## Options

`-u, --siteUrl <siteUrl>`
: URL of the site for which to retrieve the recycle bin items

`--type [type]`
: type of items which should be retrieved (1 - list items, 3 - folder, 5 - files)

`--secondary`
: use this switch to retrieve items from secondary recycle bin

--8<-- "docs/cmd/_global.md"

## Remarks

When type is not specified then the command will return all items in the recycle bin

## Examples

Lists all files, items and folders from recycle bin for site _https://contoso.sharepoint.com/site

```sh
m365 spo site recyclebinitem list --siteUrl https://contoso.sharepoint.com/site
```

Lists only files from recycle bin for site _https://contoso.sharepoint.com/site

```sh
m365 spo site recyclebinitem list --siteUrl https://contoso.sharepoint.com/site --type 1
```