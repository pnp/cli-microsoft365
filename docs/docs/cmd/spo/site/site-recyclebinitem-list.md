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
: Type of items which should be retrieved (listItems, folders, files)

`--secondary`
: Use this switch to retrieve items from secondary recycle bin

--8<-- "docs/cmd/_global.md"

## Remarks

When type is not specified then the command will return all items in the recycle bin

## Examples

Lists all files, items and folders from recycle bin for site _https://contoso.sharepoint.com/site_

```sh
m365 spo site recyclebinitem list --siteUrl https://contoso.sharepoint.com/site
```

Lists only files from recycle bin for site _https://contoso.sharepoint.com/site_

```sh
m365 spo site recyclebinitem list --siteUrl https://contoso.sharepoint.com/site --type files
```
