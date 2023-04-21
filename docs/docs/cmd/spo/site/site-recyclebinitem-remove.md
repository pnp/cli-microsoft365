# spo site recyclebinitem remove

Permanently deletes specific items from the site recycle bin

## Usage

```sh
m365 spo site recyclebinitem remove [options]
```

## Options

`-u, --siteUrl <siteUrl>`
: URL of the site where the recycle bin is located.

`-i, --ids [ids]`
: Comma separated list of item IDs.

`--confirm`
: Don't prompt for confirmation.

--8<-- "docs/cmd/_global.md"

## Examples

Permanently remove 2 specific items from the recycle bin

```sh
m365 spo site recyclebinitem remove --siteUrl https://contoso.sharepoint.com/sites/sales --ids "06ca4fe4-3048-4b76-bd41-296fed4c9881,d679c17b-d7b8-429a-9307-34e1d9e631e7"
```

Permanently remove 2 specific items from the recycle bin and skip the confirmation prompt

```sh
m365 spo site recyclebinitem remove --siteUrl https://contoso.sharepoint.com/sites/sales --ids "06ca4fe4-3048-4b76-bd41-296fed4c9881,d679c17b-d7b8-429a-9307-34e1d9e631e7" --confirm
```

## Response

The command won't return a response on success.
