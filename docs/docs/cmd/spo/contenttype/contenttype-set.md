# spo contenttype set

Update existing content type

## Usage

```sh
m365 spo contenttype set [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site where the content type to update is defined.

`-i, --id [id]`
: ID of the content type to update. Specify `id` or `name` but not both, one is required.

`-n, name [name]`
: Name of the content type to update. Specify the `id` or the `name` but not both, one is required.

`--listTitle [listTitle]`
: Title of the list if you want to update a list content type. Specify either `listTitle`, `listId` or `listUrl`.

`--listId [listId]`
: ID of the list if you want to update a list content type. Specify either `listTitle`, `listId` or `listUrl`.

`--listUrl [listUrl]`
: URL of the list if you want to update a list content type. Specify either `listTitle`, `listId` or `listUrl`.

--8<-- "docs/cmd/_global.md"

## Examples

Move site content type to a different group

```sh
m365 spo contenttype set --id 0x001001 --webUrl https://contoso.sharepoint.com --Group "My group"
```

Rename list content type

```sh
m365 spo contenttype set --name "My old item" --webUrl https://contoso.sharepoint.com --listTitle "My list" --Name "My item"
```
