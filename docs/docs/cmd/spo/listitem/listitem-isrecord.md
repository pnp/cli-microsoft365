# spo listitem isrecord

Checks if the specified list item is a record

## Usage

```sh
m365 spo listitem isrecord [options]
```

## Options

`-u, --webUrl <webUrl>`
: The URL of the site where the list is located

`-i, --id <id>`
: The ID of the list item to check if it is a record

`-l, --listId [listId]`
: ID of the list where the item should be added. Specify either `listTitle`, `listId` or `listUrl`

`-t, --listTitle [listTitle]`
: Title of the list where the item should be added. Specify either `listTitle`, `listId` or `listUrl`

`--listUrl [listUrl]`
: Server- or site-relative URL of the list. Specify either `listTitle`, `listId` or `listUrl`

--8<-- "docs/cmd/_global.md"

## Examples

Check whether the document with id _1_ in list with title _Documents_ located in site _https://contoso.sharepoint.com/sites/project-x_ is a record

```sh
m365 spo listitem isrecord --webUrl https://contoso.sharepoint.com/sites/project-x --listTitle 'Documents' --id 1
```

Check whether the document with id _1_ in list with id _0cd891ef-afce-4e55-b836-fce03286cccf_ located in site _https://contoso.sharepoint.com/sites/project-x_ is a record

```sh
m365 spo listitem isrecord --webUrl https://contoso.sharepoint.com/sites/project-x --listId 0cd891ef-afce-4e55-b836-fce03286cccf --id 1
```

Check whether a document with a specific id in a list retrieved by server-relative URL in a specific site is a record

```sh
m365 spo listitem isrecord --webUrl https://contoso.sharepoint.com/sites/project-x --listUrl /sites/project-x/documents --id 1
```

