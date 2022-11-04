# spo listitem remove

Removes the specified list item

## Usage

```sh
m365 spo listitem remove [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site where the list to remove is located

`-i, --id <id>`
: The ID of the list item to remove.

`-l, --listId [listId]`
: ID of the list where the item should be added. Specify either `listTitle`, `listId` or `listUrl`

`-t, --listTitle [listTitle]`
: Title of the list where the item should be added. Specify either `listTitle`, `listId` or `listUrl`

`--listUrl [listUrl]`
: Server- or site-relative URL of the list. Specify either `listTitle`, `listId` or `listUrl`

`--recycle`
: Recycle the list item

`--confirm`
: Don't prompt for confirming removing the list item

--8<-- "docs/cmd/_global.md"

## Examples

Remove the list item with ID _1_ from list with ID  _0cd891ef-afce-4e55-b836-fce03286cccf_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo listitem remove --webUrl https://contoso.sharepoint.com/sites/project-x --listId 0cd891ef-afce-4e55-b836-fce03286cccf --id 1
```

Remove the list item with ID _1_ from list with title _List 1_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo listitem remove --webUrl https://contoso.sharepoint.com/sites/project-x --listTitle 'List 1' --id 1
```

Remove the list item with a specific ID from a list retrieved by server-relative URL located in a specific site

```sh
m365 spo listitem remove --webUrl https://contoso.sharepoint.com/sites/project-x --listUrl /sites/project-x/lists/TestList --id 1
```
