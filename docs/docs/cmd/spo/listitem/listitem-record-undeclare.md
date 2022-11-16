# spo listitem record undeclare

Undeclares list item as a record

## Usage

```sh
m365 spo listitem record undeclare [options]
```

## Options

`-u, --webUrl <webUrl>`
: The URL of the site where the list is located

`-i, --listItemId <listItemId>`
: ID of the list item to be undeclared as a record.

`-l, --listId [listId]`
: The ID of the list where the item is located. Specify either `listTitle`, `listId` or `listUrl`

`-t, --listTitle [listTitle]`
: The title of the list where the item is located. Specify either `listTitle`, `listId` or `listUrl`

`--listUrl [listUrl]`
: Server- or site-relative URL of the list where the item is located. Specify either `listTitle`, `listId` or `listUrl`

--8<-- "docs/cmd/_global.md"

## Examples

Undeclare the list item as a record with ID _1_ from list with ID _0cd891ef-afce-4e55-b836-fce03286cccf_ located in site _https://contoso.sharepoint.com/sites/project-x

```sh_
spo listitem record undeclare --webUrl https://contoso.sharepoint.com/sites/project-x --listId 0cd891ef-afce-4e55-b836-fce03286cccf --listItemId 1
```

Undeclare the list item as a record with ID _1_ from list with title _List 1_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo listitem record undeclare --webUrl https://contoso.sharepoint.com/sites/project-x --listTitle 'List 1' --listItemId 1
```

Undeclare a list item with a specific id as a record from a list retrieved by server-relative URL located in a specific site

```sh
m365 spo listitem record undeclare --webUrl https://contoso.sharepoint.com/sites/project-x --listUrl '/sites/project-x/Lists/Lists 1' --id 1
```

## Response

The command won't return a response on success.
