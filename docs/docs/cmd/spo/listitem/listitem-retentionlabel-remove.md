# spo listitem retentionlabel remove

Clears the retention label from a list item

## Usage

```sh
m365 spo listitem retentionlabel remove [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site where the retentionlabel from a listitem to remove is located

`-i, --listItemId <listItemId>`
: The ID of the list item for which the retention label should be removed.

`-l, --listId [listId]`
: ID of the list where the retention label should be removed. Specify either `listTitle`, `listId` or `listUrl`

`-t, --listTitle [listTitle]`
: Title of the list where the retention label should be removed. Specify either `listTitle`, `listId` or `listUrl`

`--listUrl [listUrl]`
: Server- or site-relative URL of the list. Specify either `listTitle`, `listId` or `listUrl`

`--confirm`
: Don't prompt for confirming removing the list item

--8<-- "docs/cmd/_global.md"

## Examples

Removes the retention label from a list item in a given site based on the list id

```sh
m365 spo listitem retentionlabel remove --webUrl https://contoso.sharepoint.com/sites/project-x --listId 0cd891ef-afce-4e55-b836-fce03286cccf --id 1
```

Removes the retention label from a list item in a given site based on the list title

```sh
m365 spo listitem retentionlabel remove --webUrl https://contoso.sharepoint.com/sites/project-x --listTitle 'List 1' --id 1
```

Removes the retention label from a list item in a given site based on the server relative list url

```sh
m365 spo listitem retentionlabel remove --webUrl https://contoso.sharepoint.com/sites/project-x --listUrl /sites/project-x/lists/TestList --id 1
```

## Response

The command won't return a response on success.
