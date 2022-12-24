# spo listitem retentionlabel record lock

Locks the list item record

## Usage

```sh
m365 spo listitem retentionlabel record lock [options]
```

## Options

`-u, --webUrl <webUrl>`
: The url of the web

`-i, --listItemId <listItemId>`
: ID of the list item that should be locked

`-l, --listId [listId]`
: ID of the list. Specify either `listTitle`, `listId` or `listUrl`

`-t, --listTitle [listTitle]`
: Title of the list. Specify either `listTitle`, `listId` or `listUrl`

`--listUrl [listUrl]`
: Server- or site-relative URL of the list. Specify either `listTitle`, `listId` or `listUrl`

--8<-- "docs/cmd/_global.md"

## Examples

Locks the list item record in a given site based on the list id

```sh
m365 spo listitem retentionlabel record lock --webUrl https://contoso.sharepoint.com/sites/project-x --listId 0cd891ef-afce-4e55-b836-fce03286cccf --listItemId 1
```

Locks the list item record in a given site based on the list title

```sh
m365 spo listitem retentionlabel record lock --webUrl https://contoso.sharepoint.com/sites/project-x --listTitle 'List 1' --listItemId 1
```

Locks the list item record in a given site based on the server relative list url

```sh
m365 spo listitem retentionlabel record lock --webUrl https://contoso.sharepoint.com/sites/project-x --listUrl /sites/project-x/lists/TestList --listItemId 1
```

## Response

The command won't return a response on success.
