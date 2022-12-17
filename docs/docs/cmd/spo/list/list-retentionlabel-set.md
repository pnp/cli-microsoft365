# spo list retentionlabel set

Sets a default retention label on the specified list or library.

## Usage

```sh
m365 spo list retentionlabel set [options]
```

## Alias

```sh
m365 spo list label set [options]
```

## Options

`-u, --webUrl <webUrl>`
: The URL of the site where the list is located.

`--label <label>`
: The label to set on the list.

`-t, --listTitle [listTitle]`
: The title of the list on which to set the label. Specify either `listTitle`, `listId`, or `listUrl` but not multiple.

`-l, --listId [listId]`
: The ID of the list on which to set the label. Specify either `listTitle`, `listId`, or `listUrl` but not multiple.

`--listUrl [listUrl]`
: Server- or web-relative URL of the list on which to set the label. Specify either `listTitle`, `listId`, or `listUrl` but not multiple.

`--syncToItems`
: Specify, to set the label on all existing items in the list.

`--blockDelete`
: (deprecated) Specify, to disallow deleting items in the list.

`--blockEdit`
: (deprecated) Specify, to disallow editing items in the list.

--8<-- "docs/cmd/_global.md"

## Remarks

A list retention label is a default label that will be applied to all new items in the list. If you specify `syncToItems`, it is also synced to existing items. 

## Examples

Sets retention label on the list with specified site-relative URL located in the specified site.

```sh
m365 spo list retentionlabel set --webUrl https://contoso.sharepoint.com/sites/project-x --listUrl 'Shared Documents' --label 'Some label'
```

Sets retention label and disables editing and deleting items on the list and all existing items for list with specified title located the specified site.

```sh
m365 spo list retentionlabel set --webUrl https://contoso.sharepoint.com/sites/project-x --listTitle 'Documents' --label 'Some label' --blockEdit --blockDelete --syncToItems
```

## Response

The command won't return a response on success.
