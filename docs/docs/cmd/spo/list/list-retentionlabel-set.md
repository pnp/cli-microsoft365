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
: The URL of the site where the list is located

`--name <name>`
: The label to set on the list

`-t, --listTitle [listTitle]`
: The title of the list on which to set the label. Specify only one of `listTitle`, `listId` or `listUrl`

`-l, --listId [listId]`
: The ID of the list on which to set the label. Specify only one of `listTitle`, `listId` or `listUrl`

`--listUrl [listUrl]`
: Server- or web-relative URL of the list on which to set the label. Specify only one of `listTitle`, `listId` or `listUrl`

`--syncToItems`
: Specify, to set the label on all existing items in the list

`--blockDelete`
: Specify, to disallow deleting items in the list

`--blockEdit`
: Specify, to disallow editing items in the list

--8<-- "docs/cmd/_global.md"

## Remarks

A list retention label is a default label that will be applied to all new items in the list. If you specify `syncToItems`, it is also synced to existing items. 

## Examples

Sets a retention label on a given list

```sh
m365 spo list retentionlabel set --webUrl https://contoso.sharepoint.com/sites/project-x --listUrl 'Shared Documents' --name 'Some label'
```

Sets a retention label and disables editing and deleting items on the list and all existing items for a given list

```sh
m365 spo list retentionlabel set --webUrl https://contoso.sharepoint.com/sites/project-x --listTitle 'Documents' --name 'Some label' --blockEdit --blockDelete --syncToItems
```

## Response

The command won't return a response on success.
