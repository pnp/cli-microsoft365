# spo list retentionlabel ensure

Sets a default retention label on the specified list or library.

## Usage

```sh
m365 spo list retentionlabel ensure [options]
```

## Alias

```sh
m365 spo list label set [options]
```

## Options

`-u, --webUrl <webUrl>`
: The URL of the site where the list is located

`--label <label>`
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

Sets retention label "Some label" on the list _Shared Documents_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo list retentionlabel ensure --webUrl https://contoso.sharepoint.com/sites/project-x --listUrl 'Shared Documents' --label 'Some label'
```

Sets retention label "Some label" and disables editing and deleting items on the list and all existing items for list for list _Documents_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo list retentionlabel ensure --webUrl https://contoso.sharepoint.com/sites/project-x --listTitle 'Documents' --label 'Some label' --blockEdit --blockDelete --syncToItems
```

## Response

The command won't return a response on success.
