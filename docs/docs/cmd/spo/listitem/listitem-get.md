# spo listitem add

Gets a list item from the specified list

## Usage

```sh
m365 spo listitem get [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site from which the item should be retrieved

`-i, --id <id>`
: ID of the item to retrieve.

`-l, --listId [listId]`
: ID of the list from which to retrieve the item. Specify `listId` or `listTitle` but not both

`-t, --listTitle [listTitle]`
: Title of the list from which to retrieve the item. Specify `listId` or `listTitle` but not both

`-f, --fields [fields]`
: Comma-separated list of fields to retrieve. Will retrieve all fields if not specified and json output is requested

--8<-- "docs/cmd/_global.md"

## Examples

Get an item with ID _147_ from list with title _Demo List_ in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo listitem get --listTitle "Demo List" --id 147 --webUrl https://contoso.sharepoint.com/sites/project-x
```


Get an items Title and Created column and with ID _147_ from list with title _Demo List_ in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo listitem get --listTitle "Demo List" --id 147 --webUrl https://contoso.sharepoint.com/sites/project-x --fields "Title,Created"
```
