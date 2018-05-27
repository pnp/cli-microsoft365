# spo listitem add

Gets a list item from the specified list

## Usage

```sh
spo listitem get [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --webUrl <webUrl>`|URL of the site where the item should be retrieved
`-l, --listId [listId]`|ID of the list where the item should be retrieved. Specify `listId` or `listTitle` but not both
`-t, --listTitle [listTitle]`|Title of the list where the item should be retrieved. Specify `listId` or `listTitle` but not both
`-i, --id [listItemId]`|ID of the item to retrieve
`-f, --field [fields]`|Comma-separated list of fields to retrieve. Will retrieve all fields if not specified and json output is requested
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to a SharePoint Online site, using the [spo connect](../connect.md) command.

## Remarks

To get an item from a list, you have to first connect to a SharePoint Online site using the [spo connect](../connect.md) command, eg. `spo connect https://contoso.sharepoint.com`.

## Examples

Get an item with ID _147_ from list with title _Demo List_ in site _https://contoso.sharepoint.com/sites/project-x_

```sh
spo listitem get --listTitle "Demo List" --id 147 --webUrl https://contoso.sharepoint.com/sites/project-x
```

