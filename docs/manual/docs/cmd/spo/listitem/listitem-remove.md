# spo listitem remove

Removes the specified list item

## Usage

```sh
spo listitem remove [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --webUrl <webUrl>`|URL of the site where the list to remove is located
`-i, --id <id>`|The ID of the list item to remove.
`-l, --listId [listId]`|List id of the list to remove. Specify either `listId` or `listTitle` but not both
`-t, --listTitle [listTitle]`|Title of the list to remove. Specify either `listId` or `listTitle` but not both
`--recycle`|Recycle the list item
`--confirm`|Don't prompt for confirming removing the list item
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to a SharePoint Online site, using the [spo connect](../connect.md) command.

## Remarks

To remove a list item, you have to first connect to a SharePoint Online site using the [spo connect](../connect.md) command, eg. `spo connect https://contoso.sharepoint.com`.

## Examples

Remove the list item with ID _1_ from list with ID  _0cd891ef-afce-4e55-b836-fce03286cccf_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
spo listitem remove --webUrl https://contoso.sharepoint.com/sites/project-x --listId 0cd891ef-afce-4e55-b836-fce03286cccf -id 1
```

Remove the list item with ID _1_ from list with title _List 1_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
spo listitem remove --webUrl https://contoso.sharepoint.com/sites/project-x --listTitle 'List 1' --id 1
```