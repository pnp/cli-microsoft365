# spo listitem record isrecord

Check to see if the specified list item is a record

## Usage

```sh
spo listitem record isrecord [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --webUrl <webUrl>`|URL of the site where the list is located
`-l, --listId [listId]`|The ID of the list where the item is located. Specify `listId` or `listTitle` but not both
`-t, --listTitle [listTitle]`|The title of the list where the item is located. Specify `listId` or `listTitle` but not both
`-i, --id <id>`|The ID of the list item to check if it is a record
`--verbose`|Runs command with verbose logging
`--debug`| Runs command with debug logging

!!! important
    Before using this command, log in to a SharePoint Online site, using the [spo login](../login.md) command.

## Remarks

To check whether an item is a record, you have to first log in to a SharePoint site using the [spo login](../login.md) command, eg. `spo login https://contoso.sharepoint.com`.

## Examples

Check whether a document with id _1_ is a record in list with title _Demo List_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
spo listitem record isrecord --webUrl https://contoso.sharepoint.com/sites/project-x --listTitle "Demo List" --id 1
```

Check whether a document with id _1_ is a record in list with id _ea8e1109-2013-1a69-bc05-1403201257fc_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
spo listitem record isrecord --webUrl https://contoso.sharepoint.com/sites/project-x --listId ea8e1109-2013-1a69-bc05-1403201257fc --id 1
```