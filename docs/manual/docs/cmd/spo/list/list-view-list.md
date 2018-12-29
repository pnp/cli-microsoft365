# spo list view set

Gets all existing list views

## Usage

```sh
spo list view list [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --webUrl <webUrl>`|URL of the site where the list is located
`--listId [listId]`|ID of the list where the view is located. Specify `listTitle` or `listId` but not both
`--listTitle [listTitle]`|Title of the list where the view is located. Specify `listTitle` or `listId` but not both
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to a SharePoint Online site, using the [spo login](../login.md) command.

## Remarks

To gets all list views for target list, you have to first log in to a SharePoint Online site using the [spo login](../login.md) command, eg. `spo login https://contoso.sharepoint.com`.

## Examples

Gets all list views for target list

```sh
spo list view get --webUrl https://contoso.sharepoint.com/sites/project-x --listTitle 'My List'
```
