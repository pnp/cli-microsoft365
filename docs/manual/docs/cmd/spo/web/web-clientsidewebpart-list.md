# spo web clientsidewebpart list

Lists all the available clientsidewebparts for the web

## Usage

```sh
spo web clientsidewebpart list [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --webUrl <webUrl>`|URL of the site for which to retrieve the information`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to a SharePoint Online site, using the [spo connect](../connect.md) command.

## Remarks

To create a subsite, you have to first connect to a SharePoint site using the [spo connect](../connect.md) command, eg. `spo connect https://contoso.sharepoint.com`.

## Examples

Lists all the available clientsidewebparts for the web

```sh
spo web clientsidewebpart list --webUrl https://contoso.sharepoint.com
```