# spo page get

Gets information about the specific modern page

## Usage

```sh
spo page get [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-n, --name <name>`|Name of the page to retrieve
`-u, --webUrl <webUrl>`|URL of the site where the page to retrieve is located
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to a SharePoint Online site, using the [spo connect](../connect.md) command.

## Remarks

To get information about a modern page, you have to first connect to a SharePoint site using the [spo connect](../connect.md) command, eg. `spo connect https://contoso.sharepoint.com`.

If the specified name doesn't refer to an existing modern page, you will get a `File doesn't exists` error.

## Examples

Get information about the modern page with name _home.aspx_

```sh
spo page get --webUrl https://contoso.sharepoint.com/sites/team-a --name home.aspx
```