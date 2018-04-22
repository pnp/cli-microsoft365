# spo page control get

Gets information about the specific control on a modern page

## Usage

```sh
spo page control get [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-i, --id <id>`|ID of the control to retrieve information for
`-n, --name <name>`|Name of the page where the control is located
`-u, --webUrl <webUrl>`|URL of the site where the page to retrieve is located
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to a SharePoint Online site, using the [spo connect](../connect.md) command.

## Remarks

To get information about a control on a modern page, you have to first connect to a SharePoint site using the [spo connect](../connect.md) command, eg. `spo connect https://contoso.sharepoint.com`.

If the specified `name` doesn't refer to an existing modern page, you will get a `File doesn't exists` error.

## Examples

Get information about the control with ID _3ede60d3-dc2c-438b-b5bf-cc40bb2351e1_ placed on a modern page with name _home.aspx_

```sh
spo page control get --id 3ede60d3-dc2c-438b-b5bf-cc40bb2351e1 --webUrl https://contoso.sharepoint.com/sites/team-a --name home.aspx
```