# spo page section list

List sections in the specific modern page

## Usage

```sh
spo page section list [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --webUrl <webUrl>`|URL of the site where the page to retrieve is located
`-n, --name <name>`|Name of the page to list sections of
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to a SharePoint Online site, using the [spo connect](../connect.md) command.

## Remarks

To list sections of a modern page, you have to first connect to a SharePoint site using the [spo connect](../connect.md) command, eg. `spo connect https://contoso.sharepoint.com`.

If the specified name doesn't refer to an existing modern page, you will get a _File doesn't exists_ error.

## Examples

List sections of a modern page named _home.aspx_

```sh
spo page section list --webUrl https://contoso.sharepoint.com/sites/team-a --name home.aspx
```