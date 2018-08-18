# spo page column list

Get information about the specified modern page section

## Usage

```sh
spo page section get [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --webUrl <webUrl>`|URL of the site where the page to retrieve is located
`-n, --name <name>`|Name of the page to get section information of
`-s, --section <sectionId>`|ID of the section for which to retrieve information
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to a SharePoint Online site, using the [spo connect](../connect.md) command.

## Remarks

To get information about a modern page section, you have to first connect to a SharePoint site using the [spo connect](../connect.md) command, eg. `spo connect https://contoso.sharepoint.com`.

If the specified name doesn't refer to an existing modern page, you will get a _File doesn't exists_ error.

## Examples

Get information about the specified section of the modern page named _home.aspx_

```sh
spo page section get --webUrl https://contoso.sharepoint.com/sites/team-a --name home.aspx --section 1
```