# spo page column list

Lists columns in the specific section of a modern page

## Usage

```sh
spo page column list [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --webUrl <webUrl>`|URL of the site where the page to retrieve is located
`-n, --name <name>`|Name of the page to list columns of
`-s, --section <sectionId>`|ID of the section for which to list columns
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to a SharePoint Online site, using the [spo login](../login.md) command.

## Remarks

To list columns of the specific section of a modern page, you have to first log in to a SharePoint site using the [spo login](../login.md) command, eg. `spo login https://contoso.sharepoint.com`.

If the specified name doesn't refer to an existing modern page, you will get a _File doesn't exists_ error.

## Examples

List columns in the first section of a modern page with name _home.aspx_

```sh
spo page column list --webUrl https://contoso.sharepoint.com/sites/team-a --name home.aspx --section 1
```