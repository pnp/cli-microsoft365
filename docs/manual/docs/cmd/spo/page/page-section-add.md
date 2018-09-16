# spo page section add

Adds section to modern page

## Usage

```sh
spo page section add [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
-n`, --name <name>`|Name of the page to add section to
`-u, --webUrl <webUrl>`|URL of the site where the page to retrieve is located
`-t, --sectionTemplate <sectionTemplate>`|Type of section to add. Allowed values `OneColumn|OneColumnFullWidth|TwoColumn|ThreeColumn|TwoColumnLeft|TwoColumnRight`
`--order [order]`|Order of the section to add
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to a SharePoint Online site, using the [spo login](../login.md) command.

## Remarks

To add a section to the modern page, you have to first log in to a SharePoint site using the [spo login](../login.md) command, eg. `spo login https://contoso.sharepoint.com`.

If the specified `name` doesn't refer to an existing modern page, you will get a _File doesn't exists_ error.

## Examples

Add section to the modern page named _home.aspx_

```sh
spo page section add --name home.aspx --webUrl https://contoso.sharepoint.com/sites/newsletter  --sectionTemplate OneColumn --order 1
```