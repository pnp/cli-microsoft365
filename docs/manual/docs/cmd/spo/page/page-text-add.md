# spo page text add

Adds text to a modern page

## Usage

```sh
spo page text add [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --webUrl <webUrl>`|URL of the site where the page to add the text to is located
`-n, --pageName <pageName>`|Name of the page to which add the text
`-t, --text <text>`|Text to add to the page
`--section [section]`|Number of the section to which the text should be added (1 or higher)
`--column [column]`|Number of the column in which the text should be added (1 or higher)
`--order [order]`|Order of the text in the column
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to a SharePoint Online site, using the [spo login](../login.md) command.

## Remarks

To add text to a modern page, you have to first log in to a SharePoint site using the [spo login](../login.md) command, eg. `spo login https://contoso.sharepoint.com`.

If the specified `pageName` doesn't refer to an existing modern page, you will get a _File doesn't exists_ error.

## Examples

Add text to a modern page in the first available location on the page

```sh
spo page text add --webUrl https://contoso.sharepoint.com/sites/a-team --pageName page.aspx --text 'Hello world'
```

Add text to a modern page in the third column of the second section

```sh
spo page text add --webUrl https://contoso.sharepoint.com/sites/a-team --pageName page.aspx --text 'Hello world' --section 2 --column 3
```

Add text at the beginning of the default column on a modern page

```sh
spo page text add --webUrl https://contoso.sharepoint.com/sites/a-team --pageName page.aspx --text 'Hello world' --order 1
```