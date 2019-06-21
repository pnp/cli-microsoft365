# spo apppage add

Creates a single-part app page

## Usage

```sh
spo apppage add [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --webUrl <webUrl>`|The URL of the site where the page should be created
`-t, --title <title>`|The title of the page to be created
` -d, --webPartData <webPartData>`|JSON string of the web part to put on the page
`--addToQuickLaunch`|Set, to add the page to the quick launch
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to a SharePoint Online site, using the [spo login](../login.md) command.

## Remarks

To create a single-part app page, you have to first log in to a SharePoint site using the spo login command,
eg. o365$ spo login https://contoso.sharepoint.com.
If you want to add the single-part app page to quicklaunch, use the addToQuickLaunch
flag.

## Examples

Create a single-part app page in a site with url https://contoso.sharepoint.com, webpart data are stored in the $webPartData variable

```sh
spo apppage add --title "Contoso" --webUrl "https://contoso.sharepoint.com" --webPartData $webPartData --addToQuickLaunch 
```