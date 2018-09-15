# spo hubsite get

Gets information about the specified hub site

!!! attention
    This command is based on a SharePoint API that is currently in preview and is subject to change once the API reached general availability.

## Usage

```sh
spo hubsite get [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-i, --id <id>`|Hub site ID
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to a SharePoint Online site, using the [spo login](../login.md) command.

## Remarks

To get information about a hub site, you have to first log in to a SharePoint site using the [spo login](../login.md) command, eg. `spo login https://contoso.sharepoint.com`.

If the specified `id` doesn't refer to an existing hub site, you will get a `ResourceNotFoundException` error.

## Examples

Get information about the hub site with ID _2c1ba4c4-cd9b-4417-832f-92a34bc34b2a_

```sh
spo hubsite get --id 2c1ba4c4-cd9b-4417-832f-92a34bc34b2a
```

## More information

- SharePoint hub sites new in Office 365: [https://techcommunity.microsoft.com/t5/SharePoint-Blog/SharePoint-hub-sites-new-in-Office-365/ba-p/109547](https://techcommunity.microsoft.com/t5/SharePoint-Blog/SharePoint-hub-sites-new-in-Office-365/ba-p/109547)