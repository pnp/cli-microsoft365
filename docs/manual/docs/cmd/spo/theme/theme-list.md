# spo theme list

Retrieves the list of custom themes

## Usage

```sh
spo theme list [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to a SharePoint Online site, using the [spo login](../login.md) command.

## Remarks

To get the list of available themes, you have to first log in to a SharePoint admin site using the [spo login](../login.md) command, eg. `spo login https://contoso-admin.sharepoint.com`.

## Examples

List available themes

```sh
spo theme list
```

## More information

- SharePoint site theming: [https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-theming/sharepoint-site-theming-overview](https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-theming/sharepoint-site-theming-overview)