# spo hidedefaultthemes get

Gets the current value of the HideDefaultThemes setting

## Usage

```sh
spo hidedefaultthemes get [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to a SharePoint Online tenant admin site, using the [spo login](../login.md) command.

## Remarks

To get the current value of the HideDefaultThemes setting, you have to first log in to a tenant admin site using the [spo login](../login.md) command, eg. `spo login https://contoso-admin.sharepoint.com`.

## Examples

Get the current value of the HideDefaultThemes setting

```sh
spo hidedefaultthemes get
```

## More information

- SharePoint site theming: [https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-theming/sharepoint-site-theming-overview](https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-theming/sharepoint-site-theming-overview)