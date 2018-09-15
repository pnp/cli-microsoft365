# spo hidedefaultthemes set

Sets the value of the HideDefaultThemes setting

## Usage

```sh
spo hidedefaultthemes set [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-h, --hideDefaultThemes <hideDefaultThemes>`|Set to `true` to hide default themes and to `false` to show them
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to a SharePoint Online tenant admin site, using the [spo login](../login.md) command.

## Remarks

To set the value of the HideDefaultThemes setting, you have to first log in to a tenant admin site using the [spo login](../login.md) command, eg. `spo login https://contoso-admin.sharepoint.com`.

## Examples

Hide default themes and allow users to use organization themes only

```sh
spo hidedefaultthemes set --hideDefaultThemes true
```

## More information

- SharePoint site theming: [https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-theming/sharepoint-site-theming-overview](https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-theming/sharepoint-site-theming-overview)