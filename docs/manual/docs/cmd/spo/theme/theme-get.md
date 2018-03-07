# spo theme get

Get custom theme details from tenant for the given theme name

## Usage

```sh
spo theme get [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-n, --name <name>`|Name of the theme
`-o, --output [output]`|Output type `json|text` Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to a SharePoint Online admin site, using the [spo connect](../connect.md) command.

## Remarks

To get company theme, you have to first connect to a SharePoint admin site using the [spo connect](../connect.md) command, eg. `spo connect https://contoso-admin.sharepoint.com\`.

## Examples:
    
To get theme from tenant
```sh
o365$ SPO theme get -n Contoso-Blue
```

## More information:

- Refer to [SharePoint site theming overview.](https://github.com/SharePoint/sp-dev-docs/blob/master/docs/declarative-customization/site-theming/sharepoint-site-theming-overview.md)
