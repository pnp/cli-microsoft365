# spo theme list

Add or update theme to tenant with the given palette

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
    Before using this command, connect to a SharePoint Online site, using the [spo connect](../connect.md) command.

## Remarks

To list company theme(s), you have to first connect to a SharePoint admin site using the [spo connect](../connect.md) command, eg. `spo connect https://contoso-admin.sharepoint.com`.

## Examples:
    
Returns themes from the tenant store
```sh
O365 SPO theme list
```

Returns themes from the tenant store as JSON
```sh
O365 SPO theme list -o json   
```

## More information:

- Refer to [SharePoint site theming overview.](https://github.com/SharePoint/sp-dev-docs/blob/master/docs/declarative-customization/site-theming/sharepoint-site-theming-overview.md)

