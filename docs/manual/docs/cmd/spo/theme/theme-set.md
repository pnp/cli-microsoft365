# spo theme set

Add or update theme to tenant with the given palette

## Usage

```sh
spo theme set [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-n, --name <name>`|Name of the theme
`-p, --filePath <filePath>`|File path of theme json file
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to a SharePoint Online site, using the [spo connect](../connect.md) command.

## Remarks

To set company theme, you have to first connect to a SharePoint admin site using the [spo connect](../connect.md) command, eg. `spo connect https://contoso-admin.sharepoint.com`.

## Examples:
    
To add or update theme to the tenant from absolute or relative path of given theme json file
```sh
o365$ SPO theme set -n Contoso-Blue -p /Users/rjesh/themes/contoso-blue.json --isInverted false
```

## More information:

- Refer to [SharePoint site theming overview.](https://github.com/SharePoint/sp-dev-docs/blob/master/docs/declarative-customization/site-theming/sharepoint-site-theming-overview.md)
- Create custom theme using [Office Fabric theme generator tool](https://developer.microsoft.com/en-us/fabric#/styles/themegenerator), copy the JSON output and save as JSON file.
