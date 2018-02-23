# spo theme set

Add or update a theme

## Usage

```sh
spo theme set [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-n, --name <name>`|Name of the theme to add or update
`-p, --filePath <filePath>`|Absolute or relative path to the theme json file
`--isInverted`|Set to specify that the theme is inverted
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to a SharePoint Online tenant admin site, using the [spo connect](../connect.md) command.

## Remarks

To add or update a theme, you have to first connect to a tenant admin site using the [spo connect](../connect.md) command, eg. `spo connect https://contoso-admin.sharepoint.com`.

## Examples

Add or update a theme from a theme JSON file

```sh
spo theme set -n Contoso-Blue -p /Users/rjesh/themes/contoso-blue.json
```

Add or update an inverted theme from a theme JSON file

```sh
spo theme set -n Contoso-Blue -p /Users/rjesh/themes/contoso-blue.json --isInverted
```

## More information

- SharePoint site theming: [https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-theming/sharepoint-site-theming-overview](https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-theming/sharepoint-site-theming-overview)
- Theme Generator: [https://developer.microsoft.com/en-us/fabric#/styles/themegenerator](https://developer.microsoft.com/en-us/fabric#/styles/themegenerator)