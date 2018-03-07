# spo theme get

Gets custom theme information

## Usage

```sh
spo theme get [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-n, --name <name>`|The name of the theme to retrieve
`-o, --output [output]`|Output type `json|text` Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to a SharePoint Online tenant admin site, using the [spo connect](../connect.md) command.

## Remarks

To get information about a theme, you have to first connect to a tenant admin site using the [spo connect](../connect.md) command, eg. `spo connect https://contoso-admin.sharepoint.com`.

## Examples

Get information about a theme

```sh
spo theme get --name Contoso-Blue
```

## More information

- SharePoint site theming: [https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-theming/sharepoint-site-theming-overview](https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-theming/sharepoint-site-theming-overview)