# spo theme remove

Removes existing theme

## Usage

```sh
spo theme remove [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-n, --name <name>`|Name of the theme to remove
`--confirm`|Do not prompt for confirmation before removing theme
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to a SharePoint Online tenant admin site, using the [spo connect](../connect.md) command.

## Remarks

To remove a theme, you have to first connect to a tenant admin site using the [spo connect](../connect.md) command, eg. `spo connect https://contoso-admin.sharepoint.com`.

## Examples

Remove theme. Will prompt for confirmation before removing the theme

```sh
spo theme remove --name Contoso-Blue
```

Remove theme without prompting for confirmation

```sh
spo theme remove --name Contoso-Blue --confirm
```

## More information

- SharePoint site theming: [https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-theming/sharepoint-site-theming-overview](https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-theming/sharepoint-site-theming-overview)