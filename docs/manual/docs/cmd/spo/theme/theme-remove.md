# spo theme remove

Removes existing company theme from tenant with the given name.

## Usage

```sh
spo theme remove [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-n, --name <name>`|Name of the theme
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to a SharePoint Online site, using the [spo connect](../connect.md) command.

## Remarks

- To remove company theme, you have to first connect to a SharePoint admin site using the [spo connect](../connect.md) command, eg. `spo connect https://contoso-admin.sharepoint.com`.

## Examples:

Remove theme without prompting for confirmation
```sh
o365$ SPO theme remove -n Contoso-Blue --confirm
```

## More information:

- Refer to [SharePoint site theming overview.](https://github.com/SharePoint/sp-dev-docs/blob/master/docs/declarative-customization/site-theming/sharepoint-site-theming-overview.md)
