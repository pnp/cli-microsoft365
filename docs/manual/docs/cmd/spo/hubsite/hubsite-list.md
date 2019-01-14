# spo hubsite list

Lists hub sites in the current tenant

!!! attention
    This command is based on a SharePoint API that is currently in preview and is subject to change once the API reached general availability.

## Usage

```sh
spo hubsite list [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-i, --includeAssociatedSites`|Include the associated sites in the result (only in JSON output)
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to a SharePoint Online site, using the [spo login](../login.md) command.

## Remarks

To list hub sites, you have to first log in to a SharePoint site using the [spo login](../login.md) command, eg. `spo login https://contoso.sharepoint.com`.

When using the text output type (default), the command lists only the values of the `ID`, `SiteUrl` and `Title` properties of the hub site. When setting the output type to JSON, all available properties are included in the command output.

## Examples

List hub sites in the current tenant

```sh
spo hubsite list
```

List hub sites, including their associated sites, in the current tenant. Associated site info is only shown in JSON output.

```sh
spo hubsite list --includeAssociatedSites --output json
```

## More information

- SharePoint hub sites new in Office 365: [https://techcommunity.microsoft.com/t5/SharePoint-Blog/SharePoint-hub-sites-new-in-Office-365/ba-p/109547](https://techcommunity.microsoft.com/t5/SharePoint-Blog/SharePoint-hub-sites-new-in-Office-365/ba-p/109547)