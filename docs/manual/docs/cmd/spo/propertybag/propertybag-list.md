# spo propertybag list

Gets property bag values

## Usage

```sh
spo propertybag list [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-w, --webUrl <webUrl>`|The URL of the site from which the property bag value should be retrieved
`-f, --folder [folder]`|Server or site relative URL of the folder from which to retrieve property bag value. Case sensitive
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to a SharePoint Online site, using the [spo connect](../connect.md) command.

## Remarks

To retrieve property bag values, you have to first connect to a SharePoint Online site using the [spo connect](../connect.md) command, eg. `spo connect https://contoso.sharepoint.com`.

## Examples

Return propery bag values located in site _https://contoso.sharepoint.com/sites/test_

```sh
spo propertybag list -w https://contoso.sharepoint.com/sites/test
```

Return propery bag values located in site root folder _https://contoso.sharepoint.com/sites/test_

```sh
spo propertybag list -w https://contoso.sharepoint.com/sites/test -f /
```

Return propery bag values located in site document library _https://contoso.sharepoint.com/sites/test_

```sh
spo propertybag list --webUrl https://contoso.sharepoint.com/sites/test --folder '/Shared Documents'
```

Return propery bag values located in folder in site document library _https://contoso.sharepoint.com/sites/test_

```sh
spo propertybag list -w https://contoso.sharepoint.com/sites/test -f '/Shared Documents/MyFolder'
```

Return propery bag values located in site list _https://contoso.sharepoint.com/sites/test_

```sh
spo propertybag list --webUrl https://contoso.sharepoint.com/sites/test --folder /Lists/MyList
```

## More information

- SharePoint Patterns and Practices (PnP) PowerShell, Get-PnPProperty: [https://docs.microsoft.com/en-us/powershell/module/sharepoint-pnp/get-pnppropertybag?view=sharepoint-ps](https://docs.microsoft.com/en-us/powershell/module/sharepoint-pnp/get-pnppropertybag?view=sharepoint-ps)