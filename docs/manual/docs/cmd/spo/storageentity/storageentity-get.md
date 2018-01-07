# spo storageentity get

Get details for the specified tenant property

## Usage

```sh
spo storageentity get [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-k, --key <key>`|Name of the tenant property to retrieve
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to a SharePoint Online site, using the [spo connect](../connect.md) command.

## Remarks

To get details of a tenant property, you have to first connect to a SharePoint site using the [spo connect](../connect.md) command, eg. `spo connect https://contoso.sharepoint.com`.

Tenant properties are stored in the app catalog site associated with the site to which you are currently connected. When retrieving the specified tenant property, SharePoint will automatically find the associated app catalog and try to retrieve the property from it.

## Examples

Show the value, description and comment of the _AnalyticsId_ tenant property

```sh
spo storageentity get -k AnalyticsId
```

## More information

- SharePoint Framework Tenant Properties: [https://docs.microsoft.com/en-us/sharepoint/dev/spfx/tenant-properties](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/tenant-properties)
