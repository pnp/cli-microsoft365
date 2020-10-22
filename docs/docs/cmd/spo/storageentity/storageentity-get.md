# spo storageentity get

Get details for the specified tenant property

## Usage

```sh
m365 spo storageentity get [options]
```

## Options

`-k, --key <key>`
: Name of the tenant property to retrieve

--8<-- "docs/cmd/_global.md"

## Remarks

Tenant properties are stored in the app catalog site associated with the site to which you are currently connected. When retrieving the specified tenant property, SharePoint will automatically find the associated app catalog and try to retrieve the property from it.

## Examples

Show the value, description and comment of the _AnalyticsId_ tenant property

```sh
m365 spo storageentity get -k AnalyticsId
```

## More information

- SharePoint Framework Tenant Properties: [https://docs.microsoft.com/en-us/sharepoint/dev/spfx/tenant-properties](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/tenant-properties)
