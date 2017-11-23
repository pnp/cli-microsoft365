# spo storageentity list

Lists tenant properties stored on the specified SharePoint Online app catalog

## Usage

```sh
spo storageentity list [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --appCatalogUrl <appCatalogUrl>`|URL of the app catalog site
`--verbose`|Runs command with verbose logging

!!! important
    Before using this command, connect to a SharePoint Online site, using the [spo connect](../connect.md) command.

## Remarks

To list tenant properties, you have to first connect to a SharePoint site using the
[spo connect](../connect.md) command, eg. `spo connect https://contoso.sharepoint.com`.

Tenant properties are stored in the app catalog site. To list all tenant properties, you have to specify the absolute URL of the app catalog site. If you specify an incorrect URL, or the site at the given URL is not an app catalog site, no properties will be retrieved.

## Examples

```sh
spo storageentity list -u https://contoso.sharepoint.com/sites/appcatalog
```

list all tenant properties stored in the https://contoso.sharepoint.com/sites/appcatalog app catalog site

## More information

- SharePoint Framework Tenant Properties: [https://docs.microsoft.com/en-us/sharepoint/dev/spfx/tenant-properties](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/tenant-properties)