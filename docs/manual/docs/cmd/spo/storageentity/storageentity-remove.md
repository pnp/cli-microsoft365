# spo storageentity remove

Removes tenant property stored on the specified SharePoint Online app catalog

## Usage

```sh
spo storageentity remove [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --appCatalogUrl <appCatalogUrl>`|URL of the app catalog site
`-k, --key <key>`|Name of the tenant property to retrieve
`--confirm`|Don't prompt for confirming removal of a tenant property
`-o, --output <output>`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to a SharePoint Online tenant admin site, using the [spo connect](../connect.md) command.

## Remarks

To remove a tenant property, you have to first connect to a tenant admin site using the
[spo connect](../connect.md) command, eg. `spo connect https://contoso-admin.sharepoint.com`.
If you are connected to a different site and will try to manage tenant properties,
you will get an error.

Tenant properties are stored in the app catalog site associated with that tenant.
To remove a property, you have to specify the absolute URL of the app catalog site.
If you specify the URL of a site different than the app catalog, you will get an access denied error.

## Examples

```sh
spo storageentity remove -k AnalyticsId -u https://contoso.sharepoint.com/sites/appcatalog
```

remove the AnalyticsId tenant property. Yields a confirmation prompt before actually
removing the property

```sh
spo storageentity remove -k AnalyticsId --confirm -u https://contoso.sharepoint.com/sites/appcatalog
```

remove the AnalyticsId tenant property. Suppresses the confirmation prompt

## More information

- SharePoint Framework Tenant Properties: [https://docs.microsoft.com/en-us/sharepoint/dev/spfx/tenant-properties](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/tenant-properties)