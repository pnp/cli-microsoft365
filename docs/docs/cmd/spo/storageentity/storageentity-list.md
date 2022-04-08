# spo storageentity list

Lists tenant properties stored on the specified SharePoint Online app catalog

## Usage

```sh
m365 spo storageentity list [options]
```

## Options

`-u, --appCatalogUrl <appCatalogUrl>`
: URL of the app catalog site

--8<-- "docs/cmd/_global.md"

## Remarks

Tenant properties are stored in the app catalog site. To list all tenant properties, you have to specify the absolute URL of the app catalog site. If you specify an incorrect URL, or the site at the given URL is not an app catalog site, no properties will be retrieved.

## Examples

List all tenant properties stored in the _https://contoso.sharepoint.com/sites/appcatalog_ app catalog site

```sh
m365 spo storageentity list -u https://contoso.sharepoint.com/sites/appcatalog
```

## More information

- SharePoint Framework Tenant Properties: [https://docs.microsoft.com/en-us/sharepoint/dev/spfx/tenant-properties](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/tenant-properties)
