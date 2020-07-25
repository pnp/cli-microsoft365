# spo storageentity remove

Removes tenant property stored on the specified SharePoint Online app catalog

## Usage

```sh
m365 spo storageentity remove [options]
```

## Options

`-h, --help`
: output usage information

`-u, --appCatalogUrl <appCatalogUrl>`
: URL of the app catalog site

`-k, --key <key>`
: Name of the tenant property to retrieve

`--confirm`
: Don't prompt for confirming removal of a tenant property

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

!!! important
    To use this command you have to have permissions to access the tenant admin site.

## Remarks

Tenant properties are stored in the app catalog site associated with that tenant. To remove a property, you have to specify the absolute URL of the app catalog site. If you specify the URL of a site different than the app catalog, you will get an access denied error.

## Examples

Remove the _AnalyticsId_ tenant property. Yields a confirmation prompt before actually removing the property

```sh
m365 spo storageentity remove -k AnalyticsId -u https://contoso.sharepoint.com/sites/appcatalog
```

Remove the _AnalyticsId_ tenant property. Suppresses the confirmation prompt

```sh
m365 spo storageentity remove -k AnalyticsId --confirm -u https://contoso.sharepoint.com/sites/appcatalog
```

## More information

- SharePoint Framework Tenant Properties: [https://docs.microsoft.com/en-us/sharepoint/dev/spfx/tenant-properties](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/tenant-properties)