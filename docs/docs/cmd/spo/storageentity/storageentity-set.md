# spo storageentity set

Sets tenant property on the specified SharePoint Online app catalog

## Usage

```sh
m365 spo storageentity set [options]
```

## Options

`-h, --help`
: output usage information

`-u, --appCatalogUrl <appCatalogUrl>`
: URL of the app catalog site

`-k, --key <key>`
: Name of the tenant property to retrieve

`-v, --value <value>`
: Value to set for the property

`-d, --description [description]`
: Description to set for the property (optional)

`-c, --comment [comment]`
: Comment to set for the property (optional)

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

Tenant properties are stored in the app catalog site associated with that tenant. To set a property, you have to specify the absolute URL of the app catalog site without a trailing slash. If you specify the URL with trailing slash you get the error `The managed path sites/apps is not a managed path in this tenant.`

If you specify the URL of a site different than the app catalog, you will get an access denied error.

## Examples

Set _123_ as the value of the _AnalyticsId_ tenant property. Also include a description and a comment for additional clarification of the usage of the property.

```sh
m365 spo storageentity set -k AnalyticsId -v 123 -d 'Web analytics ID' -c 'Use on all sites' -u https://contoso.sharepoint.com/sites/appcatalog
```

## More information

- SharePoint Framework Tenant Properties: [https://docs.microsoft.com/en-us/sharepoint/dev/spfx/tenant-properties](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/tenant-properties)
