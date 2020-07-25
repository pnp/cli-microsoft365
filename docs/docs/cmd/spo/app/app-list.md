# spo app list

Lists apps from the specified app catalog

## Usage

```sh
m365 spo app list [options]
```

## Options

`-h, --help`
: Output usage information.

`-s, --scope [scope]`
: Target app catalog. `tenant,sitecollection`. Default `tenant`

`-u, --appCatalogUrl [appCatalogUrl]`
: URL of the tenant or site collection app catalog. It must be specified when the scope is `sitecollection`

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

When listing information about apps available in the tenant app catalog, it's not necessary to specify the tenant app catalog URL. When the URL is not specified, the CLI will try to resolve the URL itself. Specifying the app catalog URL is required when you want to list information about apps in a site collection app catalog.

When specifying site collection app catalog, you can specify the URL either with our without the _AppCatalog_ part, for example `https://contoso.sharepoint.com/sites/team-a/AppCatalog` or `https://contoso.sharepoint.com/sites/team-a`. CLI will accept both formats.

When using the text output type (default), the command lists only the values of the `Title`, `ID`, `Deployed` and `AppCatalogVersion` properties of the app. When setting the output type to JSON, all available properties are included in the command output.

## Examples

Return the list of available apps from the tenant app catalog. Show the installed version in the site if applicable.

```sh
m365 spo app list
```

Return the list of available apps from a site collection app catalog of site _https://contoso.sharepoint.com/sites/site1_.

```sh
m365 spo app list --scope sitecollection --appCatalogUrl https://contoso.sharepoint.com/sites/site1
```

## More information

- Application Lifecycle Management (ALM) APIs: [https://docs.microsoft.com/en-us/sharepoint/dev/apis/alm-api-for-spfx-add-ins](https://docs.microsoft.com/en-us/sharepoint/dev/apis/alm-api-for-spfx-add-ins)