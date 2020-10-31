# spo app get

Gets information about the specific app from the specified app catalog

## Usage

```sh
m365 spo app get [options]
```

## Options

`-h, --help`
: output usage information

`-i, --id [id]`
: ID of the app to retrieve information for. Specify the `id` or the `name` but not both

`-n, --name [name]`
: Name of the app to retrieve information for. Specify the `id` or the `name` but not both

`-u, --appCatalogUrl [appCatalogUrl]`
: URL of the tenant or site collection app catalog. It must be specified when the scope is `sitecollection`

`-s, --scope [scope]`
: Scope of the app catalog: `tenant,sitecollection`. Default `tenant`

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

When getting information about an app from the tenant app catalog, it's not necessary to specify the tenant app catalog URL. When the URL is not specified, the CLI will try to resolve the URL itself. Specifying the app catalog URL is required when you want to get information about an app from a site collection app catalog.

When specifying site collection app catalog, you can specify the URL either with our without the _AppCatalog_ part, for example `https://contoso.sharepoint.com/sites/team-a/AppCatalog` or `https://contoso.sharepoint.com/sites/team-a`. CLI will accept both formats.

## Examples

Return details about the app with ID _b2307a39-e878-458b-bc90-03bc578531d6_ available in the tenant app catalog.

```sh
m365 spo app get --id b2307a39-e878-458b-bc90-03bc578531d6
```

Return details about the app with name _solution.sppkg_ available in the tenant app catalog. Will try to detect the app catalog URL

```sh
m365 spo app get --name solution.sppkg
```

Return details about the app with name _solution.sppkg_ available in the tenant app catalog using the specified app catalog URL

```sh
m365 spo app get --name solution.sppkg --appCatalogUrl https://contoso.sharepoint.com/sites/apps
```

Return details about the app with ID _b2307a39-e878-458b-bc90-03bc578531d6_ available in the site collection app catalog of site _https://contoso.sharepoint.com/sites/site1_.

```sh
m365 spo app get --id b2307a39-e878-458b-bc90-03bc578531d6 --scope sitecollection --appCatalogUrl https://contoso.sharepoint.com/sites/site1
```

## More information

- Application Lifecycle Management (ALM) APIs: [https://docs.microsoft.com/en-us/sharepoint/dev/apis/alm-api-for-spfx-add-ins](https://docs.microsoft.com/en-us/sharepoint/dev/apis/alm-api-for-spfx-add-ins)